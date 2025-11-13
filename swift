#region Close only Sirsi/JWF javaw.exe (PSADT logging via Write-ADTLogEntry) — fixed $PID collision

# --- SETTINGS ---
$JreBinPath  = 'C:\Program Files (x86)\Sirsi\JWF\JRE\bin'   # bundled JRE\bin path
$GraceWaitMs = 8000

# --- PSADT-friendly logger wrapper ---
function Write-LogA {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet(1,2,3)][int]$Severity = 1  # 1=Info, 2=Warn, 3=Error
    )
    $prefix = switch ($Severity) { 1 {'[INFO] '} 2 {'[WARN] '} 3 {'[ERR ] '} }
    Write-ADTLogEntry -Message ($prefix + $Message) -Source $adtsession.InstallPhase
}

# Normalize path with trailing backslash for StartsWith()
$normalized = ([System.IO.Path]::GetFullPath($JreBinPath)).TrimEnd('\') + '\'
Write-LogA "Scanning for javaw.exe under: $normalized"

# 1) Enumerate all javaw.exe for diagnostics
$allJavaw = Get-CimInstance Win32_Process -Filter "Name='javaw.exe'" -ErrorAction SilentlyContinue
if (-not $allJavaw) { Write-LogA "No javaw.exe processes currently running."; return }

foreach ($p in $allJavaw) {
    Write-LogA ("javaw PID {0}; Path='{1}'; Cmd='{2}'" -f $p.ProcessId, $p.ExecutablePath, $p.CommandLine)
}

# 2) Filter to JUST the bundled JRE instances by ExecutablePath
$targets = $allJavaw | Where-Object {
    $_.ExecutablePath -and $_.ExecutablePath.StartsWith($normalized, [System.StringComparison]::InvariantCultureIgnoreCase)
}

if (-not $targets) { Write-LogA "No Sirsi/JWF javaw.exe instances matched path '$normalized'." 2; return }

$pidList = $targets.ProcessId
Write-LogA "Matched Sirsi/JWF javaw.exe PID(s): $($pidList -join ', ')"

# 3) Close/kill loop (avoid $PID collision by using $procId)
foreach ($t in $targets) {
    $procId = $t.ProcessId

    # Parent info (helps if a service relaunches it)
    try {
        $pp = Get-CimInstance Win32_Process -Filter "ProcessId=$($t.ParentProcessId)" -ErrorAction Stop
        Write-LogA ("PID {0} parent: {1} ({2})" -f $procId, $t.ParentProcessId, $pp.Name)
        if ($pp.Name -match '^(services\.exe|nssm\.exe|srvany\.exe|wrapper\.exe)$') {
            Write-LogA "PID $procId likely launched by a service/wrapper. Consider stopping that service first." 2
        }
    } catch {
        Write-LogA "PID $procId parent lookup failed: $($_.Exception.Message)" 2
    }

    try {
        $gp = Get-Process -Id $procId -ErrorAction Stop

        # 3a) Graceful close if windowed
        if ($gp.MainWindowHandle -ne 0) {
            Write-LogA "PID $procId: attempting CloseMainWindow()"
            [void]$gp.CloseMainWindow()
            if ($gp.WaitForExit($GraceWaitMs)) {
                Write-LogA "PID $procId exited gracefully."
                continue
            } else {
                Write-LogA "PID $procId did not exit after CloseMainWindow()." 2
            }
        } else {
            Write-LogA "PID $procId has no main window; skipping graceful close."
        }

        # 3b) Force kill
        try {
            Write-LogA "PID $procId: Stop-Process -Force" 2
            Stop-Process -Id $procId -Force -ErrorAction Stop
            Start-Sleep -Milliseconds 500
        } catch {
            Write-LogA "PID $procId Stop-Process failed: $($_.Exception.Message)" 2
        }

        # 3c) WMI Terminate if still alive
        if (Get-Process -Id $procId -ErrorAction SilentlyContinue) {
            Write-LogA "PID $procId still alive; trying WMI Terminate()" 2
            try {
                $ci  = Get-CimInstance Win32_Process -Filter "ProcessId=$procId" -ErrorAction Stop
                $res = Invoke-CimMethod -InputObject $ci -MethodName Terminate -Arguments @{ Reason = 5 } -ErrorAction Stop
                Write-LogA "PID $procId WMI Terminate() returned $($res.ReturnValue)."
            } catch {
                Write-LogA "PID $procId WMI Terminate() failed: $($_.Exception.Message)" 2
            }
        }

        # 3d) taskkill /T /F as last resort
        if (Get-Process -Id $procId -ErrorAction SilentlyContinue) {
            Write-LogA "PID $procId still alive; taskkill /T /F" 3
            $tk = Start-Process -FilePath "$env:WINDIR\System32\taskkill.exe" -ArgumentList "/PID $procId /T /F" -PassThru -Wait -WindowStyle Hidden
            Write-LogA "taskkill exit code for PID $procId: $($tk.ExitCode)"
        }

        if (Get-Process -Id $procId -ErrorAction SilentlyContinue) {
            Write-LogA "PID $procId is STILL running after all methods. Likely being relaunched; stop its parent service/wrapper." 3
        } else {
            Write-LogA "PID $procId terminated."
        }
    } catch {
        Write-LogA "PID $procId handling error: $($_.Exception.Message)" 3
    }
}

#endregion
