#region Close only Sirsi/JWF javaw.exe (PSADT logging via Write-ADTLogEntry)

# --- SETTINGS (edit as needed) ---
$JreBinPath = 'C:\Program Files (x86)\Sirsi\JWF\JRE\bin'   # bundled JRE\bin path
$GraceWaitMs = 8000                                        # wait after CloseMainWindow

# --- Helper for consistent logging ---
function Write-LogA {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet(1,2,3)][int]$Severity = 1            # 1=Info, 2=Warn, 3=Error
    )
    # Prefix severity tag for readability in your logs
    $prefix = switch ($Severity) { 1 {'[INFO] '} 2 {'[WARN] '} 3 {'[ERR ] '} }
    Write-ADTLogEntry -Message ($prefix + $Message) -Source $adtsession.InstallPhase
}

# Normalize path with trailing backslash for StartsWith()
$normalized = ([System.IO.Path]::GetFullPath($JreBinPath)).TrimEnd('\') + '\'
Write-LogA "Scanning for javaw.exe under: $normalized"

# 1) Enumerate ALL javaw.exe for diagnostics
$allJavaw = Get-CimInstance Win32_Process -Filter "Name='javaw.exe'" -ErrorAction SilentlyContinue
if (-not $allJavaw) {
    Write-LogA "No javaw.exe processes currently running."
    return
}
foreach ($p in $allJavaw) {
    Write-LogA ("javaw PID {0}; Path='{1}'; Cmd='{2}'" -f $p.ProcessId, $p.ExecutablePath, $p.CommandLine)
}

# 2) Filter to JUST the bundled JRE instances by ExecutablePath
$targets = $allJavaw | Where-Object {
    $_.ExecutablePath -and $_.ExecutablePath.StartsWith($normalized, [System.StringComparison]::InvariantCultureIgnoreCase)
}

if (-not $targets) {
    Write-LogA "No Sirsi/JWF javaw.exe instances matched path '$normalized'." 2
    return
}

$pidList = $targets.ProcessId
Write-LogA "Matched Sirsi/JWF javaw.exe PID(s): $($pidList -join ', ')"

# 3) Attempt close/kill per PID, and report parent (for relaunch clues)
foreach ($t in $targets) {
    $pid = $t.ProcessId
    # Try to identify the parent process (may be a service wrapper)
    try {
        $pp = Get-CimInstance Win32_Process -Filter "ProcessId=$($t.ParentProcessId)" -ErrorAction Stop
        Write-LogA ("PID {0} parent: {1} ({2})" -f $pid, $t.ParentProcessId, $pp.Name)
        if ($pp.Name -match '^(services\.exe|nssm\.exe|srvany\.exe|wrapper\.exe)$') {
            Write-LogA "PID $pid likely launched by a service/wrapper. Consider stopping that service first." 2
        }
    } catch {
        Write-LogA "PID $pid parent lookup failed: $($_.Exception.Message)" 2
    }

    try {
        $gp = Get-Process -Id $pid -ErrorAction Stop

        # 3a) Graceful close if there is a window
        if ($gp.MainWindowHandle -ne 0) {
            Write-LogA "PID $pid: attempting CloseMainWindow()"
            [void]$gp.CloseMainWindow()
            if ($gp.WaitForExit($GraceWaitMs)) {
                Write-LogA "PID $pid exited gracefully."
                continue
            } else {
                Write-LogA "PID $pid did not exit after CloseMainWindow()." 2
            }
        } else {
            Write-LogA "PID $pid has no main window; skipping graceful close."
        }

        # 3b) Force kill with Stop-Process
        try {
            Write-LogA "PID $pid: Stop-Process -Force" 2
            Stop-Process -Id $pid -Force -ErrorAction Stop
            Start-Sleep -Milliseconds 500
        } catch {
            Write-LogA "PID $pid Stop-Process failed: $($_.Exception.Message)" 2
        }

        # 3c) If still alive, use WMI Terminate
        if (Get-Process -Id $pid -ErrorAction SilentlyContinue) {
            Write-LogA "PID $pid still alive; trying WMI Terminate()" 2
            try {
                $ci  = Get-CimInstance Win32_Process -Filter "ProcessId=$pid" -ErrorAction Stop
                $res = Invoke-CimMethod -InputObject $ci -MethodName Terminate -Arguments @{ Reason = 5 } -ErrorAction Stop
                Write-LogA "PID $pid WMI Terminate() returned $($res.ReturnValue)."
            } catch {
                Write-LogA "PID $pid WMI Terminate() failed: $($_.Exception.Message)" 2
            }
        }

        # 3d) If somehow still alive, nuke tree with taskkill
        if (Get-Process -Id $pid -ErrorAction SilentlyContinue) {
            Write-LogA "PID $pid still alive; taskkill /T /F" 3
            $tk = Start-Process -FilePath "$env:WINDIR\System32\taskkill.exe" -ArgumentList "/PID $pid /T /F" -PassThru -Wait -WindowStyle Hidden
            Write-LogA "taskkill exit code for PID $pid: $($tk.ExitCode)"
        }

        if (Get-Process -Id $pid -ErrorAction SilentlyContinue) {
            Write-LogA "PID $pid is STILL running after all methods. Likely being relaunched; stop its parent service/wrapper." 3
        } else {
            Write-LogA "PID $pid terminated."
        }
    } catch {
        Write-LogA "PID $pid handling error: $($_.Exception.Message)" 3
    }
}

#endregion
