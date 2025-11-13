#region Close only Sirsi/JWF javaw.exe
# Path to the bundled JRE\bin folder (trailing backslash not required)
$JreBinPath   = 'C:\Program Files (x86)\Sirsi\JWF\JRE\bin'
$TargetExe    = Join-Path $JreBinPath 'javaw.exe'

try {
    # Get only javaw.exe processes
    $javawList = Get-CimInstance -ClassName Win32_Process -Query "
        SELECT ProcessId, ExecutablePath, CommandLine FROM Win32_Process
        WHERE Name='javaw.exe'
    " -ErrorAction Stop

    # Filter to just the bundled JRE instances
    $target = $javawList | Where-Object {
        $_.ExecutablePath -and (
            $_.ExecutablePath -ieq $TargetExe -or
            $_.ExecutablePath.StartsWith($JreBinPath, [System.StringComparison]::InvariantCultureIgnoreCase)
        )
    }

    if (-not $target) {
        Write-Log -Message "No Sirsi/JWF javaw.exe instances found." -Severity 1
    } else {
        $pids = $target.ProcessId
        Write-Log -Message "Found Sirsi/JWF javaw.exe PID(s): $($pids -join ', ')" -Severity 1

        foreach ($pid in $pids) {
            try {
                $proc = Get-Process -Id $pid -ErrorAction Stop

                if ($proc.MainWindowHandle -ne 0) {
                    Write-Log -Message "Attempting graceful close for PID $pid (windowed)." -Severity 1
                    [void]$proc.CloseMainWindow()
                    if (-not $proc.WaitForExit(10 * 1000)) {
                        Write-Log -Message "Graceful close timed out for PID $pid. Forcing termination." -Severity 2
                        Stop-Process -Id $pid -Force -ErrorAction Stop
                    } else {
                        Write-Log -Message "PID $pid exited gracefully." -Severity 1
                    }
                } else {
                    Write-Log -Message "PID $pid has no window. Forcing termination." -Severity 2
                    Stop-Process -Id $pid -Force -ErrorAction Stop
                }
            } catch {
                Write-Log -Message "Failed to close PID $pid. $_" -Severity 3
            }
        }
    }
} catch {
    Write-Log -Message "Error enumerating javaw.exe processes. $_" -Severity 3
}
#endregion
