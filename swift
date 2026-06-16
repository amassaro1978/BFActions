$Installer = "$adtSession.DirFiles\nomachine-enterprise-client_9.6.3_x64.exe"
$Args = '/VERYSILENT /NORESTART /SUPPRESSMSGBOXES /SP- /LOG="C:\Windows\Temp\NoMachineEC_First.log"'

Write-Log "Starting NoMachine Enterprise Client first-pass install."

$proc = Start-Process -FilePath $Installer -ArgumentList $Args -PassThru -WindowStyle Hidden

$deadline = (Get-Date).AddMinutes(10)
$Installed = $false

do {
    Start-Sleep -Seconds 10

    $Installed = Get-ItemProperty `
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' `
        -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -match 'NoMachine' -and
            $_.Publisher -match 'NoMachine'
        }

} until ($Installed -or $proc.HasExited -or (Get-Date) -gt $deadline)

if ($Installed -and -not $proc.HasExited) {
    Write-Log "NoMachine was detected, but first-pass installer is still running. Killing installer PID $($proc.Id)."
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
}
elseif (-not $Installed -and -not $proc.HasExited) {
    Write-Log "NoMachine was not detected and installer exceeded timeout. Killing installer PID $($proc.Id)."
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    throw "NoMachine Enterprise Client first-pass install did not complete."
}
elseif ($proc.HasExited) {
    Write-Log "First-pass installer exited with code $($proc.ExitCode)."
}

Start-Sleep -Seconds 5
