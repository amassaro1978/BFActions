# Update existing JSON file if a setting is missing or incorrect.
$FULLPATH = Get-ChildItem -Path 'C:\Users' -Filter 'Settings.json' -Recurse -ErrorAction SilentlyContinue -Force | ForEach-Object { $_.FullName }

foreach ($JSONPATH in $FULLPATH) {
    # Only target VS Code settings.json under ...\Code\User\
    if ($JSONPATH -notmatch '\\Code\\User\\Settings\.json$') { continue }

    $USERSVPATH       = Split-Path -Path $JSONPATH
    $GLOBALSTORAGEDIR = Join-Path -Path $USERSVPATH -ChildPath 'GlobalStorage'

    # Desired VS Code settings
    $desiredSettings = @{
        'telemetry.enableTelemetry'             = $false
        'update.enableWindowsBackgroundUpdates' = $false
        'update.mode'                           = 'none'
        'telemetry.enableCrashReporter'         = $false
    }

    # Read existing JSON (if any)
    $jsonObject = $null
    $fileContent = $null

    try {
        if (Test-Path -LiteralPath $JSONPATH) {
            $fileContent = Get-Content -LiteralPath $JSONPATH -Raw -ErrorAction SilentlyContinue
        }

        if ([string]::IsNullOrWhiteSpace($fileContent)) {
            # Empty or missing file – start with an empty object
            $jsonObject = [PSCustomObject]@{}
        }
        else {
            $jsonObject = $fileContent | ConvertFrom-Json -ErrorAction Stop
        }
    }
    catch {
        # If parsing fails, log and start fresh
        Write-ADTLogEntry -Message "Failed to parse JSON from '$JSONPATH': $($_.Exception.Message). Initializing empty object." -Source $adtSession.InstallPhase
        $jsonObject = [PSCustomObject]@{}
    }

    if ($null -eq $jsonObject) {
        $jsonObject = [PSCustomObject]@{}
    }

    $changed = $false

    # Ensure each desired setting exists and is set to the correct value
    foreach ($setting in $desiredSettings.GetEnumerator()) {
        $name  = $setting.Key
        $value = $setting.Value

        $prop = $jsonObject.PSObject.Properties[$name]

        if ($null -eq $prop) {
            # Property doesn't exist – add it
            $jsonObject | Add-Member -MemberType NoteProperty -Name $name -Value $value
            $changed = $true
            Write-ADTLogEntry -Message "Added '$name' = '$value' to '$JSONPATH'." -Source $adtSession.InstallPhase
        }
        else {
            # Property exists – update if different
            if ($prop.Value -ne $value) {
                $prop.Value = $value
                $changed = $true
                Write-ADTLogEntry -Message "Updated '$name' to '$value' in '$JSONPATH'." -Source $adtSession.InstallPhase
            }
        }
    }

    # Write back only if something changed
    if ($changed) {
        try {
            $jsonObject | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $JSONPATH -Encoding UTF8
            Write-ADTLogEntry -Message "Updated VS Code settings at '$JSONPATH'." -Source $adtSession.InstallPhase
        }
        catch {
            Write-ADTLogEntry -Message "Failed to write '$JSONPATH': $($_.Exception.Message)." -Source $adtSession.InstallPhase
        }
    }
    else {
        Write-ADTLogEntry -Message "No changes required for '$JSONPATH'." -Source $adtSession.InstallPhase
    }
}
