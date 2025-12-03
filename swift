	# Update existing JSON file if a setting is missing or incorrect.
	$FULLPATH=Get-ChildItem -Path C:\Users -Filter Settings.json -Recurse -ErrorAction SilentlyContinue -Force | %{$_.FullName}
	$USERSVPATH=Split-Path -Path $FULLPATH
	ForEach ($ONEPATH in $USERSVPATH) 
	{ 
		If ($ONEPATH -ilike '*\code\*')
		{
			$JSONPATH=$ONEPATH + '\Settings.json'
			$GLOBALSTORAGEDIR=$ONEPATH + '\GlobalStorage'
			$json = Get-Content -Path $JSONPATH | ConvertFrom-Json -ErrorAction SilentlyContinue
				if ($json."telemetry.enableTelemetry" -ne "false" -or $json."update.enableWindowsBackgroundUpdates" -ne "false" -or $json."update.mode" -ne "none" -or $json."telemetry.enableCrashReporter" -ne "false") 
				{
					# Update the settings if they are incorrect

                    $json | Add-Member -MemberType NoteProperty -Name telemetry -Value ([PSCustomObject]@{})
                    $json.telemetry | Add-Member -MemberType NoteProperty -Name enableTelemetry -Value $false
                    $json.telemetry | Add-Member -MemberType NoteProperty -Name enableCrashReporter -Value $false

                    $json | Add-Member -MemberType NoteProperty -Name update -Value ([PSCustomObject]@{})
                    $json.update | Add-Member -MemberType NoteProperty -Name enableWindowsBackgroundUpdates -Value $false
                    $json.update | Add-Member -MemberType NoteProperty -Name mode -Value "none"

					$json.telemetry.enableTelemetry = $false
					$json.update.enableWindowsBackgroundUpdates = $false
					$json.update.mode = "none"
					$json.telemetry.enableCrashReporter = $false
					Write-ADTLogEntry -Message "Updated $JSONPATH" -Source $adtSession.InstallPhase
				}
			else 
			{
				# Create a new JSON object with the desired settings
				$newJson = @{				
					"telemetry.enableTelemetry" = "false";
					"update.enableWindowsBackgroundUpdates" = "false";
					"update.mode" = "none";
					"telemetry.enableCrashReporter" = "false"
				}
				# If the file is empty, overwrite it with the new JSON object
				$json = $newJson
			}
			# Convert the PowerShell object back to JSON and save it to the file
			$json | ConvertTo-Json -Depth 100 | Set-Content -Path $JSONPATH
