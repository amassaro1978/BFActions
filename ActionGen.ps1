Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(700, 550)
$form.StartPosition = "CenterScreen"

# Labels and TextBoxes
$labelServer = New-Object System.Windows.Forms.Label -Property @{ Text = "Server:"; Location = '10,20'; AutoSize = $true }
$textServer = New-Object System.Windows.Forms.TextBox -Property @{ Location = '100,18'; Width = 550 }

$labelUser = New-Object System.Windows.Forms.Label -Property @{ Text = "Username:"; Location = '10,50'; AutoSize = $true }
$textUser = New-Object System.Windows.Forms.TextBox -Property @{ Location = '100,48'; Width = 550 }

$labelPass = New-Object System.Windows.Forms.Label -Property @{ Text = "Password:"; Location = '10,80'; AutoSize = $true }
$textPass = New-Object System.Windows.Forms.TextBox -Property @{ Location = '100,78'; Width = 550; UseSystemPasswordChar = $true }

$labelFixlet = New-Object System.Windows.Forms.Label -Property @{ Text = "Fixlet Name:"; Location = '10,120'; AutoSize = $true }
$textFixlet = New-Object System.Windows.Forms.TextBox -Property @{ Location = '100,118'; Width = 550 }

$labelFixletID = New-Object System.Windows.Forms.Label -Property @{ Text = "Fixlet ID:"; Location = '10,150'; AutoSize = $true }
$textFixletID = New-Object System.Windows.Forms.TextBox -Property @{ Location = '100,148'; Width = 550 }

$labelDate = New-Object System.Windows.Forms.Label -Property @{ Text = "Select Wednesday:"; Location = '10,190'; AutoSize = $true }
$comboDate = New-Object System.Windows.Forms.ComboBox -Property @{ Location = '130,188'; Width = 200; DropDownStyle = "DropDownList" }

# Populate 8 upcoming Wednesdays
$today = [datetime]::Today
$wedCount = 0
for ($i = 1; $i -le 60 -and $wedCount -lt 8; $i++) {
    $day = $today.AddDays($i)
    if ($day.DayOfWeek -eq 'Wednesday') {
        $comboDate.Items.Add($day.ToString("yyyy-MM-dd"))
        $wedCount++
    }
}

$labelTime = New-Object System.Windows.Forms.Label -Property @{ Text = "Select Time:"; Location = '350,190'; AutoSize = $true }
$comboTime = New-Object System.Windows.Forms.ComboBox -Property @{ Location = '440,188'; Width = 100; DropDownStyle = "DropDownList" }

foreach ($time in @(
    "8:00 PM", "8:15 PM", "8:30 PM", "8:45 PM",
    "9:00 PM", "9:15 PM", "9:30 PM", "9:45 PM",
    "10:00 PM", "10:15 PM", "10:30 PM", "10:45 PM",
    "11:00 PM", "11:15 PM", "11:30 PM", "11:45 PM"
)) {
    $comboTime.Items.Add($time)
}

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Create Actions"
$btnGenerate.Size = '150,30'
$btnGenerate.Location = '270,240'

$btnGenerate.Add_Click({
    $server = $textServer.Text.Trim().TrimEnd("/")
    $user = $textUser.Text.Trim()
    $pass = $textPass.Text
    $fixletName = $textFixlet.Text.Trim()
    $fixletID = $textFixletID.Text.Trim()
    $selectedDate = $comboDate.SelectedItem
    $selectedTime = $comboTime.SelectedItem

    if (-not ($server -and $user -and $pass -and $fixletName -and $fixletID -and $selectedDate -and $selectedTime)) {
        [System.Windows.Forms.MessageBox]::Show("All fields must be filled.")
        return
    }

    $startDT = [datetime]::ParseExact("$selectedDate $selectedTime", "yyyy-MM-dd h:mm tt", $null)

    if ($fixletName -match "^(.+?)\s+(.+?)\s+(\d[\d\.]*)") {
        $vendor = $matches[1]
        $app = $matches[2]
        $version = $matches[3]
    } else {
        [System.Windows.Forms.MessageBox]::Show("Unable to parse vendor/app/version from Fixlet name.")
        return
    }

    $securePass = ConvertTo-SecureString $pass -AsPlainText -Force
    $cred = New-Object PSCredential ($user, $securePass)

    $deadline = $startDT.AddDays(7)
    $groupID = 12345

    $actions = @(
        @{ Name = "Pilot"; Start=$startDT; End=$startDT.AddDays(1).AddHours(6).AddMinutes(59); Deadline=$null; Message=$null },
        @{ Name = "Deploy"; Start=$startDT.AddDays(1); End=$startDT.AddDays(6).AddHours(6).AddMinutes(59); Deadline=$null; Message=$null },
        @{ Name = "Force"; Start=$startDT.AddDays(6).AddHours(7); End=$startDT.AddDays(6).AddYears(1); Deadline=$deadline;
            Message="Update: $vendor $app $version will be enforced on $($deadline.ToString('MM/dd/yyyy h:mm tt')). Please leave your machine on overnight to get the automated update. Otherwise, please close the application and run the update now. When the deadline is reached, the action will run automatically." },
        @{ Name = "Conference/Training Rooms"; Start=$startDT.AddDays(1); End=$startDT.AddDays(6).AddHours(6).AddMinutes(59); Deadline=$null; Message=$null }
    )

    $xml = @"
<BES>
  <MultipleActionGroup>
    <Title>${fixletName} - Action Group</Title>
    <Relevance>TRUE</Relevance>
    <UseCustomGroup>true</UseCustomGroup>
    <ActionGroupCreationTime>${([datetime]::Now.ToString("yyyy-MM-dd'T'HH:mm:ss"))}</ActionGroupCreationTime>
    <SiteName>CustomSite</SiteName>
    <SourcedFixletID>$fixletID</SourcedFixletID>
    <CustomGroupTarget>
      <ComputerGroupID>$groupID</ComputerGroupID>
    </CustomGroupTarget>
"@

    foreach ($a in $actions) {
        $xml += @"
    <SourcedFixletAction>
      <Title>${fixletName}: $($a.Name)</Title>
      <Relevance>TRUE</Relevance>
      <StartDateTimeLocal>$($a.Start.ToString("yyyy-MM-dd'T'HH:mm:ss"))</StartDateTimeLocal>
      <EndDateTimeLocal>$($a.End.ToString("yyyy-MM-dd'T'HH:mm:ss"))</EndDateTimeLocal>
      <UI>
        <ShowActionButton>true</ShowActionButton>
        <HasRunningMessage>true</HasRunningMessage>
        <ActionRunningMessage>Installing $vendor $app $version. Please wait...</ActionRunningMessage>
"@
        if ($a.Message) {
            $xml += @"
        <PreActionShowUI>true</PreActionShowUI>
        <PreActionMessage>$($a.Message)</PreActionMessage>
        <PreActionAskToSaveWork>true</PreActionAskToSaveWork>
        <Deadline>$($a.Deadline.ToString("yyyy-MM-dd'T'HH:mm:ss"))</Deadline>
"@
        }
        $xml += @"
      </UI>
      <Settings>
        <Reapply>true</Reapply>
        <RetryCount>3</RetryCount>
        <RetryWait>1</RetryWait>
        <ActiveUserRequirement>NoRequirement</ActiveUserRequirement>
      </Settings>
    </SourcedFixletAction>
"@
    }

    $xml += "</MultipleActionGroup></BES>"

    $url = "$server/api/actions"
    $logPath = "$env:TEMP\BigFix_ActionGen.log"

    "`n---`n[$(Get-Date)] Posting to: $url" | Out-File -Append $logPath
    $xml | Out-File -Append $logPath

    try {
        $resp = Invoke-RestMethod -Uri $url -Method Post -Body $xml -Credential $cred -ContentType "application/xml"
        "Result: SUCCESS" | Out-File -Append $logPath
        [System.Windows.Forms.MessageBox]::Show("All actions created successfully.")
    } catch {
        "Result: FAILED - $($_.Exception.Message)" | Out-File -Append $logPath
        [System.Windows.Forms.MessageBox]::Show("Failed to create actions: $($_.Exception.Message)")
    }
})

$form.Controls.AddRange(@(
    $labelServer, $textServer,
    $labelUser, $textUser,
    $labelPass, $textPass,
    $labelFixlet, $textFixlet,
    $labelFixletID, $textFixletID,
    $labelDate, $comboDate,
    $labelTime, $comboTime,
    $btnGenerate
))

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
