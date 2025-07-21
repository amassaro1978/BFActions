Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(700, 540)
$form.StartPosition = "CenterScreen"

# Server Info
$labelServer = New-Object System.Windows.Forms.Label -Property @{ Text = "Server:"; Location = New-Object System.Drawing.Point(10, 20); AutoSize = $true }
$textServer = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 18); Width = 550 }

$labelUser = New-Object System.Windows.Forms.Label -Property @{ Text = "Username:"; Location = New-Object System.Drawing.Point(10, 50); AutoSize = $true }
$textUser = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 48); Width = 550 }

$labelPass = New-Object System.Windows.Forms.Label -Property @{ Text = "Password:"; Location = New-Object System.Drawing.Point(10, 80); AutoSize = $true }
$textPass = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 78); Width = 550; UseSystemPasswordChar = $true }

# Fixlet Name
$labelFixlet = New-Object System.Windows.Forms.Label -Property @{ Text = "Fixlet Name:"; Location = New-Object System.Drawing.Point(10, 120); AutoSize = $true }
$textFixlet = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 118); Width = 550 }

# Fixlet ID
$labelFixletID = New-Object System.Windows.Forms.Label -Property @{ Text = "Fixlet ID:"; Location = New-Object System.Drawing.Point(10, 150); AutoSize = $true }
$textFixletID = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 148); Width = 550 }

# Date Selector (future Wednesdays only)
$labelDate = New-Object System.Windows.Forms.Label -Property @{ Text = "Select Wednesday:"; Location = New-Object System.Drawing.Point(10, 190); AutoSize = $true }
$comboDate = New-Object System.Windows.Forms.ComboBox -Property @{ Location = New-Object System.Drawing.Point(130, 188); Width = 200; DropDownStyle = "DropDownList" }

$today = [datetime]::Today
for ($i = 1; $i -le 60; $i++) {
    $day = $today.AddDays($i)
    if ($day.DayOfWeek -eq 'Wednesday') {
        $comboDate.Items.Add($day.ToString("yyyy-MM-dd"))
    }
}

# Time Selector
$labelTime = New-Object System.Windows.Forms.Label -Property @{ Text = "Select Time:"; Location = New-Object System.Drawing.Point(350, 190); AutoSize = $true }
$comboTime = New-Object System.Windows.Forms.ComboBox -Property @{ Location = New-Object System.Drawing.Point(440, 188); Width = 100; DropDownStyle = "DropDownList" }

foreach ($time in @(
    "8:00 PM", "8:15 PM", "8:30 PM", "8:45 PM",
    "9:00 PM", "9:15 PM", "9:30 PM", "9:45 PM",
    "10:00 PM", "10:15 PM", "10:30 PM", "10:45 PM",
    "11:00 PM", "11:15 PM", "11:30 PM", "11:45 PM"
)) {
    $comboTime.Items.Add($time)
}

# Button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Create Actions"
$btnGenerate.Size = New-Object System.Drawing.Size(150, 30)
$btnGenerate.Location = New-Object System.Drawing.Point(270, 240)

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
    $cred = New-Object System.Management.Automation.PSCredential($user, $securePass)

    $actions = @(
        @{
            Name = "${fixletName}: Pilot"
            Start = $startDT
            End = $startDT.Date.AddDays(1).AddHours(6).AddMinutes(59)
            GroupID = "12345"
            Message = ""
            Deadline = $null
            ShowUI = $false
        },
        @{
            Name = "${fixletName}: Deploy"
            Start = $startDT.AddDays(1)
            End = $startDT.AddDays(6).Date.AddHours(6).AddMinutes(59)
            GroupID = "12345"
            Message = ""
            Deadline = $null
            ShowUI = $false
        },
        @{
            Name = "${fixletName}: Force"
            Start = $startDT.AddDays(6).Date.AddHours(7)
            End = $startDT.AddDays(6).Date.AddYears(1)
            GroupID = "12345"
            Deadline = $startDT.AddDays(7)
            Message = "Update: $vendor $app $version will be enforced on $($startDT.AddDays(7).ToString('MM/dd/yyyy h:mm tt')). Please leave your machine on overnight to get the automated update. Otherwise, please close the application and run the update now. When the deadline is reached, the action will run automatically."
            ShowUI = $true
        },
        @{
            Name = "${fixletName}: Conference/Training Rooms"
            Start = $startDT.AddDays(1)
            End = $startDT.AddDays(6).Date.AddHours(6).AddMinutes(59)
            GroupID = "12345"
            Message = ""
            Deadline = $null
            ShowUI = $false
        }
    )

    foreach ($action in $actions) {
        $uiXml = @"
<UI>
  <ShowActionButton>true</ShowActionButton>
  <ShowMessage>true</ShowMessage>
  <HasRunningMessage>true</HasRunningMessage>
  <ActionRunningMessage>Updating to $vendor $app $version. Please wait...</ActionRunningMessage>
"@
        if ($action.ShowUI) {
            $uiXml += @"
  <PreActionShowUI>true</PreActionShowUI>
  <PreActionMessage>$($action.Message)</PreActionMessage>
  <PreActionAskToSaveWork>true</PreActionAskToSaveWork>
  <Deadline>$($action.Deadline.ToString("yyyy-MM-dd'T'HH:mm:ss"))</Deadline>
"@
        }
        $uiXml += "</UI>"

        $xml = @"
<BES>
  <SourcedFixletAction>
    <SourceFixletID>$fixletID</SourceFixletID>
    <Title>$($action.Name)</Title>
    <Relevance>TRUE</Relevance>
    <StartDateTimeLocal>$($action.Start.ToString("yyyy-MM-dd'T'HH:mm:ss"))</StartDateTimeLocal>
    <EndDateTimeLocal>$($action.End.ToString("yyyy-MM-dd'T'HH:mm:ss"))</EndDateTimeLocal>
    $uiXml
    <Settings>
      <RetryCount>3</RetryCount>
      <RetryWait>1</RetryWait>
      <Reapply>true</Reapply>
      <ActiveUserRequirement>NoRequirement</ActiveUserRequirement>
    </Settings>
    <Target>
      <ComputerGroupID>$($action.GroupID)</ComputerGroupID>
    </Target>
  </SourcedFixletAction>
</BES>
"@

        $url = "$server/api/actions"
        try {
            Invoke-RestMethod -Uri $url -Method Post -Body $xml -Credential $cred -ContentType "application/xml"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to create action: $($_.Exception.Message)")
        }
    }

    [System.Windows.Forms.MessageBox]::Show("All actions created successfully.")
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
