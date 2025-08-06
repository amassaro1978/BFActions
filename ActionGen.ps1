Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(500, 500)
$form.StartPosition = "CenterScreen"

# Create labels and input fields
$labels = @("BigFix Server URL:", "Username:", "Password:", "Fixlet Name:", "Fixlet ID:")
$inputs = @{}
$y = 20

foreach ($labelText in $labels) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(10, $y)
    $label.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(170, $y)
    $textbox.Size = New-Object System.Drawing.Size(280, 20)
    if ($labelText -eq "Password:") { $textbox.UseSystemPasswordChar = $true }

    $form.Controls.Add($textbox)
    $inputs[$labelText.TrimEnd(":")] = $textbox
    $y += 30
}

# Date ComboBox - only future Wednesdays
$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Text = "Select Wednesday Date:"
$labelDate.Location = New-Object System.Drawing.Point(10, $y)
$labelDate.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelDate)

$dateComboBox = New-Object System.Windows.Forms.ComboBox
$dateComboBox.Location = New-Object System.Drawing.Point(170, $y)
$dateComboBox.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($dateComboBox)

$futureWednesdays = 0..30 | ForEach-Object {
    $date = (Get-Date).AddDays($_)
    if ($date.DayOfWeek -eq 'Wednesday') { $date.ToString("yyyy-MM-dd") }
}
$dateComboBox.Items.AddRange($futureWednesdays)
$y += 30

# Time ComboBox - 12-hour format from 8:00 PM to 11:45 PM
$labelTime = New-Object System.Windows.Forms.Label
$labelTime.Text = "Select Start Time:"
$labelTime.Location = New-Object System.Drawing.Point(10, $y)
$labelTime.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelTime)

$timeComboBox = New-Object System.Windows.Forms.ComboBox
$timeComboBox.Location = New-Object System.Drawing.Point(170, $y)
$timeComboBox.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($timeComboBox)

for ($h = 20; $h -le 23; $h++) {
    foreach ($m in 0, 15, 30, 45) {
        $timeComboBox.Items.Add((Get-Date -Hour $h -Minute $m -Format "hh:mm tt"))
    }
}
$y += 40

# Submit button
$submitBtn = New-Object System.Windows.Forms.Button
$submitBtn.Text = "Generate Actions"
$submitBtn.Location = New-Object System.Drawing.Point(170, $y)
$submitBtn.Size = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($submitBtn)

# Output log
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.Size = New-Object System.Drawing.Size(460, 120)
$logBox.Location = New-Object System.Drawing.Point(10, 350)
$form.Controls.Add($logBox)

# Submit click logic
$submitBtn.Add_Click({
    $server = $inputs["BigFix Server URL"].Text.TrimEnd("/")
    $username = $inputs["Username"].Text
    $password = $inputs["Password"].Text
    $fixletName = $inputs["Fixlet Name"].Text
    $fixletID = $inputs["Fixlet ID"].Text
    $selectedDate = $dateComboBox.SelectedItem
    $selectedTime = $timeComboBox.SelectedItem

    if (-not ($server -and $username -and $password -and $fixletName -and $fixletID -and $selectedDate -and $selectedTime)) {
        [System.Windows.Forms.MessageBox]::Show("All fields must be completed.")
        return
    }

    $startDateTime = Get-Date "$selectedDate $selectedTime"
    $formattedStart = $startDateTime.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
    $deadline = $startDateTime.AddHours(24).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")

    if ($fixletName -match "^(.*) - (.*) ([\d\.]+)$") {
        $vendor = $matches[1]
        $app = $matches[2]
        $ver = $matches[3]
    } else {
        $vendor = $app = $ver = "Unknown"
    }

    $actions = @(
        @{ Title = "Pilot"; GroupID = 12345; RunBetween = $true },
        @{ Title = "Deploy"; GroupID = 12345; RunBetween = $true },
        @{ Title = "Force"; GroupID = 12345; RunBetween = $false },
        @{ Title = "Conference/Training Rooms"; GroupID = 12345; RunBetween = $true }
    )

    $siteName = "actionsite"
    $siteURL = "http://sync.bigfix.com/cgi-bin/bfgather.exe/actionsite"

    $actionsXml = foreach ($a in $actions) {
        $runBetweenXml = if ($a.RunBetween) {
            "<StartDateTime>${formattedStart}</StartDateTime>
            <EndDateTime>$([datetime]::Parse($selectedDate).AddDays(1).ToUniversalTime().ToString("yyyy-MM-dd'T'06:59:00'Z'"))</EndDateTime>"
        } elseif ($a.Title -eq "Force") {
            "<StartDateTime>${formattedStart}</StartDateTime>
            <Deadline>${deadline}</Deadline>
            <PreActionShowUI>true</PreActionShowUI>
            <PreAction>
                <Message>⚠️ This update will be enforced in 24 hours.</Message>
            </PreAction>"
        }

@"
<Action>
  <SourcedFixletAction>
    <SourceFixletID>${fixletID}</SourceFixletID>
    <SourceSiteURL>${siteURL}</SourceSiteURL>
    <SourceSiteName>${siteName}</SourceSiteName>
    <Target>
        <ComputerGroupID>${($a.GroupID)}</ComputerGroupID>
    </Target>
    <Title>${fixletName}: $($a.Title)</Title>
    <Settings>
        <HasRunningMessage>false</HasRunningMessage>
        <HasTimeRange>false</HasTimeRange>
        <HasStartTime>false</HasStartTime>
        <HasEndTime>false</HasEndTime>
        <HasDayOfWeekConstraint>false</HasDayOfWeekConstraint>
        <HasReapply>false</HasReapply>
        <HasRetry>false</HasRetry>
        <HasTemporalDistribution>false</HasTemporalDistribution>
        <HasAllowNoFixlets>false</HasAllowNoFixlets>
        <HasWhose>false</HasWhose>
        <HasCustomRelevance>false</HasCustomRelevance>
        ${runBetweenXml}
    </Settings>
  </SourcedFixletAction>
</Action>
"@
    }

    $fullXml = @"
<MultipleActionGroup>
  <Title>${fixletName}: All Actions</Title>
  $($actionsXml -join "`n")
</MultipleActionGroup>
"@

    $encodedUrl = "$server/api/actions"
    $logBox.AppendText("POST to: $encodedUrl`r`n")
    $logBox.AppendText("XML:`r`n$fullXml`r`n")

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($fullXml)
    $headers = @{
        Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
        "Content-Type" = "application/xml"
    }

    try {
        $response = Invoke-RestMethod -Uri $encodedUrl -Method Post -Headers $headers -Body $bytes
        $logBox.AppendText("✅ Success`r`n")
    } catch {
        $logBox.AppendText("❌ Failed: $_`r`n")
    }
})

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
