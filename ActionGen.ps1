Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ====== CONFIG ======
$ServerURL = "https://YourBigFixServer:52311"
$CustomSiteName = "Test Group Managed (Workstations)"
$LogFile = "C:\temp\BigFixActionGen.log"

# Hardcoded ComputerGroupIDs (string to preserve leading zeros)
$GroupIDs = @{
    "Pilot"                    = "00-12345"
    "Deploy"                   = "00-12346"
    "Force"                    = "00-12347"
    "Conference/Training Rooms"= "00-12348"
}

# ====== LOGGING ======
function Write-Log {
    param([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "$timestamp $message"
    $LogBox.AppendText("$entry`r`n")
    Add-Content -Path $LogFile -Value $entry
}

# ====== CLEAN TITLE ======
function Parse-FixletTitleToProduct {
    param([string]$Title)
    $clean = $Title -replace '^Update:\s*', ''
    $clean = $clean -replace '\s+Win$', ''
    return $clean.Trim()
}

# ====== GET FIXLET XML ======
function Get-FixletData {
    param(
        [string]$ServerURL,
        [string]$SiteName,
        [string]$FixletID,
        [string]$Username,
        [string]$Password
    )
    $encodedSite = [System.Web.HttpUtility]::UrlEncode($SiteName)
    $uri = "$ServerURL/api/fixlet/custom/$encodedSite/$FixletID"
    Write-Log "Fetching Fixlet from: $uri"

    try {
        $credPair = "$Username`:$Password"
        $credBytes = [System.Text.Encoding]::ASCII.GetBytes($credPair)
        $credBase64 = [System.Convert]::ToBase64String($credBytes)

        $headers = @{ Authorization = "Basic $credBase64" }

        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
        return $response
    }
    catch {
        Write-Log "❌ Failed to retrieve Fixlet: $_"
        return $null
    }
}

# ====== EXTRACT ACTIONSCRIPT & RELEVANCE ======
function Extract-ActionAndRelevance {
    param($FixletXML)
    try {
        $ns = New-Object System.Xml.XmlNamespaceManager($FixletXML.NameTable)
        $ns.AddNamespace("BES", $FixletXML.DocumentElement.NamespaceURI)

        $titleNode = $FixletXML.SelectSingleNode("//BES:Title", $ns)
        $title = $titleNode.InnerText

        $relevanceNodes = $FixletXML.SelectNodes("//BES:Relevance", $ns)
        $relevanceList = $relevanceNodes | ForEach-Object { $_.InnerText }

        $actionNode = $FixletXML.SelectSingleNode("//BES:Action", $ns)
        if (-not $actionNode) {
            throw "No <Action> node found in Fixlet."
        }

        $actionScriptNode = $actionNode.SelectSingleNode("BES:ActionScript", $ns)
        $actionScript = $actionScriptNode.InnerText

        return @{
            Title = $title
            Relevance = $relevanceList
            ActionScript = $actionScript
        }
    }
    catch {
        Write-Log "❌ Error extracting Action/Relevance: $_"
        return $null
    }
}

# ====== BUILD ACTION XML ======
function Build-ActionXML {
    param(
        [string]$DisplayName,
        [string]$FixletID,
        [string]$ActionScript,
        [string[]]$Relevance,
        [datetime]$StartTime,
        [string]$GroupID,
        [switch]$IsForce
    )

    $deadlineString = ""
    if ($IsForce) {
        $deadline = $StartTime.AddHours(24)
        $deadlineString = "<EndDateTimeLocalOffset>$($deadline.ToString("yyyy-MM-ddTHH:mm:sszzz"))</EndDateTimeLocalOffset>"
    }

    $relevanceXML = ($Relevance | ForEach-Object { "<Relevance>$_</Relevance>" }) -join "`n"

    $xml = @"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SourcedFixletAction>
    <SourceFixlet>
      <Sitename>CustomSite</Sitename>
      <FixletID>$FixletID</FixletID>
      <Action>Action1</Action>
    </SourceFixlet>
    <Target>
      <ComputerGroupID>$GroupID</ComputerGroupID>
    </Target>
    $relevanceXML
    <ActionScript MIMEType="application/x-Fixlet-Windows-Shell">$([System.Security.SecurityElement]::Escape($ActionScript))</ActionScript>
    <Settings>
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage>
        <Text>Updating to $DisplayName. Please wait...</Text>
      </RunningMessage>
      <HasTimeRange>true</HasTimeRange>
      <StartDateTimeLocalOffset>$($StartTime.ToString("yyyy-MM-ddTHH:mm:sszzz"))</StartDateTimeLocalOffset>
      $deadlineString
    </Settings>
  </SourcedFixletAction>
</BES>
"@
    return $xml
}

# ====== POST ACTION ======
function Post-Action {
    param(
        [string]$ServerURL,
        [string]$Username,
        [string]$Password,
        [string]$ActionXML
    )
    try {
        $credPair = "$Username`:$Password"
        $credBytes = [System.Text.Encoding]::ASCII.GetBytes($credPair)
        $credBase64 = [System.Convert]::ToBase64String($credBytes)
        $headers = @{ Authorization = "Basic $credBase64" }

        Write-Log "Posting action to $ServerURL/api/actions"
        $result = Invoke-RestMethod -Uri "$ServerURL/api/actions" -Headers $headers -Method Post -Body $ActionXML -ContentType "application/xml"
        Write-Log "✅ Action posted successfully."
    }
    catch {
        Write-Log "❌ Failed to post action: $_"
    }
}

# ====== GUI ======
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(750,600)
$form.StartPosition = "CenterScreen"

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = "BigFix Server URL:"
$lblServer.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($lblServer)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Text = $ServerURL
$txtServer.Location = New-Object System.Drawing.Point(150,18)
$txtServer.Width = 300
$form.Controls.Add($txtServer)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Username:"
$lblUser.Location = New-Object System.Drawing.Point(10,60)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(150,58)
$txtUser.Width = 200
$form.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.Location = New-Object System.Drawing.Point(10,100)
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.MaskedTextBox
$txtPass.Location = New-Object System.Drawing.Point(150,98)
$txtPass.Width = 200
$txtPass.PasswordChar = '*'
$form.Controls.Add($txtPass)

$lblFixletID = New-Object System.Windows.Forms.Label
$lblFixletID.Text = "Fixlet ID:"
$lblFixletID.Location = New-Object System.Drawing.Point(10,140)
$form.Controls.Add($lblFixletID)

$txtFixletID = New-Object System.Windows.Forms.TextBox
$txtFixletID.Location = New-Object System.Drawing.Point(150,138)
$txtFixletID.Width = 100
$form.Controls.Add($txtFixletID)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Schedule Date:"
$lblDate.Location = New-Object System.Drawing.Point(10,180)
$form.Controls.Add($lblDate)

$cbDate = New-Object System.Windows.Forms.ComboBox
$cbDate.Location = New-Object System.Drawing.Point(150,178)
$cbDate.Width = 150
# Populate only future Wednesdays
$today = Get-Date
for ($i=1; $i -le 60; $i++) {
    $date = $today.AddDays($i)
    if ($date.DayOfWeek -eq 'Wednesday') {
        $cbDate.Items.Add($date.ToString("yyyy-MM-dd"))
    }
}
$form.Controls.Add($cbDate)

$lblTime = New-Object System.Windows.Forms.Label
$lblTime.Text = "Start Time:"
$lblTime.Location = New-Object System.Drawing.Point(10,220)
$form.Controls.Add($lblTime)

$cbTime = New-Object System.Windows.Forms.ComboBox
$cbTime.Location = New-Object System.Drawing.Point(150,218)
$cbTime.Width = 100
# Populate 15-min increments from 8:00 PM to 11:45 PM
for ($h=20; $h -le 23; $h++) {
    foreach ($m in 0,15,30,45) {
        $time = (Get-Date -Hour $h -Minute $m -Second 0).ToString("hh:mm tt")
        $cbTime.Items.Add($time)
    }
}
$form.Controls.Add($cbTime)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Generate & Post Actions"
$btnRun.Location = New-Object System.Drawing.Point(10,260)
$btnRun.Width = 200
$form.Controls.Add($btnRun)

$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Location = New-Object System.Drawing.Point(10,300)
$LogBox.Size = New-Object System.Drawing.Size(700,240)
$LogBox.Multiline = $true
$LogBox.ScrollBars = "Vertical"
$form.Controls.Add($LogBox)

# ====== BUTTON CLICK ======
$btnRun.Add_Click({
    $fixletID = $txtFixletID.Text
    $date = $cbDate.SelectedItem
    $time = $cbTime.SelectedItem
    if (-not $fixletID -or -not $date -or -not $time) {
        Write-Log "❌ Please fill all fields."
        return
    }
    $startDT = Get-Date "$date $time"

    $fixletXML = Get-FixletData -ServerURL $txtServer.Text -SiteName $CustomSiteName -FixletID $fixletID -Username $txtUser.Text -Password $txtPass.Text
    if (-not $fixletXML) { return }

    $data = Extract-ActionAndRelevance -FixletXML $fixletXML
    if (-not $data) { return }

    $displayName = Parse-FixletTitleToProduct -Title $data.Title

    foreach ($action in $GroupIDs.Keys) {
        $isForce = $false
        if ($action -eq "Force") { $isForce = $true }
        $xmlBody = Build-ActionXML -DisplayName $displayName -FixletID $fixletID -ActionScript $data.ActionScript -Relevance $data.Relevance -StartTime $startDT -GroupID $GroupIDs[$action] -IsForce:$isForce
        Write-Log "Generated XML for $action:`r`n$xmlBody"
        Post-Action -ServerURL $txtServer.Text -Username $txtUser.Text -Password $txtPass.Text -ActionXML $xmlBody
    }
})

[void]$form.ShowDialog()
