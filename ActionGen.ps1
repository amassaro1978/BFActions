Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG
# =========================
$SiteName = "Test Group Managed (Workstations)"  # hardcoded site

# Map each action to its Computer Group ID (keep as strings to preserve 00- prefix).
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12345"
    "Force"                     = "00-12345"
    "Conference/Training Rooms" = "00-12345"
}

# =========================
# ENCODING / URL / AUTH HELPERS
# =========================
function Encode-SiteName {
    param([string]$Name)
    # Encode, then normalize + to %20 and encode parentheses (to match what works in curl)
    $enc = [System.Web.HttpUtility]::UrlEncode($Name, [System.Text.Encoding]::UTF8)
    $enc = $enc -replace '\+','%20' -replace '\(','%28' -replace '\)','%29'
    return $enc
}
function Get-BaseUrl([string]$ServerInput) {
    if (-not $ServerInput) { throw "Server is empty." }
    $s = $ServerInput.Trim()
    if ($s -match '^(?i)https?://') { return ($s.TrimEnd('/')) }
    $s = $s.Trim('/')
    if ($s -match ':\d+$') { "https://$s" } else { "https://$s:52311" }
}
function Join-ApiUrl([string]$BaseUrl,[string]$RelativePath) {
    $rp = if ($RelativePath.StartsWith("/")) { $RelativePath } else { "/$RelativePath" }
    $BaseUrl.TrimEnd('/') + $rp
}
function Get-AuthHeader([string]$Username,[string]$Password) {
    $pair  = "$Username`:$Password"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    "Basic " + [Convert]::ToBase64String($bytes)
}

# =========================
# HTTP (curl-like) HELPERS
# =========================
function HttpGetXml {
    param([string]$Url,[string]$AuthHeader)
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = "GET"
    $req.Accept = "application/xml"
    $req.Headers["Authorization"] = $AuthHeader
    $req.ProtocolVersion = [Version]"1.1"
    $req.PreAuthenticate = $true
    $req.AllowAutoRedirect = $false
    $req.Timeout = 30000
    try {
        $resp = $req.GetResponse()
        try {
            $stream = $resp.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $content = $reader.ReadToEnd()
            $reader.Close()
        } finally { $resp.Close() }
        return $content
    } catch {
        throw ($_.Exception.GetBaseException().Message)
    }
}
function HttpPostXml {
    param([string]$Url,[string]$AuthHeader,[string]$XmlBody)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($XmlBody)
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = "POST"
    $req.Accept = "application/xml"
    $req.ContentType = "application/xml; charset=utf-8"
    $req.Headers["Authorization"] = $AuthHeader
    $req.ProtocolVersion = [Version]"1.1"
    $req.PreAuthenticate = $true
    $req.AllowAutoRedirect = $false
    $req.Timeout = 60000
    $req.ContentLength = $bytes.Length
    try {
        $rs = $req.GetRequestStream()
        try { $rs.Write($bytes, 0, $bytes.Length) } finally { $rs.Close() }
        $resp = $req.GetResponse(); $resp.Close()
    } catch {
        throw ($_.Exception.GetBaseException().Message)
    }
}

# =========================
# BIGFIX PARSING HELPERS (robust)
# =========================
function Get-FixletContainer {
    param([xml]$Xml)
    if ($Xml.BES.Fixlet)   { return @{ Type="Fixlet";   Node=$Xml.BES.Fixlet } }
    if ($Xml.BES.Task)     { return @{ Type="Task";     Node=$Xml.BES.Task } }
    if ($Xml.BES.Baseline) { return @{ Type="Baseline"; Node=$Xml.BES.Baseline } }
    throw "Unknown BES content type (no <Fixlet>, <Task>, or <Baseline> root)."
}
function Get-ActionAndRelevance {
    param($ContainerNode)
    # Relevance
    $relevance = @()
    foreach ($r in $ContainerNode.Relevance) { $relevance += [string]$r }

    # Action node: prefer explicit <Action>, fall back to <DefaultAction>
    $actionNode = $null
    if ($ContainerNode.Action) { $actionNode = $ContainerNode.Action | Select-Object -First 1 }
    if (-not $actionNode -and $ContainerNode.DefaultAction) { $actionNode = $ContainerNode.DefaultAction }

    if (-not $actionNode) { throw "No <Action> or <DefaultAction> block found." }

    # Extract ActionScript text
    $script = $null
    if ($actionNode.ActionScript) {
        $script = [string]$actionNode.ActionScript.'#text'
        if ([string]::IsNullOrWhiteSpace($script)) {
            $script = $actionNode.ActionScript.InnerText  # fallback
        }
    }
    if (-not $script) { throw "Action found but no <ActionScript> content present." }

    return @{ Relevance = $relevance; ActionScript = $script }
}

# =========================
# SCHEDULING HELPERS
# =========================
function Get-NextWednesdays {
    $dates = @()
    $today = Get-Date
    $daysUntilWed = (3 - [int]$today.DayOfWeek + 7) % 7  # Wednesday=3
    $nextWed = $today.AddDays($daysUntilWed)
    for ($i = 0; $i -lt 20; $i++) { $dates += $nextWed.AddDays(7*$i).ToString("yyyy-MM-dd") }
    return $dates
}
function Get-TimeSlots {
    $slots = @()
    $start = Get-Date "20:00"; $end = Get-Date "23:45"
    while ($start -le $end) { $slots += $start.ToString("h:mm tt"); $start = $start.AddMinutes(15) }
    return $slots
}
function Format-LocalBESDateTime([datetime]$dt) { $dt.ToString("yyyyMMdd'T'HHmmss") }

# =========================
# XML BUILDER
# =========================
function Build-SingleActionXml {
    param(
        [string]$ActionTitle,[string]$DisplayName,[string[]]$RelevanceBlocks,
        [string]$ActionScript,[datetime]$StartLocal,[bool]$SetDeadline = $false,
        [datetime]$DeadlineLocal = $null,[string]$GroupId
    )
    $titleText = "$($DisplayName): $ActionTitle"
    $titleEsc  = [System.Security.SecurityElement]::Escape($titleText)
    $dispEsc   = [System.Security.SecurityElement]::Escape($DisplayName)

@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SingleAction>
    <Title>$titleEsc</Title>
    $(($RelevanceBlocks | ForEach-Object { "    <Relevance>$([System.Security.SecurityElement]::Escape($_))</Relevance>" }) -join "`r`n")
    <ActionScript MIMEType="application/x-Fixlet-Windows-Shell"><![CDATA[
$ActionScript
]]></ActionScript>
    <Settings>
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage>Updating to $dispEsc. Please wait....</RunningMessage>
      <HasTimeRange>true</HasTimeRange>
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocal>$(Format-LocalBESDateTime $StartLocal)</StartDateTimeLocal>
      <HasEndTime>false</HasEndTime>
      <HasDeadline>$([string]$SetDeadline)</HasDeadline>
      $(if ($SetDeadline -and $DeadlineLocal) { "<Deadline>$(Format-LocalBESDateTime $DeadlineLocal)</Deadline>" } else { "" })
      <HasReapply>false</HasReapply>
      <HasRetry>false</HasRetry>
      <HasTemporalDistribution>false</HasTemporalDistribution>
      <HasAllowNoFixlets>false</HasAllowNoFixlets>
    </Settings>
    <Target>
      <CustomRelevance>member of group whose (id of it as string = "$GroupId")</CustomRelevance>
    </Target>
  </SingleAction>
</BES>
"@
}

# =========================
# BIGFIX API (compose with helpers)
# =========================
function Get-FixletDetails {
    param (
        [Parameter(Mandatory=$true)][string]$Server,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$Password,
        [Parameter(Mandatory=$true)][string]$FixletID
    )
    $base = Get-BaseUrl $Server
    $encodedSite = Encode-SiteName $SiteName
    $path = "/api/fixlet/custom/$encodedSite/$FixletID"
    $url  = Join-ApiUrl -BaseUrl $base -RelativePath $path
    $auth = Get-AuthHeader -Username $Username -Password $Password
    $content = HttpGetXml -Url $url -AuthHeader $auth
    [pscustomobject]@{ Url = $url; Content = $content }
}
function Post-ActionXml {
    param([string]$Server,[string]$Username,[string]$Password,[string]$XmlBody)
    $base = Get-BaseUrl $Server
    $url  = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
    $auth = Get-AuthHeader -Username $Username -Password $Password
    HttpPostXml -Url $url -AuthHeader $auth -XmlBody $XmlBody | Out-Null
    return $url
}

# =========================
# GUI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(560, 760)
$form.StartPosition = "CenterScreen"

$y = 20
function Add-Field {
    param([string]$LabelText,[bool]$IsPassword = $false,[ref]$OutTextbox)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Location = New-Object System.Drawing.Point(10, $script:y)
    $label.Size = New-Object System.Drawing.Size(140, 22)
    $form.Controls.Add($label)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(160, $script:y)
    $tb.Size = New-Object System.Drawing.Size(360, 22)
    if ($IsPassword) { $tb.UseSystemPasswordChar = $true }
    $form.Controls.Add($tb)

    $OutTextbox.Value = $tb
    $script:y += 34
}

$serverTB = $null; Add-Field -LabelText "BigFix Server:" -OutTextbox ([ref]$serverTB)
$userTB   = $null; Add-Field -LabelText "Username:"     -OutTextbox ([ref]$userTB)
$passTB   = $null; Add-Field -LabelText "Password:"     -IsPassword $true -OutTextbox ([ref]$passTB)
$fixTB    = $null; Add-Field -LabelText "Fixlet ID:"    -OutTextbox ([ref]$fixTB)

# Date (future Wednesdays only)
$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Text = "Schedule Date (Wed):"
$labelDate.Location = New-Object System.Drawing.Point(10, $y)
$labelDate.Size = New-Object System.Drawing.Size(140, 22)
$form.Controls.Add($labelDate)

$dateCB = New-Object System.Windows.Forms.ComboBox
$dateCB.Location = New-Object System.Drawing.Point(160, $y)
$dateCB.Size = New-Object System.Drawing.Size(160, 22)
$dateCB.DropDownStyle = 'DropDownList'
$dateCB.Items.AddRange((Get-NextWednesdays))
$form.Controls.Add($dateCB)
$y += 34

# Time (8:00 PM → 11:45 PM)
$labelTime = New-Object System.Windows.Forms.Label
$labelTime.Text = "Schedule Time:"
$labelTime.Location = New-Object System.Drawing.Point(10, $y)
$labelTime.Size = New-Object System.Drawing.Size(140, 22)
$form.Controls.Add($labelTime)

$timeCB = New-Object System.Windows.Forms.ComboBox
$timeCB.Location = New-Object System.Drawing.Point(160, $y)
$timeCB.Size = New-Object System.Drawing.Size(160, 22)
$timeCB.DropDownStyle = 'DropDownList'
$timeCB.Items.AddRange((Get-TimeSlots))
$form.Controls.Add($timeCB)
$y += 42

# Generate button
$goBtn = New-Object System.Windows.Forms.Button
$goBtn.Text = "Generate & Post Actions"
$goBtn.Location = New-Object System.Drawing.Point(160, $y)
$goBtn.Size = New-Object System.Drawing.Size(220, 32)
$form.Controls.Add($goBtn)
$y += 42

# Log box
$log = New-Object System.Windows.Forms.TextBox
$log.Multiline = $true
$log.ScrollBars = "Vertical"
$log.ReadOnly = $false
$log.WordWrap = $false
$log.Location = New-Object System.Drawing.Point(10, $y)
$log.Size = New-Object System.Drawing.Size(510, 520)
$log.Anchor = "Top,Left,Right,Bottom"
$form.Controls.Add($log)

# Right-click menu for log
$cmenu = New-Object System.Windows.Forms.ContextMenu
$miCopy = New-Object System.Windows.Forms.MenuItem "Copy"
$miAll  = New-Object System.Windows.Forms.MenuItem "Select All"
$cmenu.MenuItems.AddRange(@($miCopy, $miAll))
$log.ContextMenu = $cmenu
$miCopy.add_Click({ $log.Copy() })
$miAll.add_Click({ $log.SelectAll() })

# =========================
# ACTION GENERATION
# =========================
$goBtn.Add_Click({
    $log.Clear()
    $logFile = Join-Path $env:TEMP "BigFixActionGenerator.log"
    $append = {
        param($text)
        $log.AppendText($text + "`r`n")
        Add-Content -Path $logFile -Value $text
        $log.SelectionStart = $log.Text.Length
        $log.ScrollToCaret()
    }

    $server   = $serverTB.Text.Trim()
    $user     = $userTB.Text.Trim()
    $pass     = $passTB.Text
    $fixletId = $fixTB.Text.Trim()
    $dateStr  = $dateCB.SelectedItem
    $timeStr  = $timeCB.SelectedItem

    if (-not ($server -and $user -and $pass -and $fixletId -and $dateStr -and $timeStr)) {
        [System.Windows.Forms.MessageBox]::Show("All fields must be completed (including date and time).")
        return
    }

    try {
        # Log the fully encoded Fixlet GET URL before calling it
        $base        = Get-BaseUrl $server
        $encodedSite = Encode-SiteName $SiteName
        $fixletPath  = "/api/fixlet/custom/$encodedSite/$fixletId"
        $fixletUrl   = Join-ApiUrl -BaseUrl $base -RelativePath $fixletPath

        $append.Invoke(("Server base URL: {0}" -f $base))
        $append.Invoke(("Encoded Fixlet GET URL: {0}" -f $fixletUrl))

        # Call GET
        $resp = $null
        try {
            $auth = Get-AuthHeader -Username $user -Password $pass
            $content = HttpGetXml -Url $fixletUrl -AuthHeader $auth
            $resp = [pscustomobject]@{ Url = $fixletUrl; Content = $content }
        } catch {
            $append.Invoke(("❌ GET failed: {0}" -f $_))
            throw
        }
        $append.Invoke(("GET URL (used): {0}" -f $resp.Url))

        # Parse XML robustly (Fixlet/Task/Baseline, Action/DefaultAction)
        $fixletXml = $resp.Content
        $xml = [xml]$fixletXml

        $contInfo = Get-FixletContainer -Xml $xml
        $container = $contInfo.Node
        $append.Invoke(("Detected BES content type: {0}" -f $contInfo.Type))

        $titleRaw = $container.Title
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw

        $parsed = Get-ActionAndRelevance -ContainerNode $container
        $relevance    = $parsed.Relevance
        $actionScript = $parsed.ActionScript

        $append.Invoke(("Parsed title: {0}" -f $displayName))
        $append.Invoke(("Relevance count: {0}" -f $relevance.Count))
        $append.Invoke(("Action script length: {0} chars" -f $actionScript.Length))

        # Build start (local)
        $startLocal = Get-Date "$dateStr $timeStr"
        $append.Invoke(("Start (local): {0}" -f $startLocal))

        # Force deadline = +24h absolute
        $deadlineLocal = $startLocal.AddHours(24)
        $append.Invoke(("Force deadline (local): {0}" -f $deadlineLocal))

        # Define the four actions
        $actions = @("Pilot","Deploy","Force","Conference/Training Rooms")

        foreach ($a in $actions) {
            $groupId = "$($GroupMap[$a])"
            if (-not $groupId) { throw "No group mapping found for action: $a" }
            $isForce = ($a -eq "Force")

            $xmlBody = Build-SingleActionXml `
                -ActionTitle $a -DisplayName $displayName `
                -RelevanceBlocks $relevance -ActionScript $actionScript `
                -StartLocal $startLocal -SetDeadline:$isForce `
                -DeadlineLocal $deadlineLocal -GroupId $groupId

            $append.Invoke(("---- XML for {0} ----" -f $a))
            $append.Invoke($xmlBody)

            try {
                $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
                $append.Invoke(("Encoded POST URL: {0}" -f $postUrl))
                try {
                    $auth = Get-AuthHeader -Username $user -Password $pass
                    HttpPostXml -Url $postUrl -AuthHeader $auth -XmlBody $xmlBody
                    $append.Invoke(("✅ {0} created successfully." -f $a))
                } catch {
                    $append.Invoke(("❌ POST failed: {0}" -f $_))
                }
            } catch {
                $append.Invoke(("❌ Failed to create {0}: {1}" -f $a, $_))
            }
        }

        $append.Invoke(("All actions attempted. See log: {0}" -f $logFile))
    }
    catch {
        $append.Invoke(("❌ Fatal error: {0}" -f $_))
    }
})

$form.Topmost = $false
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
