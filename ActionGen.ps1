Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG (EDIT THESE)
# =========================
$LogFile = Join-Path $env:TEMP "BigFixActionGenerator.log"

# Site that hosts BOTH the Fixlet content and the Computer Groups
$CustomSiteName = "Test Group Managed (Workstations)"

# Action -> Computer Group ID (you can keep the 00- prefix; we'll strip it)
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12346"
    "Force"                     = "00-12347"
    "Conference/Training Rooms" = "00-12348"
}

# =========================
# UTIL / LOGGING
# =========================
function Encode-SiteName([string]$Name) {
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
function Get-AuthHeader([string]$User,[string]$Pass) {
    $pair  = "$User`:$Pass"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    "Basic " + [Convert]::ToBase64String($bytes)
}
function LogLine($txt) {
    $line = "{0}  {1}" -f (Get-Date -Format 'u'), $txt
    $LogBox.AppendText($line + "`r`n")
    Add-Content -Path $LogFile -Value $line
    $LogBox.SelectionStart = $LogBox.Text.Length
    $LogBox.ScrollToCaret()
}
function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { return $Matches[1] }
    return ($GroupIdWithPrefix -replace '[^\d]','') # fallback
}
function Format-BESLocal([datetime]$dt) {
    (Get-Date $dt).ToString("yyyyMMdd'T'HHmmss",
        [System.Globalization.CultureInfo]::InvariantCulture)
}

# =========================
# HTTP (curl-like)
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
            $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
            $content = $sr.ReadToEnd(); $sr.Close()
        } finally { $resp.Close() }
        return $content
    } catch {
        throw ($_.Exception.GetBaseException().Message)
    }
}
function HttpPostXml {
    param([string]$Url,[string]$AuthHeader,[string]$XmlBody)
    $bytes = [Text.Encoding]::UTF8.GetBytes($XmlBody)
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
        $rs = $req.GetRequestStream(); $rs.Write($bytes,0,$bytes.Length); $rs.Close()
        $resp = $req.GetResponse(); $resp.Close()
    } catch {
        throw ($_.Exception.GetBaseException().Message)
    }
}

# =========================
# FIXLET PARSING
# =========================
function Get-FixletContainer { param([xml]$Xml)
    if ($Xml.BES.Fixlet)   { return @{ Type="Fixlet";   Node=$Xml.BES.Fixlet } }
    if ($Xml.BES.Task)     { return @{ Type="Task";     Node=$Xml.BES.Task } }
    if ($Xml.BES.Baseline) { return @{ Type="Baseline"; Node=$Xml.BES.Baseline } }
    throw "Unknown BES content type (no <Fixlet>, <Task>, or <Baseline>)."
}
function Get-ActionAndRelevance {
    param($ContainerNode)
    $rels = @(); foreach ($r in $ContainerNode.Relevance) { $rels += [string]$r }
    $act = $null
    if ($ContainerNode.Action) { $act = $ContainerNode.Action | Select-Object -First 1 }
    if (-not $act -and $ContainerNode.DefaultAction) { $act = $ContainerNode.DefaultAction }
    if (-not $act) { throw "No <Action> or <DefaultAction> block found." }
    $script = $null
    if ($act.ActionScript) {
        $script = [string]$act.ActionScript.'#text'
        if ([string]::IsNullOrWhiteSpace($script)) { $script = $act.ActionScript.InnerText }
    }
    if (-not $script) { throw "Action found but no <ActionScript> content present." }
    return @{ Relevance=$rels; ActionScript=$script }
}
function Parse-FixletTitleToProduct([string]$Title) {
    $clean = $Title -replace '^Update:\s*','' -replace '\s+Win$',''
    return $clean.Trim()
}

# =========================
# SINGLE ACTION XML (Local timestamps, CDATA relevance, RunningMessage <Text>)
# =========================
function Build-SingleActionXml {
    param(
        [string]$ActionTitle,          # Pilot/Deploy/Force/Conference...
        [string]$DisplayName,          # Vendor App Version
        [string[]]$RelevanceBlocks,    # relevance strings
        [string]$ActionScript,         # action script
        [datetime]$StartLocal,         # scheduled local start
        [bool]$SetDeadline = $false,   # true only for Force
        [datetime]$DeadlineLocal = $null,
        [string]$GroupSiteName,        # same site as fixlet
        [string]$GroupIdNumeric        # numeric ID (no 00-)
    )

    $titleText = "$($DisplayName): $ActionTitle"
    $titleEsc  = [System.Security.SecurityElement]::Escape($titleText)
    $dispEsc   = [System.Security.SecurityElement]::Escape($DisplayName)

    # Relevance as CDATA (avoid &quot; etc.). Guard against ']]>'.
    $rels = ($RelevanceBlocks | ForEach-Object {
        $safe = $_ -replace ']]>', ']]]]><![CDATA[>'
        "    <Relevance><![CDATA[$safe]]></Relevance>"
    }) -join "`r`n"

    $startStr    = Format-BESLocal $StartLocal
    $deadlineStr = if ($SetDeadline -and $DeadlineLocal) { Format-BESLocal $DeadlineLocal } else { $null }
    $deadlineBlock = if ($deadlineStr) { "<Deadline>$deadlineStr</Deadline>" } else { "" }

@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SingleAction>
    <Title>$titleEsc</Title>
$rels
    <ActionScript MIMEType="application/x-Fixlet-Windows-Shell"><![CDATA[
$ActionScript
]]></ActionScript>
    <Settings>
      <ActionUITitle>$titleEsc</ActionUITitle>

      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage>
        <Text>Updating to $dispEsc. Please wait....</Text>
      </RunningMessage>

      <HasTimeRange>true</HasTimeRange>
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocal>$startStr</StartDateTimeLocal>
      <HasEndTime>false</HasEndTime>

      <HasDeadline>$([string]$SetDeadline)</HasDeadline>
      $deadlineBlock

      <HasReapply>false</HasReapply>
      <HasRetry>false</HasRetry>
      <HasTemporalDistribution>false</HasTemporalDistribution>
      <HasAllowNoFixlets>false</HasAllowNoFixlets>
    </Settings>
    <Target>
      <ComputerGroup>
        <SiteName>$([System.Security.SecurityElement]::Escape($GroupSiteName))</SiteName>
        <ID>$GroupIdNumeric</ID>
      </ComputerGroup>
    </Target>
  </SingleAction>
</BES>
"@
}

# =========================
# GUI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(620, 760)
$form.StartPosition = "CenterScreen"

$y = 20
function Add-Field([string]$Label,[bool]$IsPassword,[ref]$OutTB) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = New-Object System.Drawing.Point(10,$script:y)
    $lbl.Size = New-Object System.Drawing.Size(140,22)
    $form.Controls.Add($lbl)

    if ($IsPassword) {
        $tb = New-Object System.Windows.Forms.MaskedTextBox
        $tb.PasswordChar = '*'
    } else {
        $tb = New-Object System.Windows.Forms.TextBox
    }
    $tb.Location = New-Object System.Drawing.Point(160,$script:y)
    $tb.Size = New-Object System.Drawing.Size(420,22)
    $form.Controls.Add($tb)
    $OutTB.Value = $tb
    $script:y += 34
}

$tbServer = $null; Add-Field "BigFix Server:" $false ([ref]$tbServer)
$tbUser   = $null; Add-Field "Username:"      $false ([ref]$tbUser)
$tbPass   = $null; Add-Field "Password:"      $true  ([ref]$tbPass)
$tbFixlet = $null; Add-Field "Fixlet ID:"     $false ([ref]$tbFixlet)

# Date (future Wednesdays)
$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Schedule Date (Wed):"
$lblDate.Location = New-Object System.Drawing.Point(10,$y)
$lblDate.Size = New-Object System.Drawing.Size(140,22)
$form.Controls.Add($lblDate)

$cbDate = New-Object System.Windows.Forms.ComboBox
$cbDate.DropDownStyle = 'DropDownList'
$cbDate.Location = New-Object System.Drawing.Point(160,$y)
$cbDate.Size = New-Object System.Drawing.Size(160,22)
$form.Controls.Add($cbDate)
$y += 34
# populate next 20 Wednesdays
$today = Get-Date
$daysUntilWed = (3 - [int]$today.DayOfWeek + 7) % 7
$nextWed = $today.AddDays($daysUntilWed)
for ($i=0;$i -lt 20;$i++) { [void]$cbDate.Items.Add($nextWed.AddDays(7*$i).ToString("yyyy-MM-dd")) }

# Time (8:00 PM – 11:45 PM, 15m)
$lblTime = New-Object System.Windows.Forms.Label
$lblTime.Text = "Schedule Time:"
$lblTime.Location = New-Object System.Drawing.Point(10,$y)
$lblTime.Size = New-Object System.Drawing.Size(140,22)
$form.Controls.Add($lblTime)

$cbTime = New-Object System.Windows.Forms.ComboBox
$cbTime.DropDownStyle = 'DropDownList'
$cbTime.Location = New-Object System.Drawing.Point(160,$y)
$cbTime.Size = New-Object System.Drawing.Size(160,22)
$form.Controls.Add($cbTime)
$y += 42
$start = Get-Date "20:00"; $end = Get-Date "23:45"
while ($start -le $end) { [void]$cbTime.Items.Add($start.ToString("h:mm tt")); $start = $start.AddMinutes(15) }

# Button
$btn = New-Object System.Windows.Forms.Button
$btn.Text = "Generate & Post 4 Single Actions"
$btn.Location = New-Object System.Drawing.Point(160,$y)
$btn.Size = New-Object System.Drawing.Size(280,32)
$form.Controls.Add($btn)
$y += 42

# Log
$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = "Vertical"
$LogBox.ReadOnly = $false
$LogBox.WordWrap = $false
$LogBox.Location = New-Object System.Drawing.Point(10,$y)
$LogBox.Size = New-Object System.Drawing.Size(570,520)
$LogBox.Anchor = "Top,Left,Right,Bottom"
$form.Controls.Add($LogBox)
# Context menu for easy copy/select all
$cm = New-Object System.Windows.Forms.ContextMenu
$miCopy = New-Object System.Windows.Forms.MenuItem "Copy"
$miAll  = New-Object System.Windows.Forms.MenuItem "Select All"
$cm.MenuItems.AddRange(@($miCopy,$miAll))
$LogBox.ContextMenu = $cm
$miCopy.add_Click({ $LogBox.Copy() })
$miAll.add_Click({ $LogBox.SelectAll() })

# =========================
# ACTION
# =========================
$btn.Add_Click({
    $LogBox.Clear()
    $server = $tbServer.Text.Trim()
    $user   = $tbUser.Text.Trim()
    $pass   = $tbPass.Text
    $fixId  = $tbFixlet.Text.Trim()
    $dStr   = $cbDate.SelectedItem
    $tStr   = $cbTime.SelectedItem

    if (-not ($server -and $user -and $pass -and $fixId -and $dStr -and $tStr)) {
        LogLine "❌ Please fill in Server, Username, Password, Fixlet ID, Date, and Time."
        return
    }

    try {
        $base = Get-BaseUrl $server
        $encodedSite = Encode-SiteName $CustomSiteName
        $fixletUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/fixlet/custom/$encodedSite/$fixId"
        LogLine "Encoded Fixlet GET URL: $fixletUrl"

        $auth = Get-AuthHeader -User $user -Pass $pass
        $fixletContent = HttpGetXml -Url $fixletUrl -AuthHeader $auth
        $xml = [xml]$fixletContent

        $cont = Get-FixletContainer -Xml $xml
        LogLine ("Detected BES content type: {0}" -f $cont.Type)

        $titleRaw = $cont.Node.Title
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw

        $parsed = Get-ActionAndRelevance -ContainerNode $cont.Node
        $relevance = $parsed.Relevance
        $actionScript = $parsed.ActionScript

        LogLine "Parsed title: $displayName"
        LogLine ("Relevance count: {0}" -f $relevance.Count)
        LogLine ("Action script length: {0}" -f $actionScript.Length)

        $startLocal    = Get-Date "$dStr $tStr"
        $deadlineLocal = $startLocal.AddHours(24)  # for Force

        $actions = @("Pilot","Deploy","Force","Conference/Training Rooms")
        $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        LogLine "POST URL: $postUrl"

        foreach ($a in $actions) {
            $groupIdRaw = "$($GroupMap[$a])"
            if (-not $groupIdRaw) { LogLine "❌ Missing group map for $a"; continue }
            $groupIdNumeric = Get-NumericGroupId $groupIdRaw
            if (-not $groupIdNumeric) { LogLine "❌ Could not parse numeric ID from '$groupIdRaw' for $a"; continue }

            $isForce = ($a -eq "Force")

            $xmlBody = Build-SingleActionXml `
                -ActionTitle $a `
                -DisplayName $displayName `
                -RelevanceBlocks $relevance `
                -ActionScript $actionScript `
                -StartLocal $startLocal `
                -SetDeadline:$isForce `
                -DeadlineLocal $deadlineLocal `
                -GroupSiteName $CustomSiteName `
                -GroupIdNumeric $groupIdNumeric

            LogLine ("---- SingleAction XML for {0} ----" -f $a)
            LogLine $xmlBody

            try {
                HttpPostXml -Url $postUrl -AuthHeader $auth -XmlBody $xmlBody
                LogLine ("✅ {0} posted successfully." -f $a)
            } catch {
                LogLine ("❌ POST failed for {0}: {1}" -f $a, $_)
            }
        }

        LogLine "All actions attempted. Log file: $LogFile"
    }
    catch {
        LogLine ("❌ Fatal error: {0}" -f ($_.Exception.GetBaseException().Message))
    }
})

$form.Topmost = $false
[void]$form.ShowDialog()
