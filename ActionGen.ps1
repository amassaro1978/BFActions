Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG (EDIT THESE)
# =========================
$LogFile = Join-Path $env:TEMP "BigFixActionGenerator.log"

# Site that hosts the Fixlet and (ideally) the Computer Groups
$CustomSiteName = "Test Group Managed (Workstations)"

# Action -> Computer Group ID (keep 00- prefix; we'll strip to numeric for API)
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12346"
    "Force"                     = "00-12347"
    "Conference/Training Rooms" = "00-12348"
}

# Which Fixlet action to invoke in the Fixlet (must exist)
$FixletActionNameMap = @{
    "Pilot"                     = "Action1"
    "Deploy"                    = "Action1"
    "Force"                     = "Action1"
    "Conference/Training Rooms" = "Action1"
}

# Behavior toggles
$IgnoreCertErrors           = $true
$DumpFetchedXmlToTemp       = $true
$SaveActionXmlToTemp        = $true

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
    if ($s -notmatch '^(?i)https?://') {
        if ($s -match ':\d+$') { $s = "https://$s" } else { $s = "https://$s:52311" }
    }
    return $s.TrimEnd('/')
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
function Write-Utf8NoBom([string]$Path,[string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}
function Get-NextWeekday([datetime]$base,[System.DayOfWeek]$weekday) {
    $anchor = $base.Date
    $delta = ([int]$weekday - [int]$anchor.DayOfWeek + 7) % 7
    if ($delta -le 0) { $delta += 7 }
    return $anchor.AddDays($delta)
}
function IsoTimePart([TimeSpan]$ts) {
    if ($ts.Minutes -gt 0) { return "PT{0}H{1}M" -f $ts.Hours, $ts.Minutes }
    else { return "PT{0}H" -f $ts.Hours }
}

# =========================
# HTTP
# =========================
if ($IgnoreCertErrors) { try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch { } }
[System.Net.ServicePointManager]::Expect100Continue = $false

function HttpGetXml {
    param([string]$Url,[string]$AuthHeader)
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = "GET"
    $req.Accept = "application/xml"
    $req.Headers["Authorization"] = $AuthHeader
    $req.ProtocolVersion = [Version]"1.1"
    $req.PreAuthenticate = $true
    $req.AllowAutoRedirect = $false
    $req.Timeout = 45000
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

function Post-XmlFile {
    param([string]$Url,[string]$User,[string]$Pass,[string]$XmlFilePath)
    try {
        $pair  = "$User`:$Pass"
        $basic = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
        $resp = Invoke-WebRequest -Method Post -Uri $Url `
            -Headers @{ "Authorization" = $basic } `
            -ContentType "application/xml" `
            -InFile $XmlFilePath -UseBasicParsing
        if ($resp.Content) { LogLine "POST response: $($resp.Content)" }
    } catch {
        if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream) {
            $sr = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream(), [Text.Encoding]::UTF8)
            $errBody = $sr.ReadToEnd(); $sr.Close()
            throw "Invoke-WebRequest POST failed :: $errBody"
        } else {
            throw ($_.Exception.Message)
        }
    }
}

# =========================
# FIXLET & GROUP PARSING
# =========================
function Get-FixletContainer { param([xml]$Xml)
    if ($Xml.BES.Fixlet)   { return @{ Type="Fixlet";   Node=$Xml.BES.Fixlet } }
    if ($Xml.BES.Task)     { return @{ Type="Task";     Node=$Xml.BES.Task } }
    if ($Xml.BES.Baseline) { return @{ Type="Baseline"; Node=$Xml.BES.Baseline } }
    throw "Unknown BES content type (no <Fixlet>, <Task>, or <Baseline>)."
}
function Get-ActionAndRelevance { param($ContainerNode)
    $rels = @()
    foreach ($r in $ContainerNode.Relevance) {
        $t = ($r.InnerText).Trim()
        if ($t) { $rels += $t }
    }
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
    ($Title -replace '^Update:\s*','' -replace '\s+Win$','').Trim()
}

# Build client relevance for a group:
# 1) Prefer <Relevance>
# 2) Else AND-join <SearchComponentRelevance> blocks
function Get-GroupClientRelevance {
    param(
        [string]$BaseUrl,
        [string]$AuthHeader,
        [string]$SiteName,
        [string]$GroupIdNumeric
    )

    $encSite = Encode-SiteName $SiteName
    $candidates = @(
        "/api/computergroup/custom/$encSite/$GroupIdNumeric",               # custom site
        "/api/computergroup/master/$GroupIdNumeric",                        # master site
        "/api/computergroup/operator/$($env:USERNAME)/$GroupIdNumeric"      # operator site (best guess)
    )

    foreach ($relPath in $candidates) {
        $url = Join-ApiUrl -BaseUrl $BaseUrl -RelativePath $relPath
        try {
            $xmlStr = HttpGetXml -Url $url -AuthHeader $AuthHeader
            if ($DumpFetchedXmlToTemp) {
                $tmp = Join-Path $env:TEMP ("BES_ComputerGroup_{0}.xml" -f $GroupIdNumeric)
                Write-Utf8NoBom -Path $tmp -Content $xmlStr
                LogLine "Saved fetched group XML to: $tmp"
            }
            $x = [xml]$xmlStr

            $top = $x.BES.ComputerGroup.Relevance
            if ($top -and (-not [string]::IsNullOrWhiteSpace($top))) {
                $snippet = $top.Substring(0, [Math]::Min(200, $top.Length))
                LogLine "Found group <Relevance> at ${url} :: ${snippet}..."
                return [string]$top
            }

            $fragments = @()
            $nodes = $x.SelectNodes("//ComputerGroup/SearchComponentRelevance")
            if ($nodes -and $nodes.Count -gt 0) {
                foreach ($n in $nodes) {
                    $txt = $n.InnerText
                    if ($txt -and -not [string]::IsNullOrWhiteSpace($txt)) {
                        $fragments += $txt.Trim()
                    }
                }
            }
            if ($fragments.Count -gt 0) {
                $joined = ($fragments | ForEach-Object { "($_)" }) -join " AND "
                $snippet = $joined.Substring(0, [Math]::Min(200, $joined.Length))
                LogLine "Built relevance from SearchComponentRelevance at ${url} :: ${snippet}..."
                return $joined
            }

            LogLine "No usable relevance at ${url}"
        } catch {
            LogLine "Fetch failed at ${url}: $($_.Exception.Message)"
        }
    }

    throw "No relevance found or derivable for group ${GroupIdNumeric} in custom/master/operator."
}

# =========================
# ACTION XML (SourcedFixletAction; schema order; absolute times; nullable params)
# =========================
function Build-SourcedFixletActionXml {
    param(
        [string]$ActionTitle,              # Pilot/Deploy/Force/Conference...
        [string]$UiBaseTitle,              # Full Fixlet title ("Update: ... Win")
        [string]$DisplayName,              # For user messages ("The GIMP Team GIMP 3.0.4")
        [string]$SiteName,                 # Custom site name
        [string]$FixletId,                 # Fixlet ID
        [string]$FixletActionName,         # "Action1" or a named action in the Fixlet
        [string]$GroupRelevance,           # Group filter
        [datetime]$StartLocal,             # absolute start (local)
        [Nullable[datetime]]$EndLocal = $null,          # optional absolute end (local)
        [Nullable[datetime]]$DeadlineLocal = $null,     # optional absolute deadline (local) (Force)
        [bool]$HasTimeRange = $false,
        [Nullable[TimeSpan]]$TimeRangeStart = $null,    # time-of-day from midnight
        [Nullable[TimeSpan]]$TimeRangeEnd   = $null,    # time-of-day from midnight
        [bool]$ShowPreActionUI = $false,
        [string]$PreActionText = "",
        [bool]$AskToSaveWork = $false
    )

    # Console action name (": Pilot", etc.)
    $fullTitle = ("{0}: {1}" -f $UiBaseTitle, $ActionTitle)
    $uiTitle   = [System.Security.SecurityElement]::Escape($fullTitle)
    $dispEsc   = [System.Security.SecurityElement]::Escape($DisplayName)

    # Snap times to exact :00 seconds
    if ($StartLocal)             { $StartLocal    = $StartLocal.Date.AddHours($StartLocal.Hour).AddMinutes($StartLocal.Minute) }
    if ($EndLocal.HasValue)      { $EndLocal      = $EndLocal.Value.Date.AddHours($EndLocal.Value.Hour).AddMinutes($EndLocal.Value.Minute) }
    if ($DeadlineLocal.HasValue) { $DeadlineLocal = $DeadlineLocal.Value.Date.AddHours($DeadlineLocal.Value.Hour).AddMinutes($DeadlineLocal.Value.Minute) }

    # Group relevance (safe CDATA)
    $groupSafe = if ([string]::IsNullOrWhiteSpace($GroupRelevance)) { "" } else { $GroupRelevance }
    $groupSafe = $groupSafe -replace ']]>', ']]]]><![CDATA[>'

    # End time line
    $hasEnd = $EndLocal.HasValue
    $endLine = if ($hasEnd) { "      <EndDateTimeLocal>$($EndLocal.Value.ToString('yyyy-MM-ddTHH:mm:ss'))</EndDateTimeLocal>`n" } else { "" }

    # TimeRange block
    if ($HasTimeRange) {
        if (-not $TimeRangeStart.HasValue) { $TimeRangeStart = [TimeSpan]::FromHours(19) } # 7:00 PM
        if (-not $TimeRangeEnd.HasValue)   { $TimeRangeEnd   = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59)) } # 6:59 AM
        $trs = IsoTimePart $TimeRangeStart.Value
        $tre = IsoTimePart $TimeRangeEnd.Value
        $timeRangeBlock = @"
      <HasTimeRange>true</HasTimeRange>
      <TimeRange>
        <StartTime>$trs</StartTime>
        <EndTime>$tre</EndTime>
      </TimeRange>
"@
    } else {
        $timeRangeBlock = "      <HasTimeRange>false</HasTimeRange>"
    }

    # PreAction (if enabled); absolute deadline belongs inside PreAction for SourcedFixletAction
    $preActionBlock = ""
    if ($ShowPreActionUI) {
        $preEsc = [System.Security.SecurityElement]::Escape($PreActionText)
        $deadlineInner = ""
        if ($DeadlineLocal.HasValue) {
$deadlineInner = @"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineLocalTime>$($DeadlineLocal.Value.ToString('yyyy-MM-ddTHH:mm:ss'))</DeadlineLocalTime>
"@
        }
$preActionBlock = @"
      <PreAction>
        <Text>$preEsc</Text>
        <AskToSaveWork>$($AskToSaveWork.ToString().ToLower())</AskToSaveWork>
        <ShowActionButton>false</ShowActionButton>
        <ShowCancelButton>false</ShowCancelButton>
$deadlineInner        <ShowConfirmation>false</ShowConfirmation>
      </PreAction>
"@
    }

@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SourcedFixletAction>
    <SourceFixlet>
      <Sitename>$SiteName</Sitename>
      <FixletID>$FixletId</FixletID>
      <Action>$FixletActionName</Action>
    </SourceFixlet>
    <Target>
      <CustomRelevance><![CDATA[$groupSafe]]></CustomRelevance>
    </Target>
    <Settings>
      <ActionUITitle>$uiTitle</ActionUITitle>
$preActionBlock
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>Updating to $dispEsc... Please wait.</Text></RunningMessage>
$timeRangeBlock
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocal>$($StartLocal.ToString('yyyy-MM-ddTHH:mm:ss'))</StartDateTimeLocal>
      <HasEndTime>$($hasEnd.ToString().ToLower())</HasEndTime>
$endLine      <UseUTCTime>false</UseUTCTime>
      <ActiveUserRequirement>NoRequirement</ActiveUserRequirement>
      <ActiveUserType>AllUsers</ActiveUserType>
      <HasWhose>false</HasWhose>
      <PreActionCacheDownload>false</PreActionCacheDownload>
      <Reapply>true</Reapply>
      <HasReapplyLimit>false</HasReapplyLimit>
      <HasReapplyInterval>false</HasReapplyInterval>
      <HasRetry>true</HasRetry>
      <RetryCount>3</RetryCount>
      <RetryWait Behavior="WaitForInterval">PT1H</RetryWait>
      <HasTemporalDistribution>false</HasTemporalDistribution>
      <ContinueOnErrors>true</ContinueOnErrors>
      <PostActionBehavior Behavior="Nothing"></PostActionBehavior>
      <IsOffer>false</IsOffer>
    </Settings>
    <Title>$uiTitle</Title>
  </SourcedFixletAction>
</BES>
"@
}

# =========================
# GUI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(640, 780)
$form.StartPosition = "CenterScreen"

$y = 20
function Add-Field([string]$Label,[bool]$IsPassword,[ref]$OutTB) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = New-Object System.Drawing.Point(10,$script:y)
    $lbl.Size = New-Object System.Drawing.Size(140,22)
    $form.Controls.Add($lbl)

    if ($IsPassword) { $tb = New-Object System.Windows.Forms.MaskedTextBox; $tb.PasswordChar = '*' }
    else { $tb = New-Object System.Windows.Forms.TextBox }
    $tb.Location = New-Object System.Drawing.Point(160,$script:y)
    $tb.Size = New-Object System.Drawing.Size(440,22)
    $form.Controls.Add($tb)
    $OutTB.Value = $tb
    $script:y += 34
}

$tbServer = $null; Add-Field "BigFix Server:" $false ([ref]$tbServer)
$tbUser   = $null; Add-Field "Username:"      $false ([ref]$tbUser)
$tbPass   = $null; Add-Field "Password:"      $true  ([ref]$tbPass)
$tbFixlet = $null; Add-Field "Fixlet ID:"     $false ([ref]$tbFixlet)

# Date (future Wednesdays)
$lblDate = New-Object Windows.Forms.Label
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
$btn.Text = "Generate & Post 4 Actions (Pilot/Deploy/Force/Conf)"
$btn.Location = New-Object System.Drawing.Point(160,$y)
$btn.Size = New-Object System.Drawing.Size(320,32)
$form.Controls.Add($btn)
$y += 42

# Log
$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = "Vertical"
$LogBox.ReadOnly = $false
$LogBox.WordWrap = $false
$LogBox.Location = New-Object System.Drawing.Point(10,$y)
$LogBox.Size = New-Object System.Drawing.Size(600,520)
$LogBox.ContextMenu = New-Object System.Windows.Forms.ContextMenu
$LogBox.ContextMenu.MenuItems.AddRange(@(
    (New-Object System.Windows.Forms.MenuItem "Copy",      { $LogBox.Copy() }),
    (New-Object System.Windows.Forms.MenuItem "Select All", { $LogBox.SelectAll() })
))
$LogBox.Anchor = "Top,Left,Right,Bottom"
$form.Controls.Add($LogBox)

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
        LogLine "Encoded Fixlet GET URL: ${fixletUrl}"

        $auth = Get-AuthHeader -User $user -Pass $pass
        $fixletContent = HttpGetXml -Url $fixletUrl -AuthHeader $auth
        if ($DumpFetchedXmlToTemp) {
            $tmpFix = Join-Path $env:TEMP ("BES_Fixlet_{0}.xml" -f $fixId)
            Write-Utf8NoBom -Path $tmpFix -Content $fixletContent
            LogLine "Saved fetched fixlet XML to: $tmpFix"
        }

        $fixletXml = [xml]$fixletContent
        $cont = Get-FixletContainer -Xml $fixletXml
        $titleRaw     = [string]$cont.Node.Title
        $displayName  = Parse-FixletTitleToProduct -Title $titleRaw   # e.g., "The GIMP Team GIMP 3.0.4"

        # Sanity
        [void](Get-ActionAndRelevance -ContainerNode $cont.Node)

        # Build concrete schedule from UI (seconds snapped to :00)
        $pilotStart = [datetime]::ParseExact("$dStr $tStr","yyyy-MM-dd h:mm tt",$null)
        $pilotStart = $pilotStart.Date.AddHours($pilotStart.Hour).AddMinutes($pilotStart.Minute)

        $deployStart     = $pilotStart.AddDays(1)                                              # +24h
        $confStart       = $pilotStart.AddDays(1)
        $pilotEnd        = $pilotStart.Date.AddDays(1).AddHours(6).AddMinutes(59)             # next day 6:59 AM
        $deployEnd       = $deployStart.Date.AddDays(1).AddHours(6).AddMinutes(55)            # next morning 6:55 AM

        # Force: next Tuesday 7:00 AM after Pilot, deadline = Wednesday 7:00 AM
        $forceStartDate  = Get-NextWeekday -base $pilotStart -weekday ([DayOfWeek]::Tuesday)
        $forceStart      = $forceStartDate.AddHours(7)                                         # Tue 7:00 AM
        $forceEnforce    = $forceStart.AddDays(1)                                              # Wed 7:00 AM

        # Run-between window 7:00 PM – 6:59 AM
        $trStart = [TimeSpan]::FromHours(19)                                                   # 19:00
        $trEnd   = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59))                   # 06:59

        $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        LogLine "POST URL: ${postUrl}"

        $actions = @(
            @{ Name="Pilot"; Start=$pilotStart; End=$pilotEnd; TR=$true;  TRS=$trStart; TRE=$trEnd; UI=$false; Msg="";    Save=$false; Deadline=$null },
            @{ Name="Deploy";Start=$deployStart;End=$deployEnd;TR=$true;  TRS=$trStart; TRE=$trEnd; UI=$false; Msg="";    Save=$false; Deadline=$null },
            @{ Name="Conference/Training Rooms"; Start=$confStart; End=$null; TR=$true; TRS=$trStart; TRE=$trEnd; UI=$false; Msg=""; Save=$false; Deadline=$null },
            @{ Name="Force"; Start=$forceStart; End=$null; TR=$false; TRS=$null; TRE=$null; UI=$true;
               Msg=("{0} update will be enforced on {1}.  Please leave your machine on overnight to get the automated update.  Otherwise, please close the application and run the update now" -f `
                    $displayName, $forceEnforce.ToString("M/d/yyyy h:mm tt"));
               Save=$true; Deadline=$forceEnforce }
        )

        foreach ($cfg in $actions) {
            $a = $cfg.Name
            $groupIdRaw = "$($GroupMap[$a])"
            if (-not $groupIdRaw) { LogLine "❌ Missing group id for $a"; continue }
            $groupIdNumeric = Get-NumericGroupId $groupIdRaw
            if (-not $groupIdNumeric) { LogLine "❌ Could not parse numeric ID from '${groupIdRaw}' for $a"; continue }

            # fetch group's client relevance and combine with Fixlet via Target
            $groupRel = ""
            try {
                $groupRel = Get-GroupClientRelevance -BaseUrl $base -AuthHeader $auth -SiteName $CustomSiteName -GroupIdNumeric $groupIdNumeric
                LogLine ("Group relevance len ({0}): {1}" -f $a, $groupRel.Length)
            } catch {
                LogLine "❌ Could not fetch/build group relevance for $($a): $($_.Exception.Message)"
                continue  # do NOT post without group relevance
            }

            $fixletActionName = ($FixletActionNameMap[$a]); if (-not $fixletActionName) { $fixletActionName = "Action1" }

            # Build parameters with splatting; avoid passing $null to typed params
            $paramMap = @{
                ActionTitle      = $a
                UiBaseTitle      = $titleRaw
                DisplayName      = $displayName
                SiteName         = $CustomSiteName
                FixletId         = $fixId
                FixletActionName = $fixletActionName
                GroupRelevance   = $groupRel
                StartLocal       = $cfg.Start
                HasTimeRange     = $cfg.TR
                ShowPreActionUI  = $cfg.UI
                PreActionText    = $cfg.Msg
                AskToSaveWork    = $cfg.Save
            }
            if ($cfg.End)      { $paramMap['EndLocal']      = $cfg.End }
            if ($cfg.Deadline) { $paramMap['DeadlineLocal'] = $cfg.Deadline }
            if ($cfg.TR -and $cfg.TRS -ne $null -and $cfg.TRE -ne $null) {
                $paramMap['TimeRangeStart'] = $cfg.TRS
                $paramMap['TimeRangeEnd']   = $cfg.TRE
            }

            $xmlBody = Build-SourcedFixletActionXml @paramMap

            $safeTitle = ($a -replace '[^\w\-. ]','_') -replace '\s+','_'
            $tmpAction = Join-Path $env:TEMP ("BES_Action_{0}_{1:yyyyMMdd_HHmmss}.xml" -f $safeTitle,(Get-Date))
            if ($SaveActionXmlToTemp) {
                Write-Utf8NoBom -Path $tmpAction -Content $xmlBody
                LogLine "Saved action XML for $a to: $tmpAction"
                LogLine ("curl -k -u USER:PASS -H `"Content-Type: application/xml`" -d @`"$tmpAction`" {0}" -f $postUrl)
            }

            try {
                LogLine "Posting $a..."
                Post-XmlFile -Url $postUrl -User $user -Pass $pass -XmlFilePath $tmpAction
                LogLine ("✅ {0} posted successfully." -f $a)
            } catch {
                LogLine ("❌ POST failed for {0}: {1}" -f $a, $_.Exception.Message)
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
