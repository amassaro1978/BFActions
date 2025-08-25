Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

$ErrorActionPreference = "Stop"

# =========================
# CONFIG (EDIT THESE)
# =========================
$LogFile = Join-Path $env:TEMP "BigFixActionGenerator.log"

# Site that hosts the Fixlet and the Computer Groups
$CustomSiteName = "Test Group Managed (Workstations)"

# Action -> Computer Group ID (keep 00- prefix; we'll strip to numeric)
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12346"
    "Force"                     = "00-12347"
    "Conference/Training Rooms" = "00-12348"
}

# Which Fixlet action in the Fixlet to invoke (must exist)
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
    $enc -replace '\+','%20' -replace '\(','%28' -replace '\)','%29'
}
function Get-BaseUrl([string]$ServerInput) {
    if (-not $ServerInput) { throw "Server is empty." }
    $s = $ServerInput.Trim()
    if ($s -notmatch '^(?i)https?://') {
        if ($s -match ':\d+$') { $s = "https://$s" } else { $s = "https://$s:52311" }
    }
    $s.TrimEnd('/')
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
function Write-Utf8NoBom([string]$Path,[string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path,$Content,$utf8NoBom)
}
function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { $Matches[1] } else { ($GroupIdWithPrefix -replace '[^\d]','') }
}
function Get-NextWeekday([datetime]$base,[System.DayOfWeek]$weekday) {
    $anchor = $base.Date
    $delta = ([int]$weekday - [int]$anchor.DayOfWeek + 7) % 7
    if ($delta -le 0) { $delta += 7 }
    $anchor.AddDays($delta)
}
function IsoTimePart([TimeSpan]$ts) {
    if ($ts.Minutes -gt 0) { "PT{0}H{1}M" -f $ts.Hours,$ts.Minutes } else { "PT{0}H" -f $ts.Hours }
}

# Defensive helpers
function Require-NotNull($val,[string]$what,[string]$stage) {
    if ($null -eq $val -or ($val -is [string] -and [string]::IsNullOrWhiteSpace($val))) {
        throw "Stage '$stage': required value '$what' is null/empty."
    }
    $true
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
        if (-not $content) { throw "Empty response" }
        return $content
    } catch {
        throw "HttpGetXml failed ($Url): " + ($_.Exception.GetBaseException().Message)
    }
}

function Post-XmlFile {
    param([string]$Url,[string]$User,[string]$Pass,[string]$XmlFilePath)
    try {
        Require-NotNull (Test-Path $XmlFilePath) "XML file @ $XmlFilePath" "Post-XmlFile" | Out-Null
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

function Get-GroupClientRelevance {
    param([string]$BaseUrl,[string]$AuthHeader,[string]$SiteName,[string]$GroupIdNumeric)
    $encSite = Encode-SiteName $SiteName
    $candidates = @(
        "/api/computergroup/custom/$encSite/$GroupIdNumeric",
        "/api/computergroup/master/$GroupIdNumeric",
        "/api/computergroup/operator/$($env:USERNAME)/$GroupIdNumeric"
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
                LogLine "Found group <Relevance> :: ${snippet}..."
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
                LogLine "Built relevance from SearchComponentRelevance :: ${snippet}..."
                return $joined
            }
            LogLine "No usable relevance at ${url}"
        } catch {
            LogLine "Group relevance fetch failed at ${url}: $($_.Exception.Message)"
        }
    }
    throw "No relevance found for group ${GroupIdNumeric} in custom/master/operator."
}

# =========================
# ACTION XML (SourcedFixletAction)
# =========================
function Build-SourcedFixletActionXml {
    param(
        [string]$ActionTitle, [string]$UiBaseTitle, [string]$DisplayName,
        [string]$SiteName, [string]$FixletId, [string]$FixletActionName,
        [string]$GroupRelevance, [datetime]$StartLocal,
        [Nullable[datetime]]$EndLocal = $null, [Nullable[datetime]]$DeadlineLocal = $null,
        [bool]$HasTimeRange = $false, [Nullable[TimeSpan]]$TimeRangeStart = $null, [Nullable[TimeSpan]]$TimeRangeEnd = $null,
        [bool]$ShowPreActionUI = $false, [string]$PreActionText = "", [bool]$AskToSaveWork = $false
    )

    # Title & message display string
    $uiTitle = [System.Security.SecurityElement]::Escape(("{0}: {1}" -f $UiBaseTitle, $ActionTitle))
    $dispEsc = [System.Security.SecurityElement]::Escape($DisplayName)

    # Snap to :00
    $StartLocal = $StartLocal.Date.AddHours($StartLocal.Hour).AddMinutes($StartLocal.Minute)
    if ($EndLocal.HasValue)      { $EndLocal      = $EndLocal.Value.Date.AddHours($EndLocal.Value.Hour).AddMinutes($EndLocal.Value.Minute) }
    if ($DeadlineLocal.HasValue) { $DeadlineLocal = $DeadlineLocal.Value.Date.AddHours($DeadlineLocal.Value.Hour).AddMinutes($DeadlineLocal.Value.Minute) }

    # Group relevance (safe CDATA)
    $groupSafe = if ([string]::IsNullOrWhiteSpace($GroupRelevance)) { "" } else { $GroupRelevance }
    $groupSafe = $groupSafe -replace ']]>', ']]]]><![CDATA[>'

    # TimeRange block (defaults if requested)
    if ($HasTimeRange) {
        if (-not $TimeRangeStart.HasValue) { $TimeRangeStart = [TimeSpan]::FromHours(19) }
        if (-not $TimeRangeEnd.HasValue)   { $TimeRangeEnd   = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59)) }
        $trs = IsoTimePart $TimeRangeStart.Value
        $tre = IsoTimePart $TimeRangeEnd.Value
        $timeRangeBlock = @"
      <HasTimeRange>true</HasTimeRange>
      <TimeRange>
        <StartTime>$trs</StartTime>
        <EndTime>$tre</EndTime>
      </TimeRange>
"@
    } else { $timeRangeBlock = "      <HasTimeRange>false</HasTimeRange>" }

    # End lines
    $hasEnd  = $EndLocal.HasValue
    $endLine = if ($hasEnd) { "      <EndDateTimeLocal>$($EndLocal.Value.ToString('yyyy-MM-ddTHH:mm:ss'))</EndDateTimeLocal>`n" } else { "" }

    # PreAction & optional absolute Deadline
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
      <HasEndTime>$([string]$hasEnd).ToLower()</HasEndTime>
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

    try {
        # --- Stage: Read UI ---
        $server = $tbServer.Text
        $user   = $tbUser.Text
        $pass   = $tbPass.Text
        $fixId  = $tbFixlet.Text
        $dStr   = $cbDate.SelectedItem
        $tStr   = $cbTime.SelectedItem

        Require-NotNull $server "Server" "UI"
        Require-NotNull $user   "Username" "UI"
        Require-NotNull $pass   "Password" "UI"
        Require-NotNull $fixId  "Fixlet ID" "UI"
        Require-NotNull $dStr   "Date (Wed)" "UI"
        Require-NotNull $tStr   "Time" "UI"

        $server = $server.Trim(); $user=$user.Trim(); $fixId = $fixId.Trim()

        # --- Stage: Build URLs / Auth ---
        $base = Get-BaseUrl $server
        $encodedSite = Encode-SiteName $CustomSiteName
        $fixletUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/fixlet/custom/$encodedSite/$fixId"
        $auth = Get-AuthHeader -User $user -Pass $pass
        LogLine "GET Fixlet URL: ${fixletUrl}"

        # --- Stage: Fetch Fixlet ---
        $fixletContent = HttpGetXml -Url $fixletUrl -AuthHeader $auth
        Require-NotNull $fixletContent "Fixlet XML response" "FetchFixlet"
        if ($DumpFetchedXmlToTemp) {
            $tmpFix = Join-Path $env:TEMP ("BES_Fixlet_{0}.xml" -f $fixId)
            Write-Utf8NoBom -Path $tmpFix -Content $fixletContent
            LogLine "Saved fetched fixlet XML: $tmpFix"
        }

        # --- Stage: Parse Fixlet ---
        $fixletXml = [xml]$fixletContent
        $cont = Get-FixletContainer -Xml $fixletXml
        Require-NotNull $cont "Fixlet container" "ParseFixlet"

        $titleRaw     = [string]$cont.Node.Title
        $displayName  = Parse-FixletTitleToProduct -Title $titleRaw
        LogLine "Fixlet title: $titleRaw"
        LogLine "Display name (for messages): $displayName"

        # Sanity parse action and relevance (we don't inject here since we're Sourced)
        [void](Get-ActionAndRelevance -ContainerNode $cont.Node)

        # --- Stage: Build schedule ---
        $pilotStart = [datetime]::ParseExact("$dStr $tStr","yyyy-MM-dd h:mm tt",$null)
        $pilotStart = $pilotStart.Date.AddHours($pilotStart.Hour).AddMinutes($pilotStart.Minute)

        $deployStart     = $pilotStart.AddDays(1)
        $confStart       = $pilotStart.AddDays(1)
        $pilotEnd        = $pilotStart.Date.AddDays(1).AddHours(6).AddMinutes(59)
        $deployEnd       = $deployStart.Date.AddDays(1).AddHours(6).AddMinutes(55)

        # Force: next Tuesday 7:00 AM; deadline Wed 7:00 AM
        $forceStartDate  = Get-NextWeekday -base $pilotStart -weekday ([DayOfWeek]::Tuesday)
        $forceStart      = $forceStartDate.AddHours(7)
        $forceEnforce    = $forceStart.AddDays(1)

        # Run-between window 7:00 PM – 6:59 AM
        $trStart = [TimeSpan]::FromHours(19)
        $trEnd   = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59))

        $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        LogLine "POST URL: ${postUrl}"

        # --- Stage: Build actions list ---
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
            LogLine "---- Building $a ----"

            $groupIdRaw = "$($GroupMap[$a])"
            Require-NotNull $groupIdRaw "Group ID mapping for $a" "Action:$a"
            $groupIdNumeric = Get-NumericGroupId $groupIdRaw
            Require-NotNull $groupIdNumeric "Parsed numeric Group ID for $a" "Action:$a"

            # fetch group relevance
            $groupRel = Get-GroupClientRelevance -BaseUrl $base -AuthHeader $auth -SiteName $CustomSiteName -GroupIdNumeric $groupIdNumeric
            Require-NotNull $groupRel "Group Relevance for $a" "Action:$a"

            $fixletActionName = ($FixletActionNameMap[$a]); if (-not $fixletActionName) { $fixletActionName = "Action1" }

            # Build params w/o passing nulls into typed params
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
