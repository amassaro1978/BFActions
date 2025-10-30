# =========================================================
# BigFix Action Generator — Baseline-2025-09-24-ForceCascade-RunWindow22-ConfirmDialog
# Force timing: Next Tuesday 07:00 after selected Wednesday; deadline = ForceStart + 24h
# Starts/Ends absolute local (:00 seconds) — no drift
# Targeting: (member of group <id> of sites)
# =========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG
# =========================
$LogFile = "C:\temp\BigFixActionGenerator.log"

# Site that hosts the Fixlet
$CustomSiteName = "Test Group Managed (Workstations)"

# Action -> Computer Group ID (keep 00- prefix; we'll strip to numeric)
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12345"
    "Force"                     = "00-12345"
    "Conference/Training Rooms" = "00-12345"
}

# Fixlet Action name to invoke inside the Fixlet
$FixletActionNameMap = @{
    "Pilot"                     = "Action1"
    "Deploy"                    = "Action1"
    "Force"                     = "Action1"
    "Conference/Training Rooms" = "Action1"
}

# Targeting mode
$UseDirectGroupMembershipRelevance = $true
$UseSitesPlural = $true                         # true => (member of group <id> of sites)
$GroupSiteNameForSpecificMode = $CustomSiteName # only if $UseSitesPlural = $false

# Behavior toggles
$IgnoreCertErrors           = $true
$DumpFetchedXmlToTemp       = $true
$SaveActionXmlToTemp        = $true
$PostUsingInvokeWebRequest  = $true

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
    try {
        $line = "{0}  {1}" -f (Get-Date -Format 'u'), $txt
        if ($LogBox) { $LogBox.AppendText($line + "`r`n"); $LogBox.SelectionStart = $LogBox.Text.Length; $LogBox.ScrollToCaret() }
        Add-Content -Path $LogFile -Value $line
    } catch {}
}
function Fmt($v) { if ($null -eq $v) { return "<null>" } else { return $v } }
function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { return $Matches[1] }
    return ($GroupIdWithPrefix -replace '[^\d]','')
}
# Snap to exact minute (:00 seconds)
function Snap-ToExactMinute([datetime]$dt) {
    $d = $dt
    if ($d.Second -ne 0) { $d = $d.AddSeconds(-$d.Second) }
    if ($d.Millisecond -ne 0) { $d = $d.AddMilliseconds(-$d.Millisecond) }
    return $d
}
function IsoLocal([datetime]$dt) {
    # Absolute local time with explicit :00 seconds
    return (Snap-ToExactMinute $dt).ToString("yyyy-MM-dd'T'HH:mm:ss")
}
function Get-NextWeekday([datetime]$base,[System.DayOfWeek]$weekday) {
    $delta = ([int]$weekday - [int]$base.DayOfWeek + 7) % 7
    if ($delta -le 0) { $delta += 7 }
    $base.Date.AddDays($delta)
}
function SafeEscape([string]$s) {
    if ($null -eq $s) { return "" }
    [System.Security.SecurityElement]::Escape($s)
}
function Write-Utf8NoBom([string]$Path,[string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    if ($null -eq $Content) { $Content = "" }
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}
# Build ISO-8601 duration rounded to nearest second (kept in case needed elsewhere)
function To-IsoDurationRounded([TimeSpan]$ts) {
    $totalSec = [Math]::Round($ts.TotalSeconds, 0, [System.MidpointRounding]::AwayFromZero)
    if ($totalSec -lt 60) { $totalSec = 60 }
    $days  = [int]([Math]::Floor($totalSec / 86400))
    $rem   = $totalSec - ($days * 86400)
    $hours = [int]([Math]::Floor($rem / 3600)); $rem -= ($hours * 3600)
    $mins  = [int]([Math]::Floor($rem / 60));   $rem -= ($mins * 60)
    $secs  = [int]$rem
    $dPart = if ($days -gt 0) { "P{0}D" -f $days } else { "P" }
    $tParts = @()
    if ($hours -gt 0) { $tParts += ("{0}H" -f $hours) }
    if ($mins  -gt 0) { $tParts += ("{0}M" -f $mins) }
    if ($secs  -gt 0 -or $tParts.Count -eq 0) { $tParts += ("{0}S" -f $secs) }
    return $dPart + "T" + ($tParts -join "")
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
    $req.Headers["Accept-Encoding"] = "gzip, deflate"
    $req.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    if ($AuthHeader) { $req.Headers["Authorization"] = $AuthHeader }
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

function Post-XmlFile-InFile {
    param([string]$Url,[string]$User,[string]$Pass,[string]$XmlFilePath)
    try {
        $pair  = "$User`:$Pass"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
        $basic = "Basic " + [Convert]::ToBase64String($bytes)
        $resp = Invoke-WebRequest -Method Post -Uri $Url `
            -Headers @{ "Authorization" = $basic } `
            -ContentType "application/xml" `
            -InFile $XmlFilePath `
            -UseBasicParsing `
            -ErrorAction Stop
        if ($resp.Content) { LogLine "POST response: $($resp.Content)" }
    } catch {
        $respErr = $_.Exception.Response
        if ($respErr -and $respErr.GetResponseStream) {
            $rs = $respErr.GetResponseStream()
            $sr = New-Object IO.StreamReader($rs, [Text.Encoding]::UTF8)
            $errBody = $sr.ReadToEnd(); $sr.Close()
            $errFile = Join-Path $env:TEMP ("BES_Post_Error_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))
            try { [System.IO.File]::WriteAllText($errFile, $errBody, [Text.Encoding]::UTF8) } catch {}
            LogLine ("❌ Server Error body (first 2000 chars): {0}" -f ($errBody.Substring(0,[Math]::Min(2000,$errBody.Length))))
            LogLine ("Saved full error to: {0}" -f $errFile)
            throw "Invoke-WebRequest POST failed :: $errBody"
        }
        throw ($_.Exception.Message)
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
function Parse-FixletTitleToProduct([string]$Title) {
    ($Title -replace '^Update:\s*','' -replace '\s+Win$','').Trim()
}

# Direct membership relevance
function Build-GroupMembershipRelevance([string]$SiteName,[string]$GroupIdNumeric,[bool]$UseSitesPluralLocal = $UseSitesPlural) {
    if ($UseSitesPluralLocal) {
        return "(member of group $GroupIdNumeric of sites)"
    } else {
        $siteEsc = $SiteName.Replace('"','\"')
        return "(member of group $GroupIdNumeric of site whose (name of it = `"$siteEsc`"))"
    }
}

# =========================
# ACTION XML BUILDER (ABSOLUTE START/END; ABSOLUTE DEADLINE OPTION)
# =========================
function Build-SourcedFixletActionXml {
    param(
        [string]$ActionTitle,
        [string]$UiBaseTitle,
        [string]$DisplayName,
        [string]$SiteName,
        [string]$FixletId,
        [string]$FixletActionName,
        [string]$GroupRelevance,
        [string]$StartLocal,            # yyyy-MM-ddTHH:mm:ss
        [string]$HasEndText,            # "true"/"false"
        [string]$EndLocal,              # yyyy-MM-ddTHH:mm:ss or ""
        [string]$HasTimeRangeText,      # "true"/"false"
        [string]$TRStartStr,            # "HH:mm:ss" or ""
        [string]$TREndStr,              # "HH:mm:ss" or ""
        [string]$ShowPreActionUIText,   # "true"/"false"
        [string]$PreActionText,
        [string]$AskToSaveWorkText,     # "true"/"false"
        [string]$DeadlineLocal,         # yyyy-MM-ddTHH:mm:ss or ""
        [string]$DeadlineOffset         # PT… or ""
    )

    $consoleTitle    = SafeEscape(("{0}: {1}" -f $UiBaseTitle, $ActionTitle))
    $uiTitleMessage  = SafeEscape(("Update: {0}" -f $DisplayName))
    $dispEsc         = SafeEscape($DisplayName)
    $siteEsc         = SafeEscape($SiteName)
    $fixletIdEsc     = SafeEscape($FixletId)
    $actionNameEsc   = SafeEscape($FixletActionName)
    $preTextEsc      = SafeEscape($PreActionText)

    $groupSafe = if ([string]::IsNullOrWhiteSpace($GroupRelevance)) { "" } else { $GroupRelevance }
    $groupSafe = $groupSafe -replace ']]>', ']]]]><![CDATA[>'

    $timeRangeBlock = ""
    if ($HasTimeRangeText -ieq "true") {
        $trStartLine = if ($TRStartStr) { "        <StartTime>$TRStartStr</StartTime>" } else { "" }
        $trEndLine   = if ($TREndStr)   { "        <EndTime>$TREndStr</EndTime>" }     else { "" }
$timeRangeBlock = @"
      <TimeRange>
$trStartLine
$trEndLine
      </TimeRange>
"@
    }

    $deadlineInner = ""
    if ($DeadlineLocal) {
$deadlineInner = @"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineDateTimeLocal>$DeadlineLocal</DeadlineDateTimeLocal>
"@
    } elseif ($DeadlineOffset) {
$deadlineInner = @"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineLocalOffset>$DeadlineOffset</DeadlineLocalOffset>
"@
    }

    $endLine = ""
    if ($HasEndText -ieq "true" -and $EndLocal) {
        $endLine = "      <EndDateTimeLocal>$EndLocal</EndDateTimeLocal>`n"
    }

    $startLine = "      <StartDateTimeLocal>$StartLocal</StartDateTimeLocal>"

@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SourcedFixletAction>
    <SourceFixlet>
      <Sitename>$siteEsc</Sitename>
      <FixletID>$fixletIdEsc</FixletID>
      <Action>$actionNameEsc</Action>
    </SourceFixlet>
    <Target>
      <CustomRelevance><![CDATA[$groupSafe]]></CustomRelevance>
    </Target>
    <Settings>
      <ActionUITitle>$uiTitleMessage</ActionUITitle>
      <PreActionShowUI>$ShowPreActionUIText</PreActionShowUI>
      $(if ($ShowPreActionUIText -ieq "true") { @"
      <PreAction>
        <Text>$preTextEsc</Text>
        <AskToSaveWork>$AskToSaveWorkText</AskToSaveWork>
        <ShowActionButton>false</ShowActionButton>
        <ShowCancelButton>false</ShowCancelButton>
$deadlineInner        <ShowConfirmation>false</ShowConfirmation>
      </PreAction>
"@ } else { "" })
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>Updating to $dispEsc... Please wait.</Text></RunningMessage>
      <HasTimeRange>$HasTimeRangeText</HasTimeRange>
$timeRangeBlock      <HasStartTime>true</HasStartTime>
$startLine
      <HasEndTime>$HasEndText</HasEndTime>
$endLine      <UseUTCTime>false</UseUTCTime>
      <ActiveUserRequirement>NoRequirement</ActiveUserRequirement>
      <ActiveUserType>AllUsers</ActiveUserType>
      <HasWhose>false</HasWhose>
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
    <Title>$consoleTitle</Title>
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
function Add-Field([string]$Label,[bool]$IsPassword,[ref]$OutTB,[string]$DefaultValue="",$ReadOnly=$false) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = New-Object System.Drawing.Point(10,$script:y)
    $lbl.Size = New-Object System.Drawing.Size(140,22)
    $form.Controls.Add($lbl)

    if ($IsPassword) { $tb = New-Object System.Windows.Forms.MaskedTextBox; $tb.PasswordChar = '*' }
    else { $tb = New-Object System.Windows.Forms.TextBox }

    $tb.Location = New-Object System.Drawing.Point(160,$script:y)
    $tb.Size = New-Object System.Drawing.Size(440,22)
    if ($DefaultValue) { $tb.Text = $DefaultValue }
    if ($ReadOnly) { $tb.ReadOnly = $true }
    $form.Controls.Add($tb)
    $OutTB.Value = $tb
    $script:y += 34
}

# Pre-populate BigFix Server with default and lock it
$tbServer = $null; Add-Field "BigFix Server:" $false ([ref]$tbServer) "https://test.server:52311" $true
$tbUser   = $null; Add-Field "Username:"      $false ([ref]$tbUser)
$tbPass   = $null; Add-Field "Password:"      $true  ([ref]$tbPass)
$tbFixlet = $null; Add-Field "Fixlet ID:"     $false ([ref]$tbFixlet)

# Date (future Wednesdays only)
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

# populate next 20 Wednesdays
$today = Get-Date
$daysUntilWed = (3 - [int]$today.DayOfWeek + 7) % 7
$nextWed = $today.AddDays($daysUntilWed)
for ($i=0;$i -lt 20;$i++) { [void]$cbDate.Items.Add($nextWed.AddDays(7*$i).ToString("yyyy-MM-dd")) }

# Timeslot selector dormant; default is 11:00 PM
$DefaultAnchorTime = "11:00 PM"

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
    LogLine "== Begin click handler =="

    $server = if ($tbServer.Text) { $tbServer.Text.Trim() } else { "" }
    $user   = if ($tbUser.Text)   { $tbUser.Text.Trim()   } else { "" }
    $pass   = $tbPass.Text
    $fixId  = if ($tbFixlet.Text) { $tbFixlet.Text.Trim() } else { "" }
    $dStr   = $cbDate.SelectedItem
    $tStr   = $DefaultAnchorTime

    LogLine ("Fields: server='{0}' user='{1}' fixId='{2}' date='{3}' time='{4}'" -f (Fmt $server),(Fmt $user),(Fmt $fixId),(Fmt $dStr),(Fmt $tStr))

    if (-not ($server -and $user -and $pass -and $fixId -and $dStr)) {
        LogLine "❌ Please fill in Server, Username, Password, Fixlet ID, and Date."
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
        LogLine ("Detected BES content type: {0}" -f $cont.Type)

        $titleRaw = [string]$cont.Node.Title
        if ($null -eq $titleRaw) { $titleRaw = "" }
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw
        LogLine "Parsed title (console): $titleRaw"
        LogLine "Display name (messages): $displayName"

        # ---- Absolute desired times, snapped to exact minute (:00) ----
        $PilotStart  = Snap-ToExactMinute([datetime]::ParseExact("$dStr $tStr","yyyy-MM-dd h:mm tt",[System.Globalization.CultureInfo]::InvariantCulture))
        $DeployStart = Snap-ToExactMinute($PilotStart.AddDays(1))
        $ConfStart   = Snap-ToExactMinute($PilotStart.AddDays(1))

        # Run window 22:00–06:59 for Pilot/Deploy/Conf
        $TRStartStr   = "22:00:00"
        $TREndStr     = "06:59:00"

        # Pilot ends next morning 06:59
        $PilotEnd = Snap-ToExactMinute($PilotStart.Date.AddDays(1).AddHours(6).AddMinutes(59))

        # Deploy ends the following Tuesday at 06:55 AM (relative to anchor Wed)
        $nextTueAfterPilot = Get-NextWeekday -base $PilotStart -weekday ([DayOfWeek]::Tuesday)
        if ($nextTueAfterPilot -le $PilotStart) { $nextTueAfterPilot = $nextTueAfterPilot.AddDays(7) }
        $DeployEnd = Snap-ToExactMinute($nextTueAfterPilot.AddHours(6).AddMinutes(55))

        # -------- FORCE: Next Tuesday 07:00 AFTER the selected Wednesday ----------
        $ForceStart = Snap-ToExactMinute($nextTueAfterPilot.AddHours(7))  # Tue 07:00 after anchor Wed
        # Desired absolute deadline = ForceStart + 24h (i.e., Wed 07:00) — pin seconds to :00
        $ForceDeadlineAbs = Snap-ToExactMinute($ForceStart.AddDays(1))

        # 1-year end times for Conference & Force
        $ConfEnd  = Snap-ToExactMinute($ConfStart.AddYears(1))
        $ForceEnd = Snap-ToExactMinute($ForceStart.AddYears(1))

        $actions = @(
            @{ Name="Pilot"; AbsStart=$PilotStart;  AbsEnd=$PilotEnd;    HasEnd="true";  HasTR="true";  TRS=$TRStartStr; TRE=$TREndStr; ShowUI="false"; Msg=""; SaveAsk="false"; DeadlineLocal="";  DeadlineOffset=""      },
            @{ Name="Deploy";AbsStart=$DeployStart; AbsEnd=$DeployEnd;   HasEnd="true";  HasTR="true";  TRS=$TRStartStr; TRE=$TREndStr; ShowUI="false"; Msg=""; SaveAsk="false"; DeadlineLocal="";  DeadlineOffset=""      },
            @{ Name="Conference/Training Rooms"; AbsStart=$ConfStart; AbsEnd=$ConfEnd; HasEnd="true"; HasTR="true"; TRS=$TRStartStr; TRE=$TREndStr; ShowUI="false"; Msg=""; SaveAsk="false"; DeadlineLocal="";  DeadlineOffset="" },
            @{ Name="Force"; AbsStart=$ForceStart; AbsEnd=$ForceEnd; HasEnd="true"; HasTR="false"; TRS=""; TRE=""; ShowUI="true";
               Msg=("{0} update will be enforced on {1}.  Please leave your machine on overnight to get the automated update.  Otherwise, please close the application and run the update now" -f `
                    $displayName, $ForceDeadlineAbs.ToString("M/d/yyyy h:mm tt"));
               SaveAsk="true"; DeadlineLocal=(IsoLocal $ForceDeadlineAbs); DeadlineOffset="" }
        )

        $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        LogLine "POST URL: ${postUrl}"

        foreach ($cfg in $actions) {
            $a = $cfg.Name
            LogLine "---- Building: $a ----"

            $groupIdRaw = "$($GroupMap[$a])"
            if (-not $groupIdRaw) { LogLine "❌ Missing group id for $a"; continue }
            $groupIdNumeric = Get-NumericGroupId $groupIdRaw
            if (-not $groupIdNumeric) { LogLine ("❌ Could not parse numeric ID from '{0}' for {1}" -f $groupIdRaw, $a); continue }

            # === Targeting ===
            $groupRel = if ($UseDirectGroupMembershipRelevance) {
                $siteForSpecific = if ([string]::IsNullOrWhiteSpace($GroupSiteNameForSpecificMode)) { $CustomSiteName } else { $GroupSiteNameForSpecificMode }
                Build-GroupMembershipRelevance -SiteName $siteForSpecific -GroupIdNumeric $groupIdNumeric -UseSitesPluralLocal:$UseSitesPlural
            } else {
                throw "Direct membership relevance is required by current baseline."
            }
            LogLine ("Using direct membership relevance: {0}" -f $groupRel)

            # ==== ABSOLUTE TIME STRINGS ====
            $startLocal    = IsoLocal $cfg.AbsStart
            $endLocal      = if ($cfg.HasEnd -ieq "true" -and $cfg.AbsEnd) { IsoLocal $cfg.AbsEnd } else { "" }

            $deadlineLocal = $cfg.DeadlineLocal
            $deadlineOff   = $cfg.DeadlineOffset

            $xmlBody = Build-SourcedFixletActionXml `
                -ActionTitle          $a `
                -UiBaseTitle          $titleRaw `
                -DisplayName          $displayName `
                -SiteName             $CustomSiteName `
                -FixletId             $fixId `
                -FixletActionName     $FixletActionNameMap[$a] `
                -GroupRelevance       $groupRel `
                -StartLocal           $startLocal `
                -HasEndText           $cfg.HasEnd `
                -EndLocal             $endLocal `
                -HasTimeRangeText     $cfg.HasTR `
                -TRStartStr           $cfg.TRS `
                -TREndStr             $cfg.TRE `
                -ShowPreActionUIText  $cfg.ShowUI `
                -PreActionText        $cfg.Msg `
                -AskToSaveWorkText    $cfg.SaveAsk `
                -DeadlineLocal        $deadlineLocal `
                -DeadlineOffset       $deadlineOff

            $safeTitle = ($a -replace '[^\w\-. ]','_') -replace '\s+','_'
            $tmpAction = Join-Path $env:TEMP ("BES_Action_{0}_{1:yyyyMMdd_HHmmss}.xml" -f $safeTitle,(Get-Date))
            if ($SaveActionXmlToTemp) {
                Write-Utf8NoBom -Path $tmpAction -Content $xmlBody
                LogLine "Saved action XML for $a to: $tmpAction"
                LogLine ("curl -k -u USER:PASS -H `"Content-Type: application/xml`" -d @`"$tmpAction`" {0}" -f $postUrl)
            }

            # Confirmation dialog (once, on first/ Pilot)
            if ($a -eq "Pilot") {
                $dlg = [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    ("Fixlet: {0}`r`nCreate the 4 actions (Pilot/Deploy/Force/Conf) now?" -f $titleRaw),
                    "Confirm — Create Actions",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question,
                    [System.Windows.Forms.MessageBoxDefaultButton]::Button2
                )
                if ($dlg -ne [System.Windows.Forms.DialogResult]::Yes) {
                    LogLine "🚫 User canceled."
                    return
                }
            }

            try {
                if ($PostUsingInvokeWebRequest -and (Test-Path $tmpAction)) {
                    Post-XmlFile-InFile -Url $postUrl -User $user -Pass $pass -XmlFilePath $tmpAction
                } else {
                    LogLine "⚠️ Direct POST path disabled; enable if needed."
                }
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
