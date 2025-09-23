Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG (EDIT THESE)
# =========================
$LogFile = "C:\temp\BigFixActionGenerator.log"

# Site that hosts the Fixlet + (ideally) the Computer Groups
$CustomSiteName = "Test Group Managed (Workstations)"

# Action -> Computer Group ID (keep 00- prefix; we'll strip to numeric)
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12345"
    "Force"                     = "00-12345"
    "Conference/Training Rooms" = "00-12345"
}

# Fixlet Action name to invoke (use your actual Fixlet action name(s))
$FixletActionNameMap = @{
    "Pilot"                     = "Action1"
    "Deploy"                    = "Action1"
    "Force"                     = "Action1"
    "Conference/Training Rooms" = "Action1"
}

# Behavior toggles
$IgnoreCertErrors           = $true
$DumpFetchedXmlToTemp       = $true
$AggressiveRegexFallback    = $true
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
# Round a DateTime to exact minute (seconds & ms -> 0)
function Round-ToMinute([datetime]$dt) { $dt.Date.AddHours($dt.Hour).AddMinutes($dt.Minute) }

# Build ISO-8601 duration but **rounded to nearest second** (no truncation drift)
function To-IsoDurationRounded([TimeSpan]$ts) {
    if ($ts.Ticks -lt 0) { $ts = [TimeSpan]::Zero }
    $totalSec = [Math]::Round($ts.TotalSeconds, 0, [System.MidpointRounding]::AwayFromZero)
    if ($totalSec -lt 0) { $totalSec = 0 }
    $days  = [int]([Math]::Floor($totalSec / 86400))
    $rem   = $totalSec - ($days * 86400)
    $hours = [int]([Math]::Floor($rem / 3600))
    $rem   = $rem - ($hours * 3600)
    $mins  = [int]([Math]::Floor($rem / 60))
    $secs  = [int]($rem - ($mins * 60))
    $dPart = if ($days -gt 0) { "P{0}D" -f $days } else { "P" }
    $tParts = @()
    if ($hours -gt 0) { $tParts += ("{0}H" -f $hours) }
    if ($mins  -gt 0) { $tParts += ("{0}M" -f $mins) }
    if ($secs  -gt 0 -or $tParts.Count -eq 0) { $tParts += ("{0}S" -f $secs) }
    return $dPart + "T" + ($tParts -join "")
}
function Write-Utf8NoBom([string]$Path,[string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    if ($null -eq $Content) { $Content = "" }
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
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
            -UseBasicParsing
        if ($resp.Content) { LogLine "POST response: $($resp.Content)" }
    } catch {
        $respErr = $_.Exception.Response
        if ($respErr -and $respErr.GetResponseStream) {
            $rs = $respErr.GetResponseStream()
            $sr = New-Object IO.StreamReader($rs, [Text.Encoding]::UTF8)
            $errBody = $sr.ReadToEnd(); $sr.Close()
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
function Get-ActionAndRelevance {
    param($ContainerNode)
    $rels = @()
    $direct = $ContainerNode.SelectNodes("./*[local-name()='Relevance']")
    if ($direct) { foreach ($n in $direct) { $t = ($n.InnerText).Trim(); if ($t) { $rels += $t } } }
    if ($rels.Count -eq 0) {
        $any = $ContainerNode.SelectNodes(".//*[local-name()='Relevance']")
        if ($any) { foreach ($n in $any) { $t = ($n.InnerText).Trim(); if ($t) { $rels += $t } } }
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
    LogLine ("Fixlet relevance nodes found: {0}" -f $rels.Count)
    return @{ Relevance=$rels; ActionScript=$script }
}
function Parse-FixletTitleToProduct([string]$Title) {
    ($Title -replace '^Update:\s*','' -replace '\s+Win$','').Trim()
}
# GROUP relevance helpers
function Extract-AllRelevanceFromXmlString {
    param([string]$XmlString,[string]$Context = "Unknown")
    $all = @()
    try {
        $x = [xml]$XmlString
        $cgRels = $x.SelectNodes("//*[local-name()='ComputerGroup']//*[local-name()='Relevance']")
        if ($cgRels) { foreach ($n in $cgRels) { $t = ($n.InnerText).Trim(); if ($t) { $all += $t } } }
        if ($all.Count -eq 0) {
            $globalRels = $x.SelectNodes("//*[local-name()='Relevance']")
            if ($globalRels) { foreach ($n in $globalRels) { $t = ($n.InnerText).Trim(); if ($t) { $all += $t } } }
        }
    } catch { LogLine "[$Context] XML parse failed: $($_.Exception.Message)" }
    if ($AggressiveRegexFallback -and $all.Count -eq 0) {
        try {
            $regex = [regex]'(?is)<Relevance\b[^>]*>(.*?)</Relevance>'
            foreach ($mm in $regex.Matches($XmlString)) { $t = ($mm.Groups[1].Value).Trim(); if ($t) { $all += $t } }
        } catch { LogLine "[$Context] Regex relevance fallback failed: $($_.Exception.Message)" }
    }
    return ,$all
}
function Extract-SCRFragments {
    param([string]$XmlString,[string]$Context="Unknown")
    $frags = @()
    try {
        $x = [xml]$XmlString
        $scrNodes = $x.SelectNodes("//*[local-name()='SearchComponentRelevance']")
        if ($scrNodes) {
            foreach ($n in $scrNodes) {
                $innerR = $n.SelectNodes(".//*[local-name()='Relevance']")
                if ($innerR -and $innerR.Count -gt 0) {
                    foreach ($ir in $innerR) { $t = ($ir.InnerText).Trim(); if ($t) { $frags += $t } }
                } else {
                    $t = ($n.InnerText).Trim(); if ($t) { $frags += $t }
                }
            }
        }
    } catch { LogLine "[$Context] SCR parse failed: $($_.Exception.Message)" }
    return ,$frags
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
            $rels = Extract-AllRelevanceFromXmlString -XmlString $xmlStr -Context "Group:$GroupIdNumeric"
            if ($rels.Count -gt 0) {
                $joined = ($rels | ForEach-Object { "($_)" }) -join " AND "
                $snippet = $joined.Substring(0, [Math]::Min(200, $joined.Length))
                LogLine "Using group relevance from <Relevance> nodes :: ${snippet}..."
                return $joined
            }
            $frags = Extract-SCRFragments -XmlString $xmlStr -Context "Group:$GroupIdNumeric"
            if ($frags.Count -gt 0) {
                $joined = ($frags | ForEach-Object { "($_)" }) -join " AND "
                $snippet = $joined.Substring(0, [Math]::Min(200, $joined.Length))
                LogLine "Built relevance from SearchComponentRelevance :: ${snippet}..."
                return $joined
            }
            LogLine "No usable relevance at ${url}"
        } catch {
            LogLine ("❌ Group relevance fetch failed ({0}): {1}" -f $GroupIdNumeric, $_.Exception.Message)
        }
    }
    throw "No relevance found or derivable for group ${GroupIdNumeric} in custom/master/operator."
}

# =========================
# ACTION XML BUILDER
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
        [string]$StartOffset,            # PT…
        [string]$HasEndText,             # "true"/"false"
        [string]$EndOffset,              # PT… or empty
        [string]$HasTimeRangeText,       # "true"/"false"
        [string]$TRStartStr,             # "HH:mm:ss" or ""
        [string]$TREndStr,               # "HH:mm:ss" or ""
        [string]$ShowPreActionUIText,    # "true"/"false"
        [string]$PreActionText,
        [string]$AskToSaveWorkText,      # "true"/"false"
        [string]$DeadlineOffset          # PT… or empty (Force only)
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
    if ($DeadlineOffset) {
$deadlineInner = @"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineLocalOffset>$DeadlineOffset</DeadlineLocalOffset>
"@
    }

    $preActionBlock = ""
    if ($ShowPreActionUIText -ieq "true") {
$preActionBlock = @"
      <PreActionShowUI>true</PreActionShowUI>
      <PreAction>
        <Text>$preTextEsc</Text>
        <AskToSaveWork>$AskToSaveWorkText</AskToSaveWork>
        <ShowActionButton>false</ShowActionButton>
        <ShowCancelButton>false</ShowCancelButton>
$deadlineInner        <ShowConfirmation>false</ShowConfirmation>
      </PreAction>
"@
    } else {
$preActionBlock = @"
      <PreActionShowUI>false</PreActionShowUI>
"@
    }

    $endLine = ""
    if ($HasEndText -ieq "true" -and $EndOffset) {
        $endLine = "      <EndDateTimeLocalOffset>$EndOffset</EndDateTimeLocalOffset>`n"
    }

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
$preActionBlock      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>Updating to $dispEsc... Please wait.</Text></RunningMessage>
      <HasTimeRange>$HasTimeRangeText</HasTimeRange>
$timeRangeBlock      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocalOffset>$StartOffset</StartDateTimeLocalOffset>
      <HasEndTime>$HasEndText</HasEndTime>
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
# next 20 Wednesdays
$today = Get-Date
$daysUntilWed = (3 - [int]$today.DayOfWeek + 7) % 7
$nextWed = $today.AddDays($daysUntilWed)
for ($i=0;$i -lt 20;$i++) { [void]$cbDate.Items.Add($nextWed.AddDays(7*$i).ToString("yyyy-MM-dd")) }

# Time (CHANGED: fixed list 11:00 PM → 12:45 AM, 8 slots)
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

# CHANGED: explicit 8-slot list (added 11:45 PM to meet “8 slots”)
$cbTime.Items.AddRange(@(
    "11:00 PM","11:15 PM","11:30 PM","11:45 PM",
    "12:00 AM","12:15 AM","12:30 AM","12:45 AM"
))

# NEW: red warning label (shown only for 12:00 AM and later)
$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$lblWarn.Location = New-Object System.Drawing.Point(160,$y)
$lblWarn.Size = New-Object System.Drawing.Size(440,34)
$lblWarn.Visible = $false
$form.Controls.Add($lblWarn)
$y += 34

# Button
$btn = New-Object System.Windows.Forms.Button
$btn.Text = "Generate & Post 4 Actions (Pilot/Deploy/Force/Conf)"
$btn.Location = New-Object System.Drawing.Point(160,$y)
$btn.Size = New-Object System.Drawing.Size(320,32)  # unchanged baseline
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

# --- helper for warning text ---
function Update-WarnLabel {
    param([string]$slot,[string]$dateStr)
    if ($slot -and $slot -like '12:* AM') {
        try {
            $wed = [datetime]::ParseExact($dateStr,'yyyy-MM-dd',$null)
            $thu = $wed.AddDays(1)
            $lblWarn.Text = "⚠️ Please note: Since you selected $slot, this deployment will technically start on Thursday, " + $thu.ToString('yyyy-MM-dd')
            $lblWarn.Visible = $true
        } catch { $lblWarn.Visible = $false }
    } else {
        $lblWarn.Visible = $false
    }
}
$cbTime.add_SelectedIndexChanged({ Update-WarnLabel $cbTime.SelectedItem $cbDate.SelectedItem })
$cbDate.add_SelectedIndexChanged({ Update-WarnLabel $cbTime.SelectedItem $cbDate.SelectedItem })

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
    $tStr   = $cbTime.SelectedItem

    LogLine ("Fields: server='{0}' user='{1}' fixId='{2}' date='{3}' time='{4}'" -f (Fmt $server),(Fmt $user),(Fmt $fixId),(Fmt $dStr),(Fmt $tStr))

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
        LogLine ("Detected BES content type: {0}" -f $cont.Type)

        $titleRaw = [string]$cont.Node.Title
        if ($null -eq $titleRaw) { $titleRaw = "" }
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw
        LogLine "Parsed title (console): $titleRaw"
        LogLine "Display name (messages): $displayName"

        # ---- Absolute desired times (seconds = 0) ----
        # CHANGED: roll to Thursday if timeslot is 12:00 AM or later
        $PilotStart = [datetime]::ParseExact("$dStr $tStr","yyyy-MM-dd h:mm tt",$null)
        if ($tStr -like '12:* AM') { $PilotStart = $PilotStart.AddDays(1) }
        $PilotStart   = Round-ToMinute($PilotStart)

        $DeployStart  = Round-ToMinute($PilotStart.AddDays(1))
        $ConfStart    = Round-ToMinute($PilotStart.AddDays(1))
        $PilotEnd     = Round-ToMinute($PilotStart.Date.AddDays(1).AddHours(6).AddMinutes(59))
        $DeployEnd    = Round-ToMinute($DeployStart.Date.AddDays(1).AddHours(6).AddMinutes(55))
        $ForceStart   = Round-ToMinute((Get-NextWeekday -base $PilotStart -weekday ([DayOfWeek]::Tuesday)).AddHours(7))
        $ForceDeadline= Round-ToMinute($ForceStart.AddDays(1))     # Wed 7:00 AM

        # 1-year end times for Conference & Force
        $ConfEnd      = Round-ToMinute($ConfStart.AddYears(1))
        $ForceEnd     = Round-ToMinute($ForceStart.AddYears(1))

        # Run between window strings
        $TRStartStr  = "19:00:00"
        $TREndStr    = "06:59:00"

        $actions = @(
            @{ Name="Pilot"; AbsStart=$PilotStart;  AbsEnd=$PilotEnd;    HasEnd="true";  HasTR="true";  TRS=$TRStartStr; TRE=$TREndStr; ShowUI="false"; Msg=""; SaveAsk="false"; AbsDeadline=$null },
            @{ Name="Deploy";AbsStart=$DeployStart; AbsEnd=$DeployEnd;   HasEnd="true";  HasTR="true";  TRS=$TRStartStr; TRE=$TREndStr; ShowUI="false"; Msg=""; SaveAsk="false"; AbsDeadline=$null },
            @{ Name="Conference/Training Rooms"; AbsStart=$ConfStart; AbsEnd=$ConfEnd; HasEnd="true"; HasTR="true"; TRS=$TRStartStr; TRE=$TREndStr; ShowUI="false"; Msg=""; SaveAsk="false"; AbsDeadline=$null },
            @{ Name="Force"; AbsStart=$ForceStart; AbsEnd=$ForceEnd; HasEnd="true"; HasTR="false"; TRS=""; TRE=""; ShowUI="true";
               Msg=("{0} update will be enforced on {1}.  Please leave your machine on overnight to get the automated update.  Otherwise, please close the application and run the update now" -f `
                    $displayName, $ForceDeadline.ToString("M/d/yyyy h:mm tt"));
               SaveAsk="true"; AbsDeadline=$ForceDeadline }
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

            # fetch group relevance
            try {
                LogLine "Fetching group relevance for $a (group $groupIdNumeric)"
                $groupRel = Get-GroupClientRelevance -BaseUrl $base -AuthHeader $auth -SiteName $CustomSiteName -GroupIdNumeric $groupIdNumeric
                $grLen = 0; if ($groupRel) { $grLen = $groupRel.Length }
                LogLine ("Group relevance len ({0}): {1}" -f $a, $grLen)
            } catch {
                LogLine ("❌ Could not fetch/build group relevance for {0}: {1}" -f $a, $_.Exception.Message)
                continue
            }

            $fixletActionName = ($FixletActionNameMap[$a]); if (-not $fixletActionName) { $fixletActionName = "Action1" }

            # Fresh offsets right before POST — with proper rounding to whole seconds
            $postNow = Get-Date
            $startOff   = To-IsoDurationRounded ($cfg.AbsStart - $postNow)
            $endOff     = if ($cfg.HasEnd -ieq "true" -and $cfg.AbsEnd) { To-IsoDurationRounded ($cfg.AbsEnd - $postNow) } else { "" }
            $deadlineOff= if ($cfg.AbsDeadline) { To-IsoDurationRounded ($cfg.AbsDeadline - $postNow) } else { "" }

            $xmlBody = Build-SourcedFixletActionXml `
                -ActionTitle          $a `
                -UiBaseTitle          $titleRaw `
                -DisplayName          $displayName `
                -SiteName             $CustomSiteName `
                -FixletId             $fixId `
                -FixletActionName     $fixletActionName `
                -GroupRelevance       $groupRel `
                -StartOffset          $startOff `
                -HasEndText           $cfg.HasEnd `
                -EndOffset            $endOff `
                -HasTimeRangeText     $cfg.HasTR `
                -TRStartStr           $cfg.TRS `
                -TREndStr             $cfg.TRE `
                -ShowPreActionUIText  $cfg.ShowUI `
                -PreActionText        $cfg.Msg `
                -AskToSaveWorkText    $cfg.SaveAsk `
                -DeadlineOffset       $deadlineOff

            $xmlBodyToSend = $xmlBody

            $safeTitle = ($a -replace '[^\w\-. ]','_') -replace '\s+','_'
            $tmpAction = Join-Path $env:TEMP ("BES_Action_{0}_{1:yyyyMMdd_HHmmss}.xml" -f $safeTitle,(Get-Date))
            if ($SaveActionXmlToTemp) {
                Write-Utf8NoBom -Path $tmpAction -Content $xmlBodyToSend
                LogLine "Saved action XML for $a to: $tmpAction"
                LogLine ("curl -k -u USER:PASS -H `"Content-Type: application/xml`" -d @`"$tmpAction`" {0}" -f $postUrl)
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
