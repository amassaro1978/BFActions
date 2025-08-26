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
    "Pilot"                       = "00-12345"
    "Deploy"                      = "00-12345"
    "Force"                       = "00-12345"
    "Conference/Training Rooms"   = "00-12345"
}

# Map rollout to the existing Fixlet Action name to invoke
$FixletActionNameMap = @{
    "Pilot"                       = "Action1"
    "Deploy"                      = "Action1"
    "Force"                       = "Action1"
    "Conference/Training Rooms"   = "Action1"
}

# Always use Sourced (lives under the Fixlet's site). Single kept for completeness.
$ActionMode = 'Sourced'   # 'Sourced' or 'Single'

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
function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { return $Matches[1] }
    return ($GroupIdWithPrefix -replace '[^\d]','')
}
# ISO-8601 duration (SingleAction fallback)
function To-IsoDuration([TimeSpan]$ts) {
    if ($ts.Ticks -lt 0) { $ts = [TimeSpan]::Zero }
    $days = [int]$ts.TotalDays
    $hours = $ts.Hours
    $mins  = $ts.Minutes
    $secs  = $ts.Seconds
    $dPart = if ($days -gt 0) { "P{0}D" -f $days } else { "P" }
    $hPart = if ($hours -gt 0) { "{0}H" -f $hours } else { "" }
    $mPart = if ($mins  -gt 0) { "{0}M" -f $mins  } else { "" }
    $sPart = if ($secs  -gt 0) { "{0}S" -f $secs  } else { "" }
    if ($hPart -eq "" -and $mPart -eq "" -and $sPart -eq "") { $sPart = "0S" }
    return $dPart + "T" + $hPart + $mPart + $sPart
}
function Normalize-XmlForPost([string]$s) {
    if (-not $s) { return $s }
    $noBom = $s -replace "^\uFEFF",""
    $noLeadWs = $noBom -replace '^\s+',''
    return $noLeadWs
}
function Write-Utf8NoBom([string]$Path,[string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    if ($null -eq $Content) { $Content = "" }
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}
function Get-FirstBytesHex([string]$s, [int]$n = 32) {
    if (-not $s) { return "" }
    $bytes = [Text.Encoding]::UTF8.GetBytes($s)
    $take = [Math]::Min($n, $bytes.Length)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $take; $i++) { [void]$sb.AppendFormat("{0:X2} ", $bytes[$i]) }
    $sb.ToString().TrimEnd()
}
function Get-NextWeekday([datetime]$base,[System.DayOfWeek]$weekday) {
    $delta = ([int]$weekday - [int]$base.DayOfWeek + 7) % 7
    if ($delta -le 0) { $delta += 7 }
    return $base.Date.AddDays($delta)
}
function SafeEscape([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Security.SecurityElement]::Escape($s)
}
function SafeIsoLocal([Nullable[datetime]]$dt) {
    if ($null -eq $dt) { return "" }
    return $dt.Value.ToString('yyyy-MM-ddTHH:mm:ss')
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
# ACTION XML BUILDER (Sourced)
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
        [datetime]$StartLocal,
        [Nullable[datetime]]$EndLocal = $null,
        [Nullable[datetime]]$DeadlineLocal = $null,
        [bool]$HasTimeRange = $false,
        [object]$TimeRangeStart = $null,
        [object]$TimeRangeEnd   = $null,
        [bool]$ShowPreActionUI = $false,
        [string]$PreActionText = "",
        [bool]$AskToSaveWork = $false
    )

    # ---- Basic sanitization ----
    if ($null -eq $UiBaseTitle) { $UiBaseTitle = "" }
    if ($null -eq $DisplayName) { $DisplayName = "" }
    if ($null -eq $SiteName)    { $SiteName    = "" }
    if ($null -eq $FixletId)    { $FixletId    = "" }
    if ($null -eq $FixletActionName) { $FixletActionName = "Action1" }
    if ($null -eq $PreActionText) { $PreActionText = "" }

    # Console action name (with suffix)
    $fullTitle = ("{0}: {1}" -f $UiBaseTitle, $ActionTitle)
    $uiTitle   = SafeEscape($fullTitle)

    # UI message/title (without suffix) -> "Update: $DisplayName"
    $actionUiTitleRaw = ("Update: {0}" -f $DisplayName)
    $actionUiTitleEsc = SafeEscape($actionUiTitleRaw)
    $dispEsc          = SafeEscape($DisplayName)

    # Snap scheduled times to :00 secs (only when present)
    if ($null -ne $StartLocal)  { $StartLocal   = $StartLocal.Date.AddHours($StartLocal.Hour).AddMinutes($StartLocal.Minute) }
    if ($null -ne $EndLocal)    { $EndLocal     = $EndLocal.Value.Date.AddHours($EndLocal.Value.Hour).AddMinutes($EndLocal.Value.Minute) }
    if ($null -ne $DeadlineLocal){ $DeadlineLocal= $DeadlineLocal.Value.Date.AddHours($DeadlineLocal.Value.Hour).AddMinutes($DeadlineLocal.Value.Minute) }

    $startAbs = if ($null -ne $StartLocal) { $StartLocal.ToString('yyyy-MM-ddTHH:mm:ss') } else { "" }

    # End block
    $hasEnd     = ($null -ne $EndLocal)
    $hasEndText = ([string]$hasEnd).ToLower()
    $endLine    = if ($hasEnd) { "      <EndDateTimeLocal>$([string](SafeIsoLocal($EndLocal)))</EndDateTimeLocal>`n" } else { "" }

    # Group relevance (safe CDATA)
    $groupSafe = if ([string]::IsNullOrWhiteSpace($GroupRelevance)) { "" } else { $GroupRelevance }
    $groupSafe = $groupSafe -replace ']]>', ']]]]><![CDATA[>'

    # TimeRange block -> HH:mm:ss
    $timeRangeBlock = "      <HasTimeRange>false</HasTimeRange>"
    if ($HasTimeRange) {
        # defaults if null
        if ($null -eq $TimeRangeStart) { $trsSpan = [TimeSpan]::FromHours(19) }
        elseif ($TimeRangeStart -is [TimeSpan]) { $trsSpan = $TimeRangeStart }
        else { $trsSpan = [TimeSpan]::Parse($TimeRangeStart.ToString()) }

        if ($null -eq $TimeRangeEnd) { $treSpan = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59)) }
        elseif ($TimeRangeEnd -is [TimeSpan]) { $treSpan = $TimeRangeEnd }
        else { $treSpan = [TimeSpan]::Parse($TimeRangeEnd.ToString()) }

        $trs = "{0:00}:{1:00}:{2:00}" -f $trsSpan.Hours, $trsSpan.Minutes, $trsSpan.Seconds
        $tre = "{0:00}:{1:00}:{2:00}" -f $treSpan.Hours, $treSpan.Minutes, $treSpan.Seconds
$timeRangeBlock = @"
      <HasTimeRange>true</HasTimeRange>
      <TimeRange>
        <StartTime>$trs</StartTime>
        <EndTime>$tre</EndTime>
      </TimeRange>
"@
    }

    # PreAction block (Force uses this + absolute DeadlineLocalTime)
    $preActionBlock = ""
    if ($ShowPreActionUI) {
        $preEsc = SafeEscape($PreActionText)
        $deadlineInner = ""
        if ($null -ne $DeadlineLocal) {
$deadlineInner = @"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineLocalTime>$([string](SafeIsoLocal($DeadlineLocal)))</DeadlineLocalTime>
"@
        }
$preActionBlock = @"
      <PreActionShowUI>true</PreActionShowUI>
      <PreAction>
        <Text>$preEsc</Text>
        <AskToSaveWork>$(([string]$AskToSaveWork).ToLower())</AskToSaveWork>
        <ShowActionButton>false</ShowActionButton>
        <ShowCancelButton>false</ShowCancelButton>
$deadlineInner        <ShowConfirmation>false</ShowConfirmation>
      </PreAction>
"@
    } else {
        $preActionBlock = "      <PreActionShowUI>false</PreActionShowUI>"
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
      <ActionUITitle>$actionUiTitleEsc</ActionUITitle>
$preActionBlock
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>Updating to $dispEsc... Please wait.</Text></RunningMessage>
$timeRangeBlock
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocal>$startAbs</StartDateTimeLocal>
      <HasEndTime>$hasEndText</HasEndTime>
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
# next 20 Wednesdays
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

    LogLine "== Begin click handler =="
    $server = if ($tbServer.Text) { $tbServer.Text.Trim() } else { "" }
    $user   = if ($tbUser.Text)   { $tbUser.Text.Trim()   } else { "" }
    $pass   = $tbPass.Text
    $fixId  = if ($tbFixlet.Text) { $tbFixlet.Text.Trim() } else { "" }
    $dStr   = $cbDate.SelectedItem
    $tStr   = $cbTime.SelectedItem

    LogLine "Fields: server='$server' user='$user' fixId='$fixId' date='$dStr' time='$tStr'"

    if (-not ($server -and $user -and $pass -and $fixId -and $dStr -and $tStr)) {
        LogLine "❌ Please fill in Server, Username, Password, Fixlet ID, Date, and Time."
        return
    }

    try {
        LogLine "Phase: Build base URLs"
        $base = Get-BaseUrl $server
        $encodedSite = Encode-SiteName $CustomSiteName
        $fixletUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/fixlet/custom/$encodedSite/$fixId"
        LogLine "Encoded Fixlet GET URL: ${fixletUrl}"

        LogLine "Phase: Auth + GET fixlet"
        $auth = Get-AuthHeader -User $user -Pass $pass
        $fixletContent = HttpGetXml -Url $fixletUrl -AuthHeader $auth
        if ($DumpFetchedXmlToTemp) {
            $tmpFix = Join-Path $env:TEMP ("BES_Fixlet_{0}.xml" -f $fixId)
            Write-Utf8NoBom -Path $tmpFix -Content $fixletContent
            LogLine "Saved fetched fixlet XML to: $tmpFix"
        }

        LogLine "Phase: Parse fixlet XML"
        $fixletXml = [xml]$fixletContent
        $cont = Get-FixletContainer -Xml $fixletXml
        LogLine ("Detected BES content type: {0}" -f $cont.Type)

        $titleRaw = [string]$cont.Node.Title
        if ($null -eq $titleRaw) { $titleRaw = "" }
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw

        $parsed = Get-ActionAndRelevance -ContainerNode $cont.Node
        $fixletRelevance = @(); if ($parsed.Relevance) { $fixletRelevance = $parsed.Relevance }
        $actionScript = $parsed.ActionScript

        LogLine "Parsed title (console): ${titleRaw}"
        LogLine "Display name (messages): ${displayName}"
        LogLine ("Fixlet relevance count: {0}" -f $fixletRelevance.Count)
        LogLine ("Action script length: {0}" -f $actionScript.Length)

        LogLine "Phase: Build schedules"
        # Absolute schedule (user picks local date/time) — snap to :00 seconds
        $PilotStart = [datetime]::ParseExact("$dStr $tStr","yyyy-MM-dd h:mm tt",$null)
        $PilotStart = $PilotStart.Date.AddHours($PilotStart.Hour).AddMinutes($PilotStart.Minute)

        # Derived schedules per action
        $DeployStart     = $PilotStart.AddDays(1)
        $confStart       = $PilotStart.AddDays(1)
        $PilotEnd        = $PilotStart.Date.AddDays(1).AddHours(6).AddMinutes(59) # next day 6:59 AM
        $DeployEnd       = $DeployStart.Date.AddDays(1).AddHours(6).AddMinutes(55) # next morning 6:55 AM

        # Force: next Tuesday 7:00 AM after Pilot, with deadline Wednesday 7:00 AM
        $forceStartDate  = Get-NextWeekday -base $PilotStart -weekday ([DayOfWeek]::Tuesday)
        $forceStart      = $forceStartDate.AddHours(7) # Tue 7:00 AM
        $forceEnforce    = $forceStart.AddDays(1)      # Wed 7:00 AM

        # TimeRange window (7:00 PM–6:59 AM)
        $trStart = [TimeSpan]::FromHours(19)
        $trEnd   = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59))

        $actions = @(
            @{ Name="Pilot"; Start=$PilotStart; End=$PilotEnd; TR=$true;  TRS=$trStart; TRE=$trEnd; UI=$false; Msg=""; Save=$false; Deadline=$null },
            @{ Name="Deploy";Start=$DeployStart;End=$DeployEnd;TR=$true;  TRS=$trStart; TRE=$trEnd; UI=$false; Msg=""; Save=$false; Deadline=$null },
            @{ Name="Conference/Training Rooms"; Start=$confStart; End=$null; TR=$true; TRS=$trStart; TRE=$trEnd; UI=$false; Msg=""; Save=$false; Deadline=$null },
            @{ Name="Force"; Start=$forceStart; End=$null; TR=$false; TRS=$null; TRE=$null; UI=$true;
               Msg=("{0} update will be enforced on {1}.  Please leave your machine on overnight to get the automated update.  Otherwise, please close the application and run the update now" -f `
                    $displayName, $forceEnforce.ToString("M/d/yyyy h:mm tt"));
               Save=$true; Deadline=$forceEnforce }
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
                $grLen = if ($null -eq $groupRel) { 0 } else { $groupRel.Length }
                LogLine ("Group relevance len ({0}): {1}" -f $a, $grLen)
            } catch {
                LogLine ("❌ Could not fetch/build group relevance for {0}: {1}" -f $a, $_.Exception.Message)
                continue
            }

            $fixletActionName = ($FixletActionNameMap[$a]); if (-not $fixletActionName) { $fixletActionName = "Action1" }

            # Dump parameters for this action (helps catch nulls)
            LogLine ("Params for {0}: Start={1} End={2} Deadline={3} TR={4} TRS={5} TRE={6}" -f `
                $a, ($cfg.Start), ($cfg.End), ($cfg.Deadline), ($cfg.TR), ($cfg.TRS), ($cfg.TRE))

            if ($ActionMode -ieq 'Sourced') {
                LogLine "Assembling SourcedFixletAction XML for $a"
                $xmlBody = Build-SourcedFixletActionXml `
                    -ActionTitle      $a `
                    -UiBaseTitle      $titleRaw `
                    -DisplayName      $displayName `
                    -SiteName         $CustomSiteName `
                    -FixletId         $fixId `
                    -FixletActionName $fixletActionName `
                    -GroupRelevance   $groupRel `
                    -StartLocal       $cfg.Start `
                    -EndLocal         $cfg.End `
                    -DeadlineLocal    $cfg.Deadline `
                    -HasTimeRange     $cfg.TR `
                    -TimeRangeStart   $cfg.TRS `
                    -TimeRangeEnd     $cfg.TRE `
                    -ShowPreActionUI  $cfg.UI `
                    -PreActionText    $cfg.Msg `
                    -AskToSaveWork    $cfg.Save
            } else {
                # SingleAction path (not used by default)
                $allRel = @(); $allRel += $fixletRelevance; if ($groupRel) { $allRel += $groupRel }
                $xmlBody = Build-SingleActionXml `
                    -ActionTitle     $a `
                    -UiBaseTitle     $titleRaw `
                    -DisplayName     $displayName `
                    -RelevanceBlocks $allRel `
                    -ActionScript    $actionScript `
                    -StartLocal      $cfg.Start `
                    -IsForce:($a -eq "Force")
            }

            $xmlBodyToSend = Normalize-XmlForPost $xmlBody
            $hex = Get-FirstBytesHex $xmlBodyToSend 32
            LogLine ("First 32 bytes (hex) for {0}: {1}" -f $a, $hex)

            $safeTitle = ($a -replace '[^\w\-. ]','_') -replace '\s+','_'
            $tmpAction = Join-Path $env:TEMP ("BES_Action_{0}_{1:yyyyMMdd_HHmmss}.xml" -f $safeTitle,(Get-Date))
            if ($SaveActionXmlToTemp) {
                Write-Utf8NoBom -Path $tmpAction -Content $xmlBodyToSend
                LogLine "Saved action XML for $a to: $tmpAction"
                LogLine ("curl -k -u USER:PASS -H `"Content-Type: application/xml`" -d @`"$tmpAction`" {0}" -f $postUrl)
            }

            try {
                if ($PostUsingInvokeWebRequest -and (Test-Path $tmpAction)) {
                    LogLine "Posting via Invoke-WebRequest (file): $tmpAction"
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
