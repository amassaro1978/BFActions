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
    "Deploy"                       = "00-12345"
    "Force"                     = "00-12345"
    "Conference/Training Rooms" = "00-12345"
}

# Map rollout to the **existing Fixlet Action name** to invoke.
# Leave as "Action1" if your Fixlet only has one action; otherwise match your named action(s).
$FixletActionNameMap = @{
    "Pilot"                       = "Action1"
    "Deploy"                       = "Action1"
    "Force"                     = "Action1"
    "Conference/Training Rooms" = "Action1"
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
    $line = "{0}  {1}" -f (Get-Date -Format 'u'), $txt
    $LogBox.AppendText($line + "`r`n")
    Add-Content -Path $LogFile -Value $line
    $LogBox.SelectionStart = $LogBox.Text.Length
    $LogBox.ScrollToCaret()
}
function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { return $Matches[1] }
    return ($GroupIdWithPrefix -replace '[^\d]','')
}
# ISO-8601 duration PnDTnHnMnS (positive from now)
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
# Time-of-day as ISO 8601 duration from midnight, e.g., 19:00 -> PT19H, 6:59 -> PT6H59M
function To-IsoTimeOfDay([int]$hour,[int]$minute=0,[int]$second=0) {
    $parts = @()
    if ($hour   -gt 0) { $parts += ("{0}H" -f $hour) }
    if ($minute -gt 0) { $parts += ("{0}M" -f $minute) }
    if ($second -gt 0 -or $parts.Count -eq 0) { $parts += ("{0}S" -f $second) }
    return "PT" + ($parts -join "")
}
# Normalize leading BOM/whitespace so POST starts with '<'
function Normalize-XmlForPost([string]$s) {
    if (-not $s) { return $s }
    $noBom = $s -replace "^\uFEFF",""
    $noLeadWs = $noBom -replace '^\s+',''
    return $noLeadWs
}
# UTF8 (no BOM) file write
function Write-Utf8NoBom([string]$Path,[string]$Content) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}
# Hex preview (debug)
function Get-FirstBytesHex([string]$s, [int]$n = 32) {
    if (-not $s) { return "" }
    $bytes = [Text.Encoding]::UTF8.GetBytes($s)
    $take = [Math]::Min($n, $bytes.Length)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $take; $i++) { [void]$sb.AppendFormat("{0:X2} ", $bytes[$i]) }
    $sb.ToString().TrimEnd()
}
# Next specific weekday AFTER the given date (never same day)
function Get-NextWeekday([datetime]$base,[System.DayOfWeek]$weekday) {
    $delta = ([int]$weekday - [int]$base.DayOfWeek + 7) % 7
    if ($delta -le 0) { $delta += 7 }
    return $base.Date.AddDays($delta)
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

# curl-like POST using a file on disk
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
        if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream) {
            $sr = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream(), [Text.Encoding]::UTF8)
            $errBody = $sr.ReadToEnd(); $sr.Close()
            throw "Invoke-WebRequest POST failed :: $errBody"
        } else {
            throw ($_.Exception.Message)
        }
    }
}

# Direct-byte POST (fallback)
function HttpPostXml {
    param([string]$Url,[string]$AuthHeader,[string]$XmlBody)
    $bytes = [Text.Encoding]::UTF8.GetBytes($XmlBody)
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = "POST"
    $req.Accept = "application/xml"
    $req.ContentType = "application/xml; charset=utf-8"
    $req.UserAgent = "BigFixActionGenerator/1.0"
    $req.KeepAlive = $false
    $req.Headers["Authorization"] = $AuthHeader
    $req.ProtocolVersion = [Version]"1.1"
    $req.PreAuthenticate = $true
    $req.AllowAutoRedirect = $false
    $req.Timeout = 60000
    $req.ContentLength = $bytes.Length
    try {
        $rs = $req.GetRequestStream(); $rs.Write($bytes,0,$bytes.Length); $rs.Close()
        $resp = $req.GetResponse()
        try {
            $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
            $body = $sr.ReadToEnd(); $sr.Close()
            if ($body) { LogLine "POST response: $body" }
        } finally { $resp.Close() }
    } catch {
        throw ($_.Exception.GetBaseException().Message)
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
            LogLine ("❌ Could not fetch/build group relevance for {0}: {1}" -f $a, $_.Exception.Message)
        }
    }
    throw "No relevance found or derivable for group ${GroupIdNumeric} in custom/master/operator."
}

# =========================
# ACTION XML BUILDERS
# =========================
function Build-SingleActionXml {  # unchanged, kept for completeness
    param(
        [string]$ActionTitle,
        [string]$UiBaseTitle,
        [string]$DisplayName,
        [string[]]$RelevanceBlocks,
        [string]$ActionScript,
        [datetime]$StartLocal,
        [bool]$IsForce=$false
    )
    $fullTitle = "${UiBaseTitle}: $ActionTitle"
    $titleEsc  = [System.Security.SecurityElement]::Escape($fullTitle)
    $dispEsc   = [System.Security.SecurityElement]::Escape($DisplayName)

    $relevanceCombined = ""
    if ($RelevanceBlocks -and $RelevanceBlocks.Count -gt 0) {
        $relevanceCombined = ($RelevanceBlocks | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { "($_)" }) -join " AND "
    }
    $relSafe = $relevanceCombined -replace ']]>', ']]]]><![CDATA[>'
    $rels = if ([string]::IsNullOrWhiteSpace($relevanceCombined)) { "" } else { "    <Relevance><![CDATA[$relSafe]]></Relevance>" }

    $now = Get-Date
    $startOffset = To-IsoDuration ($StartLocal - $now)
@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SingleAction>
    <Title>$titleEsc</Title>
$rels
    <ActionScript MIMEType="application/x-Fixlet-Windows-Shell"><![CDATA[
$ActionScript
]]></ActionScript>
    <SuccessCriteria Option="RunToCompletion" />
    <Settings>
      <ActionUITitle>$titleEsc</ActionUITitle>
      <PreActionShowUI>false</PreActionShowUI>
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>Updating to $dispEsc...please wait.</Text></RunningMessage>
      <HasTimeRange>false</HasTimeRange>
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocalOffset>$startOffset</StartDateTimeLocalOffset>
      <HasEndTime>false</HasEndTime>
      <HasDayOfWeekConstraint>false</HasDayOfWeekConstraint>
      <UseUTCTime>false</UseUTCTime>
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
    <Target><AllComputers>true</AllComputers></Target>
  </SingleAction>
</BES>
"@
}

# SourcedFixletAction builder — emits only the elements we request for each action
function Build-SourcedFixletActionXml {
    param(
        [string]$ActionTitle,       # Pilot/Deploy/Force/Conference...
        [string]$UiBaseTitle,       # Full Fixlet title ("Update: ... Win")
        [string]$DisplayName,       # Product-friendly name for messages
        [string]$SiteName,          # Custom site name
        [string]$FixletId,          # Fixlet ID
        [string]$FixletActionName,  # "Action1" or a named action that exists in the Fixlet
        [string]$GroupRelevance,    # Group filter to AND with fixlet relevance
        [datetime]$StartLocal,      # scheduled local start (absolute)
        [NulDeployle[datetime]]$EndLocal = $null,   # optional absolute end
        [NulDeployle[datetime]]$DeadlineLocal = $null, # optional absolute deadline
        [bool]$HasTimeRange = $false,
        [NulDeployle[TimeSpan]]$TimeRangeStart = $null, # time-of-day from midnight
        [NulDeployle[TimeSpan]]$TimeRangeEnd = $null,   # time-of-day from midnight
        [bool]$ShowPreActionUI = $false,
        [string]$PreActionText = "",   # shown only if ShowPreActionUI=true
        [bool]$AskToSaveWork = $false  # only meaningful when ShowPreActionUI=true
    )

    $fullTitle = "${UiBaseTitle}: $ActionTitle"
    $uiTitle   = [System.Security.SecurityElement]::Escape($fullTitle)
    $dispEsc   = [System.Security.SecurityElement]::Escape($DisplayName)

    # Sanitize relevance
    if ([string]::IsNullOrWhiteSpace($GroupRelevance)) { $groupSafe = "" } else { $groupSafe = $GroupRelevance }
    $groupSafe = $groupSafe -replace ']]>', ']]]]><![CDATA[>'

    # Offsets
    $now = Get-Date
    $startOffset = To-IsoDuration ($StartLocal - $now)
    $hasEnd = $false; $endOffsetLine = ""
    if ($EndLocal.HasValue) {
        $hasEnd = $true
        $endOffsetLine = "      <EndDateTimeLocalOffset>$((To-IsoDuration ($EndLocal.Value - $now)))</EndDateTimeLocalOffset>`n"
    }

    # Deadline (Force)
    $deadlineBlock = ""
    if ($DeadlineLocal.HasValue) {
        $deadlineOffset = To-IsoDuration ($DeadlineLocal.Value - $now)
        $deadlineBlock = @"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineLocalOffset>$deadlineOffset</DeadlineLocalOffset>
"@
    }

    # TimeRange block
    $timeRangeBlock = ""
    if ($HasTimeRange -and $TimeRangeStart.HasValue -and $TimeRangeEnd.HasValue) {
        $trs = To-IsoTimeOfDay($TimeRangeStart.Value.Hours, $TimeRangeStart.Value.Minutes, $TimeRangeStart.Value.Seconds)
        $tre = To-IsoTimeOfDay($TimeRangeEnd.Value.Hours,   $TimeRangeEnd.Value.Minutes,   $TimeRangeEnd.Value.Seconds)
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

    # PreAction (only if requested)
    $preActionBlock = ""
    if ($ShowPreActionUI) {
        $preEsc = [System.Security.SecurityElement]::Escape($PreActionText)
        $preActionBlock = @"
      <PreActionShowUI>true</PreActionShowUI>
      <PreAction>
        <Text>$preEsc</Text>
        <AskToSaveWork>$($AskToSaveWork.ToString().ToLower())</AskToSaveWork>
        <ShowActionButton>false</ShowActionButton>
        <ShowCancelButton>false</ShowCancelButton>$deadlineBlock        <ShowConfirmation>false</ShowConfirmation>
      </PreAction>
"@
    } else {
        # No pre-action UI
        $preActionBlock = "      <PreActionShowUI>false</PreActionShowUI>"
        if ($deadlineBlock) {
            # Deadline without showing UI: still allowed; add only the deadline elements under Settings
            $preActionBlock += "`n      " + $deadlineBlock.Trim()
        }
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
      <RunningMessage><Text>Updating to $dispEsc...please wait.</Text></RunningMessage>
$timeRangeBlock
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocalOffset>$startOffset</StartDateTimeLocalOffset>
      <HasEndTime>$($hasEnd.ToString().ToLower())</HasEndTime>
$endOffsetLine      <HasDayOfWeekConstraint>false</HasDayOfWeekConstraint>
      <UseUTCTime>false</UseUTCTime>
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
function Add-Field([string]$Deployel,[bool]$IsPassword,[ref]$OutTB) {
    $lbl = New-Object System.Windows.Forms.Deployel
    $lbl.Text = $Deployel
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
$lblDate = New-Object Windows.Forms.Deployel
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
$lblTime = New-Object System.Windows.Forms.Deployel
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
        LogLine ("Detected BES content type: {0}" -f $cont.Type)

        $titleRaw = [string]$cont.Node.Title
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw

        $parsed = Get-ActionAndRelevance -ContainerNode $cont.Node
        $fixletRelevance = @(); if ($parsed.Relevance) { $fixletRelevance = $parsed.Relevance }
        $actionScript = $parsed.ActionScript

        LogLine "Parsed title (console): ${titleRaw}"
        LogLine "Display name (messages): ${displayName}"
        LogLine ("Fixlet relevance count: {0}" -f $fixletRelevance.Count)
        LogLine ("Action script length: {0}" -f $actionScript.Length)

        # Absolute schedule (user picks local date/time)
        $PilotStart = Get-Date "$dStr $tStr"

        # Derived schedules per action
        $DeployStart     = $PilotStart.AddDays(1)
        $confStart       = $PilotStart.AddDays(1)
        $PilotEnd        = $PilotStart.Date.AddDays(1).AddHours(6).AddMinutes(59) # next day 6:59 AM
        $DeployEnd       = $DeployStart.Date.AddDays(1).AddHours(6).AddMinutes(55) # next morning 6:55 AM
        $forceStartDate  = Get-NextWeekday -base $PilotStart -weekday ([DayOfWeek]::Tuesday)
        $forceStart      = $forceStartDate.AddHours(7) # 7:00 AM
        $forceEnforce    = $forceStart.AddHours(24)    # message date + deadline

        # TimeRange window (7:00 PM–6:59 AM)
        $trStart = [TimeSpan]::FromHours(19)               # 7:00 PM -> PT19H
        $trEnd   = [TimeSpan]::FromHours(6).Add([TimeSpan]::FromMinutes(59))  # 6:59 AM -> PT6H59M

        $actions = @(
            @{ Name="Pilot"; Start=$PilotStart; End=$PilotEnd; TR=$true;  TRS=$trStart; TRE=$trEnd; UI=$false; Msg="";    Save=$false; Deadline=$null },
            @{ Name="Deploy";Start=$DeployStart;End=$DeployEnd;TR=$false; TRS=$null;   TRE=$null; UI=$false; Msg="";    Save=$false; Deadline=$null },
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
            $groupIdRaw = "$($GroupMap[$a])"
            if (-not $groupIdRaw) { LogLine "❌ Missing group id for $a"; continue }
            $groupIdNumeric = Get-NumericGroupId $groupIdRaw
            if (-not $groupIdNumeric) { LogLine ("❌ Could not parse numeric ID from '{0}' for {1}" -f $groupIdRaw, $a); continue }

            # fetch group relevance
            try {
                $groupRel = Get-GroupClientRelevance -BaseUrl $base -AuthHeader $auth -SiteName $CustomSiteName -GroupIdNumeric $groupIdNumeric
                LogLine ("Group relevance len ({0}): {1}" -f $a, $groupRel.Length)
            } catch {
                LogLine ("❌ Could not fetch/build group relevance for {0}: {1}" -f $a, $_.Exception.Message)
                continue
            }

            $fixletActionName = ($FixletActionNameMap[$a]); if (-not $fixletActionName) { $fixletActionName = "Action1" }

            if ($ActionMode -ieq 'Sourced') {
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
                # (SingleAction path retained but not used by default)
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
                    LogLine "Posting via Invoke-WebRequest (curl-like) using file: $tmpAction"
                    Post-XmlFile-InFile -Url $postUrl -User $user -Pass $pass -XmlFilePath $tmpAction
                } else {
                    LogLine "Posting via HttpWebRequest body (direct bytes)"
                    $authHeader = Get-AuthHeader -User $user -Pass $pass
                    HttpPostXml -Url $postUrl -AuthHeader $authHeader -XmlBody $xmlBodyToSend
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
