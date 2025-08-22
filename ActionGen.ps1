Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG (EDIT THESE)
# =========================
$LogFile = Join-Path $env:TEMP "BigFixActionGenerator.log"

# Your site that hosts the Fixlet and (ideally) the Computer Groups
$CustomSiteName = "Test Group Managed (Workstations)"

# Action -> Computer Group ID (keep 00- prefix; we'll strip to numeric for API)
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12346"
    "Force"                     = "00-12347"
    "Conference/Training Rooms" = "00-12348"
}

# Match curl -k behavior if your server uses an untrusted cert
$IgnoreCertErrors        = $true
# Dump fetched XMLs to temp for side-by-side compare with curl output
$DumpFetchedXmlToTemp    = $true
# Last-resort regex extractor for <Relevance>…</Relevance> if XML shape is odd
$AggressiveRegexFallback = $true

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
# Build an ISO-8601 duration like PnDTnHnMnS (positive from "now")
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

# =========================
# HTTP
# =========================
if ($IgnoreCertErrors) {
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch { }
}

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
# FIXLET & GROUP PARSING
# =========================
function Get-FixletContainer { param([xml]$Xml)
    if ($Xml.BES.Fixlet)   { return @{ Type="Fixlet";   Node=$Xml.BES.Fixlet } }
    if ($Xml.BES.Task)     { return @{ Type="Task";     Node=$Xml.BES.Task } }
    if ($Xml.BES.Baseline) { return @{ Type="Baseline"; Node=$Xml.BES.Baseline } }
    throw "Unknown BES content type (no <Fixlet>, <Task>, or <Baseline>)."
}

# FIXED: use XPath + InnerText (no more System.Xml.XmlElement)
function Get-ActionAndRelevance { 
    param($ContainerNode)

    # --- Relevance (robust) ---
    $rels = @()

    # 1) Direct children named "Relevance"
    $direct = $ContainerNode.SelectNodes("./*[local-name()='Relevance']")
    if ($direct -and $direct.Count -gt 0) {
        foreach ($n in $direct) {
            $t = ($n.InnerText).Trim()
            if ($t) { $rels += $t }
        }
    }

    # 2) If none, search anywhere under the Fixlet/Task/Baseline node
    if ($rels.Count -eq 0) {
        $any = $ContainerNode.SelectNodes(".//*[local-name()='Relevance']")
        if ($any -and $any.Count -gt 0) {
            foreach ($n in $any) {
                $t = ($n.InnerText).Trim()
                if ($t) { $rels += $t }
            }
        }
    }

    # --- ActionScript (safe) ---
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

# =========================
# RELEVANCE EXTRACTION (GROUPS - HARDENED)
# =========================
function Extract-AllRelevanceFromXmlString {
    param(
        [string]$XmlString,
        [string]$Context = "Unknown"
    )

    $all = @()

    try {
        $x = [xml]$XmlString

        # 1) Under any ComputerGroup node, namespace-agnostic
        $cgRels = $x.SelectNodes("//*[local-name()='ComputerGroup']//*[local-name()='Relevance']")
        $cnt1 = if ($cgRels) { $cgRels.Count } else { 0 }
        LogLine "[$Context] Top-level ComputerGroup Relevance nodes: $cnt1"
        if ($cnt1 -gt 0) {
            foreach ($n in $cgRels) {
                $t = ($n.InnerText).Trim()
                if ($t) { $all += $t }
            }
        }

        # 2) If still empty, try global //Relevance anywhere
        if ($all.Count -eq 0) {
            $globalRels = $x.SelectNodes("//*[local-name()='Relevance']")
            $cnt2 = if ($globalRels) { $globalRels.Count } else { 0 }
            LogLine "[$Context] Global Relevance nodes: $cnt2"
            if ($cnt2 -gt 0) {
                foreach ($n in $globalRels) {
                    $t = ($n.InnerText).Trim()
                    if ($t) { $all += $t }
                }
            }
        }
    } catch {
        LogLine "[$Context] XML parse failed: $($_.Exception.Message)"
    }

    # 3) Regex fallback if requested and still nothing
    if ($AggressiveRegexFallback -and $all.Count -eq 0) {
        try {
            $regex = [regex]'(?is)<Relevance\b[^>]*>(.*?)</Relevance>'
            $m = $regex.Matches($XmlString)
            $cnt3 = $m.Count
            LogLine "[$Context] Regex-extracted Relevance nodes: $cnt3"
            if ($cnt3 -gt 0) {
                foreach ($mm in $m) {
                    $t = ($mm.Groups[1].Value).Trim()
                    if ($t) { $all += $t }
                }
            }
        } catch {
            LogLine "[$Context] Regex relevance fallback failed: $($_.Exception.Message)"
        }
    }

    return ,$all
}

function Extract-SCRFragments {
    param([string]$XmlString,[string]$Context="Unknown")
    $frags = @()
    try {
        $x = [xml]$XmlString
        $scrNodes = $x.SelectNodes("//*[local-name()='SearchComponentRelevance']")
        $cnt = if ($scrNodes) { $scrNodes.Count } else { 0 }
        LogLine "[$Context] SearchComponentRelevance nodes: $cnt"
        if ($cnt -gt 0) {
            foreach ($n in $scrNodes) {
                $innerR = $n.SelectNodes(".//*[local-name()='Relevance']")
                if ($innerR -and $innerR.Count -gt 0) {
                    foreach ($ir in $innerR) {
                        $t = ($ir.InnerText).Trim()
                        if ($t) { $frags += $t }
                    }
                } else {
                    $t = ($n.InnerText).Trim()
                    if ($t) { $frags += $t }
                }
            }
        }
    } catch {
        LogLine "[$Context] SCR parse failed: $($_.Exception.Message)"
    }
    return ,$frags
}

# Build client relevance for a group (tries multiple endpoints)
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
                Set-Content -Path $tmp -Value $xmlStr -Encoding UTF8
                LogLine "Saved fetched group XML to: $tmp"
            }

            $len = $xmlStr.Length
            $head = $xmlStr.Substring(0, [Math]::Min(800, $len)).Replace("`r"," ").Replace("`n"," ")
            $tail = if ($len -gt 240) { $xmlStr.Substring([Math]::Max(0, $len-240)).Replace("`r"," ").Replace("`n"," ") } else { "" }
            LogLine "Fetched group XML (${url}) len=$len head: $head"
            if ($tail) { LogLine "Fetched group XML tail: $tail" }

            # Try relevance extraction paths
            $rels = Extract-AllRelevanceFromXmlString -XmlString $xmlStr -Context "Group:$GroupIdNumeric"
            if ($rels.Count -gt 0) {
                $joined = ($rels | ForEach-Object { "($_)" }) -join " AND "
                $snippet = $joined.Substring(0, [Math]::Min(200, $joined.Length))
                LogLine "Using group relevance from <Relevance> nodes :: ${snippet}..."
                return $joined
            }

            # SCR fallback
            $frags = Extract-SCRFragments -XmlString $xmlStr -Context "Group:$GroupIdNumeric"
            if ($frags.Count -gt 0) {
                $joined = ($frags | ForEach-Object { "($_)" }) -join " AND "
                $snippet = $joined.Substring(0, [Math]::Min(200, $joined.Length))
                LogLine "Built relevance from SearchComponentRelevance :: ${snippet}..."
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
# SINGLE ACTION XML
# =========================
function Build-SingleActionXml {
    param(
        [string]$ActionTitle,            # Pilot/Deploy/Force/Conference...
        [string]$DisplayName,            # Vendor App Version
        [string[]]$RelevanceBlocks,      # Fixlet relevance + group relevance
        [string]$ActionScript,           # Action script
        [datetime]$StartLocal,           # scheduled local start (absolute)
        [bool]$IsForce = $false          # Force adds end offset (start+24h)
    )

    $titleText = "$($DisplayName): $ActionTitle"
    $titleEsc  = [System.Security.SecurityElement]::Escape($titleText)
    $dispEsc   = [System.Security.SecurityElement]::Escape($DisplayName)

    # Relevance (each in CDATA)
    $rels = ""
    if ($RelevanceBlocks -and $RelevanceBlocks.Count -gt 0) {
        $rels = ($RelevanceBlocks | ForEach-Object {
            $safe = $_ -replace ']]>', ']]]]><![CDATA[>'
            "    <Relevance><![CDATA[$safe]]></Relevance>"
        }) -join "`r`n"
    }

    # Durations from now
    $now = Get-Date
    $startTs = $StartLocal - $now
    $startOffset = To-IsoDuration $startTs

    $hasEnd = $false
    $endOffset = $null
    $endOffsetLine = ""
    if ($IsForce) {
        $endAbs = $StartLocal.AddHours(24)
        $endTs  = $endAbs - $now
        $endOffset = To-IsoDuration $endTs
        $hasEnd = $true
        $endOffsetLine = "      <EndDateTimeLocalOffset>$endOffset</EndDateTimeLocalOffset>`n"
    }

    # Only include deadline nodes when Force is true (avoid invalid 'None' enums)
    $deadlineBehaviorBlock = if ($IsForce) {
@"
        <DeadlineBehavior>RunAutomatically</DeadlineBehavior>
        <DeadlineType>Absolute</DeadlineType>
        <DeadlineLocalOffset>$endOffset</DeadlineLocalOffset>
"@
    } else { "" }

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

      <PreActionShowUI>true</PreActionShowUI>
      <PreAction>
        <Text>$dispEsc update will be enforced on $((Get-Date $StartLocal).ToString('M/d/yy h:mm tt')). Please save your work.</Text>
        <AskToSaveWork>true</AskToSaveWork>
        <ShowActionButton>false</ShowActionButton>
        <ShowCancelButton>false</ShowCancelButton>$deadlineBehaviorBlock        <ShowConfirmation>false</ShowConfirmation>
      </PreAction>

      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>Updating to $dispEsc...please wait.</Text></RunningMessage>

      <HasTimeRange>false</HasTimeRange>
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
    <Target>
      <AllComputers>true</AllComputers>
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
$form.Size = New-Object System.Drawing.Size(640, 780)
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
            Set-Content -Path $tmpFix -Value $fixletContent -Encoding UTF8
            LogLine "Saved fetched fixlet XML to: $tmpFix"
        }

        $fixLen  = $fixletContent.Length
        $fixHead = $fixletContent.Substring(0, [Math]::Min(800, $fixLen)).Replace("`r"," ").Replace("`n"," ")
        $fixTail = if ($fixLen -gt 240) { $fixletContent.Substring([Math]::Max(0, $fixLen-240)).Replace("`r"," ").Replace("`n"," ") } else { "" }
        LogLine "Fetched fixlet XML len=$fixLen head: $fixHead"
        if ($fixTail) { LogLine "Fetched fixlet XML tail: $fixTail" }

        $fixletXml = [xml]$fixletContent

        $cont = Get-FixletContainer -Xml $fixletXml
        LogLine ("Detected BES content type: {0}" -f $cont.Type)

        $titleRaw = $cont.Node.Title
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw

        $parsed = Get-ActionAndRelevance -ContainerNode $cont.Node
        $fixletRelevance = @()
        if ($parsed.Relevance) { $fixletRelevance = $parsed.Relevance }
        $actionScript = $parsed.ActionScript

        LogLine "Parsed title: ${displayName}"
        LogLine ("Fixlet relevance count: {0}" -f $fixletRelevance.Count)
        LogLine ("Action script length: {0}" -f $actionScript.Length)

        # Absolute schedule (user picks local date/time)
        $startLocal = Get-Date "$dStr $tStr"

        $actions = @("Pilot","Deploy","Force","Conference/Training Rooms")
        $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        LogLine "POST URL: ${postUrl}"

        foreach ($a in $actions) {
            $groupIdRaw = "$($GroupMap[$a])"
            if (-not $groupIdRaw) { LogLine "❌ Missing group id for $($a)"; continue }
            $groupIdNumeric = Get-NumericGroupId $groupIdRaw
            if (-not $groupIdNumeric) { LogLine "❌ Could not parse numeric ID from '${groupIdRaw}' for $($a)"; continue }

            try {
                $groupRel = Get-GroupClientRelevance -BaseUrl $base -AuthHeader $auth -SiteName $CustomSiteName -GroupIdNumeric $groupIdNumeric
                LogLine ("Group relevance len ({0}): {1}" -f $a, $groupRel.Length)
            } catch {
                LogLine "❌ Could not fetch/build group relevance for $($a): $($_.Exception.Message)"
                continue  # do NOT post without group relevance
            }

            $allRel = @()
            $allRel += $fixletRelevance
            if ($groupRel) { $allRel += $groupRel }

            $isForce = ($a -eq "Force")

            $xmlBody = Build-SingleActionXml `
                -ActionTitle $a `
                -DisplayName $displayName `
                -RelevanceBlocks $allRel `
                -ActionScript $actionScript `
                -StartLocal $startLocal `
                -IsForce:$isForce

            LogLine ("---- SingleAction XML for {0} ----" -f $a)
            LogLine $xmlBody

            try {
                HttpPostXml -Url $postUrl -AuthHeader $auth -XmlBody $xmlBody
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
