Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# CONFIG
# =========================
$SiteName = "Test Group Managed (Workstations)"  # hardcoded custom site name

# Map each action to its Computer Group ID (keep as strings to preserve 00- prefix).
$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12345"
    "Force"                     = "00-12345"
    "Conference/Training Rooms" = "00-12345"
}

# Toggle this to $true ONLY if your test server uses a self-signed/invalid cert
$BypassCertValidation = $false

# =========================
# NETWORK/TLS HARDENING
# =========================
try {
    # Prefer TLS 1.2; optionally include 1.1/1.0 if your server still needs it
    [System.Net.ServicePointManager]::SecurityProtocol =
        [System.Net.SecurityProtocolType]::Tls12 `
        -bor [System.Net.SecurityProtocolType]::Tls11 `
        -bor [System.Net.SecurityProtocolType]::Tls

    # Avoid 100-Continue stalls
    [System.Net.ServicePointManager]::Expect100Continue = $false

    # Use system proxy creds if any
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
} catch {}

if ($BypassCertValidation) {
    if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
        Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) { return true; }
}
"@
    }
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

# =========================
# ENCODING / URL / AUTH HELPERS
# =========================
Add-Type -AssemblyName System.Web
function Encode-SiteName {
    param([string]$Name)
    # Encode, then normalize space and parentheses
    $enc = [System.Web.HttpUtility]::UrlEncode($Name, [System.Text.Encoding]::UTF8)
    $enc = $enc -replace '\+','%20' -replace '\(','%28' -replace '\)','%29'
    return $enc
}

function Get-BaseUrl {
    param([string]$ServerInput)
    if (-not $ServerInput) { throw "Server is empty." }
    $s = $ServerInput.Trim()
    if ($s -match '^(?i)https?://') { return ($s.TrimEnd('/')) }
    $s = $s.Trim('/')
    if ($s -match ':\d+$') { "https://$s" } else { "https://$s:52311" }
}

function Join-ApiUrl {
    param([string]$BaseUrl,[string]$RelativePath)
    $rp = if ($RelativePath.StartsWith("/")) { $RelativePath } else { "/$RelativePath" }
    $BaseUrl.TrimEnd('/') + $rp
}

function Get-AuthHeader {
    param([string]$Username, [string]$Password)
    $pair  = "$Username`:$Password"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    "Basic " + [Convert]::ToBase64String($bytes)
}

# =========================
# BIGFIX HELPERS
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

    try {
        $resp = Invoke-WebRequest -Uri $url -Headers @{ Authorization=$auth } -UseBasicParsing -ErrorAction Stop
        [pscustomobject]@{ Url = $url; Content = $resp.Content }
    } catch {
        throw ("GET failed: " + ($_.Exception.GetBaseException().Message))
    }
}

function Parse-FixletTitleToProduct {
    param([string]$Title)
    # "Update: Vendor AppName Version Win" -> "Vendor AppName Version"
    if ($Title -match "^Update:\s*(.+?)\s+Win$") { return $matches[1] }
    return $Title
}

function Get-NextWednesdays {
    $dates = @()
    $today = Get-Date
    # Wednesday = 3 (Sunday=0)
    $daysUntilWed = (3 - [int]$today.DayOfWeek + 7) % 7
    $nextWed = $today.AddDays($daysUntilWed)
    for ($i = 0; $i -lt 20; $i++) { $dates += $nextWed.AddDays(7*$i).ToString("yyyy-MM-dd") }
    return $dates
}

function Get-TimeSlots {
    $slots = @()
    $start = Get-Date "20:00"   # 8:00 PM
    $end   = Get-Date "23:45"   # 11:45 PM
    while ($start -le $end) {
        $slots += $start.ToString("h:mm tt")
        $start = $start.AddMinutes(15)
    }
    return $slots
}

function Format-LocalBESDateTime {
    param([datetime]$dt)
    # Local (no Z) yyyymmddTHHMMSS format
    return $dt.ToString("yyyyMMdd'T'HHmmss")
}

function Build-SingleActionXml {
    param(
        [string]$ActionTitle,          # "Pilot" | "Deploy" | "Force" | "Conference/Training Rooms"
        [string]$DisplayName,          # "Vendor AppName Version"
        [string[]]$RelevanceBlocks,    # relevance strings
        [string]$ActionScript,         # action script text
        [datetime]$StartLocal,         # scheduled local start
        [bool]$SetDeadline = $false,   # only true for "Force"
        [datetime]$DeadlineLocal = $null,
        [string]$GroupId               # "00-12345"
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

function Post-ActionXml {
    param(
        [string]$Server,
        [string]$Username,
        [string]$Password,
        [string]$XmlBody
    )
    $base = Get-BaseUrl $Server
    $url  = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
    $auth = Get-AuthHeader -Username $Username -Password $Password
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($XmlBody)

    try {
        Invoke-RestMethod -Uri $url -Method Post -Headers @{
            Authorization = $auth
            "Content-Type" = "application/xml"
        } -Body $bodyBytes -ErrorAction Stop
        return $url
    } catch {
        throw ("POST failed: " + ($_.Exception.GetBaseException().Message))
    }
}

# =========================
# GUI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(560, 780)
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

# TLS bypass toggle (so you can switch without editing script)
$sslChk = New-Object System.Windows.Forms.CheckBox
$sslChk.Text = "Bypass SSL certificate validation (unsafe)"
$sslChk.Checked = $BypassCertValidation
$sslChk.Location = New-Object System.Drawing.Point(10, $y)
$sslChk.Size = New-Object System.Drawing.Size(360, 24)
$form.Controls.Add($sslChk)
$y += 34

# Log box
$log = New-Object System.Windows.Forms.TextBox
$log.Multiline = $true
$log.ScrollBars = "Vertical"
$log.ReadOnly = $false
$log.WordWrap = $false
$log.Location = New-Object System.Drawing.Point(10, $y)
# Slightly taller now that we added the checkbox
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

    # If user checked the bypass box at runtime, enable it now
    if ($sslChk.Checked -and -not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
        Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) { return true; }
}
"@
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        $append.Invoke("⚠️ SSL certificate validation is DISABLED for this session.")
    }

    try {
        # Build and log the fully encoded Fixlet GET URL before calling it
        $base        = Get-BaseUrl $server
        $encodedSite = Encode-SiteName $SiteName
        $fixletPath  = "/api/fixlet/custom/$encodedSite/$fixletId"
        $fixletUrl   = Join-ApiUrl -BaseUrl $base -RelativePath $fixletPath

        $append.Invoke(("Server base URL: {0}" -f $base))
        $append.Invoke(("Encoded Fixlet GET URL: {0}" -f $fixletUrl))

        # Now call the API
        try {
            $resp = Get-FixletDetails -Server $server -Username $user -Password $pass -FixletID $fixletId
        } catch {
            $append.Invoke(("❌ TLS/Send error on GET: {0}" -f ($_.Exception.GetBaseException().Message)))
            throw
        }
        $append.Invoke(("GET URL (from func): {0}" -f $resp.Url))

        $fixletXml = $resp.Content
        $xml = [xml]$fixletXml

        $titleRaw = $xml.BES.Fixlet.Title
        $displayName = Parse-FixletTitleToProduct -Title $titleRaw

        # Collect relevance
        $relevance = @()
        foreach ($rel in $xml.BES.Fixlet.Relevance) { $relevance += [string]$rel }

        # First Action's ActionScript
        $actionNode = $xml.BES.Fixlet.Action | Select-Object -First 1
        if (-not $actionNode) { throw "No <Action> block found in Fixlet." }
        $actionScript = [string]$actionNode.ActionScript.'#text'

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
                -ActionTitle $a `
                -DisplayName $displayName `
                -RelevanceBlocks $relevance `
                -ActionScript $actionScript `
                -StartLocal $startLocal `
                -SetDeadline:$isForce `
                -DeadlineLocal $deadlineLocal `
                -GroupId $groupId

            $append.Invoke(("---- XML for {0} ----" -f $a))
            $append.Invoke($xmlBody)

            try {
                $postBase = Get-BaseUrl $server
                $postUrl  = Join-ApiUrl -BaseUrl $postBase -RelativePath "/api/actions"
                $append.Invoke(("Encoded POST URL: {0}" -f $postUrl))

                try {
                    $postedUrl = Post-ActionXml -Server $server -Username $user -Password $pass -XmlBody $xmlBody
                    $append.Invoke(("✅ {0} created successfully." -f $a))
                } catch {
                    $append.Invoke(("❌ TLS/Send error on POST: {0}" -f ($_.Exception.GetBaseException().Message)))
                }
            } catch {
                $append.Invoke(("❌ Failed to create {0}: {1}" -f $a, $_))
            }
        }

        $append.Invoke(("All actions attempted. See log: {0}" -f $logFile))
    }
    catch {
        $append.Invoke(("❌ Fatal error: {0}" -f ($_.Exception.GetBaseException().Message)))
    }
})

$form.Topmost = $false
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
