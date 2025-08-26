Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG
# =========================
$LogFile = "C:\temp\BigFixActionGenerator.log"
$CustomSiteName = "Test Group Managed (Workstations)"

$GroupMap = @{
    "Pilot"                     = "00-12345"
    "Deploy"                    = "00-12345"
    "Force"                     = "00-12345"
    "Conference/Training Rooms" = "00-12345"
}
$FixletActionNameMap = @{
    "Pilot"                     = "Action1"
    "Deploy"                    = "Action1"
    "Force"                     = "Action1"
    "Conference/Training Rooms" = "Action1"
}

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
function Get-NextWeekday([datetime]$base,[System.DayOfWeek]$weekday) {
    $delta = ([int]$weekday - [int]$base.DayOfWeek + 7) % 7
    if ($delta -le 0) { $delta += 7 }
    return $base.Date.AddDays($delta)
}
function SafeEscape([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Security.SecurityElement]::Escape($s)
}

# (HTTP + Fixlet/Group parsing functions remain unchanged here …)

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

# (rest of the GUI setup, buttons, logging text box, and action logic unchanged …)

$form.Topmost = $false
[void]$form.ShowDialog()
