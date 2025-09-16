# BigFix Offer Scheduler (API-based with live offers, offers only)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- BigFix API Config (change these) ---
$server = "https://bigfixserver:52311"
$username = "your-operator"
$password = "your-password"

# --- Helper: Call BigFix API ---
function Invoke-BigFixAPI {
    param(
        [string]$Method,
        [string]$Endpoint,
        [string]$Body = $null
    )
    $uri = "$server$Endpoint"
    $secpasswd = ConvertTo-SecureString $password -AsPlainText -Force
    $creds = New-Object System.Management.Automation.PSCredential ($username, $secpasswd)

    if ($Body) {
        Invoke-RestMethod -Uri $uri -Method $Method -Credential $creds -Body $Body -ContentType "application/xml"
    } else {
        Invoke-RestMethod -Uri $uri -Method $Method -Credential $creds
    }
}

# --- Query Offers with Category = Update ---
function Get-BigFixUpdateOffers {
    $relevance = "(id of it, name of it, site name of it) of relevant fixlets whose (category of it = ""Update"" and exists action of it whose (offer flag of it = true))"
    $encoded   = [System.Uri]::EscapeDataString($relevance)
    $endpoint  = "/api/query?relevance=$encoded"

    $xml = Invoke-BigFixAPI -Method Get -Endpoint $endpoint
    $offers = @()

    foreach ($result in $xml.Query.Result.Tuple) {
        $id   = $result.Answer[0].'#text'
        $name = $result.Answer[1].'#text'
        $site = $result.Answer[2].'#text'
        $offers += [PSCustomObject]@{
            FixletID = $id
            Name     = $name
            SiteName = $site
        }
    }
    return $offers
}

# --- Take Action on Fixlet ---
function Invoke-BigFixOffer {
    param(
        [string]$FixletID,
        [string]$SiteName,
        [string]$ComputerName,
        [datetime]$ScheduledTime = $null
    )

    $xml = @"
<BES>
  <SourcedFixletAction>
    <SourceFixlet>
      <Sitename>$SiteName</Sitename>
      <FixletID>$FixletID</FixletID>
      <Action>Action1</Action>
    </SourceFixlet>
    <Target>
      <ComputerName>$ComputerName</ComputerName>
    </Target>
"@

    if ($ScheduledTime) {
        $xml += @"
    <Settings>
      <StartDateTimeLocal>$($ScheduledTime.ToString("yyyyMMddTHHmmss"))</StartDateTimeLocal>
    </Settings>
"@
    }

    $xml += @"
  </SourcedFixletAction>
</BES>
"@

    Invoke-BigFixAPI -Method Post -Endpoint "/api/actions" -Body $xml
}

# --- GUI Helper: Date/Time Picker ---
function Show-DateTimePicker {
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Select Date & Time"
    $popup.Size = New-Object System.Drawing.Size(400,300)
    $popup.StartPosition = "CenterParent"

    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.MaxSelectionCount = 1
    $calendar.Location = New-Object System.Drawing.Point(10,10)
    $popup.Controls.Add($calendar)

    $timePicker = New-Object System.Windows.Forms.DateTimePicker
    $timePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
    $timePicker.ShowUpDown = $true
    $timePicker.Location = New-Object System.Drawing.Point(250,30)
    $popup.Controls.Add($timePicker)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(100,200)
    $btnOK.Add_Click({
        $popup.Tag = $calendar.SelectionStart.Date + $timePicker.Value.TimeOfDay
        $popup.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $popup.Close()
    })
    $popup.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(200,200)
    $btnCancel.Add_Click({
        $popup.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $popup.Close()
    })
    $popup.Controls.Add($btnCancel)

    $result = $popup.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $popup.Tag }
    return $null
}

# --- Main GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Updates"
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "Available 'Update' Offers:"
$label.Location = New-Object System.Drawing.Point(10,10)
$label.AutoSize = $true
$form.Controls.Add($label)

$listbox = New-Object System.Windows.Forms.ListBox
$listbox.Location = New-Object System.Drawing.Point(10,40)
$listbox.Size = New-Object System.Drawing.Size(550,200)
$listbox.SelectionMode = "MultiExtended"
$form.Controls.Add($listbox)

# Load offers from API
try {
    $offers = Get-BigFixUpdateOffers
    foreach ($offer in $offers) {
        $listbox.Items.Add("$($offer.FixletID) | $($offer.SiteName) | $($offer.Name)")
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to load offers: $_")
}

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = "Install Now"
$btnInstall.Location = New-Object System.Drawing.Point(10,260)
$btnInstall.Add_Click({
    foreach ($item in $listbox.SelectedItems) {
        $parts = $item.Split("|")
        $fixletId = $parts[0].Trim()
        $siteName = $parts[1].Trim()
        $comp = $env:COMPUTERNAME
        Invoke-BigFixOffer -FixletID $fixletId -SiteName $siteName -ComputerName $comp
    }
    [System.Windows.Forms.MessageBox]::Show("Actions triggered now.")
})
$form.Controls.Add($btnInstall)

$btnDefer = New-Object System.Windows.Forms.Button
$btnDefer.Text = "Defer..."
$btnDefer.Location = New-Object System.Drawing.Point(120,260)
$btnDefer.Add_Click({
    $when = Show-DateTimePicker
    if ($when) {
        foreach ($item in $listbox.SelectedItems) {
            $parts = $item.Split("|")
            $fixletId = $parts[0].Trim()
            $siteName = $parts[1].Trim()
            $comp = $env:COMPUTERNAME
            Invoke-BigFixOffer -FixletID $fixletId -SiteName $siteName -ComputerName $comp -ScheduledTime $when
        }
        [System.Windows.Forms.MessageBox]::Show("Deferred installs scheduled for $when.")
    }
})
$form.Controls.Add($btnDefer)

[void]$form.ShowDialog()
