Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Generator"
$form.Size = New-Object System.Drawing.Size(700, 540)
$form.StartPosition = "CenterScreen"

# Server Info
$labelServer = New-Object System.Windows.Forms.Label -Property @{ Text = "Server:"; Location = New-Object System.Drawing.Point(10, 20); AutoSize = $true }
$textServer = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 18); Width = 550 }

$labelUser = New-Object System.Windows.Forms.Label -Property @{ Text = "Username:"; Location = New-Object System.Drawing.Point(10, 50); AutoSize = $true }
$textUser = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 48); Width = 550 }

$labelPass = New-Object System.Windows.Forms.Label -Property @{ Text = "Password:"; Location = New-Object System.Drawing.Point(10, 80); AutoSize = $true }
$textPass = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 78); Width = 550; UseSystemPasswordChar = $true }

# Fixlet Name
$labelFixlet = New-Object System.Windows.Forms.Label -Property @{ Text = "Fixlet Name:"; Location = New-Object System.Drawing.Point(10, 120); AutoSize = $true }
$textFixlet = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 118); Width = 550 }

# Fixlet ID
$labelFixletID = New-Object System.Windows.Forms.Label -Property @{ Text = "Fixlet ID:"; Location = New-Object System.Drawing.Point(10, 150); AutoSize = $true }
$textFixletID = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(100, 148); Width = 550 }

# Date Selector (future Wednesdays)
$labelDate = New-Object System.Windows.Forms.Label -Property @{ Text = "Select Wednesday:"; Location = New-Object System.Drawing.Point(10, 190); AutoSize = $true }
$comboDate = New-Object System.Windows.Forms.ComboBox -Property @{ Location = New-Object System.Drawing.Point(130, 188); Width = 200; DropDownStyle = "DropDownList" }

$today = [datetime]::Today
for ($i = 1; $i -le 60; $i++) {
    $day = $today.AddDays($i)
    if ($day.DayOfWeek -eq 'Wednesday') {
        $comboDate.Items.Add($day.ToString("yyyy-MM-dd"))
    }
}

# Time Selector
$labelTime = New-Object System.Windows.Forms.Label -Property @{ Text = "Select Time:"; Location = New-Object System.Drawing.Point(350, 190); AutoSize = $true }
$comboTime = New-Object System.Windows.Forms.ComboBox -Property @{ Location = New-Object System.Drawing.Point(440, 188); Width = 100; DropDownStyle = "DropDownList" }

foreach ($time in @(
    "8:00 PM", "8:15 PM", "8:30 PM", "8:45 PM",
    "9:00 PM", "9:15 PM", "9:30 PM", "9:45 PM",
    "10:00 PM", "10:15 PM", "10:30 PM", "10:45 PM",
    "11:00 PM", "11:15 PM", "11:30 PM", "11:45 PM"
)) {
    $comboTime.Items.Add($time)
}

# Generate Button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Create Actions"
$btnGenerate.Size = New-Object System.Drawing.Size(150, 30)
$btnGenerate.Location = New-Object System.Drawing.Point(270, 240)

# Add controls to form
$form.Controls.AddRange(@(
    $labelServer, $textServer,
    $labelUser, $textUser,
    $labelPass, $textPass,
    $labelFixlet, $textFixlet,
    $labelFixletID, $textFixletID,
    $labelDate, $comboDate,
    $labelTime, $comboTime,
    $btnGenerate
))

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
