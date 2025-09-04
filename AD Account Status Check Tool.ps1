Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Ensure ActiveDirectory Module is Installed ---
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ActiveDirectory module not found. Attempting to install..." -ForegroundColor Yellow
    $osVersion = (Get-CimInstance Win32_OperatingSystem).Caption

    try {
        if ($osVersion -like "*Server*") {
            Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeAllSubFeature -ErrorAction Stop
        } else {
            Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
        }
        Write-Host "ActiveDirectory module installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install ActiveDirectory module. Please install RSAT manually." -ForegroundColor Red
        exit
    }
}

# --- Import Module ---
Import-Module ActiveDirectory

# --- Logging Setup with UK Date Format ---
$logPath = "$env:USERPROFILE\Desktop\AD_Status_Log.txt"
function Log-Message($message) {
    $timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    Add-Content -Path $logPath -Value "$timestamp - $message"
}

# --- Create Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Account Status Checker"
$form.Size = New-Object System.Drawing.Size(700,600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f2f2f2")  # Light grey

# --- Optional Logo Setup (Edit or remove path as needed) ---
$logoPath = "C:\Path\To\Your\Logo.png"  # Update or remove this path
if (Test-Path $logoPath) {
    try {
        $logo = New-Object System.Windows.Forms.PictureBox
        $logo.Image = [System.Drawing.Image]::FromFile($logoPath)
        $logo.SizeMode = "StretchImage"
        $logo.Size = New-Object System.Drawing.Size(200, 80)
        $logo.Location = New-Object System.Drawing.Point(250,10)
        $form.Controls.Add($logo)
    } catch {
        Write-Host "Logo could not be loaded. Continuing without it." -ForegroundColor Blue
    }
}

# --- Textbox for usernames ---
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$textBox.Size = New-Object System.Drawing.Size(300,120)
$textBox.Location = New-Object System.Drawing.Point(10,100)
$textBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($textBox)

# --- Generic Button Styling ---
function New-CustomButton($text, $x, $y) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Size = New-Object System.Drawing.Size(220,40)
    $btn.Location = New-Object System.Drawing.Point($x, $y)
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#004080")  # Custom blue
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    return $btn
}

$checkButton = New-CustomButton "Check Status" 320 100
$uploadButton = New-CustomButton "Upload File" 320 150
$clearButton  = New-CustomButton "Clear Results" 320 200

$form.Controls.AddRange(@($checkButton, $uploadButton, $clearButton))

# --- RichTextBox for Results ---
$resultsBox = New-Object System.Windows.Forms.RichTextBox
$resultsBox.ReadOnly = $true
$resultsBox.Size = New-Object System.Drawing.Size(660,180)
$resultsBox.Location = New-Object System.Drawing.Point(10,260)
$resultsBox.BackColor = [System.Drawing.Color]::White
$resultsBox.ForeColor = [System.Drawing.Color]::Black
$resultsBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($resultsBox)

# --- Footer ---
$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = "By Liam M, V1.0.0"
$footerLabel.AutoSize = $true
$footerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$footerLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#004080")
$footerLabel.Location = New-Object System.Drawing.Point(10,470)
$form.Controls.Add($footerLabel)

# --- File Upload Event ---
$uploadButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text Files (*.txt;*.csv)|*.txt;*.csv"
    $openFileDialog.Title = "Select Username File"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $fileContent = Get-Content $openFileDialog.FileName
        $textBox.Text = ($textBox.Text + "`n" + ($fileContent -join "`n")).Trim()
    }
})

# --- Clear Results Event ---
$clearButton.Add_Click({
    $resultsBox.Clear()
})

# --- Check Button Event ---
$checkButton.Add_Click({
    $resultsBox.Clear()
    $usernames = $textBox.Text -split "`n"
    foreach ($username in $usernames) {
        $username = $username.Trim()
        if ($username -eq "") { continue }

        try {
            $user = Get-ADUser -Identity $username -Properties Enabled
            if ($user.Enabled -eq $false) {
                $msg = "${username}: DISABLED"
                $resultsBox.SelectionColor = 'Red'
                $resultsBox.AppendText("$msg`n")
                Log-Message $msg
            } else {
                $msg = "${username}: ENABLED"
                $resultsBox.SelectionColor = 'Green'
                $resultsBox.AppendText("$msg`n")
                Log-Message $msg
            }
        } catch {
            $msg = "${username}: NOT FOUND or ERROR"
            $resultsBox.SelectionColor = 'Orange'
            $resultsBox.AppendText("$msg`n")
            Log-Message $msg
        }
    }
})

# --- Run the Form ---
[void]$form.ShowDialog()
