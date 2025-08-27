
# Import the new configuration module
Import-Module "$PSScriptRoot\ZDAConfiguration.psm1"

# Start transcript for configuration
Start-TranscriptForApp -name "configure"

# Read user configuration
$user_config = Get-UserConfiguration

$taskName = "Zoomdownloader"


# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing



# Global configuration storage initialization
if (-not $global:Config) { $global:Config = @{} }
if (-not $global:Config.Zoom) { $global:Config.Zoom = @{} }
if (-not $global:Config.Database) { $global:Config.Database = @{} }
if (-not $global:Config.Storage) { $global:Config.Storage = @{} }
if (-not $global:Config.Schedule) { $global:Config.Schedule = @{} }

# Current page tracking
$script:currentPage = 1

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zoom Downloader Setup Wizard"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.TopMost = $true

# Focus event
$form.Add_Shown({
    $form.TopMost = $false
    $form.Activate()
    $form.Focus()
})

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 20)
$progressBar.Size = New-Object System.Drawing.Size(540, 20)
$progressBar.Minimum = 1
$progressBar.Maximum = 7  # Changed to 6 pages total
$progressBar.Value = 1
$form.Controls.Add($progressBar)

# Page title
$lblPageTitle = New-Object System.Windows.Forms.Label
$lblPageTitle.Location = New-Object System.Drawing.Point(20, 50)
$lblPageTitle.Size = New-Object System.Drawing.Size(540, 30)
$lblPageTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
$lblPageTitle.Text = "Welcome"
$form.Controls.Add($lblPageTitle)

# Content panel
$pnlContent = New-Object System.Windows.Forms.Panel
$pnlContent.Location = New-Object System.Drawing.Point(20, 90)
$pnlContent.Size = New-Object System.Drawing.Size(640, 400)
$pnlContent.BorderStyle = "FixedSingle"
$form.Controls.Add($pnlContent)

# Buttons
$btnPrevious = New-Object System.Windows.Forms.Button
$btnPrevious.Text = "< Previous"
$btnPrevious.Location = New-Object System.Drawing.Point(290, 520)
$btnPrevious.Size = New-Object System.Drawing.Size(80, 30)
$btnPrevious.Enabled = $false
$form.Controls.Add($btnPrevious)

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Next >"
$btnNext.Location = New-Object System.Drawing.Point(380, 520)
$btnNext.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($btnNext)

$btnFinish = New-Object System.Windows.Forms.Button
$btnFinish.Text = "Finish"
$btnFinish.Location = New-Object System.Drawing.Point(470, 520)
$btnFinish.Size = New-Object System.Drawing.Size(80, 30)
$btnFinish.Visible = $false
$form.Controls.Add($btnFinish)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(200, 520)
$btnCancel.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($btnCancel)

# Page titles
$pageTitles = @{
    1 = "Welcome to the Zoom Downloader Setup Wizard"
    2 = "Zoom Credentials"
    3 = "Storage Selection"
    4 = "Database Configuration"
    5 = "Schedule Task"
    6 = "Accounts To Download"
    7 = "Setup Summary"
}

# Function to create Welcome page
function Create-WelcomePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(640, 200)
    
    $lblWelcome = New-Object System.Windows.Forms.Label
    $lblWelcome.Text = "Welcome to the Zoom Downloader Setup Wizard!"
    $lblWelcome.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 14, [System.Drawing.FontStyle]::Bold)
    $lblWelcome.Location = New-Object System.Drawing.Point(50, 50)
    $lblWelcome.Size = New-Object System.Drawing.Size(440, 40)
    $panel.Controls.Add($lblWelcome)
    
    $lblDescription = New-Object System.Windows.Forms.Label
    $lblDescription.Text = "This wizard will guide you through setting up your Zoom downloader.`r`n`r`nYou will need to provide:`r`n* Zoom account credentials`r`n* Choose storage location (Local, OneDrive, or AWS S3)`r`n* Database connection details`r`n`r`nClick Next to continue."
    $lblDescription.Location = New-Object System.Drawing.Point(50, 100)
    $lblDescription.Size = New-Object System.Drawing.Size(440, 150)
    $lblDescription.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
    $panel.Controls.Add($lblDescription)
    
    return $panel
}

# Function to create Zoom Credentials page
function Create-ZoomCredentialsPage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(640, 300)
    
    # API Key
    $lblApiKey = New-Object System.Windows.Forms.Label
    $lblApiKey.Text = "Zoom Client ID:"
    $lblApiKey.Location = New-Object System.Drawing.Point(30, 30)
    $lblApiKey.Size = New-Object System.Drawing.Size(150, 20)
    $panel.Controls.Add($lblApiKey)
    
    $script:txtApiKey = New-Object System.Windows.Forms.TextBox
    $script:txtApiKey.Location = New-Object System.Drawing.Point(200, 30)
    $script:txtApiKey.Size = New-Object System.Drawing.Size(300, 20)
    $script:txtApiKey.Name = "txtApiKey"
    $script:txtApiKey.text = $user_config.zoom.clientId
    $panel.Controls.Add($script:txtApiKey)
    
    # API Secret
    $lblApiSecret = New-Object System.Windows.Forms.Label
    $lblApiSecret.Text = "Zoom Client Secret:"
    $lblApiSecret.Location = New-Object System.Drawing.Point(30, 70)
    $lblApiSecret.Size = New-Object System.Drawing.Size(150, 20)
    $panel.Controls.Add($lblApiSecret)
    
    $script:txtApiSecret = New-Object System.Windows.Forms.TextBox
    $script:txtApiSecret.Location = New-Object System.Drawing.Point(200, 70)
    $script:txtApiSecret.Size = New-Object System.Drawing.Size(300, 20)
    $script:txtApiSecret.UseSystemPasswordChar = $true
    $script:txtApiSecret.Name = "txtApiSecret"
    $script:txtApiSecret.text = $user_config.zoom.clientSecret
    $panel.Controls.Add($script:txtApiSecret)
    
    # Account ID
    $lblAccountId = New-Object System.Windows.Forms.Label
    $lblAccountId.Text = "Account ID:"
    $lblAccountId.Location = New-Object System.Drawing.Point(30, 110)
    $lblAccountId.Size = New-Object System.Drawing.Size(150, 20)
    $panel.Controls.Add($lblAccountId)
    
    $script:txtAccountId = New-Object System.Windows.Forms.TextBox
    $script:txtAccountId.Location = New-Object System.Drawing.Point(200, 110)
    $script:txtAccountId.Size = New-Object System.Drawing.Size(300, 20)
    $script:txtAccountId.Name = "txtAccountId"
    $script:txtAccountId.text = $user_config.zoom.accountId
    $panel.Controls.Add($script:txtAccountId)

    # Status label (add before button)
    $script:lblTestStatus = New-Object System.Windows.Forms.Label
    $script:lblTestStatus.Location = New-Object System.Drawing.Point(30, 5)
    $script:lblTestStatus.Size = New-Object System.Drawing.Size(500, 40)
    $script:lblTestStatus.Name = "lblTestStatus"
    $script:lblTestStatus.Text = ""
    $panel.Controls.Add($script:lblTestStatus)

    # Test Connection Button
    $script:btnTestZoom = New-Object System.Windows.Forms.Button
    $script:btnTestZoom.Text = "Test Connection"
    $script:btnTestZoom.Enabled = $false
    $script:btnTestZoom.Location = New-Object System.Drawing.Point(140, 180)
    $script:btnTestZoom.Size = New-Object System.Drawing.Size(100, 25)
    $panel.Controls.Add($script:btnTestZoom)

    # Enable Test Connection only if all fields are filled
    $checkFields = {
        if ($script:txtApiKey.Text -and $script:txtApiSecret.Text -and $script:txtAccountId.Text) {
            $script:btnTestZoom.Enabled = $true
        } else {
            $script:btnTestZoom.Enabled = $false
        }
    }
    $script:txtApiKey.Add_TextChanged($checkFields)
    $script:txtApiSecret.Add_TextChanged($checkFields)
    $script:txtAccountId.Add_TextChanged($checkFields)
    
    # Status label
    $script:lblTestStatus = New-Object System.Windows.Forms.Label
    $script:lblTestStatus.Location = New-Object System.Drawing.Point(140, 135)
    $script:lblTestStatus.Size = New-Object System.Drawing.Size(500, 40)
    $script:lblTestStatus.Name = "lblTestStatus"
    $script:lblTestStatus.Text = ""
    $panel.Controls.Add($script:lblTestStatus)
    
    # Test button click event
    $script:btnTestZoom.Add_Click({
        $apiKey = $script:txtApiKey.Text.Trim()
        $apiSecret = $script:txtApiSecret.Text.Trim()
        $accountId = $script:txtAccountId.Text.Trim()

        $script:lblTestStatus.Text = "Testing connection..."
        $script:lblTestStatus.ForeColor = [System.Drawing.Color]::Blue
        $script:btnTestZoom.Enabled = $false
        $form.Update()

        try {
            $result = Test-ZoomConnection -ApiKey $apiKey -ApiSecret $apiSecret -AccountId $accountId

            if ($result.Success) {
                $script:lblTestStatus.Text = "SUCCESS: Connection successful!"
                $script:lblTestStatus.ForeColor = [System.Drawing.Color]::Green
            } else {
                $script:lblTestStatus.Text = "ERROR: $($result.ErrorMessage)"
                $script:lblTestStatus.ForeColor = [System.Drawing.Color]::Red
            }
        }
        catch {
            $script:lblTestStatus.Text = "ERROR: Connection test failed"
            $script:lblTestStatus.ForeColor = [System.Drawing.Color]::Red
        }
        finally {
            $script:btnTestZoom.Enabled = $true
        }
    })
    
    $checkFields.Invoke()
    return $panel
}

# Function to create Storage Selection page
function Create-StorageSelectionPage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(640, 370)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Choose where to store downloaded recordings:"
    $lblTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(30, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(400, 20)
    $panel.Controls.Add($lblTitle)
    
    # Storage type radio buttons
    $radioLocal = New-Object System.Windows.Forms.RadioButton
    $radioLocal.Text = "Local Storage Only"
    $radioLocal.Location = New-Object System.Drawing.Point(50, 60)
    $radioLocal.Size = New-Object System.Drawing.Size(400, 20)
    $radioLocal.Checked = $true
    $radioLocal.Name = "radioLocal"
    $panel.Controls.Add($radioLocal)

    # Download path label and note
    $lblLocalPath = New-Object System.Windows.Forms.Label
    $lblLocalPath.Text = "Download Path: %AppData%\\Local\\ZoomDownloader"
    $lblLocalPath.Location = New-Object System.Drawing.Point(50, 85)
    $lblLocalPath.Size = New-Object System.Drawing.Size(500, 20)  # Increase width to prevent wrapping
    $lblLocalPath.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblLocalPath.AutoSize = $false
    $panel.Controls.Add($lblLocalPath)

    $lblLocalNote = New-Object System.Windows.Forms.Label
    $lblLocalNote.Text = "(This location is hardcoded and cannot be changed)"
    $lblLocalNote.Location = New-Object System.Drawing.Point(50, 105)
    $lblLocalNote.Size = New-Object System.Drawing.Size(350, 18)
    $lblLocalNote.ForeColor = [System.Drawing.Color]::Gray
    $panel.Controls.Add($lblLocalNote)

    # Move other radio buttons down
    $radioOneDrive = New-Object System.Windows.Forms.RadioButton
    $radioOneDrive.Text = "Microsoft OneDrive"
    $radioOneDrive.Location = New-Object System.Drawing.Point(50, 130)
    $radioOneDrive.Size = New-Object System.Drawing.Size(200, 20)
    $radioOneDrive.Name = "radioOneDrive"
    $panel.Controls.Add($radioOneDrive)

    $radioS3 = New-Object System.Windows.Forms.RadioButton
    $radioS3.Text = "Amazon S3"
    $radioS3.Location = New-Object System.Drawing.Point(50, 160)
    $radioS3.Size = New-Object System.Drawing.Size(200, 20)
    $radioS3.Name = "radioS3"
    $panel.Controls.Add($radioS3)
    
    # Configuration panels for each storage type
    
    # Local Storage Panel
    $pnlLocal = New-Object System.Windows.Forms.Panel
    $pnlLocal.Location = New-Object System.Drawing.Point(70, 150)
    $pnlLocal.Size = New-Object System.Drawing.Size(450, 60)
    $pnlLocal.Name = "pnlLocal"
    $pnlLocal.Visible = $true  # Always visible
    $panel.Controls.Add($pnlLocal)
    
    # Move lblLocalPath and lblLocalNote underneath the Local Storage Only radio button
    $lblLocalPath = New-Object System.Windows.Forms.Label
    $lblLocalPath.Text = "Download Path: %AppData%\\Local\\ZoomDownloader"
    $lblLocalPath.Location = New-Object System.Drawing.Point(50, 85)
    $lblLocalPath.Size = New-Object System.Drawing.Size(500, 20)  # Increase width to prevent wrapping
    $lblLocalPath.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblLocalPath.AutoSize = $false
    $panel.Controls.Add($lblLocalPath)

    $lblLocalNote = New-Object System.Windows.Forms.Label
    $lblLocalNote.Text = "(This location is hardcoded and cannot be changed)"
    $lblLocalNote.Location = New-Object System.Drawing.Point(50, 105)
    $lblLocalNote.Size = New-Object System.Drawing.Size(350, 18)
    $lblLocalNote.ForeColor = [System.Drawing.Color]::Gray
    $panel.Controls.Add($lblLocalNote)

    # Remove from pnlLocal
    $pnlLocal.Visible = $true  # Ensure always visible
    
    # OneDrive Panel
    $pnlOneDrive = New-Object System.Windows.Forms.Panel
    $pnlOneDrive.Location = New-Object System.Drawing.Point(70, 190)
    $pnlOneDrive.Size = New-Object System.Drawing.Size(600, 180)
    $pnlOneDrive.Visible = $false
    $pnlOneDrive.Name = "pnlOneDrive"
    $panel.Controls.Add($pnlOneDrive)
    
    # Remove OneDrive folder field
    # Add OneDrive credential fields
    $lblAppId = New-Object System.Windows.Forms.Label
    $lblAppId.Text = "App ID:"
    $lblAppId.Location = New-Object System.Drawing.Point(0, 10)
    $lblAppId.Size = New-Object System.Drawing.Size(100, 20)
    $pnlOneDrive.Controls.Add($lblAppId)

    $script:txtAppId = New-Object System.Windows.Forms.TextBox
    $script:txtAppId.Location = New-Object System.Drawing.Point(110, 10)
    $script:txtAppId.Size = New-Object System.Drawing.Size(250, 20)
    $script:txtAppId.Name = "txtAppId"
    $pnlOneDrive.Controls.Add($script:txtAppId)

    $lblClientSecret = New-Object System.Windows.Forms.Label
    $lblClientSecret.Text = "Client Secret:"
    $lblClientSecret.Location = New-Object System.Drawing.Point(0, 40)
    $lblClientSecret.Size = New-Object System.Drawing.Size(100, 20)
    $pnlOneDrive.Controls.Add($lblClientSecret)

    $script:txtClientSecret = New-Object System.Windows.Forms.TextBox
    $script:txtClientSecret.Location = New-Object System.Drawing.Point(110, 40)
    $script:txtClientSecret.Size = New-Object System.Drawing.Size(250, 20)
    $script:txtClientSecret.UseSystemPasswordChar = $true
    $script:txtClientSecret.Name = "txtClientSecret"
    $pnlOneDrive.Controls.Add($script:txtClientSecret)

    $script:lblTenantName = New-Object System.Windows.Forms.Label
    $script:lblTenantName.Text = "Tenant Name:"
    $script:lblTenantName.Location = New-Object System.Drawing.Point(0, 70)
    $script:lblTenantName.Size = New-Object System.Drawing.Size(100, 20)
    $pnlOneDrive.Controls.Add($lblTenantName)

    $script:txtTenantName = New-Object System.Windows.Forms.TextBox
    $script:txtTenantName.Location = New-Object System.Drawing.Point(110, 70)
    $script:txtTenantName.Size = New-Object System.Drawing.Size(250, 20)
    $script:txtTenantName.Name = "txtTenantName"
    $pnlOneDrive.Controls.Add($script:txtTenantName)
    
    # Add Test Connection button to OneDrive panel
    $script:lblOneDriveStatus = New-Object System.Windows.Forms.Label
    $script:lblOneDriveStatus.Name = "lblOneDriveStatus"
    $script:lblOneDriveStatus.Location = New-Object System.Drawing.Point(0, 95)
    $script:lblOneDriveStatus.Size = New-Object System.Drawing.Size(400, 35)
    $script:lblOneDriveStatus.Text = ""
    $pnlOneDrive.Controls.Add($script:lblOneDriveStatus)

    $script:btnTestOneDrive = New-Object System.Windows.Forms.Button
    $script:btnTestOneDrive.Text = "Test Connection"
    $script:btnTestOneDrive.Location = New-Object System.Drawing.Point(0, 125)
    $script:btnTestOneDrive.Size = New-Object System.Drawing.Size(120, 30)
    $script:btnTestOneDrive.Name = "btnTestOneDrive"
    $script:btnTestOneDrive.Enabled = $false

    $pnlOneDrive.Controls.Add($script:btnTestOneDrive)
    # Enable Test Connection only if all OneDrive fields are filled
    $checkOneDriveFields = {
        if ($txtAppId.Text -and $txtClientSecret.Text -and $txtTenantName.Text) {
            $btnTestOneDrive.Enabled = $true
        } else {
            $btnTestOneDrive.Enabled = $false
        }
    }
    $script:txtAppId.Add_TextChanged($checkOneDriveFields)
    $script:txtClientSecret.Add_TextChanged($checkOneDriveFields)
    $script:txtTenantName.Add_TextChanged($checkOneDriveFields)

    # S3 Panel
    $pnlS3 = New-Object System.Windows.Forms.Panel
    $pnlS3.Location = New-Object System.Drawing.Point(70, 190)
    $pnlS3.Size = New-Object System.Drawing.Size(500, 400)
    $pnlS3.Visible = $false
    $pnlS3.Name = "pnlS3"
    $panel.Controls.Add($pnlS3)
    
    $lblAccessKey = New-Object System.Windows.Forms.Label
    $lblAccessKey.Text = "Access Key ID:"
    $lblAccessKey.Location = New-Object System.Drawing.Point(0, 10)
    $lblAccessKey.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblAccessKey)
    
    $script:txtAccessKey = New-Object System.Windows.Forms.TextBox
    $script:txtAccessKey.Location = New-Object System.Drawing.Point(110, 10)
    $script:txtAccessKey.Size = New-Object System.Drawing.Size(250, 20)
    $script:txtAccessKey.Name = "txtAccessKey"
    $pnlS3.Controls.Add($script:txtAccessKey)
    
    $lblSecretKey = New-Object System.Windows.Forms.Label
    $lblSecretKey.Text = "Secret Key:"
    $lblSecretKey.Location = New-Object System.Drawing.Point(0, 40)
    $lblSecretKey.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblSecretKey)
    
    $script:txtSecretKey = New-Object System.Windows.Forms.TextBox
    $script:txtSecretKey.Location = New-Object System.Drawing.Point(110, 40)
    $script:txtSecretKey.Size = New-Object System.Drawing.Size(250, 20)
    $script:txtSecretKey.UseSystemPasswordChar = $true
    $script:txtSecretKey.Name = "txtSecretKey"
    $pnlS3.Controls.Add($script:txtSecretKey)
    
    $lblBucket = New-Object System.Windows.Forms.Label
    $lblBucket.Text = "Bucket Name:"
    $lblBucket.Location = New-Object System.Drawing.Point(0, 70)
    $lblBucket.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblBucket)
    
    $script:txtBucket = New-Object System.Windows.Forms.TextBox
    $script:txtBucket.Location = New-Object System.Drawing.Point(110, 70)
    $script:txtBucket.Size = New-Object System.Drawing.Size(250, 20)
    $script:txtBucket.Name = "txtBucket"
    $pnlS3.Controls.Add($script:txtBucket)
    
    $lblRegion = New-Object System.Windows.Forms.Label
    $lblRegion.Text = "Region:"
    $lblRegion.Location = New-Object System.Drawing.Point(0, 100)
    $lblRegion.Size = New-Object System.Drawing.Size(100, 25)
    $pnlS3.Controls.Add($lblRegion)
    
    $script:cmbRegion = New-Object System.Windows.Forms.ComboBox
    $script:cmbRegion.Items.AddRange(@("us-west-2", "us-east-1", "us-west-1", "eu-west-1", "ap-southeast-1"))
    $script:cmbRegion.Location = New-Object System.Drawing.Point(110, 100)
    $script:cmbRegion.Size = New-Object System.Drawing.Size(200, 30)
    $script:cmbRegion.DropDownStyle = "DropDownList"
    $script:cmbRegion.Name = "cmbRegion"
    $pnlS3.Controls.Add($script:cmbRegion)  
    
    $script:lblS3Status = New-Object System.Windows.Forms.Label
    $script:lblS3Status.Name = "lblS3Status"
    $script:lblS3Status.Location = New-Object System.Drawing.Point(0, 130)
    $script:lblS3Status.Size = New-Object System.Drawing.Size(400, 20)
    $script:lblS3Status.Text = ""
    $pnlS3.Controls.Add($script:lblS3Status)

    $script:btnTestS3 = New-Object System.Windows.Forms.Button
    $script:btnTestS3.Text = "Test Connection"
    $script:btnTestS3.Location = New-Object System.Drawing.Point(0, 150)
    $script:btnTestS3.Size = New-Object System.Drawing.Size(120, 30)
    $script:btnTestS3.Name = "btnTestS3"
    $script:btnTestS3.Enabled = $false
    $script:btnTestS3.Add_Click({
        $this.Enabled = $false
        $script:lbls3Status.Text = "Testing S3 connection..."
        $script:lblS3Status.ForeColor = [System.Drawing.Color]::Blue
        $form.Update()
        
        try {                        
            # Test 1: Verify credentials are valid
            Write-Host "1. Checking credential validity..." -ForegroundColor Cyan
            Get-STSCallerIdentity -AccessKey $script:txtAccessKey.Text.Trim() -SecretKey $script:txtSecretKey.Text.Trim()


            Write-Host "INFO: S3 Authentication successful."
            $script:lblS3Status.Text = "S3 Authentication successful"
            $script:lblS3Status.ForeColor = [System.Drawing.Color]::Green

            Write-Host "S3 connection test successful."
            
            $script:lblS3Status.Text = "SUCCESS: Connection successful!"
            $script:lblS3Status.ForeColor = [System.Drawing.Color]::Green   
        }
        catch {
            $script:lblS3Status.Text = "ERROR: Connection test failed., $($_.Exception.Message)"
            $script:lblS3Status.ForeColor = [System.Drawing.Color]::Red
        }
        finally {
            $script:btnTestS3.Enabled = $true    
        }
    })
    $pnlS3.Controls.Add($script:btnTestS3)

    $checkS3Fields = {
        if ($script:txtAccessKey.Text -and $script:txtSecretKey.Text -and $script:txtBucket.Text) {
            $script:btnTestS3.Enabled = $true
        } else {
            $script:btnTestS3.Enabled = $false
        }
    }
    $script:txtAccessKey.Add_TextChanged($checkS3Fields)
    $script:txtSecretKey.Add_TextChanged($checkS3Fields)
    $script:txtBucket.Add_TextChanged($checkS3Fields)

    
    $radioOneDrive.Add_CheckedChanged({
        $parentPanel = $this.Parent
        $parentPanel.Controls["pnlOneDrive"].Visible = $this.Checked
        if ($this.Checked) { $parentPanel.Controls["pnlOneDrive"].BringToFront() }
    })
    $radioS3.Add_CheckedChanged({
        $parentPanel = $this.Parent
        $parentPanel.Controls["pnlS3"].Visible = $this.Checked
        if ($this.Checked) { $parentPanel.Controls["pnlS3"].BringToFront() }
    })

    if($user_config.upload.provider -eq "S3") {
        $radioS3.Checked = $true
        $pnlS3.Visible = $true
        $pnlLocal.Visible = $false
        $pnlOneDrive.Visible = $false
        $script:txtAccessKey.Text = $user_config.upload.s3.accessKeyId
        $script:txtSecretKey.Text = $user_config.upload.s3.secretAccessKey
        $script:txtBucket.Text = $user_config.upload.s3.bucketName
        $script:cmbRegion.SelectedItem = $user_config.upload.s3.region
    } elseif($user_config.upload.provider -eq "OneDrive") {
        $radioOneDrive.Checked = $true
        $pnlOneDrive.Visible = $true
        $pnlLocal.Visible = $false
        $pnlS3.Visible = $false
        $script:txtAppId.Text = $user_config.upload.onedrive.clientId
        $script:txtClientSecret.Text = $user_config.upload.onedrive.clientSecret
        $script:txtTenantName.Text = $user_config.upload.onedrive.tenantId
    } else {
        # Default to Local Storage
        $radioLocal.Checked = $true
        $pnlLocal.Visible = $true
        $pnlOneDrive.Visible = $false
        $pnlS3.Visible = $false
    }   
    
    return $panel
}

# Function to create Database Configuration page
function Create-DatabasePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(640, 400)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Database Configuration:"
    $lblTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(30, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(300, 20)
    $panel.Controls.Add($lblTitle)
    
    # Database Type
    $lblDbType = New-Object System.Windows.Forms.Label
    $lblDbType.Text = "Database Type:"
    $lblDbType.Location = New-Object System.Drawing.Point(30, 60)
    $lblDbType.Size = New-Object System.Drawing.Size(90, 20)
    $panel.Controls.Add($lblDbType)



    $lblDbTypeValue = New-Object System.Windows.Forms.Label
    $lblDbTypeValue.Text = "SQL Server"
    $lblDbTypeValue.Location = New-Object System.Drawing.Point(130, 60)
    $lblDbTypeValue.Size = New-Object System.Drawing.Size(150, 20)
    $lblDbTypeValue.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Regular)
    $lblDbTypeValue.Name = "lblDbTypeValue"
    $panel.Controls.Add($lblDbTypeValue)
    
    #$panel.Controls.Add($cmbDatabaseType)
    
    # Server
    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Text = "Server:"
    $lblServer.Location = New-Object System.Drawing.Point(30, 100)
    $lblServer.Size = New-Object System.Drawing.Size(100, 20)

    $panel.Controls.Add($lblServer)
    
    $script:txtServer = New-Object System.Windows.Forms.TextBox
    $script:txtServer.Location = New-Object System.Drawing.Point(130, 100)
    $script:txtServer.Size = New-Object System.Drawing.Size(200, 20)
    $script:txtServer.Text = $user_config.database.server
    $script:txtServer.Name = "txtServer"
    $panel.Controls.Add($script:txtServer)
    

    # Port
    $lblPort = New-Object System.Windows.Forms.Label
    $lblPort.Text = "Port:"
    $lblPort.Location = New-Object System.Drawing.Point(30, 140)
    $lblPort.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblPort)

    $script:txtPort = New-Object System.Windows.Forms.TextBox
    $script:txtPort.Location = New-Object System.Drawing.Point(130, 140)
    $script:txtPort.Size = New-Object System.Drawing.Size(200, 20)
    $script:txtPort.Name = "txtPort"
    $script:txtPort.Text = $user_config.database.port
    $panel.Controls.Add($script:txtPort)

    # Database Name
    $lblDatabase = New-Object System.Windows.Forms.Label
    $lblDatabase.Text = "Database:"
    $lblDatabase.Location = New-Object System.Drawing.Point(30, 180)
    $lblDatabase.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblDatabase)
    
    $script:txtDatabase = New-Object System.Windows.Forms.TextBox
    $script:txtDatabase.Location = New-Object System.Drawing.Point(130, 180)
    $script:txtDatabase.Size = New-Object System.Drawing.Size(200, 20)
    $script:txtDatabase.Name = "txtDatabase"
    $script:txtDatabase.Text = $user_config.database.database
    $panel.Controls.Add($script:txtDatabase)

    # Schema Name
    $lblSchema = New-Object System.Windows.Forms.Label
    $lblSchema.Text = "Schema:"
    $lblSchema.Location = New-Object System.Drawing.Point(30, 220)
    $lblSchema.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblSchema)
    
    $script:txtSchema = New-Object System.Windows.Forms.TextBox
    $script:txtSchema.Location = New-Object System.Drawing.Point(130, 220)
    $script:txtSchema.Size = New-Object System.Drawing.Size(200, 20)
    $script:txtSchema.Name = "txtSchema"
    $script:txtSchema.Text = $user_config.database.schema
    $panel.Controls.Add($script:txtSchema)
    
    # Username
    $lblUsername = New-Object System.Windows.Forms.Label
    $lblUsername.Text = "Username:"
    $lblUsername.Location = New-Object System.Drawing.Point(30, 269)
    $lblUsername.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblUsername)
    
    $script:txtUsername = New-Object System.Windows.Forms.TextBox
    $script:txtUsername.Location = New-Object System.Drawing.Point(130, 260)
    $script:txtUsername.Size = New-Object System.Drawing.Size(200, 20)
    $script:txtUsername.Name = "txtUsername"
    $script:txtUsername.Text = $user_config.database.userid
    $panel.Controls.Add($script:txtUsername)
    
    # Password
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Text = "Password:"
    $lblPassword.Location = New-Object System.Drawing.Point(30, 300)
    $lblPassword.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblPassword)
    
    $script:txtPassword = New-Object System.Windows.Forms.TextBox
    $script:txtPassword.Location = New-Object System.Drawing.Point(130, 300)
    $script:txtPassword.Size = New-Object System.Drawing.Size(200, 20)
    $script:txtPassword.UseSystemPasswordChar = $true
    $script:txtPassword.Name = "txtPassword"
    $script:txtPassword.Text = $user_config.database.password
    $panel.Controls.Add($script:txtPassword)

    $script:lblDbStatus = New-Object System.Windows.Forms.Label
    $script:lblDbStatus.Name = "lblDbStatus"
    $script:lblDbStatus.Location = New-Object System.Drawing.Point(30, 330)
    $script:lblDbStatus.Size = New-Object System.Drawing.Size(600, 30)
    $script:lblDbStatus.Text = ""
    $panel.Controls.Add($script:lblDbStatus)

    # Test Connection Button
    $script:btnTestDb = New-Object System.Windows.Forms.Button
    $script:btnTestDb.Text = "Test Connection"
    $script:btnTestDb.Location = New-Object System.Drawing.Point(30, 360)
    $script:btnTestDb.Size = New-Object System.Drawing.Size(120, 30)
    $script:btnTestDb.Name = "btnTestDb"
    $script:btnTestDb.Enabled = $false
    $script:btnTestDb.Add_Click({ 
        $this.Enabled = $false
        $script:lblDbStatus.Text = "Testing Database connection..."
        $script:lblDbStatus.ForeColor = [System.Drawing.Color]::Blue
        $form.Update()

        $sqlserver = @{
            server = $script:txtServer.text
            port = $script:txtPort.text
            database = $script:txtDatabase.text
            schema = $script:txtSchema.text
            userid = $script:txtUsername.text
            password = $script:txtPassword.text
        }

        $ConnectionString = "Server=$($sqlserver.server),$($sqlserver.port);Database=$($sqlserver.database);User ID=$($sqlserver.userid);Password=$($sqlserver.password);TrustServerCertificate=true"
        Write-Host "Testing Connection String $ConnectionString"
        try {
            $SQLServerConnection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
            $SQLServerConnection.Open()
            $SQLServerConnection.Close()
            $script:lblDbStatus.Text = "SUCCESS: SQLServer Credentials Worked."
            $script:lblDbStatus.ForeColor = [System.Drawing.Color]::Green

        }
        catch {
            $script:lblDbStatus.Text = "FAILURE: SQLServer Credentials did not work."
            Write-Host "ERROR: $($_.Exception.Message)"
            $script:lblDbStatus.ForeColor = [System.Drawing.Color]::Red
        } 
        finally {
            $this.Enabled = $true
        }   
    })
    $panel.Controls.Add($script:btnTestDb)

    # Enable Test Connection only if all DB fields are filled
    $checkDbFields = {
        if ($script:txtServer.Text -and $script:txtDatabase.Text -and $script:txtUsername.Text -and $script:txtPassword.Text) {
            $script:btnTestDb.Enabled = $true
        } else {
            $script:btnTestDb.Enabled = $false
        }
    }
    $script:txtServer.Add_TextChanged($checkDbFields)
    $script:txtDatabase.Add_TextChanged($checkDbFields)
    $script:txtUsername.Add_TextChanged($checkDbFields)
    $script:txtPassword.Add_TextChanged($checkDbFields)
    
    $checkDbFields.Invoke()
    return $panel
}

# Function to create Summary page
function Create-SummaryPage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    
    $lblSummary = New-Object System.Windows.Forms.Label
    $lblSummary.Text = "Setup Summary - Review your configuration:"
    $lblSummary.Location = New-Object System.Drawing.Point(30, 30)
    $lblSummary.Size = New-Object System.Drawing.Size(400, 20)
    $lblSummary.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $panel.Controls.Add($lblSummary)
    
    $txtSummary = New-Object System.Windows.Forms.TextBox
    $txtSummary.Location = New-Object System.Drawing.Point(30, 60)
    $txtSummary.Size = New-Object System.Drawing.Size(480, 300)
    $txtSummary.Multiline = $true
    $txtSummary.ScrollBars = "Vertical"
    $txtSummary.ReadOnly = $true
    $txtSummary.Name = "txtSummary"
    $panel.Controls.Add($txtSummary)
    
    return $panel
}

# Function to create Schedule page
function Create-SchedulePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(640, 300)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Schedule PowerShell Task:"
    $lblTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(30, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(400, 20)
    $panel.Controls.Add($lblTitle)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Choose when the Zoom Downloader should run. Select a default schedule or enter a custom PowerShell schedule expression."
    $lblDesc.Location = New-Object System.Drawing.Point(30, 50)
    $lblDesc.Size = New-Object System.Drawing.Size(580, 40)
    $panel.Controls.Add($lblDesc)

    $lblSchedule = New-Object System.Windows.Forms.Label
    $lblSchedule.Text = "Schedule:"
    $lblSchedule.Location = New-Object System.Drawing.Point(30, 100)
    $lblSchedule.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblSchedule)

    $script:cmbSchedule = New-Object System.Windows.Forms.ComboBox
    $script:cmbSchedule.Items.AddRange(@('Every Day at Midnight','Every Hour','Every 5 Minutes','Immediate Single Run', 'Custom...'))
    $script:cmbSchedule.Location = New-Object System.Drawing.Point(140, 100)
    $script:cmbSchedule.Size = New-Object System.Drawing.Size(200, 25)
    $script:cmbSchedule.DropDownStyle = "DropDownList"
    $script:cmbSchedule.Name = "cmbSchedule"
    $script:cmbSchedule.SelectedItem = $user_config.schedule.schedule
    $panel.Controls.Add($script:cmbSchedule)

    $script:lblCustom = New-Object System.Windows.Forms.Label
    $script:lblCustom.Text = "Custom PowerShell Schedule Expression:"
    $script:lblCustom.Location = New-Object System.Drawing.Point(50, 140)
    $script:lblCustom.Size = New-Object System.Drawing.Size(300, 20)
    $script:lblCustom.Visible = $false
    $panel.Controls.Add($script:lblCustom)

    $script:txtCustom = New-Object System.Windows.Forms.TextBox
    $script:txtCustom.Location = New-Object System.Drawing.Point(50, 165)
    $script:txtCustom.Size = New-Object System.Drawing.Size(400, 20)
    $script:txtCustom.Multiline = $true
    $script:txtCustom.Name = "txtCustom"
    $script:txtCustom.Visible = $false
    $script:txtCustom.Text = $user_config.schedule.custom
    $panel.Controls.Add($script:txtCustom)

    if($script:cmbSchedule.SelectedItem -ne "Custom...") {
        $script:lblCustom.Visible = $false
        $script:txtCustom.Visible = $false
    } else {
        $script:lblCustom.Visible = $true
        $script:txtCustom.Visible = $true
    }

    $script:cmbSchedule.Add_SelectedIndexChanged({
        if ($this.SelectedItem -eq "Custom...") {
            $script:lblCustom.Visible = $true
            $script:txtCustom.Visible = $true
        } else {
            $script:lblCustom.Visible = $false
            $script:txtCustom.Visible = $false
        }
    })

        # Add range dropdown
        $lblRange = New-Object System.Windows.Forms.Label
        $lblRange.Text = "Download Range:"
        $lblRange.Location = New-Object System.Drawing.Point(30, 210)
        $lblRange.Size = New-Object System.Drawing.Size(100, 20)
        $panel.Controls.Add($lblRange)

        $script:cmbRange = New-Object System.Windows.Forms.ComboBox
        $script:cmbRange.Items.AddRange(@(
            "Last 2 weeks",
            "Last month",
            "Custom start date..."
        ))
        $script:cmbRange.Location = New-Object System.Drawing.Point(140, 210)
        $script:cmbRange.Size = New-Object System.Drawing.Size(180, 20)
        $script:cmbRange.DropDownStyle = "DropDownList"
        $script:cmbRange.Name = "cmbRange"
        $script:cmbRange.SelectedItem = $user_config.schedule.dateRange
        $panel.Controls.Add($script:cmbRange)

        $script:lblStartDate = New-Object System.Windows.Forms.Label
        $script:lblStartDate.Text = "Start Date:"
        $script:lblStartDate.Location = New-Object System.Drawing.Point(30, 240)
        $script:lblStartDate.Size = New-Object System.Drawing.Size(100, 20)
   
        $panel.Controls.Add($script:lblStartDate)

        $script:dtpStartDate = New-Object System.Windows.Forms.DateTimePicker
        $script:dtpStartDate.Location = New-Object System.Drawing.Point(140, 240)
        $script:dtpStartDate.Size = New-Object System.Drawing.Size(180, 25)
        $script:dtpStartDate.Format = "Short"
        $script:dtpStartDate.Visible = $false
        $script:dtpStartDate.Name = "dtpStartDate"
        if ($user_config.schedule.customFromDate) {
            $script:dtpStartDate.Value = [DateTime]::Parse($user_config.schedule.customFromDate)
        } 
        
        $panel.Controls.Add($script:dtpStartDate)

        if($script:cmbRange.SelectedItem -ne "Custom start date...") {
            $script:lblStartDate.Visible = $false
            $script:dtpStartDate.Visible = $false
        } else {
            $script:lblStartDate.Visible = $true
            $script:dtpStartDate.Visible = $true
        }

        $script:cmbRange.Add_SelectedIndexChanged({
            if ($this.SelectedItem -eq "Custom start date...") {
                $script:lblStartDate.Visible = $true
                $script:dtpStartDate.Visible = $true
            } else {
                $script:lblStartDate.Visible = $false
                $script:dtpStartDate.Visible = $false
            }
        })
    return $panel
}

# Function to create Accounts To Download page
function Create-AccountsToDownload {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(640, 300)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Accounts to Download:"
    $lblTitle.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(30, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(400, 20)
    $panel.Controls.Add($lblTitle)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Enter one email address per line. These accounts will be downloaded."
    $lblDesc.Location = New-Object System.Drawing.Point(30, 50)
    $lblDesc.Size = New-Object System.Drawing.Size(580, 30)
    $panel.Controls.Add($lblDesc)

    $script:txtAccounts = New-Object System.Windows.Forms.TextBox
    $script:txtAccounts.Location = New-Object System.Drawing.Point(30, 90)
    $script:txtAccounts.Size = New-Object System.Drawing.Size(580, 150)
    $script:txtAccounts.Multiline = $true
    $script:txtAccounts.ScrollBars = "Vertical"
    $script:txtAccounts.Name = "txtAccounts"
    $script:txtAccounts.Text = $user_config.accounts -join "`r`n"
    $panel.Controls.Add($script:txtAccounts)

    return $panel
}

# Page controls array
$pageControls = @(
    (Create-WelcomePage),
    (Create-ZoomCredentialsPage),
    (Create-StorageSelectionPage),
    (Create-DatabasePage),
    (Create-SchedulePage),
    (Create-AccountsToDownload),
    (Create-SummaryPage)
)

# Function to show current page
function Show-CurrentPage {
    $pnlContent.Controls.Clear()
    $pnlContent.Controls.Add($pageControls[$script:currentPage - 1])
    $lblPageTitle.Text = $pageTitles[$script:currentPage]
    $progressBar.Value = $script:currentPage
    
    # Update button states
    $btnPrevious.Enabled = $script:currentPage -gt 1
    $btnNext.Visible = $script:currentPage -lt 7
    $btnFinish.Visible = $script:currentPage -eq 7

    # Update summary if on last page
    if ($script:currentPage -eq 6) {
        Update-Summary
    }
}

# Function to validate current page
function Validate-CurrentPage {
    switch ($script:currentPage) {
        2 { # Zoom credentials
            $currentPanel = $pageControls[1]
            $apiKey = $currentPanel.Controls["txtApiKey"].Text.Trim()
            $apiSecret = $currentPanel.Controls["txtApiSecret"].Text.Trim()
            $accountId = $currentPanel.Controls["txtAccountId"].Text.Trim()
            
            if ([string]::IsNullOrWhiteSpace($apiKey)) {
                [System.Windows.Forms.MessageBox]::Show("API Key is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($apiSecret)) {
                [System.Windows.Forms.MessageBox]::Show("API Secret is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($accountId)) {
                [System.Windows.Forms.MessageBox]::Show("Account ID is required.", "Validation Error")
                return $false
            }
            
            if (-not $global:Config.Zoom) { $global:Config.Zoom = @{} }
            $global:Config.Zoom.ApiKey = $apiKey
            $global:Config.Zoom.ApiSecret = $apiSecret
            $global:Config.Zoom.AccountId = $accountId
        }
        3 { # Storage selection
            $currentPanel = $pageControls[2]
            $radioLocal = $currentPanel.Controls["radioLocal"]
            $radioOneDrive = $currentPanel.Controls["radioOneDrive"]
            $radioS3 = $currentPanel.Controls["radioS3"]
            
            if ($radioLocal.Checked) {
                $global:Config.Storage.Type = "Local"
            }
            elseif ($radioOneDrive.Checked) {
                $global:Config.Storage.Type = "OneDrive"
                $global:Config.Storage.AppId = $currentPanel.Controls["pnlOneDrive"].Controls["txtAppId"].Text.Trim()
                $global:Config.Storage.ClientSecret = $currentPanel.Controls["pnlOneDrive"].Controls["txtClientSecret"].Text.Trim()
                $global:Config.Storage.TenantName = $currentPanel.Controls["pnlOneDrive"].Controls["txtTenantName"].Text.Trim()
            }
            elseif ($radioS3.Checked) {
                $pnlS3 = $currentPanel.Controls["pnlS3"]
                $accessKey = $pnlS3.Controls["txtAccessKey"].Text.Trim()
                $secretKey = $pnlS3.Controls["txtSecretKey"].Text.Trim()
                $bucket = $pnlS3.Controls["txtBucket"].Text.Trim()
                $region = $pnlS3.Controls["cmbRegion"].SelectedItem
                
                if ([string]::IsNullOrWhiteSpace($accessKey)) {
                    [System.Windows.Forms.MessageBox]::Show("S3 Access Key is required.", "Validation Error")
                    return $false
                }
                if ([string]::IsNullOrWhiteSpace($secretKey)) {
                    [System.Windows.Forms.MessageBox]::Show("S3 Secret Key is required.", "Validation Error")
                    return $false
                }
                if ([string]::IsNullOrWhiteSpace($bucket)) {
                    [System.Windows.Forms.MessageBox]::Show("S3 Bucket name is required.", "Validation Error")
                    return $false
                }
                if ($null -eq $region) {
                    [System.Windows.Forms.MessageBox]::Show("S3 Region is required.", "Validation Error")
                    return $false
                }
                
                $global:Config.Storage.Type = "S3"
                $global:Config.Storage.AccessKey = $accessKey
                $global:Config.Storage.SecretAccessKey = $secretKey
                $global:Config.Storage.Bucket = $bucket
                $global:Config.Storage.Region = $region
            }
        }
        4 { # Database
            $currentPanel = $pageControls[3]
            $server = $currentPanel.Controls["txtServer"].Text.Trim()
            $port = $currentPanel.Controls["txtPort"].Text.Trim()
            $database = $currentPanel.Controls["txtDatabase"].Text.Trim()
            $schema = $currentPanel.Controls["txtSchema"].Text.Trim()
            $username = $currentPanel.Controls["txtUsername"].Text.Trim()
            $password = $currentPanel.Controls["txtPassword"].Text.Trim()
 
            if ([string]::IsNullOrWhiteSpace($server)) {
                [System.Windows.Forms.MessageBox]::Show("Server is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($port)) {
                [System.Windows.Forms.MessageBox]::Show("Port is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($database)) {
                [System.Windows.Forms.MessageBox]::Show("Database name is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($schema)) {
                [System.Windows.Forms.MessageBox]::Show("Schema is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($username)) {
                [System.Windows.Forms.MessageBox]::Show("Username is required.", "Validation Error")
                return $false
            }            
            if ([string]::IsNullOrWhiteSpace($password)) {
                [System.Windows.Forms.MessageBox]::Show("Password is required.", "Validation Error")
                return $false
            }

            $global:Config.Database.Server = $server
            $global:Config.Database.Port = $port
            $global:Config.Database.Database = $database
            $global:Config.Database.Schema = $schema
            $global:Config.Database.Username = $username
            $global:Config.Database.Password = $password
        }
        5 { # Schedule
            $currentPanel = $pageControls[4]
            $schedule = $script:cmbSchedule.SelectedItem.Trim()
            $dateRange = $script:cmbRange.SelectedItem.Trim()
            
            $global:Config.Schedule.schedule = $schedule
            $global:Config.Schedule.dateRange = $dateRange

            if ([string]::IsNullOrWhiteSpace($schedule)) {
                [System.Windows.Forms.MessageBox]::Show("Schedule is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($dateRange)) {
                [System.Windows.Forms.MessageBox]::Show("Date Range is required.", "Validation Error")
                return $false
            }   

            if( $schedule -eq "Custom..." ) {
                $customSchedule = $script:txtCustom.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($customSchedule)) {
                    [System.Windows.Forms.MessageBox]::Show("Custom schedule expression is required.", "Validation Error")
                    return $false
                }
                $global:Config.Schedule.custom = $customSchedule
            } else {
                $global:Config.Schedule.custom = $null
            }

            if( $dateRange -eq "Custom start date..." ) {
                $customFromDate = $script:dtpStartDate.Value.ToString("yyyy-MM-dd")
                if ([string]::IsNullOrWhiteSpace($customFromDate)) {
                    [System.Windows.Forms.MessageBox]::Show("Custom start date is required.", "Validation Error")
                    return $false
                }
                $global:Config.Schedule.customFromDate = $customFromDate
            } else {
                $global:Config.Schedule.customFromDate = $null
            }   
        }
        6 { # Accounts to Download
            $currentPanel = $pageControls[5]
            $accountsText = $currentPanel.Controls["txtAccounts"].Text.Trim()
            $global:Config.Accounts = $accountsText -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        }
    }
    return $true
}

# Function to update summary
function Update-Summary {

    
    $summaryText = @"
Zoom Configuration:
  API Key: $($global:Config.Zoom.ApiKey)
  Account ID: $($global:Config.Zoom.AccountId)

Storage Configuration:
  Type: $(if ($global:Config.Storage.Type -eq "Local") { "Local Storage" } else { $global:Config.Storage.Type })
  $(if ($global:Config.Storage.Type -eq "OneDrive") {
    "App ID: $($global:Config.Storage.AppId)"
    "$([Environment]::NewLine)"
    "Tenant Name: $($global:Config.Storage.TenantName)"
  } elseif ($global:Config.Storage.Type -eq "S3") {
    "Access Key: $($global:Config.Storage.AccessKey)"
    "$([Environment]::NewLine)"
    "Bucket Name: $($global:Config.Storage.Bucket)"
    "$([Environment]::NewLine)"
    "Region: $($global:Config.Storage.Region)"
  })

Database Configuration:
  Type: SQL Server
  Server: $($global:Config.Database.Server)
  Port: $($global:Config.Database.Port)
  Database: $($global:Config.Database.Database)
  Schema: $($global:Config.Database.Schema)
  Username: $($global:Config.Database.Username)

Schedule Configuration:
  Schedule: $($global:Config.Schedule.schedule)
  $(if ($global:Config.Schedule.schedule -eq "Custom...") {
    "   Custom Schedule: $($global:Config.Schedule.custom)"
  })
  Date Range: $($global:Config.Schedule.dateRange)
  $(if ($global:Config.Schedule.dateRange -eq "Custom start date...") {
    "   Custom start date: $($global:Config.Schedule.customFromDate)"
  })

Accounts to Download:
$(($global:Config.Accounts | ForEach-Object { "    $_" }) -join "`r`n")

Click Finish to complete the setup.
"@
    
    $currentPanel = $pageControls[6]

    $txtSummary = $currentPanel.Controls | Where-Object { $_.Name -eq "txtSummary" }
    if ($txtSummary) { $txtSummary.Text = $summaryText }    
    }

# Event handlers
$btnNext.Add_Click({
    if (Validate-CurrentPage) {
        $script:currentPage++
        Show-CurrentPage
    }
})

$btnPrevious.Add_Click({
    $script:currentPage--
    Show-CurrentPage
})


$btnFinish.Add_Click({
    if (Validate-CurrentPage) {
        # Save configuration to file
        $configJson = $global:Config | ConvertTo-Json -Depth 3

        $zoomConfig = @{
            accountId    = $global:Config.Zoom.AccountId
            clientId     = $global:Config.Zoom.ApiKey
            clientSecret = $global:Config.Zoom.ApiSecret
        }

        $type = $global:Config.Storage.Type
        if( $type -eq 'Local' ) {
            $storageConfig = @{
                provider = $global:Config.Storage.Type
            }             
        } elseif( $type -eq 'OneDrive' ) {
            $storageConfig = @{
                provider = $global:Config.Storage.Type
                onedrive = @{
                    clientId = $global:Config.Storage.AppId
                    clientSecret = $global:Config.Storage.ClientSecret
                    tenantId = $global:Config.Storage.TenantName
                }
            }             
        } else {
            $storageConfig = @{
                provider = $global:Config.Storage.Type
                s3 = @{
                    accessKeyId = $global:Config.Storage.AccessKey
                    secretAccessKey = $global:Config.Storage.SecretAccessKey
                    bucketName = $global:Config.Storage.Bucket
                    region = $global:Config.Storage.Region
                    multipartThreshold = 100000000
                    maxConcurrency = 10
                }
            }               
        }
  
        $databaseConfig = @{
            server = $global:Config.Database.Server
            port = $global:Config.Database.Port
            database = $global:Config.Database.Database
            schema = $global:Config.Database.Schema
            userid = $global:Config.Database.Username
            password = $global:Config.Database.Password
            connectionString = "Server=$($global:Config.Database.Server),$($global:Config.Database.Port);Database=$($global:Config.Database.Database);User ID=$($global:Config.Database.Username);Password=$($global:Config.Database.Password);TrustServerCertificate=true"
            tableName = "recordings.ZoomRecordings"
            idColumn = "GUID"
        }

        $schedule = @{
            schedule = $global:Config.Schedule.schedule
            custom = $global:Config.Schedule.custom
            dateRange = $global:Config.Schedule.dateRange
            customFromDate = $global:Config.Schedule.customFromDate
        }

        $accounts = $global:Config.Accounts -join "`r`n"

        $runspaces = @{
            batchSize = 200
            maxThreads = 25
            maxRecordsPerThread = 5000
            uploadDelayMs = 100
        }

        $downloads = @{
            basepath = Get-DownloadsDirectoryPath
        }

        $config = @{
            zoom     = $zoomConfig
            upload  = $storageConfig
            database = $databaseConfig
            schedule = $schedule
            accounts = $accounts
            runspaces = $runspaces
            download = $downloads
        }


        Add-LocalAppDataFolder
        $jsonString = $config | ConvertTo-Json
        Save-UserConfiguration -Json $jsonString

        CreateDatabase -ConnectionString $databaseConfig.ConnectionString
        $taskScheduled = scheduleOn
Write-Host "The results of scheduling the task was: $taskScheduled"
        if ($taskScheduled -eq $true) {
            [System.Windows.Forms.MessageBox]::Show("Configuration saved and Job Scheduled", "Setup Complete")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Configuration saved, however, Job NOT Scheduled", "Setup Complete")
        }
        

        $form.DialogResult = "OK"
        $form.Close()
    }
})

function CreateDatabase {
    param(
        [string]$ConnectionString
    )
    Write-Host "Ensuring tables exist in SQL Server database"
    $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)    
    $connection.Open()
    $table = $global:Config.Database.Schema + ".ZoomRecordings"
    $mainTableSQL = @"
    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '$($global:Config.Database.Schema)' AND TABLE_NAME = 'ZoomRecordings')
    CREATE TABLE $table (
        GUID NVARCHAR(255) NOT NULL,
        HOST_EMAIL NVARCHAR(255),
        RECORDING_START NVARCHAR(255),
        RECORDING_END NVARCHAR(255),
        FILE_SIZE NVARCHAR(255),
        DOWNLOAD_URL NVARCHAR(255),
        MEETING_ID NVARCHAR(255),
        TOPIC NVARCHAR(255),
        RECORDING_TYPE NVARCHAR(255),
        DOWNLOADED BIT,
        TRYDLAGAIN INT,
        DOWNLOAD_PATH NVARCHAR(255),
        UPLOADED BIT,
        UPLOAD_PATH NVARCHAR(255),
        UPLOAD_COMPLETED DATETIME2 NULL,
    )
"@
    #Write-Host("Query:$mainTableSQL")
    $command = $connection.CreateCommand()
    $command.CommandText = $mainTableSQL
    $command.ExecuteNonQuery()

    $table = $global:Config.Database.Schema + ".ZoomRecordings"
    $createIndexSQL = @"
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Uploaded' AND object_id = OBJECT_ID('$table'))
    CREATE INDEX IX_Uploaded 
    ON $table (Uploaded)
"@
    #Write-Host("Query:$createIndexSQL")
    $command.CommandText = $createIndexSQL
    $command.ExecuteNonQuery()
    $command.Dispose()
    $connection.Close()
}

function Get-CurrentUserInfo {
    $username = $env:USERNAME
    $computername = $env:COMPUTERNAME
    $userdomain = $env:USERDOMAIN
    $dnsdomain = $env:USERDNSDOMAIN
    
    # Check if user is domain user
    if ($userdomain -ne $computername) {
        # Domain user
        $isDomainUser = $true
        $fullUsername = "$userdomain\$username"
        
        # Alternative UPN format if available
        if ($dnsdomain) {
            $upnUsername = "$username@$dnsdomain"
        } else {
            $upnUsername = $fullUsername
        }
    } else {
        # Local user
        $isDomainUser = $false
        $fullUsername = "$computername\$username"
        $upnUsername = $fullUsername
    }
    
    return @{
        IsDomainUser = $isDomainUser
        Username = $username
        Domain = $userdomain
        ComputerName = $computername
        FullUsername = $fullUsername
        UPNUsername = $upnUsername
    }
}

function Test-UserCredentials {
    param(
        [string]$Username,
        [string]$Password
    )
    
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        
        # Determine if this is actually a domain user or local user
        # If username contains \ and domain part equals computer name, it's local
        if ($Username.Contains('\')) {
            $domain = $Username.Split('\')[0]
            $user = $Username.Split('\')[1]
            
            # Check if domain part is actually the computer name (local user)
            if ($domain -eq $env:COMPUTERNAME) {
                Write-Host "DEBUG: Detected local user '$user' (domain part matches computer name)"
                
                $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Machine
                $principalContext = $null
                
                try {
                    $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType)
                    $isValid = $principalContext.ValidateCredentials($user, $Password)
                    Write-Host "DEBUG: Local user validation result: $isValid"
                    return $isValid
                }
                finally {
                    if ($principalContext) { $principalContext.Dispose() }
                }
            } else {
                # True domain user
                Write-Host "DEBUG: Validating domain user '$user' in domain '$domain'"
                
                $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $principalContext = $null
                
                try {
                    $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType, $domain)
                    $isValid = $principalContext.ValidateCredentials($user, $Password)
                    Write-Host "DEBUG: Domain validation result: $isValid"
                    return $isValid
                }
                catch {
                    Write-Host "DEBUG: Domain validation failed, might be local user with domain format: $($_.Exception.Message)"
                    # Fallback: try as local user
                    $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Machine
                    $principalContext2 = $null
                    try {
                        $principalContext2 = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType)
                        $isValid = $principalContext2.ValidateCredentials($user, $Password)
                        Write-Host "DEBUG: Fallback local validation result: $isValid"
                        return $isValid
                    }
                    finally {
                        if ($principalContext2) { $principalContext2.Dispose() }
                    }
                }
                finally {
                    if ($principalContext) { $principalContext.Dispose() }
                }
            }
        } 
        else {
            Write-Host "DEBUG: Validating local user '$Username' (no domain part)"
            
            $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Machine
            $principalContext = $null
            
            try {
                $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType)
                $isValid = $principalContext.ValidateCredentials($Username, $Password)
                Write-Host "DEBUG: Local validation result: $isValid"
                return $isValid
            }
            finally {
                if ($principalContext) { $principalContext.Dispose() }
            }
        }
    }
    catch {
        Write-Host "DEBUG: Credential validation exception: $($_.Exception.Message)"
        Write-Host "DEBUG: Exception type: $($_.Exception.GetType().FullName)"
        return $false
    }
}

function Get-PasswordModal {
    param(
        [string]$Username,
        [string]$Title = "Enter Password"
    )
    
    do {
        $validPassword = $false
        $password = $null
        
        # Create password input form
        $passwordForm = New-Object System.Windows.Forms.Form
        $passwordForm.Text = $Title
        $passwordForm.Size = New-Object System.Drawing.Size(400, 280)
        $passwordForm.StartPosition = "CenterParent"
        $passwordForm.FormBorderStyle = "FixedDialog"
        $passwordForm.MaximizeBox = $false
        $passwordForm.MinimizeBox = $false
        $passwordForm.TopMost = $true
        
        # Username label
        $lblUser = New-Object System.Windows.Forms.Label
        $lblUser.Text = "Username: $Username"
        $lblUser.Location = New-Object System.Drawing.Point(20, 20)
        $lblUser.Size = New-Object System.Drawing.Size(350, 20)
        $lblUser.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
        $passwordForm.Controls.Add($lblUser)
        
        # Instructions label
        $lblInstructions = New-Object System.Windows.Forms.Label
        $lblInstructions.Text = "Enter your password and click Validate. You must validate successfully to continue."
        $lblInstructions.Location = New-Object System.Drawing.Point(20, 45)
        $lblInstructions.Size = New-Object System.Drawing.Size(350, 40)
        $lblInstructions.ForeColor = [System.Drawing.Color]::DarkBlue
        $passwordForm.Controls.Add($lblInstructions)
        
        # Password label
        $lblPassword = New-Object System.Windows.Forms.Label
        $lblPassword.Text = "Password:"
        $lblPassword.Location = New-Object System.Drawing.Point(20, 100)
        $lblPassword.Size = New-Object System.Drawing.Size(80, 20)
        $passwordForm.Controls.Add($lblPassword)
        
        # Password textbox
        $txtPassword = New-Object System.Windows.Forms.TextBox
        $txtPassword.Location = New-Object System.Drawing.Point(100, 100)
        $txtPassword.Size = New-Object System.Drawing.Size(250, 20)
        $txtPassword.UseSystemPasswordChar = $true
        $passwordForm.Controls.Add($txtPassword)
        
        # Status label for validation feedback
        $lblStatus = New-Object System.Windows.Forms.Label
        $lblStatus.Location = New-Object System.Drawing.Point(20, 130)
        $lblStatus.Size = New-Object System.Drawing.Size(350, 40)
        $lblStatus.Text = ""
        $passwordForm.Controls.Add($lblStatus)
        
        # Validate button
        $btnValidate = New-Object System.Windows.Forms.Button
        $btnValidate.Text = "Validate"
        $btnValidate.Location = New-Object System.Drawing.Point(100, 160)
        $btnValidate.Size = New-Object System.Drawing.Size(80, 30)
        $passwordForm.Controls.Add($btnValidate)
        
        # OK button
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "OK"
        $btnOK.Location = New-Object System.Drawing.Point(200, 200)
        $btnOK.Size = New-Object System.Drawing.Size(80, 30)
        $btnOK.DialogResult = "OK"
        $btnOK.Enabled = $false  # Disabled until validation passes
        $passwordForm.AcceptButton = $btnOK
        $passwordForm.Controls.Add($btnOK)
        
        # Cancel button
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"
        $btnCancel.Location = New-Object System.Drawing.Point(290, 200)
        $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
        $btnCancel.DialogResult = "Cancel"
        $passwordForm.CancelButton = $btnCancel
        $passwordForm.Controls.Add($btnCancel)
        
        # Variable to track if password was validated
        $script:passwordValidated = $false
        
        # Validation logic
        $btnValidate.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                $lblStatus.Text = "Please enter a password"
                $lblStatus.ForeColor = [System.Drawing.Color]::Red
                $btnOK.Enabled = $false
                $script:passwordValidated = $false
                return
            }
            
            $lblStatus.Text = "Validating credentials..."
            $lblStatus.ForeColor = [System.Drawing.Color]::Blue
            $btnValidate.Enabled = $false
            $passwordForm.Update()
            
            $isValid = Test-UserCredentials -Username $Username -Password $txtPassword.Text
            
            if ($isValid) {
                $lblStatus.Text = "Credentials validated successfully! You can now click OK."
                $lblStatus.ForeColor = [System.Drawing.Color]::Green
                $btnOK.Enabled = $true
                $script:passwordValidated = $true
            } else {
                $lblStatus.Text = "Invalid credentials. Please try again."
                $lblStatus.ForeColor = [System.Drawing.Color]::Red
                $btnOK.Enabled = $false
                $script:passwordValidated = $false
                $txtPassword.Clear()
                $txtPassword.Focus()
            }
            
            $btnValidate.Enabled = $true
        })
        
        # Auto-validate on Enter key in password field
        $txtPassword.Add_KeyDown({
            if ($_.KeyCode -eq "Enter") {
                $btnValidate.PerformClick()
            }
        })
        
        # Reset password validation when text changes
        $txtPassword.Add_TextChanged({
            if ($script:passwordValidated -and $txtPassword.Text -ne $password) {
                $lblStatus.Text = "Password changed. Please validate again."
                $lblStatus.ForeColor = [System.Drawing.Color]::Orange
                $btnOK.Enabled = $false
                $script:passwordValidated = $false
            }
        })
        
        # Focus on password field
        $passwordForm.Add_Shown({
            $txtPassword.Focus()
        })
        
        # Show the form and get result
        $result = $passwordForm.ShowDialog()
        
        if ($result -eq "OK" -and $script:passwordValidated) {
            $validPassword = $true
            $password = $txtPassword.Text
        } elseif ($result -eq "Cancel") {
            return $null  # User cancelled
        } else {
            # User clicked OK without validating - show warning and loop again
            [System.Windows.Forms.MessageBox]::Show("You must validate your password before continuing.", "Validation Required", "OK", "Warning")
        }
        
        $passwordForm.Dispose()
        
    } while (-not $validPassword)
    
    return $password
}

function scheduleOn { 
    # Remove any existing task first
    scheduleOff

    # Define the script path and action
    $scriptPath = Join-Path $PSScriptRoot 'zoomdownloader.ps1'
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    Write-Host "Creating scheduled task with arguments: $arguments"
    
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $arguments -WorkingDirectory $PSScriptRoot
    
   # Create trigger based on selected schedule
    $trigger = switch ($script:cmbSchedule.SelectedIndex) {
        0 { # Every Day at Midnight
            New-ScheduledTaskTrigger -At "00:00" -Daily
        }
        1 { # Every Hour
            New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 1825)
        }
        2 { # Every 5 Minutes
            New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 1825)
        }
        3 { # Immediate Single Run
            New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(15)
        }
        4 { # Custom
            if ([string]::IsNullOrWhiteSpace($script:txtCustom.Text)) {
                Write-Host "Custom schedule expression is required but was empty"
                return $false
            }
            try {
                Invoke-Expression "New-ScheduledTaskTrigger $($script:txtCustom.Text.Trim())"
            }
            catch {
                Write-Host "Invalid custom schedule expression: $($_.Exception.Message)"
                return $false
            }
        }
        default {
            Write-Host "Invalid schedule selection: $($script:cmbSchedule.SelectedIndex)"
            return $false
        }
    }

    Write-Host "Creating scheduled task trigger for option $($script:cmbSchedule.SelectedIndex): $($script:cmbSchedule.SelectedItem)"
  
    # Get user information
    $userInfo = Get-CurrentUserInfo
    if ($userInfo.IsDomainUser) {
        Write-Host "Task will be configured for domain user: $($userInfo.FullUsername)"
        $userForTask = $userInfo.FullUsername
    } else {
        Write-Host "Task will be configured for local user: $($userInfo.FullUsername)"
        $userForTask = $userInfo.FullUsername
    }

    $passwordPlain = Get-PasswordModal -Username $userForTask -Title "Scheduled Task Credentials Required"
    
    if ($null -eq $passwordPlain) {
        Write-Host "Password entry was cancelled. Task creation aborted."
        $null = [System.Windows.Forms.MessageBox]::Show("Task creation was cancelled. No scheduled task was created.", "Task Creation Cancelled", "Ok", "Information")
        return $false
    }
    
    try {
       # Principal (RunLevel Highest + logon type)
        $principal = New-ScheduledTaskPrincipal -UserId $userForTask -LogonType Password -RunLevel Highest

        # (optional) Settings
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

        # Build the task object
        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings

        # Register using InputObject + credentials (do NOT also pass -Action/-Trigger/-Principal here)
        $null = Register-ScheduledTask -TaskName $taskName -InputObject $task -User $userForTask -Password $passwordPlain -Force

        Write-Host "Task successfully created and configured to run as $userForTask."
        $null = [System.Windows.Forms.MessageBox]::Show("Scheduled task created successfully and configured to run as $userForTask!", "Task Created Successfully", "OK", "Information")
        return $true
    }
    catch {
        Write-Host "Failed to create scheduled task: $($_.Exception.Message)"
        $null = [System.Windows.Forms.MessageBox]::Show("Failed to create scheduled task: $($_.Exception.Message)", "Task Creation Failed", "OK", "Error")
        return $false
    }
}

function scheduleOff { 
  try {
    Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "The task '$taskName' has been unscheduled."
  }
  catch {
   
  }    
}

$btnCancel.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to cancel the setup?", "Cancel Setup", "YesNo", "Question")
    if ($result -eq "Yes") {
        $form.DialogResult = "Cancel"
        $form.Close()
    }
})

# Function to get Zoom OAuth token with retry logic
function Get-ZoomAccessToken {
    param(
        [string]$AccountId,
        [string]$ClientId,
        [string]$ClientSecret,
        [int]$MaxRetries = 3
    )
    
    $tokenUrl = "https://zoom.us/oauth/token"
    $credentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$ClientId`:$ClientSecret"))
    
    $headers = @{
        "Authorization" = "Basic $credentials"
        "Content-Type" = "application/x-www-form-urlencoded"
    }
    
    $body = @{
        "grant_type" = "account_credentials"
        "account_id" = $AccountId
    }
    
    for ($retry = 1; $retry -le $MaxRetries; $retry++) {
        try {
            $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Headers $headers -Body $body
            Write-Host "Access token obtained successfully"
            return $response.access_token
        } catch {
            Write-Host "Failed to get Zoom access token (attempt $retry/$MaxRetries): $_"
            if ($retry -eq $MaxRetries) {
                throw "Failed to get Zoom access token after $MaxRetries attempts: $_"
            }
            Start-Sleep -Seconds (2 * $retry)
        }
    }
}


function Test-OneDrive {
    param(
        [string] $appId,
        [string] $appSecret,
        [string] $tenantName
    )
      
    $Scope = "https://graph.microsoft.com/.default"
    $AuthUrl = "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" 
    $UserAgent = "NONISV|Zoom Downloader|OneDrive Upload/1.0"
    #Write-Host "Authenticating to OneDrive, $appId, $appSecret, $tenantName"
    Add-Type -AssemblyName System.Web

    # Create body
    $Body = @{
        client_id     = $appId
        client_secret = $appSecret
        scope         = $Scope
        grant_type    = 'client_credentials'
    }

    # Splat the parameters for Invoke-Restmethod for cleaner code
    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        Body = $Body
        Uri = $AuthUrl
        UserAgent = $UserAgent
    }

    # Request the token!
    try {
        $Request = Invoke-RestMethod @PostSplat
        $AccessToken = $Request.access_token
        Write-Host "Authentication To OneDrive Successful"
    } catch {
        # Catch block to handle the exception
        Write-Host "Unable to Authenticate: $($_.Exception.Message)"
        $errorMessage = "Unable to authenticate"
        $exception = New-Object System.Exception($_.Exception.Message)
        throw $exception
    }
}

$script:btnTestOneDrive.Add_Click({
    $AppId = $script:txtAppId.Text.Trim()
    $ClientSecret = $script:txtClientSecret.Text.Trim()
    $TenantName = $script:txtTenantName.Text.Trim()

    $script:lblOneDriveStatus.Text = "Testing connection..."
    $script:lblOneDriveStatus.ForeColor = [System.Drawing.Color]::Blue
    $script:btnTestOneDrive.Enabled = $false
    $form.Update()

    try {
        #$oneDriveFileStorage = [OneDriveFileStorage]::new($AppId, $ClientSecret, $TenantName)
        Test-OneDrive -appId $AppId -appSecret $ClientSecret -tenantName $TenantName
        Write-Host "SUCCESS: Simulated OneDrive connection."
        $script:lblOneDriveStatus.Text = "SUCCESS: OneDrive connection."
        $script:lblOneDriveStatus.ForeColor = [System.Drawing.Color]::Green
    }
    catch {
        $script:lblOneDriveStatus.Text = "ERROR: Connection test failed, $($_.Exception.Message)"
        $script:lblOneDriveStatus.ForeColor = [System.Drawing.Color]::Red  
    }
    finally {
        $script:btnTestOneDrive.Enabled = $true
    }
})

function Test-ZoomConnection {
    param(
        [string]$ApiKey,
        [string]$ApiSecret,
        [string]$AccountId
    )

    $configuration =  [PSCustomObject]@{
        AccountID    = $AccountId
        ClientID     = $ApiKey
        ClientSecret = $ApiSecret
    }
  
    try {
        # Basic validation
        if ([string]::IsNullOrWhiteSpace($ApiKey) -or 
            [string]::IsNullOrWhiteSpace($ApiSecret) -or 
            [string]::IsNullOrWhiteSpace($AccountId)) {
            return @{
                Success = $false
                ErrorMessage = "All credentials are required"
                AccountName = $null
            }
        }
        
        $config = @{
            zoom     = $configuration
        } 

        $accessToken = Get-ZoomAccessToken -AccountId $config.zoom.accountId -ClientId $config.zoom.clientId -ClientSecret $config.zoom.clientSecret

        return @{
            Success = $true
            ErrorMessage = "Connection to Zoom was Successful"
            AccountName = $null
        }        
    }
    catch {
        return @{
            Success = $false
            ErrorMessage = "Connection to Zoom was Unsuccessful $($_.Exception.Message)"
            AccountName = $null
        }         
        Write-Host "Unable to authenticate to Zoom $($_.Exception.Message)"
    }
    
}

# Show initial page
Show-CurrentPage

# Show the form
$result = $form.ShowDialog()

if ($result -eq "OK") {
    Write-Host "Setup completed successfully!" -ForegroundColor Green
} else {
    Write-Host "Setup was cancelled." -ForegroundColor Yellow
}
