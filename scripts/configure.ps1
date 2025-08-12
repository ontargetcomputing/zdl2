using module ../Modules/FileStorage/Classes/OneDriveFileStorage.psm1
using module ../Modules/FileStorage/Classes/S3FileStorage.psm1
using module ../Modules/Zoom/Classes/ZoomService.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1
using module ../Modules/Database/Classes/SQLServerDatabase.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("configure")

$user_config = $configuration.ReadUserConfiguration()

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

    $script:btnTestOneDrive.Add_Click({
        $AppId = $script:txtAppId.Text.Trim()
        $ClientSecret = $script:txtClientSecret.Text.Trim()
        $TenantName = $script:txtTenantName.Text.Trim()

        $script:lblOneDriveStatus.Text = "Testing connection..."
        $script:lblOneDriveStatus.ForeColor = [System.Drawing.Color]::Blue
        $script:btnTestOneDrive.Enabled = $false
        $form.Update()

        try {
            $oneDriveFileStorage = [OneDriveFileStorage]::new($AppId, $ClientSecret, $TenantName)
            $oneDriveFileStorage.Authenticate()
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
            $s3FileStorage = [S3FileStorage]::new(
                $script:txtAccessKey.Text.Trim(),
                $script:txtSecretKey.Text.Trim(),
                $script:txtBucket.Text.Trim(),
                $script:cmbRegion.SelectedItem
            )
            $s3FileStorage.Authenticate()
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

    if($user_config.storage.type -eq "S3") {
        $radioS3.Checked = $true
        $pnlS3.Visible = $true
        $pnlLocal.Visible = $false
        $pnlOneDrive.Visible = $false
        $script:txtAccessKey.Text = $user_config.storage.accessKey
        $script:txtSecretKey.Text = $user_config.storage.secretAccessKey
        $script:txtBucket.Text = $user_config.storage.bucketName
        $script:cmbRegion.SelectedItem = $user_config.storage.region
    } elseif($user_config.storage.type -eq "OneDrive") {
        $radioOneDrive.Checked = $true
        $pnlOneDrive.Visible = $true
        $pnlLocal.Visible = $false
        $pnlS3.Visible = $false
        $script:txtAppId.Text = $user_config.storage.appId
        $script:txtClientSecret.Text = $user_config.storage.clientSecret
        $script:txtTenantName.Text = $user_config.storage.tenantName
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

        $config = @{
            sqlserver = $sqlserver
        }

        $script:database = [SQLServerDatabase]::new($config, $false)  
        
        try {
            $script:database.Connect()
            $script:database.Disconnect()
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
        $script:lblStartDate.Location = New-Object System.Drawing.Point(50, 240)
        $script:lblStartDate.Size = New-Object System.Drawing.Size(120, 20)

`        
        $panel.Controls.Add($script:lblStartDate)

        $script:dtpStartDate = New-Object System.Windows.Forms.DateTimePicker
        $script:dtpStartDate.Location = New-Object System.Drawing.Point(160, 240)
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
                type = $global:Config.Storage.Type
            }             
        } elseif( $type -eq 'OneDrive' ) {
            $storageConfig = @{
                type = $global:Config.Storage.Type

                appId = $global:Config.Storage.AppId
                clientSecret = $global:Config.Storage.ClientSecret
                tenantName = $global:Config.Storage.TenantName
            }             
        } else {
            $storageConfig = @{
                type = $global:Config.Storage.Type
                accessKey = $global:Config.Storage.AccessKey
                secretAccessKey = $global:Config.Storage.SecretAccessKey
                bucketName = $global:Config.Storage.Bucket
                region = $global:Config.Storage.Region
            }               
        }

        $databaseConfig = @{
            server = $global:Config.Database.Server
            port = $global:Config.Database.Port
            database = $global:Config.Database.Database
            schema = $global:Config.Database.Schema
            userid = $global:Config.Database.Username
            password = $global:Config.Database.Password
        }

        $schedule = @{
            schedule = $global:Config.Schedule.schedule
            custom = $global:Config.Schedule.custom
            dateRange = $global:Config.Schedule.dateRange
            customFromDate = $global:Config.Schedule.customFromDate
        }

        $accounts = $global:Config.Accounts -join "`r`n"

        $config = @{
            zoom     = $zoomConfig
            storage  = $storageConfig
            database = $databaseConfig
            schedule = $schedule
            accounts = $accounts
        }

        $configuration.CreateLocalAppdataFolder()
        $jsonString = $config | ConvertTo-Json
        $configuration.SaveUserConfiguration($jsonString)

        if ($database -eq $null) {
            Write-Host "The database is null."

            $sqlserver = @{
                server = $script:txtServer.text
                port = $script:txtPort.text
                database = $script:txtDatabase.text
                schema = $script:txtSchema.text
                userid = $script:txtUsername.text
                password = $script:txtPassword.text
            }

            $config = @{
                sqlserver = $sqlserver
            }

            $script:database = [SQLServerDatabase]::new($config, $false)  
            
            try {
                $database.Connect()
                $database.InsertIntoAccountsToDownloadTable($accounts)
                $database.Disconnect()
            }
            catch {
                Write-Host "ERROR: $($_.Exception.Message)"
            } 
        } 

        scheduleOn
        [System.Windows.Forms.MessageBox]::Show("Configuration saved and Job Scheduled", "Setup Complete")

        $form.DialogResult = "OK"
        $form.Close()
    }
})


function scheduleOn { 
      scheduleOff
      $script = (Join-Path $PSScriptRoot '\zoomdownloader.ps1')
      $arguement = "-noprofile -executionpolicy bypass ", $script -join " "
      Write-Host($arguement)
      $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $arguement -WorkingDirectory $PSScriptRoot

      if ( $script:cmbSchedule.SelectedIndex -eq 0 ) {
        $commandString =  "New-ScheduledTaskTrigger", "-At 0:00 -Daily" -join " "
      } elseif ( $script:cmbSchedule.SelectedIndex -eq 1) {
        $commandString =  "New-ScheduledTaskTrigger", "-Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 60) -RepetitionDuration (New-TimeSpan -Days (5 * 365))" -join " "        
      } elseif ( $script:cmbSchedule.SelectedIndex -eq 2) {
        $commandString =  "New-ScheduledTaskTrigger", "-Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days (5 * 365))" -join " "
      } elseif ( $script:cmbSchedule.SelectedIndex -eq 3) {
        $commandString =  "New-ScheduledTaskTrigger", "-At (Get-Date) -Once" -join " "
      } elseif ( $script:cmbSchedule.SelectedIndex -eq 4) {
        $commandString =  "New-ScheduledTaskTrigger", $script.textCustom.text -join " "
      }

      Write-Host "Scheduling with '$commandString'."
      $trigger = Invoke-Expression $commandString
      $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
      Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal
      Write-Host "Scheduled the task '$taskName' to run even if the user logs out."
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
        $zoomService = [ZoomService]::new($config)
        $zoomService.GetAccessToken()
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
