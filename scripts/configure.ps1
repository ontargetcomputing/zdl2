# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global configuration storage
$global:Config = @{
    Zoom = @{}
    Database = @{}
    Storage = @{}
}

# Current page tracking
$script:currentPage = 1

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zoom Downloader Setup Wizard"
$form.Size = New-Object System.Drawing.Size(600, 500)
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
$progressBar.Maximum = 5  # Changed to 5 pages total
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
$pnlContent.Size = New-Object System.Drawing.Size(540, 320)
$pnlContent.BorderStyle = "FixedSingle"
$form.Controls.Add($pnlContent)

# Buttons
$btnPrevious = New-Object System.Windows.Forms.Button
$btnPrevious.Text = "< Previous"
$btnPrevious.Location = New-Object System.Drawing.Point(290, 420)
$btnPrevious.Size = New-Object System.Drawing.Size(80, 30)
$btnPrevious.Enabled = $false
$form.Controls.Add($btnPrevious)

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Next >"
$btnNext.Location = New-Object System.Drawing.Point(380, 420)
$btnNext.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($btnNext)

$btnFinish = New-Object System.Windows.Forms.Button
$btnFinish.Text = "Finish"
$btnFinish.Location = New-Object System.Drawing.Point(470, 420)
$btnFinish.Size = New-Object System.Drawing.Size(80, 30)
$btnFinish.Visible = $false
$form.Controls.Add($btnFinish)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(200, 420)
$btnCancel.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($btnCancel)

# Page titles
$pageTitles = @{
    1 = "Welcome to the Zoom Downloader Setup Wizard"
    2 = "Zoom Credentials"
    3 = "Storage Selection"
    4 = "Database Configuration"
    5 = "Setup Summary"
}

# Function to create Welcome page
function Create-WelcomePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    
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
    $panel.Dock = "Fill"
    
    # API Key
    $lblApiKey = New-Object System.Windows.Forms.Label
    $lblApiKey.Text = "Zoom Client ID:"
    $lblApiKey.Location = New-Object System.Drawing.Point(30, 30)
    $lblApiKey.Size = New-Object System.Drawing.Size(150, 20)
    $panel.Controls.Add($lblApiKey)
    
    $script:txtApiKey = New-Object System.Windows.Forms.TextBox
    $txtApiKey.Location = New-Object System.Drawing.Point(200, 30)
    $txtApiKey.Size = New-Object System.Drawing.Size(300, 20)
    $txtApiKey.Name = "txtApiKey"
    $panel.Controls.Add($txtApiKey)
    
    # API Secret
    $lblApiSecret = New-Object System.Windows.Forms.Label
    $lblApiSecret.Text = "Zoom Client Secret:"
    $lblApiSecret.Location = New-Object System.Drawing.Point(30, 70)
    $lblApiSecret.Size = New-Object System.Drawing.Size(150, 20)
    $panel.Controls.Add($lblApiSecret)
    
    $script:txtApiSecret = New-Object System.Windows.Forms.TextBox
    $txtApiSecret.Location = New-Object System.Drawing.Point(200, 70)
    $txtApiSecret.Size = New-Object System.Drawing.Size(300, 20)
    $txtApiSecret.UseSystemPasswordChar = $true
    $txtApiSecret.Name = "txtApiSecret"
    $panel.Controls.Add($txtApiSecret)
    
    # Account ID
    $lblAccountId = New-Object System.Windows.Forms.Label
    $lblAccountId.Text = "Account ID:"
    $lblAccountId.Location = New-Object System.Drawing.Point(30, 110)
    $lblAccountId.Size = New-Object System.Drawing.Size(150, 20)
    $panel.Controls.Add($lblAccountId)
    
    $script:txtAccountId = New-Object System.Windows.Forms.TextBox
    $txtAccountId.Location = New-Object System.Drawing.Point(200, 110)
    $txtAccountId.Size = New-Object System.Drawing.Size(300, 20)
    $txtAccountId.Name = "txtAccountId"
    $panel.Controls.Add($txtAccountId)

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

    # # Show message box when Test Connection is clicked
    # $script:btnTestZoom.Add_Click({
    #     [System.Windows.Forms.MessageBox]::Show("Zoom Connection Tested", "Connection Test", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    # })

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
                $script:lblTestStatus.Text = "SUCCESS: Connection successful! Account: $($result.AccountName)"
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
    
    return $panel
}

# Function to create Storage Selection page
function Create-StorageSelectionPage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    
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
    $pnlOneDrive.Size = New-Object System.Drawing.Size(450, 100)
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

    $txtAppId = New-Object System.Windows.Forms.TextBox
    $txtAppId.Location = New-Object System.Drawing.Point(110, 10)
    $txtAppId.Size = New-Object System.Drawing.Size(250, 20)
    $txtAppId.Name = "txtAppId"
    $pnlOneDrive.Controls.Add($txtAppId)

    $lblClientSecret = New-Object System.Windows.Forms.Label
    $lblClientSecret.Text = "Client Secret:"
    $lblClientSecret.Location = New-Object System.Drawing.Point(0, 40)
    $lblClientSecret.Size = New-Object System.Drawing.Size(100, 20)
    $pnlOneDrive.Controls.Add($lblClientSecret)

    $txtClientSecret = New-Object System.Windows.Forms.TextBox
    $txtClientSecret.Location = New-Object System.Drawing.Point(110, 40)
    $txtClientSecret.Size = New-Object System.Drawing.Size(250, 20)
    $txtClientSecret.UseSystemPasswordChar = $true
    $txtClientSecret.Name = "txtClientSecret"
    $pnlOneDrive.Controls.Add($txtClientSecret)

    $lblTenantName = New-Object System.Windows.Forms.Label
    $lblTenantName.Text = "Tenant Name:"
    $lblTenantName.Location = New-Object System.Drawing.Point(0, 70)
    $lblTenantName.Size = New-Object System.Drawing.Size(100, 20)
    $pnlOneDrive.Controls.Add($lblTenantName)

    $txtTenantName = New-Object System.Windows.Forms.TextBox
    $txtTenantName.Location = New-Object System.Drawing.Point(110, 70)
    $txtTenantName.Size = New-Object System.Drawing.Size(250, 20)
    $txtTenantName.Name = "txtTenantName"
    $pnlOneDrive.Controls.Add($txtTenantName)
    
    # Add Test Connection button to OneDrive panel
    $btnTestOneDrive = New-Object System.Windows.Forms.Button
    $btnTestOneDrive.Text = "Test Connection"
    $btnTestOneDrive.Location = New-Object System.Drawing.Point(370, 10)
    $btnTestOneDrive.Size = New-Object System.Drawing.Size(120, 30)
    $btnTestOneDrive.Name = "btnTestOneDrive"
    $pnlOneDrive.Controls.Add($btnTestOneDrive)

    # S3 Panel
    $pnlS3 = New-Object System.Windows.Forms.Panel
    $pnlS3.Location = New-Object System.Drawing.Point(70, 190)
    $pnlS3.Size = New-Object System.Drawing.Size(450, 150)
    $pnlS3.Visible = $false
    $pnlS3.Name = "pnlS3"
    $panel.Controls.Add($pnlS3)
    
    $lblAccessKey = New-Object System.Windows.Forms.Label
    $lblAccessKey.Text = "Access Key ID:"
    $lblAccessKey.Location = New-Object System.Drawing.Point(0, 10)
    $lblAccessKey.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblAccessKey)
    
    $txtAccessKey = New-Object System.Windows.Forms.TextBox
    $txtAccessKey.Location = New-Object System.Drawing.Point(110, 10)
    $txtAccessKey.Size = New-Object System.Drawing.Size(250, 20)
    $txtAccessKey.Name = "txtAccessKey"
    $pnlS3.Controls.Add($txtAccessKey)
    
    $lblSecretKey = New-Object System.Windows.Forms.Label
    $lblSecretKey.Text = "Secret Key:"
    $lblSecretKey.Location = New-Object System.Drawing.Point(0, 40)
    $lblSecretKey.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblSecretKey)
    
    $txtSecretKey = New-Object System.Windows.Forms.TextBox
    $txtSecretKey.Location = New-Object System.Drawing.Point(110, 40)
    $txtSecretKey.Size = New-Object System.Drawing.Size(250, 20)
    $txtSecretKey.UseSystemPasswordChar = $true
    $txtSecretKey.Name = "txtSecretKey"
    $pnlS3.Controls.Add($txtSecretKey)
    
    $lblBucket = New-Object System.Windows.Forms.Label
    $lblBucket.Text = "Bucket Name:"
    $lblBucket.Location = New-Object System.Drawing.Point(0, 70)
    $lblBucket.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblBucket)
    
    $txtBucket = New-Object System.Windows.Forms.TextBox
    $txtBucket.Location = New-Object System.Drawing.Point(110, 70)
    $txtBucket.Size = New-Object System.Drawing.Size(250, 20)
    $txtBucket.Name = "txtBucket"
    $pnlS3.Controls.Add($txtBucket)
    
    $lblRegion = New-Object System.Windows.Forms.Label
    $lblRegion.Text = "Region:"
    $lblRegion.Location = New-Object System.Drawing.Point(0, 100)
    $lblRegion.Size = New-Object System.Drawing.Size(100, 20)
    $pnlS3.Controls.Add($lblRegion)
    
    $cmbRegion = New-Object System.Windows.Forms.ComboBox
    $cmbRegion.Items.AddRange(@("us-east-1", "us-west-1", "us-west-2", "eu-west-1", "ap-southeast-1"))
    $cmbRegion.Location = New-Object System.Drawing.Point(110, 100)
    $cmbRegion.Size = New-Object System.Drawing.Size(150, 20)
    $cmbRegion.DropDownStyle = "DropDownList"
    $cmbRegion.Name = "cmbRegion"
    $pnlS3.Controls.Add($cmbRegion)
    
    # Add Test Connection button to S3 panel
    $btnTestS3 = New-Object System.Windows.Forms.Button
    $btnTestS3.Text = "Test Connection"
    $btnTestS3.Location = New-Object System.Drawing.Point(370, 10)
    $btnTestS3.Size = New-Object System.Drawing.Size(120, 30)
    $btnTestS3.Name = "btnTestS3"
    $pnlS3.Controls.Add($btnTestS3)
    
    # Radio button events to show/hide panels
    $radioLocal.Add_CheckedChanged({
        $parentPanel = $this.Parent
        $parentPanel.Controls["pnlLocal"].Visible = $true #always visible
    })
    
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
    
    return $panel
}

# Function to create Database Configuration page
function Create-DatabasePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    
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
    $lblDbType.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblDbType)
    
    $cmbDatabaseType = New-Object System.Windows.Forms.ComboBox
    $cmbDatabaseType.Items.AddRange(@("SQL Server", "MySQL", "PostgreSQL", "SQLite"))
    $cmbDatabaseType.Location = New-Object System.Drawing.Point(140, 60)
    $cmbDatabaseType.Size = New-Object System.Drawing.Size(150, 20)
    $cmbDatabaseType.DropDownStyle = "DropDownList"
    $cmbDatabaseType.Name = "cmbDatabaseType"
    $panel.Controls.Add($cmbDatabaseType)
    
    # Server
    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Text = "Server:"
    $lblServer.Location = New-Object System.Drawing.Point(30, 100)
    $lblServer.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblServer)
    
    $txtServer = New-Object System.Windows.Forms.TextBox
    $txtServer.Location = New-Object System.Drawing.Point(140, 100)
    $txtServer.Size = New-Object System.Drawing.Size(200, 20)
    $txtServer.Name = "txtServer"
    $panel.Controls.Add($txtServer)
    
    # Database Name
    $lblDatabase = New-Object System.Windows.Forms.Label
    $lblDatabase.Text = "Database:"
    $lblDatabase.Location = New-Object System.Drawing.Point(30, 140)
    $lblDatabase.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblDatabase)
    
    $txtDatabase = New-Object System.Windows.Forms.TextBox
    $txtDatabase.Location = New-Object System.Drawing.Point(140, 140)
    $txtDatabase.Size = New-Object System.Drawing.Size(200, 20)
    $txtDatabase.Name = "txtDatabase"
    $panel.Controls.Add($txtDatabase)
    
    # Username
    $lblUsername = New-Object System.Windows.Forms.Label
    $lblUsername.Text = "Username:"
    $lblUsername.Location = New-Object System.Drawing.Point(30, 180)
    $lblUsername.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblUsername)
    
    $txtUsername = New-Object System.Windows.Forms.TextBox
    $txtUsername.Location = New-Object System.Drawing.Point(140, 180)
    $txtUsername.Size = New-Object System.Drawing.Size(200, 20)
    $txtUsername.Name = "txtUsername"
    $panel.Controls.Add($txtUsername)
    
    # Password
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Text = "Password:"
    $lblPassword.Location = New-Object System.Drawing.Point(30, 220)
    $lblPassword.Size = New-Object System.Drawing.Size(100, 20)
    $panel.Controls.Add($lblPassword)
    
    $txtPassword = New-Object System.Windows.Forms.TextBox
    $txtPassword.Location = New-Object System.Drawing.Point(140, 220)
    $txtPassword.Size = New-Object System.Drawing.Size(200, 20)
    $txtPassword.UseSystemPasswordChar = $true
    $txtPassword.Name = "txtPassword"
    $panel.Controls.Add($txtPassword)
    
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
    $txtSummary.Size = New-Object System.Drawing.Size(480, 200)
    $txtSummary.Multiline = $true
    $txtSummary.ScrollBars = "Vertical"
    $txtSummary.ReadOnly = $true
    $txtSummary.Name = "txtSummary"
    $panel.Controls.Add($txtSummary)
    
    return $panel
}

# Page controls array
$pageControls = @(
    (Create-WelcomePage),
    (Create-ZoomCredentialsPage),
    (Create-StorageSelectionPage),
    (Create-DatabasePage),
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
    $btnNext.Visible = $script:currentPage -lt 5
    $btnFinish.Visible = $script:currentPage -eq 5
    
    # Update summary if on last page
    if ($script:currentPage -eq 5) {
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
                $localPath = $currentPanel.Controls["pnlLocal"].Controls["txtLocalPath"].Text.Trim()
                if ([string]::IsNullOrWhiteSpace($localPath)) {
                    [System.Windows.Forms.MessageBox]::Show("Local path is required.", "Validation Error")
                    return $false
                }
                $global:Config.Storage.Type = "Local"
                $global:Config.Storage.Path = $localPath
            }
            elseif ($radioOneDrive.Checked) {
                $oneDriveFolder = $currentPanel.Controls["pnlOneDrive"].Controls["txtOneDriveFolder"].Text.Trim()
                if ([string]::IsNullOrWhiteSpace($oneDriveFolder)) {
                    [System.Windows.Forms.MessageBox]::Show("OneDrive folder is required.", "Validation Error")
                    return $false
                }
                $global:Config.Storage.Type = "OneDrive"
                $global:Config.Storage.Folder = $oneDriveFolder
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
                $global:Config.Storage.SecretKey = $secretKey
                $global:Config.Storage.Bucket = $bucket
                $global:Config.Storage.Region = $region
            }
        }
        4 { # Database
            $currentPanel = $pageControls[3]
            $dbType = $currentPanel.Controls["cmbDatabaseType"].SelectedItem
            $server = $currentPanel.Controls["txtServer"].Text.Trim()
            $database = $currentPanel.Controls["txtDatabase"].Text.Trim()
            $username = $currentPanel.Controls["txtUsername"].Text.Trim()
            $password = $currentPanel.Controls["txtPassword"].Text.Trim()
            
            if ($null -eq $dbType) {
                [System.Windows.Forms.MessageBox]::Show("Please select a database type.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($server)) {
                [System.Windows.Forms.MessageBox]::Show("Server is required.", "Validation Error")
                return $false
            }
            if ([string]::IsNullOrWhiteSpace($database)) {
                [System.Windows.Forms.MessageBox]::Show("Database name is required.", "Validation Error")
                return $false
            }
            
            $global:Config.Database.Type = $dbType
            $global:Config.Database.Server = $server
            $global:Config.Database.Database = $database
            $global:Config.Database.Username = $username
            $global:Config.Database.Password = $password
        }
    }
    return $true
}

# Function to update summary
function Update-Summary {
    $storageInfo = ""
    switch ($global:Config.Storage.Type) {
        "Local" { $storageInfo = "Local Storage: $($global:Config.Storage.Path)" }
        "OneDrive" { $storageInfo = "OneDrive Folder: $($global:Config.Storage.Folder)" }
        "S3" { $storageInfo = "S3 Bucket: $($global:Config.Storage.Bucket) (Region: $($global:Config.Storage.Region))" }
        default { $storageInfo = "Not configured" }
    }
    
    $summaryText = @"
Zoom Configuration:
  API Key: $($global:Config.Zoom.ApiKey)
  Account ID: $($global:Config.Zoom.AccountId)

Storage Configuration:
  $storageInfo

Database Configuration:
  Type: $($global:Config.Database.Type)
  Server: $($global:Config.Database.Server)
  Database: $($global:Config.Database.Database)
  Username: $($global:Config.Database.Username)

Click Finish to complete the setup.
"@
    
    $currentPanel = $pageControls[4]
    $currentPanel.Controls["txtSummary"].Text = $summaryText
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
        $configPath = Join-Path $env:USERPROFILE "ZoomDownloaderConfig.json"
        $configJson | Out-File -FilePath $configPath -Encoding utf8
        
        [System.Windows.Forms.MessageBox]::Show("Configuration saved to: $configPath", "Setup Complete")
        $form.DialogResult = "OK"
        $form.Close()
    }
})

$btnCancel.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to cancel the setup?", "Cancel Setup", "YesNo", "Question")
    if ($result -eq "Yes") {
        $form.DialogResult = "Cancel"
        $form.Close()
    }
})

# Placeholder function for testing Zoom connection
function Test-ZoomConnection {
    param(
        [string]$ApiKey,
        [string]$ApiSecret,
        [string]$AccountId
    )
    
    try {
        # Simulate API call delay
        Start-Sleep -Milliseconds 1500
        
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
        
        # Simulate successful connection for demo
        if ($ApiKey.Length -gt 10 -and $ApiSecret.Length -gt 10) {
            return @{
                Success = $true
                ErrorMessage = $null
                AccountName = "Demo Account (Replace with real API call)"
            }
        } else {
            return @{
                Success = $false
                ErrorMessage = "Invalid credentials format"
                AccountName = $null
            }
        }
    }
    catch {
        return @{
            Success = $false
            ErrorMessage = $_.Exception.Message
            AccountName = $null
        }
    }
}

# Show initial page
Show-CurrentPage

# Show the form
$result = $form.ShowDialog()

if ($result -eq "OK") {
    Write-Host "Setup completed successfully!" -ForegroundColor Green
    Write-host "Configuration saved to: $(Join-Path $env:USERPROFILE "ZoomDownloaderConfig.json")" -ForegroundColor Green
} else {
    Write-Host "Setup was cancelled." -ForegroundColor Yellow
}
