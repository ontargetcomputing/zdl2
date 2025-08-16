using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1


# Import required modules
$requiredModules = @('SqlServer')
foreach ($module in $requiredModules) {
    try {
        Import-Module $module -ErrorAction Stop
    } catch {
        Write-Warning "$module module not found. Installing..."
        Install-Module $module -Force -AllowClobber -Scope CurrentUser
        Import-Module $module
    }
}

# Global variables for progress tracking
$script:TotalProcessed = 0
$script:TotalInserted = 0
$script:TotalSkipped = 0
$script:TotalErrors = 0
$script:ProgressMutex = [System.Threading.Mutex]::new($false)
$script:LogMutex = [System.Threading.Mutex]::new($false)

# Function to write thread-safe log messages
function Write-ThreadSafeLog {
    param(
        [string]$Message,
        [string]$Level = "Info",
        [ConsoleColor]$Color = "White"
    )
    
    $script:LogMutex.WaitOne() | Out-Null
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        Write-Host $logMessage -ForegroundColor $Color
        
        # Also write to log file
        $logFile = "zoom_identify_$(Get-Date -Format 'yyyyMMdd').log"
        Add-Content -Path $logFile -Value $logMessage
    } finally {
        $script:LogMutex.ReleaseMutex()
    }
}

# Function to update progress in thread-safe manner
function Update-Progress {
    param(
        [int]$Processed = 0,
        [int]$Inserted = 0,
        [int]$Skipped = 0,
        [int]$Errors = 0
    )
    
    $script:ProgressMutex.WaitOne() | Out-Null
    try {
        $script:TotalProcessed += $Processed
        $script:TotalInserted += $Inserted
        $script:TotalSkipped += $Skipped
        $script:TotalErrors += $Errors
        
        if (($script:TotalProcessed % 100) -eq 0 -or $Processed -gt 0) {
            Write-ThreadSafeLog "Progress: Processed=$($script:TotalProcessed), Inserted=$($script:TotalInserted), Skipped=$($script:TotalSkipped), Errors=$($script:TotalErrors)" -Level "PROGRESS" -Color Cyan
        }
    } finally {
        $script:ProgressMutex.ReleaseMutex()
    }
}

# Function to load configuration using ZDAConfiguration module
function Get-Configuration {
    try {
        $configuration = [ZDAConfiguration]::new()
        $config = $configuration.ReadUserConfiguration()
        
        if (-not $config) {
            throw "Failed to read user configuration"
        }
        
        Write-ThreadSafeLog "Configuration loaded successfully" -Color Green
        return $config
    } catch {
        throw "Failed to load configuration: $_"
    }
}

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
            Write-ThreadSafeLog "Access token obtained successfully" -Color Green
            return $response.access_token
        } catch {
            Write-ThreadSafeLog "Failed to get Zoom access token (attempt $retry/$MaxRetries): $_" -Level "WARNING" -Color Yellow
            if ($retry -eq $MaxRetries) {
                throw "Failed to get Zoom access token after $MaxRetries attempts: $_"
            }
            Start-Sleep -Seconds (2 * $retry)
        }
    }
}

# Function to get all users from Zoom account (for when no specific accounts are configured)
function Get-ZoomUsers {
    param(
        [string]$AccessToken,
        [int]$PageSize = 300
    )
    
    $users = @()
    $nextPageToken = $null
    
    do {
        $url = "https://api.zoom.us/v2/users?status=active&page_size=$PageSize"
        if ($nextPageToken) {
            $url += "&next_page_token=$nextPageToken"
        }
        
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Content-Type" = "application/json"
        }
        
        try {
            $response = Invoke-RestMethod -Uri $url -Method GET -Headers $headers
            if ($response.users) {
                $users += $response.users.email
            }
            $nextPageToken = $response.next_page_token
        } catch {
            Write-ThreadSafeLog "Failed to get users: $_" -Level "ERROR" -Color Red
            break
        }
        
        Start-Sleep -Milliseconds 200
        
    } while ($nextPageToken)
    
    return $users
}

# Function to get existing recordings for duplicate checking (optimized with hash table)
function Get-ExistingRecordings {
    param(
        [string]$ConnectionString,
        [string]$TableName,
        [datetime]$FromDate
    )
    
    $fromStr = $FromDate.ToString("yyyy-MM-dd")
    
    # Get existing recordings as a hash set for fast lookups
    $sql = @"
SELECT DOWNLOAD_URL, MEETING_ID, RECORDING_TYPE, FILE_SIZE 
FROM $TableName 
WHERE RECORDING_START >= '$fromStr'
"@
    
    try {
        $existing = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 120
        $hashSet = @{}
        
        foreach ($row in $existing) {
            # Create composite key for duplicate detection
            $key1 = $row.DOWNLOAD_URL
            $key2 = "$($row.MEETING_ID)|$($row.RECORDING_TYPE)|$($row.FILE_SIZE)"
            
            if ($key1) { $hashSet[$key1] = $true }
            if ($key2) { $hashSet[$key2] = $true }
        }
        
        Write-ThreadSafeLog "Loaded $($existing.Count) existing recordings for duplicate checking" -Color Cyan
        return $hashSet
    } catch {
        Write-ThreadSafeLog "Failed to get existing recordings: $_" -Level "ERROR" -Color Red
        return @{}
    }
}

# Function to create the main script block for runspace execution
function Get-WorkerScriptBlock {
    return {
        param(
            $AccessToken,
            $Account, 
            $StartDate,
            $EndDate,
            $ConnectionString,
            $TableName,
            $ExistingRecordings,
            $ThreadId
        )
        
        # Thread-safe logging function
        function Write-ThreadSafeLog {
            param(
                [string]$Message,
                [string]$Level = "Info",
                [ConsoleColor]$Color = "White"
            )
            
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [Thread-$ThreadId] [$Level] $Message"
            
            # Simple console output
            Write-Host $logMessage -ForegroundColor $Color
            
            $logFile = "scripts/zoom_identify_${ThreadId}_$(Get-Date -Format 'yyyyMMdd').log"
            Add-Content -Path $logFile -Value $logMessage -Force
                    
        }

        # Function to update progress in thread-safe manner
        function Update-ThreadProgress {
            param(
                [int]$Processed = 0,
                [int]$Inserted = 0,
                [int]$Skipped = 0,
                [int]$Errors = 0
            )
            
            $sync.Progress.Processed += $Processed
            $sync.Progress.Inserted += $Inserted
            $sync.Progress.Skipped += $Skipped
            $sync.Progress.Errors += $Errors
        }
        
        # Function to get recordings with enhanced error handling and rate limiting
        function Get-ZoomRecordings {
            param(
                [string]$AccessToken,
                [string]$UserId,
                [datetime]$From,
                [datetime]$To,
                [int]$PageSize = 300,
                [int]$MaxRetries = 3
            )
            $recordings = @()
            $nextPageToken = $null
            $pageCount = 0
            $fromStr = $From.ToString("yyyy-MM-dd")
            $toStr = $To.ToString("yyyy-MM-dd")
            Write-ThreadSafeLog "Querying recordings for $Account from $($fromStr) to $($toStr)"
            do {
                $pageCount++
                $url = "https://api.zoom.us/v2/users/me/recordings?from=$fromStr&to=$toStr&page_size=$PageSize"
                if ($nextPageToken) {
                    $url += "&next_page_token=$nextPageToken"
                }
                $headers = @{
                    "Authorization" = "Bearer $AccessToken"
                    "Content-Type" = "application/json"
                }
                $success = $false
                for ($retry = 1; $retry -le $MaxRetries; $retry++) {
                    try {
                        $response = Invoke-RestMethod -Uri $url -Method GET -Headers $headers
                        $success = $true
                        break
                    } catch {
                        $statusCode = $_.Exception.Response.StatusCode.value__
                        if ($statusCode -eq 429) {
                            # Rate limited - exponential backoff
                            $waitTime = [math]::Pow(2, $retry) * 5
                            Write-ThreadSafeLog "Rate limited for account $UserId, waiting $waitTime seconds..." -Level "WARNING"
                            Start-Sleep -Seconds $waitTime
                        } elseif ($statusCode -eq 404) {
                            # User not found or no recordings
                            Write-ThreadSafeLog "No recordings found for account $UserId" -Level "INFO"
                            return @()
                        } else {
                            Write-ThreadSafeLog "API error for account $UserId (attempt $retry/$MaxRetries): $_ (Status: $statusCode)" -Level "WARNING"
                            Start-Sleep -Seconds $retry
                        }
                    }
                }
                
                if (-not $success) {
                    Write-ThreadSafeLog "Failed to get recordings for account $UserId after $MaxRetries attempts" -Level "ERROR"
                    break
                }
            
                if ($response.meetings -and $response.meetings.Count -gt 0) {
                    $totalRecordingFiles = ($response.meetings | ForEach-Object { $_.recording_files.Count } | Measure-Object -Sum).Sum
                    Write-ThreadSafeLog "Account: $UserId, Page: $pageCount, Found: $totalRecordingFiles recordings over $($response.meetings.Count) meetings"
                    #Write-ThreadSafeLog "Full response: $(($response | ConvertTo-Json -Depth 10))" -Level "DEBUG"
                    $recordings += $response.meetings
                } 
                
                $nextPageToken = $response.next_page_token
                if (-not [string]::IsNullOrEmpty($nextPageToken)) {
                     Write-ThreadSafeLog "Account: $UserId, has another page for $fromStr to $toString - $nextPageTokenMsg"
                }
               
                # Rate limiting - be respectful to API
                Start-Sleep -Milliseconds 150
                
            } while ($nextPageToken)

            return $recordings
        }
        
        # Function for bulk insert using SqlBulkCopy for maximum performance
        function Invoke-BulkInsert {
            param(
                [string]$ConnectionString,
                [string]$TableName,
                [System.Data.DataTable]$DataTable
            )
            
            if ($DataTable.Rows.Count -eq 0) {
                Write-ThreadSafeLog "No records to insert for bulk insert for $UserId" -Level "INFO"
                return $true
            }
            Write-ThreadSafeLog "Starting bulk insert of $($DataTable.Rows.Count) records for $UserId" -Level "INFO"
            
            $connection = $null
            $bulkCopy = $null

            try {
                $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($ConnectionString)
                $bulkCopy.DestinationTableName = $TableName
                $bulkCopy.BatchSize = $BatchSize
                $bulkCopy.WriteToServer($DataTable)
                Write-ThreadSafeLog "Bulk inserted $($DataTable.Rows.Count) records for $UserId" -Level "INFO"
                return $true
            } catch {
                Write-ThreadSafeLog "Bulk insert failed for $UserId - $_" -Level "ERROR"
                if ($transaction) {
                    try { $transaction.Rollback() } catch { }
                }
                return $false
            } finally {
                if ($bulkCopy) { $bulkCopy.Dispose() }
                if ($transaction) { $transaction.Dispose() }
                if ($connection) { $connection.Close(); $connection.Dispose() }
            }
        }
        
        # Function to process recordings in batches
        function Process-RecordingsBatch {
            param(
                [array]$Meetings,
                [hashtable]$ExistingRecordings,
                [string]$ConnectionString,
                [string]$TableName
            )

            $totalRecordingFiles = ($meetings | ForEach-Object { $_.recording_files.Count } | Measure-Object -Sum).Sum
            Write-ThreadSafeLog "Processing batch of $totalRecordingFiles recordings over $($Meetings.Count) meetings" -Color White
            # Create DataTable for bulk insert
            $dataTable = New-Object System.Data.DataTable
            
            $dataTable.Columns.Add('GUID', [string]) | Out-Null
            $dataTable.Columns.Add('HOST_EMAIL', [string]) | Out-Null
            $dataTable.Columns.Add('RECORDING_START', [string]) | Out-Null
            $dataTable.Columns.Add('RECORDING_END', [string]) | Out-Null
            $dataTable.Columns.Add('FILE_SIZE', [string]) | Out-Null
            $dataTable.Columns.Add('DOWNLOAD_URL', [string]) | Out-Null
            $dataTable.Columns.Add('MEETING_ID', [string]) | Out-Null
            $dataTable.Columns.Add('TOPIC', [string]) | Out-Null
            $dataTable.Columns.Add('RECORDING_TYPE', [string]) | Out-Null
            $dataTable.Columns.Add('DOWNLOADED', [bool]) | Out-Null
            $dataTable.Columns.Add('TRYDLAGAIN', [int]) | Out-Null
            $dataTable.Columns.Add('DOWNLOAD_PATH', [string]) | Out-Null
            $dataTable.Columns.Add('UPLOADED', [bool]) | Out-Null
            $dataTable.Columns.Add('UPLOAD_PATH', [string]) | Out-Null
            $dataTable.Columns.Add('UPLOAD_STARTED', [datetime]) | Out-Null
            $dataTable.Columns.Add('UPLOAD_COMPLETED', [datetime]) | Out-Null
            $dataTable.Columns.Add('UPLOAD_THREAD', [string]) | Out-Null
            $dataTable.Columns.Add('UPLOAD_MESSAGE', [string]) | Out-Null

            $batchProcessed = 0
            $batchInserted = 0
            $batchSkipped = 0
            $batchErrors = 0
            
            try {
                foreach ($meeting in $Meetings) {
                    $meetingJson = $meeting | ConvertTo-Json -Depth 5
                    if (-not $meeting.recording_files) {
                        Write-ThreadSafeLog "No recording files found for meeting ID: $($meeting.id)" -Color Yellow -Level "WARNING"
                        continue
                    }
                    
                    foreach ($file in $meeting.recording_files) {
                        $batchProcessed++
                        $key1 = $file.download_url
                        $key2 = "$($meeting.id)|$($file.recording_type)|$($file.file_size)"
                        
                        if ($ExistingRecordings.ContainsKey($key1) -or $ExistingRecordings.ContainsKey($key2)) {
                            Write-ThreadSafeLog "Duplicate recording found for meeting ID: $($meeting.id) - Skipping" -Color Yellow -Level "WARNING"
                            $batchSkipped++
                            continue
                        }
                        
                        # Create new row
                        $row = $dataTable.NewRow()
                        $row['GUID'] = $file.id
                        $row['HOST_EMAIL'] = if($meeting.host_email) { $meeting.host_email } else { "" }
                        $row['RECORDING_START'] = if($file.recording_start) { $file.recording_start } else { $meeting.start_time }
                        $row['RECORDING_END'] = if($file.recording_end) { $file.recording_end } else { "" }
                        $row['FILE_SIZE'] = if($file.file_size) { [int]$file.file_size } else { [System.DBNull]::Value }
                        $row['DOWNLOAD_URL'] = if($file.download_url) { $file.download_url } else { "" }
                        $row['MEETING_ID'] = $meeting.id.ToString()
                        $row['TOPIC'] = if($meeting.topic) { $meeting.topic.Substring(0, [Math]::Min($meeting.topic.Length, 250)) } else { "" }
                        $row['RECORDING_TYPE'] = if($file.recording_type) { $file.recording_type } else { "" }
                        $row['DOWNLOADED'] = 0
                        $row['TRYDLAGAIN'] = 0
                        $row['DOWNLOAD_PATH'] = ""
                        $row['UPLOADED'] = 0
                        $row['UPLOAD_PATH'] = ""
                        $row['UPLOAD_STARTED'] = [System.DBNull]::Value
                        $row['UPLOAD_COMPLETED'] = [System.DBNull]::Value
                        $row['UPLOAD_THREAD'] = [System.DBNull]::Value
                        $row['UPLOAD_MESSAGE'] = [System.DBNull]::Value
                        
                        
                        $dataTable.Rows.Add($row)
                        $batchInserted++
                        
                        # Bulk insert when batch is full
                        if ($dataTable.Rows.Count -ge 1000) {
                            Write-ThreadSafeLog "Bulk inserting records for account: $Account" -Color White -Level "INFO"
                            if (Invoke-BulkInsert -ConnectionString $ConnectionString -TableName $TableName -DataTable $dataTable) {
                                # Optionally add custom progress tracking here if needed
                            } else {
                                $batchErrors += $dataTable.Rows.Count
                            }
                            $dataTable.Clear()
                        }
                    }
                }
                
                # Insert remaining records
                if ($dataTable.Rows.Count -gt 0) {
                    Write-ThreadSafeLog "Bulk inserting remaining records into account: $Account" -Color White -Level "INFO"

                    if (Invoke-BulkInsert -ConnectionString $ConnectionString -TableName $TableName -DataTable $dataTable) {
                        # Optionally add custom progress tracking here if needed
                    } else {
                        $batchErrors += $dataTable.Rows.Count
                    }
                }
                
                # Update final progress for this thread
                Update-ThreadProgress -Skipped $batchSkipped
                
                Write-ThreadSafeLog "Thread completed: Processed=$batchProcessed, Inserted=$batchInserted, Skipped=$batchSkipped, Errors=$batchErrors"
                
            } catch {
                Write-ThreadSafeLog "Thread error: $_" -Level "ERROR"
                Update-ThreadProgress -Errors $batchProcessed
            }
        }
        # Main worker execution            
        try {
            Write-ThreadSafeLog "Thread started for account: $Account"

            # Split date range into 30-day chunks (inclusive of last chunk)
            $chunkStart = $StartDate
            $chunkEnd = $null
            $allRecordings = @()
            while ($chunkStart -lt $EndDate) {
                $chunkEnd = $chunkStart.AddDays(29)
                if ($chunkEnd -gt $EndDate) {
                    $chunkEnd = $EndDate
                }
        
                try {
                    $chunkRecordings = Get-ZoomRecordings -AccessToken $AccessToken -UserId $Account -From $chunkStart -To $chunkEnd
                    # Ensure $chunkRecordings is always an array
                    if ($null -eq $chunkRecordings) {
                        $chunkRecordings = @()
                    } elseif ($chunkRecordings -isnot [System.Collections.IEnumerable] -or $chunkRecordings -is [string]) {
                        $chunkRecordings = @($chunkRecordings)
                    }
                    Write-ThreadSafeLog "Account: $Account, Chunk: $($chunkStart.ToString('yyyy-MM-dd')) to $($chunkEnd.ToString('yyyy-MM-dd')), Found: $($chunkRecordings.Count) meetings"
                    if ($chunkRecordings.Count -gt 0) {
                        $allRecordings += $chunkRecordings
                    }
                } catch {
                    Write-ThreadSafeLog "Error querying recordings for $Account from $($chunkStart.ToString('yyyy-MM-dd')) to $($chunkEnd.ToString('yyyy-MM-dd'))with . Exception: $_" -Level "ERROR"
                }
                $chunkStart = $chunkEnd
            }

            #Write-ThreadSafeLog "Full response: $(($allRecordings | ConvertTo-Json -Depth 10))" -Level "DEBUG"
            if ($allRecordings.Count -gt 0) {
                $totalRecordingFiles = ($allRecordings | ForEach-Object { $_.recording_files.Count } | Measure-Object -Sum).Sum
                Write-ThreadSafeLog "Account: $Account,  Processing: $totalRecordingFiles recordings over $($allRecordings.Count) meetings"
                Process-RecordingsBatch -Meetings $allRecordings -ExistingRecordings $ExistingRecordings -ConnectionString $ConnectionString -TableName $TableName
            } 

        } catch {
            Write-ThreadSafeLog "Error processing account $Account`: $_" -Level "ERROR"
        }
    }
}

# Main execution
try {
    Write-ThreadSafeLog "Starting Zoom Recordings Identification..." -Color Cyan
    
    # Load configuration using ZDAConfiguration module
    Write-ThreadSafeLog "Loading configuration using ZDAConfiguration module..."
    $config = Get-Configuration
    
    # Extract configuration values
    $MaxThreads = if ($config.runspaces.maxThreads) { $config.runspaces.maxThreads } else { 5 }
    $BatchSize = if ($config.runspaces.batchSize) { $config.runspaces.batchSize } else { 1000 }
    
    Write-ThreadSafeLog "Max Threads: $MaxThreads, Batch Size: $BatchSize" -Color Cyan
    
    # Determine date range from configuration
    $StartDate = $null
    $EndDate = Get-Date
    $DaysBack = 30
    
    if ($config.schedule.dateRange -eq "Custom start date..." -and $config.schedule.customFromDate) {
        $StartDate = [datetime]::Parse($config.schedule.customFromDate)
        Write-ThreadSafeLog "Using custom start date from config: $($StartDate.ToString('yyyy-MM-dd'))"
    } else {
        $StartDate = (Get-Date).AddDays(-$DaysBack)
        Write-ThreadSafeLog "Using default date range: last $DaysBack days"
    }
    
    Write-ThreadSafeLog "Date range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
    
    # Get Zoom access token
    Write-ThreadSafeLog "Getting Zoom access token..."
    $accessToken = Get-ZoomAccessToken -AccountId $config.zoom.accountId -ClientId $config.zoom.clientId -ClientSecret $config.zoom.clientSecret
    
    # Determine which accounts to process
    $accounts = @()
    if ($config.accounts) {
        if ($config.accounts -is [array]) {
            $accounts = $config.accounts
        } else {
            $accounts = @($config.accounts)
        }
    } else {
        # Get all users from Zoom account
        Write-ThreadSafeLog "Getting all users from Zoom account..."
        $accounts = Get-ZoomUsers -AccessToken $accessToken
    }
    
    # Handle resume functionality if configured
    $ResumeFromAccount = $null
    if ($config.resume -and $config.resume.fromAccount) {
        $ResumeFromAccount = $config.resume.fromAccount
    }
    
    # Test mode from configuration
    $TestMode = if ($config.runspaces.testMode) { $config.runspaces.testMode } else { $false }
    
    if ($TestMode) {
        Write-ThreadSafeLog "RUNNING IN TEST MODE - Limited processing" -Color Yellow
    }
    
    # Handle resume functionality if configured
    if ($ResumeFromAccount) {
        $resumeIndex = $accounts.IndexOf($ResumeFromAccount)
        if ($resumeIndex -ge 0) {
            $accounts = $accounts[$resumeIndex..($accounts.Count-1)]
            Write-ThreadSafeLog "Resuming from account: $ResumeFromAccount ($($accounts.Count) accounts remaining)"
        } else {
            Write-ThreadSafeLog "Resume account not found: $ResumeFromAccount" -Level "WARNING" -Color Yellow
        }
    }
    
    if ($TestMode) {
        $accounts = $accounts[0..([Math]::Min(2, $accounts.Count-1))]
        Write-ThreadSafeLog "Test mode: Processing only $($accounts.Count) accounts"
    }
    
    Write-ThreadSafeLog "Processing $($accounts.Count) account(s)"
    
    # Pre-load existing recordings for duplicate checking
    Write-ThreadSafeLog "Loading existing recordings for duplicate checking..."
    $existingRecordings = Get-ExistingRecordings -ConnectionString $config.database.connectionString -TableName $config.database.tableName -FromDate $StartDate
    
    # Process accounts using Runspace Pool for maximum performance
    Write-ThreadSafeLog "Creating runspace pool with $MaxThreads threads..."
    
    # Create synchronized hashtable for thread-safe communication
    $sync = [hashtable]::Synchronized(@{
        LogQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[string]
        Progress = @{
            Processed = 0
            Inserted = 0
            Skipped = 0
            Errors = 0
        }
    })
    
    # Create runspace pool
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $runspacePool.Open()
    
    # Get the worker script block
    $workerScript = Get-WorkerScriptBlock
    
    # Create runspaces for each account
    $runspaces = @()
    $threadId = 1
    
    foreach ($account in  $accounts -split "`r`n") {
        Write-ThreadSafeLog "Starting thread $threadId for account: $account" -Color Yellow
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
        
    # Add the script and parameters
    $null = $powershell.AddScript($workerScript)
    $null = $powershell.AddParameter("AccessToken", $accessToken)
    $null = $powershell.AddParameter("Account", $account)
    $null = $powershell.AddParameter("StartDate", $StartDate)
    $null = $powershell.AddParameter("EndDate", $EndDate)
    $null = $powershell.AddParameter("ConnectionString", $config.database.connectionString)
    $null = $powershell.AddParameter("TableName", $config.database.tableName)
    $null = $powershell.AddParameter("ExistingRecordings", $existingRecordings)
    $null = $powershell.AddParameter("ThreadId", $threadId)
    #$null = $powershell.AddParameter("sync", $sync)
        
        # Start the runspace
        $asyncResult = $powershell.BeginInvoke()
        
        $runspaceInfo = [PSCustomObject]@{
            PowerShell = $powershell
            AsyncResult = $asyncResult
            Account = $account
            ThreadId = $threadId
            StartTime = Get-Date
        }
        
        $runspaces += $runspaceInfo
        $threadId++
        
        Write-ThreadSafeLog "Started thread $($threadId-1) for account: $account" -Color Yellow
        
        # Stagger thread starts to avoid overwhelming the API
        Start-Sleep -Milliseconds 100
    }
    
    Write-ThreadSafeLog "All $($runspaces.Count) threads started. Monitoring progress..." -Color Cyan
    
    # Monitor progress and collect logs
    $progressTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastProgressUpdate = Get-Date
    
    do {
        # Update progress every 30 seconds
        if ((Get-Date) - $lastProgressUpdate -gt [TimeSpan]::FromSeconds(30)) {
            $currentProgress = $sync.Progress
            Write-ThreadSafeLog "=== PROGRESS UPDATE ===" -Color Cyan
            Write-ThreadSafeLog "Processed: $($currentProgress.Processed) | Inserted: $($currentProgress.Inserted) | Skipped: $($currentProgress.Skipped) | Errors: $($currentProgress.Errors)" -Color Cyan
            Write-ThreadSafeLog "Active Threads: $(($runspaces | Where-Object { -not $_.AsyncResult.IsCompleted }).Count)/$($runspaces.Count)" -Color Cyan
            Write-ThreadSafeLog "Runtime: $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" -Color Cyan
            $lastProgressUpdate = Get-Date
        } 
        
        # Check if any runspaces completed
        $completedRunspaces = $runspaces | Where-Object { $_.AsyncResult.IsCompleted -and $_.PowerShell }
        foreach ($completed in $completedRunspaces) {
            Write-ThreadSafeLog "Processing completed thread $($completed.ThreadId) for account: $($completed.Account)" -Color Green
            try {
                # Get any results/errors from the completed runspace
                $result = $completed.PowerShell.EndInvoke($completed.AsyncResult)
                $duration = (Get-Date) - $completed.StartTime
                Write-ThreadSafeLog "Thread $($completed.ThreadId) completed for account: $($completed.Account) (Duration: $($duration.ToString('mm\:ss')))" -Color Green
            } catch {
                Write-ThreadSafeLog "Thread $($completed.ThreadId) error for account $($completed.Account): $_" -Level "ERROR" -Color Red
            } finally {
                # Cleanup
                $completed.PowerShell.Dispose()
                $completed.PowerShell = $null
            }
        }

        # Short sleep to prevent excessive CPU usage
        Start-Sleep -Milliseconds 500
    } while (($runspaces | Where-Object { $_.AsyncResult -and -not $_.AsyncResult.IsCompleted }))
    
    # Final cleanup
    Write-ThreadSafeLog "All threads completed. Cleaning up runspace pool..." -Color Cyan
    
    # Dispose any remaining PowerShell objects
    foreach ($runspace in $runspaces) {
        if ($runspace.PowerShell) {
            try {
                $runspace.PowerShell.Dispose()
            } catch {
                # Ignore disposal errors
            }
        }
    }
    
    # Close and dispose the runspace pool
    $runspacePool.Close()
    $runspacePool.Dispose()
    
    # # Process any remaining log messages
    # while ($sync.LogQueue.Count -gt 0) {
    #     $logMessage = $null
    #     if ($sync.LogQueue.TryDequeue([ref]$logMessage)) {
    #         Write-Host $logMessage
    #     }
    # }
    
    # Final summary
    $finalProgress = $sync.Progress
    Write-ThreadSafeLog "`n=== FINAL SUMMARY ===" -Color Cyan
    Write-ThreadSafeLog "Total files processed: $($finalProgress.Processed)" -Color White
    Write-ThreadSafeLog "Total files inserted: $($finalProgress.Inserted)" -Color Green
    Write-ThreadSafeLog "Total files skipped (duplicates): $($finalProgress.Skipped)" -Color Yellow
    Write-ThreadSafeLog "Total errors: $($finalProgress.Errors)" -Color Red
    Write-ThreadSafeLog "Total runtime: $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" -Color Cyan
    Write-ThreadSafeLog "Import completed successfully!" -Color Green
    
} catch {
    Write-ThreadSafeLog "Script execution failed: $_" -Level "ERROR" -Color Red
    exit 1
} finally {
    # Cleanup is handled in the main execution block
}