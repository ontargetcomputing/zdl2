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
        $logFile = "zoom_download_$(Get-Date -Format 'yyyyMMdd').log"
        Add-Content -Path $logFile -Value $logMessage
    } finally {
        $script:LogMutex.ReleaseMutex()
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

# Function to get recordings to download from database per account
function Get-RecordingsToDownload {
    param(
        [string]$ConnectionString,
        [string]$TableName,
        [string]$HostEmail,
        [int]$MaxRecords = 5000  # Increased for better performance
    )
    
    $sql = @"
SELECT TOP ($MaxRecords) 
    GUID, HOST_EMAIL, RECORDING_START, RECORDING_END, FILE_SIZE, 
    DOWNLOAD_URL, MEETING_ID, TOPIC, RECORDING_TYPE, DOWNLOADED, 
    TRYDLAGAIN, DOWNLOAD_PATH
FROM $TableName 
WHERE HOST_EMAIL = '$HostEmail' 
    AND DOWNLOADED = 0 
    AND TRYDLAGAIN < 3
    AND DOWNLOAD_URL IS NOT NULL 
    AND DOWNLOAD_URL != ''
ORDER BY RECORDING_START DESC
"@
    
    try {
        $recordings = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 120
        Write-ThreadSafeLog "Found $($recordings.Count) recordings to download for $HostEmail" -Color Cyan
        return $recordings
    } catch {
        Write-ThreadSafeLog "Failed to get recordings for $HostEmail`: $_" -Level "ERROR" -Color Red
        return @()
    }
}

# Function to create download directory structure
function New-DownloadDirectory {
    param(
        [string]$BaseDownloadPath,
        [string]$HostEmail,
        [string]$MeetingId,
        [datetime]$RecordingStart
    )
    
    $sanitizedEmail = $HostEmail -replace '[\\/:*?"<>|]', '_'
    $dateFolder = $RecordingStart.ToString("yyyy-MM-dd")
    $downloadPath = Join-Path $BaseDownloadPath "$sanitizedEmail\$dateFolder\$MeetingId"
    
    if (-not (Test-Path $downloadPath)) {
        New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
    }
    
    return $downloadPath
}

# Function to create the main script block for runspace execution
function Get-WorkerScriptBlock {
    return {
        param(
            $AccessToken,
            $HostEmail, 
            $ConnectionString,
            $TableName,
            $BaseDownloadPath,
            $ThreadId,
            $MaxRecords,
            $BatchUpdateSize,
            $Sync
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
            
            $logFile = "scripts/zoom_download_${ThreadId}_$(Get-Date -Format 'yyyyMMdd').log"
            Add-Content -Path $logFile -Value $logMessage -Force
        }

        # Function to update progress in thread-safe manner
        function Update-ThreadProgress {
            param(
                [int]$Processed = 0,
                [int]$Downloaded = 0,
                [int]$Skipped = 0,
                [int]$Errors = 0
            )
            
            try {
                $Sync.Progress.Processed += $Processed
                $Sync.Progress.Downloaded += $Downloaded
                $Sync.Progress.Skipped += $Skipped
                $Sync.Progress.Errors += $Errors
                if (($Sync.Progress.Processed % 100) -eq 0 -or $Processed -gt 0) {
                    Write-ThreadSafeLog "Progress: Processed=$($Sync.Progress.Processed), Downloaded=$($Sync.Progress.Downloaded), Skipped=$($Sync.Progress.Skipped), Errors=$($Sync.Progress.Errors)" -Level "PROGRESS" -Color Cyan
                }
            } finally {
            }
        }

        # Function to get recordings to download from database per account
        function Get-RecordingsToDownload {
            param(
                [string]$ConnectionString,
                [string]$TableName,
                [string]$HostEmail,
                [int]$MaxRecords = 5000
            )
            
            $sql = @"
SELECT TOP ($MaxRecords) 
    GUID, HOST_EMAIL, RECORDING_START, RECORDING_END, FILE_SIZE, 
    DOWNLOAD_URL, MEETING_ID, TOPIC, RECORDING_TYPE, DOWNLOADED, 
    TRYDLAGAIN, DOWNLOAD_PATH
FROM $TableName 
WHERE HOST_EMAIL = '$HostEmail' 
    AND DOWNLOADED = 0 
    AND TRYDLAGAIN < 3
    AND DOWNLOAD_URL IS NOT NULL 
    AND DOWNLOAD_URL != ''
ORDER BY RECORDING_START DESC
"@
            
            try {
                $recordings = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 120
                Write-ThreadSafeLog "Found $($recordings.Count) recordings to download for $HostEmail" -Color Cyan
                return $recordings
            } catch {
                Write-ThreadSafeLog "Failed to get recordings for $HostEmail`: $_" -Level "ERROR" -Color Red
                return @()
            }
        }

        # Function to create download directory structure
        function New-DownloadDirectory {
            param(
                [string]$BaseDownloadPath,
                [string]$HostEmail,
                [string]$MeetingId,
                [datetime]$RecordingStart
            )
            
            $sanitizedEmail = $HostEmail -replace '[\\/:*?"<>|]', '_'
            $dateFolder = $RecordingStart.ToString("yyyy-MM-dd")
            $downloadPath = Join-Path $BaseDownloadPath "$sanitizedEmail\$dateFolder\$MeetingId"
            
            if (-not (Test-Path $downloadPath)) {
                New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
            }
            
            return $downloadPath
        }

        # Function to perform batch database updates
        function Update-DownloadStatusBatch {
            param(
                [string]$ConnectionString,
                [string]$TableName,
                [array]$DownloadResults,
                [int]$BatchSize = 100
            )
            
            if ($DownloadResults.Count -eq 0) {
                return $true
            }
            
            Write-ThreadSafeLog "Performing batch update of $($DownloadResults.Count) records" -Color Cyan
            
            try {
                # Process in chunks to avoid huge transactions
                
                for ($i = 0; $i -lt $DownloadResults.Count; $i += $BatchSize) {
                    $batch = $DownloadResults[$i..[Math]::Min($i + $BatchSize - 1, $DownloadResults.Count - 1)]
                    # Build a batch UPDATE statement using CASE
                    $whenClauses = @()
                    $guidList = @()
                    foreach ($result in $batch) {
                        $guidList += "'$($result.Guid)'"
                        $whenClauses += "WHEN '$($result.Guid)' THEN $($result.Downloaded)"
                    }
                    
                    $downloadedCase = "CASE GUID " + ($whenClauses -join " ") + " END"
                    # Similar for download path and retry count
                    $pathClauses = @()
                    $retryClauses = @()
                    foreach ($result in $batch) {
                        $escapedPath = $result.DownloadPath -replace "'", "''"  # Escape single quotes
                        $pathClauses += "WHEN '$($result.Guid)' THEN '$escapedPath'"
                        $retryClauses += "WHEN '$($result.Guid)' THEN $($result.TryDlAgain)"
                    }                    

                    
                    $pathCase = "CASE GUID " + ($pathClauses -join " ") + " END"
                    $retryCase = "CASE GUID " + ($retryClauses -join " ") + " END"
                    
                    $sql = @"
UPDATE $TableName 
SET DOWNLOADED = $downloadedCase,
    DOWNLOAD_PATH = $pathCase,
    TRYDLAGAIN = $retryCase
WHERE GUID IN ($($guidList -join ','))
"@
                    Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 120
                    Write-ThreadSafeLog "Batch updated $($batch.Count) records" -Color Green
                }
                
                return $true
            } catch {
                Write-ThreadSafeLog "Batch update failed: $_" -Level "ERROR" -Color Red
                return $false
            }
        }
        
        # Function to download a recording file with enhanced error handling and retry logic
        function Download-Recording {
            param(
                [string]$DownloadUrl,
                [string]$AccessToken,
                [string]$FilePath,
                [long]$ExpectedSize = 0,
                [int]$MaxRetries = 3
            )
            
            $headers = @{
                "Content-Type"  = "application/x-www-form-urlencoded"
                "Authorization" = "Bearer $AccessToken"
            }
            
            for ($retry = 1; $retry -le $MaxRetries; $retry++) {
                try {
                    Write-ThreadSafeLog "Downloading file (attempt $retry/$MaxRetries): $FilePath"
                    
                    # Use Invoke-WebRequest for better control over downloads
                    $response = Invoke-WebRequest -Uri $DownloadUrl -Headers $headers -OutFile $FilePath -PassThru
                    
                    # Verify file was downloaded
                    if (Test-Path $FilePath) {
                        $fileInfo = Get-Item $FilePath
                        Write-ThreadSafeLog "Download completed: $($fileInfo.Name) ($($fileInfo.Length) bytes)" -Color Green
                        
                        # Optional: Verify file size if expected size is provided
                        if ($ExpectedSize -gt 0 -and $fileInfo.Length -ne $ExpectedSize) {
                            Write-ThreadSafeLog "Warning: Downloaded file size ($($fileInfo.Length)) doesn't match expected size ($ExpectedSize)" -Level "WARNING" -Color Yellow
                        }
                        
                        return @{
                            Success = $true
                            Message = "Download successful"
                            Attempts = $retry
                        }
                    } else {
                        throw "File was not created after download"
                    }
                    
                } catch {
                    $errorMessage = $_.Exception.Message
                    if ($_.Exception.Response) {
                        $statusCode = $_.Exception.Response.StatusCode.value__
                        Write-ThreadSafeLog "Download failed (attempt $retry/$MaxRetries) - Status: $statusCode, Error: $errorMessage" -Level "WARNING" -Color Yellow
                        
                        if ($statusCode -eq 401) {
                            Write-ThreadSafeLog "Authentication failed - token may be expired" -Level "ERROR" -Color Red
                            return @{
                                Success = $false
                                Message = "Authentication failed"
                                Attempts = $retry
                            }
                        } elseif ($statusCode -eq 404) {
                            Write-ThreadSafeLog "File not found - URL may be expired or invalid" -Level "ERROR" -Color Red
                            return @{
                                Success = $false
                                Message = "File not found"
                                Attempts = $retry
                            }
                        } elseif ($statusCode -eq 429) {
                            # Rate limited - exponential backoff
                            $waitTime = [math]::Pow(2, $retry) * 5
                            Write-ThreadSafeLog "Rate limited, waiting $waitTime seconds..." -Level "WARNING" -Color Yellow
                            Start-Sleep -Seconds $waitTime
                        }
                    } else {
                        Write-ThreadSafeLog "Download failed (attempt $retry/$MaxRetries): $errorMessage" -Level "WARNING" -Color Yellow
                    }
                    
                    if ($retry -eq $MaxRetries) {
                        Write-ThreadSafeLog "Download failed after $MaxRetries attempts: $errorMessage" -Level "ERROR" -Color Red
                        return @{
                            Success = $false
                            Message = "Failed after $MaxRetries attempts: $errorMessage"
                            Attempts = $retry
                        }
                    }
                    
                    # Clean up partial download
                    if (Test-Path $FilePath) {
                        Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
                    }
                    
                    Start-Sleep -Seconds $retry
                }
            }
            
            return @{
                Success = $false
                Message = "Download failed"
                Attempts = $MaxRetries
            }
        }
        
        # Function to process recordings for download with batching
        function Process-RecordingsDownload {
            param(
                [array]$Recordings,
                [string]$AccessToken,
                [string]$BaseDownloadPath,
                [string]$ConnectionString,
                [string]$TableName
            )

            Write-ThreadSafeLog "Processing $($Recordings.Count) recordings for download" -Color White
            
            $batchProcessed = 0
            $batchDownloaded = 0
            $batchSkipped = 0
            $batchErrors = 0
            $downloadResults = @()
            
            # Performance tracking
            $startTime = Get-Date
            
            try {
                foreach ($recording in $Recordings) {
                    $batchProcessed++
                    
                    # Skip if already downloaded
                    if ($recording.DOWNLOADED -eq 1) {
                        Write-ThreadSafeLog "Recording already downloaded: $($recording.GUID) - Skipping" -Color Yellow -Level "INFO"
                        $batchSkipped++
                        continue
                    }
                    
                    # Skip if too many retry attempts
                    if ($recording.TRYDLAGAIN -ge 3) {
                        Write-ThreadSafeLog "Recording has too many failed attempts: $($recording.GUID) - Skipping" -Color Yellow -Level "WARNING"
                        $batchSkipped++
                        continue
                    }
                    
                    try {
                        # Create download directory
                        $recordingStart = [datetime]::Parse($recording.RECORDING_START)
                        $downloadDir = New-DownloadDirectory -BaseDownloadPath $BaseDownloadPath -HostEmail $recording.HOST_EMAIL -MeetingId $recording.MEETING_ID -RecordingStart $recordingStart
                        
                        # Generate filename
                        $sanitizedTopic = if ($recording.TOPIC) { 
                            ($recording.TOPIC -replace '[\\/:*?"<>|]', '_').Substring(0, [Math]::Min($recording.TOPIC.Length, 100))
                        } else { 
                            "Recording" 
                        }
                        
                        $fileExtension = switch ($recording.RECORDING_TYPE.ToLower()) {
                            "shared_screen_with_speaker_view" { ".mp4" }
                            "audio_only" { ".m4a" }
                            "chat_file" { ".txt" }
                            "transcript" { ".vtt" }
                            default { ".mp4" }
                        }
                        
                        $fileName = "$($recording.MEETING_ID)_$($recording.RECORDING_TYPE)_$sanitizedTopic$fileExtension"
                        $filePath = Join-Path $downloadDir $fileName
                        
                        # Download the file
                        Write-ThreadSafeLog "Starting download: $($recording.HOST_EMAIL) - $($recording.MEETING_ID) - $($recording.RECORDING_TYPE)"
                        
                        $downloadResult = Download-Recording -DownloadUrl $recording.DOWNLOAD_URL -AccessToken $AccessToken -FilePath $filePath -ExpectedSize $recording.FILE_SIZE
                        
                        # Collect result for batch update (success or failure)
                        if ($downloadResult.Success) {
                            $downloadResults += @{
                                Guid = $recording.GUID
                                Downloaded = 1
                                DownloadPath = $filePath
                                TryDlAgain = $recording.TRYDLAGAIN
                            }
                            $batchDownloaded++
                            Write-ThreadSafeLog "Successfully downloaded: $fileName" -Color Green
                        } else {
                            # Increment retry counter
                            $newRetryCount = $recording.TRYDLAGAIN + 1
                            $downloadResults += @{
                                Guid = $recording.GUID
                                Downloaded = 0
                                DownloadPath = ""
                                TryDlAgain = $newRetryCount
                            }
                            $batchErrors++
                            Write-ThreadSafeLog "Failed to download: $($recording.GUID) (Retry count: $newRetryCount)" -Level "ERROR" -Color Red
                        }
                        
                        # Perform batch update when we reach batch size
                        if ($downloadResults.Count -ge $BatchUpdateSize) {
                            Write-ThreadSafeLog "Performing batch database update for $($downloadResults.Count) records" -Color Cyan
                            $updateSuccess = Update-DownloadStatusBatch -ConnectionString $ConnectionString -TableName $TableName -DownloadResults $downloadResults
                            if ($updateSuccess) {
                                Write-ThreadSafeLog "Batch database update completed successfully" -Color Green
                            } else {
                                Write-ThreadSafeLog "Batch database update failed" -Level "ERROR" -Color Red
                            }
                            $downloadResults = @()  # Reset for next batch
                        }
                        
                    } catch {
                        Write-ThreadSafeLog "Error processing recording $($recording.GUID): $_" -Level "ERROR" -Color Red
                        
                        # Add error result to batch
                        $newRetryCount = $recording.TRYDLAGAIN + 1
                        $downloadResults += @{
                            Guid = $recording.GUID
                            Downloaded = 0
                            DownloadPath = ""
                            TryDlAgain = $newRetryCount
                        }
                        $batchErrors++
                    }
                    
                    # Add a small delay between downloads to be respectful
                    Start-Sleep -Milliseconds 100  # Reduced for better performance
                }
                
                # Process any remaining results in final batch
                if ($downloadResults.Count -gt 0) {
                    Write-ThreadSafeLog "Performing final batch database update for $($downloadResults.Count) records" -Color Cyan
                    $updateSuccess = Update-DownloadStatusBatch -ConnectionString $ConnectionString -TableName $TableName -DownloadResults $downloadResults
                    if ($updateSuccess) {
                        Write-ThreadSafeLog "Final batch database update completed successfully" -Color Green
                    } else {
                        Write-ThreadSafeLog "Final batch database update failed" -Level "ERROR" -Color Red
                    }
                }
                
                # Performance summary for this thread
                $duration = (Get-Date) - $startTime
                $filesPerMinute = if ($duration.TotalMinutes -gt 0) { [math]::Round($batchProcessed / $duration.TotalMinutes, 1) } else { 0 }
                Write-ThreadSafeLog "Thread performance: $filesPerMinute files/minute over $($duration.ToString('hh\:mm\:ss'))" -Color Magenta
                
                # Update final progress for this thread
                Update-ThreadProgress -Processed $batchProcessed -Downloaded $batchDownloaded -Skipped $batchSkipped -Errors $batchErrors
                
                Write-ThreadSafeLog "Thread completed: Processed=$batchProcessed, Downloaded=$batchDownloaded, Skipped=$batchSkipped, Errors=$batchErrors"
                
            } catch {
                Write-ThreadSafeLog "Thread error: $_" -Level "ERROR" -Color Red
                Update-ThreadProgress -Errors $batchProcessed
            }
        }

        # Main worker execution            
        try {
            Write-ThreadSafeLog "Thread started for account: $HostEmail"
            # Get recordings to download for this account
            $recordingsToDownload = Get-RecordingsToDownload -ConnectionString $ConnectionString -TableName $TableName -HostEmail $HostEmail -MaxRecords $MaxRecords

            Write-ThreadSafeLog "Account: $HostEmail, Found: $($recordingsToDownload.Count) recordings to download"
            
            if ($recordingsToDownload.Count -gt 0) {
                Write-ThreadSafeLog "Starting download to: $BaseDownloadPath" -Color Cyan
                Process-RecordingsDownload -Recordings $recordingsToDownload -AccessToken $AccessToken -BaseDownloadPath $BaseDownloadPath -ConnectionString $ConnectionString -TableName $TableName
            } else {
                Write-ThreadSafeLog "No recordings to download for account: $HostEmail" -Color Yellow
            }

        } catch {
            Write-ThreadSafeLog "Error processing account $HostEmail`: $_" -Level "ERROR" -Color Red
        }
    }
}

# Main execution
try {
    Write-ThreadSafeLog "Starting Zoom Recordings Download..." -Color Cyan
    
    # Load configuration using ZDAConfiguration module
    Write-ThreadSafeLog "Loading configuration using ZDAConfiguration module..."
    $config = Get-Configuration
    
    # Extract configuration values - OPTIMIZED FOR PERFORMANCE
    $MaxThreads = if ($config.runspaces.maxThreads) { $config.runspaces.maxThreads } else { 10 }
    $MaxRecordsPerThread = if ($config.runspaces.maxRecordsPerThread) { $config.runspaces.maxRecordsPerThread } else { 5000 }
    $BatchUpdateSize = if ($config.runspaces.batchUpdateSize) { $config.runspaces.batchUpdateSize } else { 100 }
    $BaseDownloadPath = if ($config.download.basepath) { $config.download.basepath } else { ".\Downloads" }
    
    Write-ThreadSafeLog "PERFORMANCE MODE: Max Threads: $MaxThreads, Records/Thread: $MaxRecordsPerThread, Batch Size: $BatchUpdateSize" -Color Cyan
    Write-ThreadSafeLog "Download Path: $BaseDownloadPath" -Color Cyan
    
    # Ensure download directory exists
    if (-not (Test-Path $BaseDownloadPath)) {
        New-Item -Path $BaseDownloadPath -ItemType Directory -Force | Out-Null
        Write-ThreadSafeLog "Created download directory: $BaseDownloadPath" -Color Green
    }
    
    # Get Zoom access token
    Write-ThreadSafeLog "Getting Zoom access token..."
    $accessToken = Get-ZoomAccessToken -AccountId $config.zoom.accountId -ClientId $config.zoom.clientId -ClientSecret $config.zoom.clientSecret
    
    # Get unique host emails from database that have recordings to download
    $sql = @"
SELECT DISTINCT HOST_EMAIL 
FROM $($config.database.tableName) 
WHERE DOWNLOADED = 0 
    AND TRYDLAGAIN < 3
    AND DOWNLOAD_URL IS NOT NULL 
    AND DOWNLOAD_URL != ''
    AND HOST_EMAIL IS NOT NULL 
    AND HOST_EMAIL != ''
ORDER BY HOST_EMAIL
"@
    
    Write-ThreadSafeLog "Getting list of accounts with recordings to download..."
    $hostEmails = Invoke-Sqlcmd -ConnectionString $config.database.connectionString -Query $sql -QueryTimeout 120 | Select-Object -ExpandProperty HOST_EMAIL
    
    if (-not $hostEmails -or $hostEmails.Count -eq 0) {
        Write-ThreadSafeLog "No accounts found with recordings to download" -Color Yellow
        exit 0
    }
    
    # Get total count for performance planning
    $totalCountSql = @"
SELECT COUNT(*) as Total
FROM $($config.database.tableName) 
WHERE DOWNLOADED = 0 
    AND TRYDLAGAIN < 3
    AND DOWNLOAD_URL IS NOT NULL 
    AND DOWNLOAD_URL != ''
"@
    $totalFiles = (Invoke-Sqlcmd -ConnectionString $config.database.connectionString -Query $totalCountSql -QueryTimeout 120).Total
    
    Write-ThreadSafeLog "PERFORMANCE ANALYSIS:" -Color Magenta
    Write-ThreadSafeLog "Total files to download: $totalFiles" -Color Magenta
    if ($totalFiles -gt 0) {
        Write-ThreadSafeLog "With $MaxThreads threads: $([math]::Round($totalFiles / $MaxThreads, 0)) files per thread average" -Color Magenta
        Write-ThreadSafeLog "Expected database calls: ~$([math]::Round($totalFiles / $BatchUpdateSize, 0)) (was $totalFiles without batching)" -Color Magenta
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
        $resumeIndex = $hostEmails.IndexOf($ResumeFromAccount)
        if ($resumeIndex -ge 0) {
            $hostEmails = $hostEmails[$resumeIndex..($hostEmails.Count-1)]
            Write-ThreadSafeLog "Resuming from account: $ResumeFromAccount ($($hostEmails.Count) accounts remaining)"
        } else {
            Write-ThreadSafeLog "Resume account not found: $ResumeFromAccount" -Level "WARNING" -Color Yellow
        }
    }
    
    if ($TestMode) {
        $hostEmails = $hostEmails[0..([Math]::Min(2, $hostEmails.Count-1))]
        Write-ThreadSafeLog "Test mode: Processing only $($hostEmails.Count) accounts"
    }
    
    Write-ThreadSafeLog "Processing $($hostEmails.Count) account(s) for download"
    
    # Process accounts using Runspace Pool for maximum performance
    Write-ThreadSafeLog "Creating runspace pool with $MaxThreads threads..."
    
    # Create synchronized hashtable for thread-safe communication
    $sync = [hashtable]::Synchronized(@{
        LogQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[string]
        Progress = @{
            Processed = 0
            Downloaded = 0
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
    $accountIndex = 0
    foreach ($hostEmail in $hostEmails) {
        # Calculate thread ID based on max threads (cycle through 1 to MaxThreads)
        $threadId = ($accountIndex % $MaxThreads) + 1
        
        Write-ThreadSafeLog "Starting thread $threadId for account: $hostEmail" -Color Yellow
        
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
                    
        # Add the script and parameters
        $null = $powershell.AddScript($workerScript)
        $null = $powershell.AddParameter("AccessToken", $accessToken)
        $null = $powershell.AddParameter("HostEmail", $hostEmail)
        $null = $powershell.AddParameter("ConnectionString", $config.database.connectionString)
        $null = $powershell.AddParameter("TableName", $config.database.tableName)
        $null = $powershell.AddParameter("BaseDownloadPath", $BaseDownloadPath)
        $null = $powershell.AddParameter("ThreadId", $threadId)
        $null = $powershell.AddParameter("MaxRecords", $MaxRecordsPerThread)
        $null = $powershell.AddParameter("BatchUpdateSize", $BatchUpdateSize)
        $null = $powershell.AddParameter("Sync", $sync)

        # Start the runspace
        $asyncResult = $powershell.BeginInvoke()
        
        $runspaceInfo = [PSCustomObject]@{
            PowerShell = $powershell
            AsyncResult = $asyncResult
            HostEmail = $hostEmail
            ThreadId = $threadId
            AccountIndex = $accountIndex
            StartTime = Get-Date
        }
        
        $runspaces += $runspaceInfo        
        Write-ThreadSafeLog "Started thread for account: $hostEmail (Thread ID: $threadId, Account Index: $accountIndex)" -Color Yellow
        $accountIndex++

        # Stagger thread starts to avoid overwhelming the API
        Start-Sleep -Milliseconds 200  # Reduced for faster startup
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
            Write-ThreadSafeLog "Processed: $($currentProgress.Processed) | Downloaded: $($currentProgress.Downloaded) | Skipped: $($currentProgress.Skipped) | Errors: $($currentProgress.Errors)" -Color Cyan
            Write-ThreadSafeLog "Active Threads: $(($runspaces | Where-Object { -not $_.AsyncResult.IsCompleted }).Count)/$($runspaces.Count)" -Color Cyan
            Write-ThreadSafeLog "Runtime: $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" -Color Cyan
            
            # Calculate download rate
            if ($currentProgress.Downloaded -gt 0 -and $progressTimer.Elapsed.TotalMinutes -gt 0) {
                $downloadRate = [math]::Round($currentProgress.Downloaded / $progressTimer.Elapsed.TotalMinutes, 1)
                Write-ThreadSafeLog "Download rate: $downloadRate files/minute" -Color Cyan
            }
            
            $lastProgressUpdate = Get-Date
        } 
        
        # Check if any runspaces completed
        $completedRunspaces = $runspaces | Where-Object { $_.AsyncResult.IsCompleted -and $_.PowerShell }
        foreach ($completed in $completedRunspaces) {
            Write-ThreadSafeLog "Processing completed thread $($completed.ThreadId) for account: $($completed.HostEmail)" -Color Green
            try {
                # Get any results/errors from the completed runspace
                $result = $completed.PowerShell.EndInvoke($completed.AsyncResult)
                $duration = (Get-Date) - $completed.StartTime
                Write-ThreadSafeLog "Thread $($completed.ThreadId) completed for account: $($completed.HostEmail) (Duration: $($duration.ToString('mm\:ss')))" -Color Green
            } catch {
                Write-ThreadSafeLog "Thread $($completed.ThreadId) error for account $($completed.HostEmail): $_" -Level "ERROR" -Color Red
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
    
    # Final summary
    $finalProgress = $sync.Progress
    Write-ThreadSafeLog "`n=== FINAL SUMMARY ===" -Color Cyan
    Write-ThreadSafeLog "Total files processed: $($finalProgress.Processed)" -Color White
    Write-ThreadSafeLog "Total files downloaded: $($finalProgress.Downloaded)" -Color Green
    Write-ThreadSafeLog "Total files skipped (already downloaded/too many retries): $($finalProgress.Skipped)" -Color Yellow
    Write-ThreadSafeLog "Total errors: $($finalProgress.Errors)" -Color Red
    Write-ThreadSafeLog "Total runtime: $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" -Color Cyan
    
    # Performance summary
    if ($finalProgress.Downloaded -gt 0 -and $progressTimer.Elapsed.TotalMinutes -gt 0) {
        $avgDownloadRate = [math]::Round($finalProgress.Downloaded / $progressTimer.Elapsed.TotalMinutes, 1)
        Write-ThreadSafeLog "Average download rate: $avgDownloadRate files/minute" -Color Cyan
        
        $estimatedDbCalls = [math]::Round($finalProgress.Downloaded / $BatchUpdateSize, 0)
        $savedDbCalls = $finalProgress.Downloaded - $estimatedDbCalls
        Write-ThreadSafeLog "Database optimization: ~$estimatedDbCalls calls (saved ~$savedDbCalls calls with batching)" -Color Cyan
    }
    
    Write-ThreadSafeLog "Download completed successfully!" -Color Green
    Write-ThreadSafeLog "Files downloaded to: $BaseDownloadPath" -Color Cyan
    
} catch {
    Write-ThreadSafeLog "Script execution failed: $_" -Level "ERROR" -Color Red
    exit 1
} finally {
    # Cleanup is handled in the main execution block
}