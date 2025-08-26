# Import the new configuration module
Import-Module "$PSScriptRoot\ZDAConfiguration.psm1"

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
        
        $logFileName = "zoom_upload_$(Get-Date -Format 'yyyyMMdd').log"
        $logPath = Get-LogsPath
        $logFile = Join-Path -Path $logPath -ChildPath $logFileName
        Add-Content -Path $logFile -Value $logMessage
    } finally {
        $script:LogMutex.ReleaseMutex()
    }
}

# Function to load configuration using ZDAConfiguration module
function Get-Configuration {
    try {
        $config = Get-UserConfiguration

        if (-not $config) {
            throw "Failed to read user configuration"
        }
        
        Write-ThreadSafeLog "Configuration loaded successfully" -Color Green
        return $config
    } catch {
        throw "Failed to load configuration: $_"
    }
}

# Function to create the main script block for runspace execution
function Get-WorkerScriptBlock {
    return {
        param(
            $HostEmail, 
            $ConnectionString,
            $TableName,
            $UploadConfig,
            $ThreadId,
            $LogPath,
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
            
            $logFileName = "zoom_upload_${ThreadId}_$(Get-Date -Format 'yyyyMMdd').log"
            $logFile = Join-Path -Path $LogPath -ChildPath $logFileName
            Add-Content -Path $logFile -Value $logMessage -Force
        }

        # Function to update progress in thread-safe manner
        function Update-ThreadProgress {
            param(
                [int]$Processed = 0,
                [int]$Uploaded = 0,
                [int]$Skipped = 0,
                [int]$Errors = 0
            )
            try {
                $Sync.Progress.Processed += $Processed
                $Sync.Progress.Uploaded += $Uploaded
                $Sync.Progress.Skipped += $Skipped
                $Sync.Progress.Errors += $Errors
                if (($Sync.Progress.Processed % 100) -eq 0 -or $Processed -gt 0) {
                    Write-ThreadSafeLog "Progress: Processed=$($Sync.Progress.Processed), Uploaded=$($Sync.Progress.Uploaded), Skipped=$($Sync.Progress.Skipped), Errors=$($Sync.Progress.Errors)" -Level "PROGRESS" -Color Cyan
                }
            } finally {
            }
        }

        # Function to get recordings to upload from database per account
        function Get-RecordingsToUpload {
            param(
                [string]$ConnectionString,
                [string]$TableName,
                [string]$HostEmail,
                [int]$MaxRecords = 5000  # Increased for performance
            )
            
            $sql = @"
SELECT TOP ($MaxRecords) 
    GUID, HOST_EMAIL, RECORDING_START, RECORDING_END, FILE_SIZE, 
    DOWNLOAD_URL, MEETING_ID, TOPIC, RECORDING_TYPE, DOWNLOADED, 
    DOWNLOAD_PATH, UPLOADED, UPLOAD_PATH, UPLOAD_COMPLETED
FROM $TableName 
WHERE HOST_EMAIL = '$HostEmail' 
    AND DOWNLOADED = 1 
    AND UPLOADED = 0
    AND DOWNLOAD_PATH IS NOT NULL 
    AND DOWNLOAD_PATH != ''
ORDER BY RECORDING_START DESC
"@
            #Write-ThreadSafeLog "Executing SQL query to get recordings $sql" -Color Yellow
            try {
                $recordings = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $sql -QueryTimeout 120
                Write-ThreadSafeLog "Found $($recordings.Count) recordings to upload for $HostEmail" -Color Cyan
                return $recordings
            } catch {
                Write-ThreadSafeLog "Failed to get recordings for $HostEmail`: $_" -Level "ERROR" -Color Red
                return @()
            }
        }

        # Function to upload to S3
        function Upload-ToS3 {
            param(
                [string]$FilePath,
                [string]$S3BucketName,
                [string]$S3Key,
                [string]$AccessKeyId,
                [string]$SecretAccessKey,
                [string]$Region = "us-east-1",
                [int]$MaxRetries = 3
            )
            
            try {
                # Set AWS credentials
                $env:AWS_ACCESS_KEY_ID = $AccessKeyId
                $env:AWS_SECRET_ACCESS_KEY = $SecretAccessKey
                $env:AWS_DEFAULT_REGION = $Region
                
                for ($retry = 1; $retry -le $MaxRetries; $retry++) {
                    try {
                        Write-ThreadSafeLog "Uploading to S3 (attempt $retry/$MaxRetries): s3://$S3BucketName/$S3Key"
                        
                        Write-S3Object -BucketName $S3BucketName -Key $S3Key -File $FilePath
                        Write-ThreadSafeLog "S3 upload completed: s3://$S3BucketName/$S3Key" -Color Green
                        return @{
                            Success = $true
                            UploadPath = "s3://$S3BucketName/$S3Key"
                            Message = "Upload successful"
                        }
                    } catch {
                        Write-ThreadSafeLog "S3 upload failed (attempt $retry/$MaxRetries): $_" -Level "WARNING" -Color Yellow
                        if ($retry -eq $MaxRetries) {
                            return @{
                                Success = $false
                                UploadPath = ""
                                Message = "S3 upload failed after $MaxRetries attempts: $_"
                            }
                        }
                        Start-Sleep -Seconds ($retry * 2)
                    }
                }
                
            } catch {
                return @{
                    Success = $false
                    UploadPath = ""
                    Message = "S3 upload error: $_"
                }
            } finally {
                # Clean up environment variables
                Remove-Item env:AWS_ACCESS_KEY_ID -ErrorAction SilentlyContinue
                Remove-Item env:AWS_SECRET_ACCESS_KEY -ErrorAction SilentlyContinue
                Remove-Item env:AWS_DEFAULT_REGION -ErrorAction SilentlyContinue
            }
        }

        # Function to upload to OneDrive using Microsoft Graph API (Admin accessing user drives)
        function Upload-ToOneDrive {
            param(
                [string]$FilePath,
                [string]$OneDrivePath,
                [string]$ClientId,
                [string]$ClientSecret,
                [string]$TenantId,
                [string]$UserEmail,  # Target user's email
                [int]$MaxRetries = 3
            )
            Write-ThreadSafeLog "Starting OneDrive upload for $UserEmail - $OneDrivePath"
            try {
                # Get access token for Microsoft Graph with admin permissions
                $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                $tokenBody = @{
                    client_id = $ClientId
                    client_secret = $ClientSecret
                    scope = "https://graph.microsoft.com/.default"  # Admin scope
                    grant_type = "client_credentials"
                }
                
                $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $tokenBody
                $accessToken = $tokenResponse.access_token
                
                for ($retry = 1; $retry -le $MaxRetries; $retry++) {
                    try {
                        Write-ThreadSafeLog "Uploading to $UserEmail's OneDrive (attempt $retry/$MaxRetries): $OneDrivePath"
                        
                        # Read file content
                        $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
                        $fileName = Split-Path $FilePath -Leaf
                        
                        # Upload file to specific user's OneDrive using admin access
                        # Note: Using /users/{email}/drive instead of /me/drive
                        $uploadUrl = "https://graph.microsoft.com/v1.0/users/$UserEmail/drive/root:/$OneDrivePath/$fileName" + ":/content"
                        $headers = @{
                            "Authorization" = "Bearer $accessToken"
                            "Content-Type" = "application/octet-stream"
                        }
                        
                        $response = Invoke-RestMethod -Uri $uploadUrl -Method PUT -Headers $headers -Body $fileBytes
                        
                        Write-ThreadSafeLog "OneDrive upload completed to $UserEmail's drive: $OneDrivePath/$fileName" -Color Green
                        return @{
                            Success = $true
                            UploadPath = "$UserEmail/OneDrive/$OneDrivePath/$fileName"
                            Message = "Upload successful to $UserEmail's OneDrive"
                        }
                        
                    } catch {
                        $errorMessage = $_.Exception.Message
                        $statusCode = $null
                        
                        if ($_.Exception.Response) {
                            $statusCode = $_.Exception.Response.StatusCode.value__
                        }
                        
                        Write-ThreadSafeLog "OneDrive upload failed for $UserEmail (attempt $retry/$MaxRetries) - Status: $statusCode, Error: $errorMessage" -Level "WARNING" -Color Yellow
                        
                        # Handle specific errors
                        if ($statusCode -eq 401) {
                            Write-ThreadSafeLog "Authentication failed - check admin permissions for accessing $UserEmail's OneDrive" -Level "ERROR" -Color Red
                            return @{
                                Success = $false
                                UploadPath = ""
                                Message = "Authentication failed for $UserEmail's OneDrive"
                            }
                        } elseif ($statusCode -eq 403) {
                            Write-ThreadSafeLog "Forbidden - admin may not have access to $UserEmail's OneDrive" -Level "ERROR" -Color Red
                            return @{
                                Success = $false
                                UploadPath = ""
                                Message = "Access forbidden to $UserEmail's OneDrive"
                            }
                        } elseif ($statusCode -eq 404) {
                            Write-ThreadSafeLog "User $UserEmail not found or OneDrive not provisioned" -Level "ERROR" -Color Red
                            return @{
                                Success = $false
                                UploadPath = ""
                                Message = "User $UserEmail not found or OneDrive not available"
                            }
                        }
                        
                        if ($retry -eq $MaxRetries) {
                            return @{
                                Success = $false
                                UploadPath = ""
                                Message = "OneDrive upload failed for $UserEmail after $MaxRetries attempts: $errorMessage"
                            }
                        }
                        Start-Sleep -Seconds ($retry * 2)
                    }
                }
                
            } catch {
                return @{
                    Success = $false
                    UploadPath = ""
                    Message = "OneDrive upload error for $UserEmail`: $_"
                }
            }
        }

        # Function to upload a recording file
        function Upload-Recording {
            param(
                [string]$FilePath,
                [object]$UploadConfig,
                [string]$HostEmail,
                [string]$MeetingId,
                [datetime]$RecordingStart,
                [string]$RecordingType,
                [int]$MaxRetries = 3
            )
            
            # Verify file exists
            if (-not (Test-Path $FilePath)) {
                return @{
                    Success = $false
                    UploadPath = ""
                    Message = "File not found: $FilePath"
                }
            }
            
            $fileInfo = Get-Item $FilePath
            $fileName = $fileInfo.Name
            $sanitizedEmail = $HostEmail -replace '[\\/:*?"<>|]', '_'
            $dateFolder = $RecordingStart.ToString("yyyy-MM-dd")
            
            # Determine upload destination based on configuration
            if ($UploadConfig.provider -eq "s3") {
                $s3Key = "$sanitizedEmail/$dateFolder/$MeetingId/$fileName"
                return Upload-ToS3 -FilePath $FilePath -S3BucketName $UploadConfig.s3.bucketName -S3Key $s3Key -AccessKeyId $UploadConfig.s3.accessKeyId -SecretAccessKey $UploadConfig.s3.secretAccessKey -Region $UploadConfig.s3.region
            }
            elseif ($UploadConfig.provider -eq "onedrive") {
                $oneDrivePath = "ZoomRecordings/$sanitizedEmail/$dateFolder/$MeetingId"
                return Upload-ToOneDrive -FilePath $FilePath -OneDrivePath $oneDrivePath -ClientId $UploadConfig.onedrive.clientId -ClientSecret $UploadConfig.onedrive.clientSecret -TenantId $UploadConfig.onedrive.tenantId -UserEmail $HostEmail
            }
            else {
                return @{
                    Success = $false
                    UploadPath = ""
                    Message = "Unknown upload provider: $($UploadConfig.provider)"
                }
            }
        }
        
        # Function to perform batch database updates (simplified)
        function Update-UploadStatusBatch {
            param(
                [string]$ConnectionString,
                [string]$TableName,
                [array]$UploadResults,
                [int]$BatchSize = 100
            )
            
            if ($UploadResults.Count -eq 0) {
                return $true
            }
            
            Write-ThreadSafeLog "Performing batch update of $($UploadResults.Count) records" -Color Cyan
            
            try {
                # Process in chunks to avoid huge transactions
                for ($i = 0; $i -lt $UploadResults.Count; $i += $BatchSize) {
                    $batch = $UploadResults[$i..[Math]::Min($i + $BatchSize - 1, $UploadResults.Count - 1)]
                    
                    # Build a batch UPDATE statement using CASE
                    $whenClauses = @()
                    $guidList = @()
                    
                    foreach ($result in $batch) {
                        $guidList += "'$($result.Guid)'"
                        $whenClauses += "WHEN '$($result.Guid)' THEN $($result.Uploaded)"
                    }
                    
                    $uploadedCase = "CASE GUID " + ($whenClauses -join " ") + " END"
                    
                    # Similar for upload path
                    $pathClauses = @()
                    

                   foreach ($result in $batch) {
                        $escapedPath = $result.UploadPath -replace "'", "''"  # Escape single quotes
                        $pathClauses += "WHEN '$($result.Guid)' THEN '$escapedPath'"
                        $retryClauses += "WHEN '$($result.Guid)' THEN $($result.TryDlAgain)"
                    }                    

                    
                    $pathCase = "CASE GUID " + ($pathClauses -join " ") + " END"
                    $retryCase = "CASE GUID " + ($retryClauses -join " ") + " END"

                    
                    foreach ($result in $batch) {
                        $escapedPath = $result.UploadPath -replace "'", "''"  # Escape single quotes
                        $pathClauses += "WHEN '$($result.Guid)' THEN '$escapedPath'"
                    }
                    
                    $pathCase = "CASE GUID " + ($pathClauses -join " ") + " END"
                    
                    $sql = @"
UPDATE $TableName 
SET UPLOADED = $uploadedCase,
    UPLOAD_PATH = $pathCase,
    UPLOAD_COMPLETED = CASE WHEN $uploadedCase = 1 THEN GETDATE() ELSE NULL END
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

        # Function to process recordings for upload
        function Process-RecordingsUpload {
            param(
                [array]$Recordings,
                [object]$UploadConfig,
                [string]$ConnectionString,
                [string]$TableName
            )

            Write-ThreadSafeLog "Processing $($Recordings.Count) recordings for upload" -Color White
            
            $batchProcessed = 0
            $batchUploaded = 0
            $batchSkipped = 0
            $batchErrors = 0
            $uploadResults = @()
            $batchUpdateSize = 200  # Larger batches for better performance
            
            # Performance tracking
            $startTime = Get-Date
            
            try {
                foreach ($recording in $Recordings) {
                    $batchProcessed++
                    
                    # Skip if already uploaded
                    if ($recording.UPLOADED -eq 1) {
                        Write-ThreadSafeLog "Recording already uploaded: $($recording.GUID) - Skipping" -Color Yellow -Level "INFO"
                        $batchSkipped++
                        continue
                    }
                    
                    # Skip if not downloaded
                    if ($recording.DOWNLOADED -eq 0 -or [string]::IsNullOrEmpty($recording.DOWNLOAD_PATH)) {
                        Write-ThreadSafeLog "Recording not downloaded yet: $($recording.GUID) - Skipping" -Color Yellow -Level "WARNING"
                        $batchSkipped++
                        continue
                    }
                    
                    try {
                        # Upload the file
                        $recordingStart = [datetime]::Parse($recording.RECORDING_START)
                        Write-ThreadSafeLog "Starting upload: $($recording.HOST_EMAIL) - $($recording.MEETING_ID) - $($recording.RECORDING_TYPE)"
                        
                        $uploadResult = Upload-Recording -FilePath $recording.DOWNLOAD_PATH -UploadConfig $UploadConfig -HostEmail $recording.HOST_EMAIL -MeetingId $recording.MEETING_ID -RecordingStart $recordingStart -RecordingType $recording.RECORDING_TYPE
                        
                        # Collect result for batch update (simplified)
                        #Write-ThreadSafeLog "Success: $($uploadResult.Success), Path: $($uploadResult.UploadPath), Message: $($uploadResult.Message)"
                        $uploadResults += @{
                            Guid = $recording.GUID
                            Uploaded = if ($uploadResult.Success) { 1 } else { 0 }
                            UploadPath = $uploadResult.UploadPath
                        }
                        
                        if ($uploadResult.Success) {
                            $batchUploaded++
                            Write-ThreadSafeLog "Successfully uploaded: $($recording.DOWNLOAD_PATH) -> $($uploadResult.UploadPath)" -Color Green
                        } else {
                            $batchErrors++
                            Write-ThreadSafeLog "Failed to upload: $($recording.GUID) - $($uploadResult.Message)" -Level "ERROR" -Color Red
                        }
                        
                        # Perform batch update when we reach batch size
                        if ($uploadResults.Count -ge $batchUpdateSize) {
                            $updateSuccess = Update-UploadStatusBatch -ConnectionString $ConnectionString -TableName $TableName -UploadResults $uploadResults
                            if ($updateSuccess) {
                                Write-ThreadSafeLog "Batch database update completed for $($uploadResults.Count) records" -Color Green
                            }
                            $uploadResults = @()  # Reset for next batch
                        }
                        
                    } catch {
                        Write-ThreadSafeLog "Error processing recording $($recording.GUID): $_" -Level "ERROR" -Color Red
                        
                        # Add error result to batch (simplified)
                        $uploadResults += @{
                            Guid = $recording.GUID
                            Uploaded = 0
                            UploadPath = ""
                        }
                        $batchErrors++
                    }
                    
                    # Reduced delay for performance - but still respectful
                    Start-Sleep -Milliseconds 100
                }
                
                # Performance summary for this thread
                $duration = (Get-Date) - $startTime
                $filesPerMinute = if ($duration.TotalMinutes -gt 0) { [math]::Round($batchProcessed / $duration.TotalMinutes, 1) } else { 0 }
                Write-ThreadSafeLog "Thread performance: $filesPerMinute files/minute over $($duration.ToString('hh\:mm\:ss'))" -Color Magenta
                
                # Process any remaining results
                if ($uploadResults.Count -gt 0) {
                    $updateSuccess = Update-UploadStatusBatch -ConnectionString $ConnectionString -TableName $TableName -UploadResults $uploadResults
                    if ($updateSuccess) {
                        Write-ThreadSafeLog "Final batch database update completed for $($uploadResults.Count) records" -Color Green
                    }
                }
                
                # Update final progress for this thread
                Update-ThreadProgress -Processed $batchProcessed -Uploaded $batchUploaded -Skipped $batchSkipped -Errors $batchErrors
                
                Write-ThreadSafeLog "Thread completed: Processed=$batchProcessed, Uploaded=$batchUploaded, Skipped=$batchSkipped, Errors=$batchErrors"
                
            } catch {
                Write-ThreadSafeLog "Thread error: $_" -Level "ERROR" -Color Red
                Update-ThreadProgress -Errors $batchProcessed
            }
        }

        # Main worker execution            
        try {
            Write-ThreadSafeLog "Thread started for account: $HostEmail"

            # Get recordings to upload for this account
            $recordingsToUpload = Get-RecordingsToUpload -ConnectionString $ConnectionString -TableName $TableName -HostEmail $HostEmail

            Write-ThreadSafeLog "Account: $HostEmail, Found: $($recordingsToUpload.Count) recordings to upload"
            
            if ($recordingsToUpload.Count -gt 0) {
                Write-ThreadSafeLog "Starting upload using provider: $($UploadConfig.provider)" -Color Cyan
                Process-RecordingsUpload -Recordings $recordingsToUpload -UploadConfig $UploadConfig -ConnectionString $ConnectionString -TableName $TableName
            } else {
                Write-ThreadSafeLog "No recordings to upload for account: $HostEmail" -Color Yellow
            }

        } catch {
            Write-ThreadSafeLog "Error processing account $HostEmail`: $_" -Level "ERROR" -Color Red
        }
    }
}

# Main execution
try {
    Write-ThreadSafeLog "Starting Zoom Recordings Upload..." -Color Cyan
    
    # Load configuration using ZDAConfiguration module
    Write-ThreadSafeLog "Loading configuration using ZDAConfiguration module..."
    $config = Get-Configuration
    
    # Extract configuration values - OPTIMIZED FOR SPEED
    $MaxThreads = if ($config.runspaces.maxThreads) { $config.runspaces.maxThreads } else { 25 }
    $MaxRecordsPerThread = if ($config.runspaces.maxRecordsPerThread) { $config.runspaces.maxRecordsPerThread } else { 5000 }
    $BatchUpdateSize = if ($config.runspaces.batchUpdateSize) { $config.runspaces.batchUpdateSize } else { 200 }
    $UploadDelayMs = if ($config.runspaces.uploadDelayMs) { $config.runspaces.uploadDelayMs } else { 100 }  # Reduce delay
    
    Write-ThreadSafeLog "PERFORMANCE MODE: Max Threads: $MaxThreads, Records/Thread: $MaxRecordsPerThread, Batch Size: $BatchUpdateSize" -Color Cyan
    
    # Validate upload configuration
    if (-not $config.upload -or -not $config.upload.provider) {
        throw "Upload configuration not found or provider not specified in configuration"
    }
    
    $uploadProvider = $config.upload.provider.ToLower()
    Write-ThreadSafeLog "Upload Provider: $uploadProvider" -Color Cyan
    
    if ($uploadProvider -eq "s3") {
        if (-not $config.upload.s3 -or -not $config.upload.s3.bucketName -or -not $config.upload.s3.accessKeyId -or -not $config.upload.s3.secretAccessKey) {
            throw "S3 configuration incomplete. Required: bucketName, accessKeyId, secretAccessKey"
        }
        Write-ThreadSafeLog "S3 Bucket: $($config.upload.s3.bucketName)" -Color Cyan
    }
    elseif ($uploadProvider -eq "onedrive") {
        if (-not $config.upload.onedrive -or -not $config.upload.onedrive.clientId -or -not $config.upload.onedrive.clientSecret -or -not $config.upload.onedrive.tenantId) {
            throw "OneDrive configuration incomplete. Required: clientId, clientSecret, tenantId"
        }
        Write-ThreadSafeLog "OneDrive Tenant: $($config.upload.onedrive.tenantId) (Admin access to user OneDrives)" -Color Cyan
    }
    else {
        throw "Unsupported upload provider: $uploadProvider. Supported providers: s3, onedrive"
    }
    
    # Get unique host emails from database that have recordings to upload
    $sql = @"
SELECT DISTINCT HOST_EMAIL 
FROM $($config.database.tableName) 
WHERE DOWNLOADED = 1 
    AND UPLOADED = 0
    AND DOWNLOAD_PATH IS NOT NULL 
    AND DOWNLOAD_PATH != ''
    AND HOST_EMAIL IS NOT NULL 
    AND HOST_EMAIL != ''
ORDER BY HOST_EMAIL
"@
    
    Write-ThreadSafeLog "Getting list of accounts with recordings to upload..."
    $hostEmails = Invoke-Sqlcmd -ConnectionString $config.database.connectionString -Query $sql -QueryTimeout 120 | Select-Object -ExpandProperty HOST_EMAIL
    
    if (-not $hostEmails -or $hostEmails.Count -eq 0) {
        Write-ThreadSafeLog "No accounts found with recordings to upload" -Color Yellow
        exit 0
    }
    
    # Get total count for performance planning
    $totalCountSql = @"
SELECT COUNT(*) as Total
FROM $($config.database.tableName) 
WHERE DOWNLOADED = 1 
    AND UPLOADED = 0
    AND DOWNLOAD_PATH IS NOT NULL 
    AND DOWNLOAD_PATH != ''
"@
    $totalFiles = (Invoke-Sqlcmd -ConnectionString $config.database.connectionString -Query $totalCountSql -QueryTimeout 120).Total
    
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
    
    Write-ThreadSafeLog "Processing $($hostEmails.Count) account(s) for upload"
    
    # Process accounts using Runspace Pool for maximum performance
    Write-ThreadSafeLog "Creating runspace pool with $MaxThreads threads..."
    
    # Create synchronized hashtable for thread-safe communication
    $sync = [hashtable]::Synchronized(@{
        LogQueue = New-Object System.Collections.Concurrent.ConcurrentQueue[string]
        Progress = @{
            Processed = 0
            Uploaded = 0
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

    # Refactored: Reuse thread IDs and log files, only MaxThreads log files created and reused
    $runspaces = @()
    $availableThreadIds = [System.Collections.Queue]::new()
    for ($i = 1; $i -le $MaxThreads; $i++) { $availableThreadIds.Enqueue($i) }
    $accountsQueue = [System.Collections.Queue]::new()
    foreach ($hostEmail in $hostEmails) { $accountsQueue.Enqueue($hostEmail) }

    # Start initial runspaces up to MaxThreads
    while ($runspaces.Count -lt $MaxThreads -and $accountsQueue.Count -gt 0) {
        $threadId = $availableThreadIds.Dequeue()
        $hostEmail = $accountsQueue.Dequeue()
        Write-ThreadSafeLog "Starting thread $threadId for account: $hostEmail" -Color Yellow
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
        $null = $powershell.AddScript($workerScript)
        $null = $powershell.AddParameter("HostEmail", $hostEmail)
        $null = $powershell.AddParameter("ConnectionString", $config.database.connectionString)
        $null = $powershell.AddParameter("TableName", $config.database.tableName)
        $null = $powershell.AddParameter("UploadConfig", $config.upload)
        $null = $powershell.AddParameter("ThreadId", $threadId)
        $null = $powershell.AddParameter("MaxRecords", $MaxRecordsPerThread)
        $null = $powershell.AddParameter("BatchUpdateSize", $BatchUpdateSize)
        $logPath = Get-LogsPath
        $null = $powershell.AddParameter("LogPath", $logPath)
        $null = $powershell.AddParameter("Sync", $sync)
        $asyncResult = $powershell.BeginInvoke()
        $runspaceInfo = [PSCustomObject]@{
            PowerShell = $powershell
            AsyncResult = $asyncResult
            HostEmail = $hostEmail
            ThreadId = $threadId
            StartTime = Get-Date
        }
        $runspaces += $runspaceInfo
        Write-ThreadSafeLog "Started thread $threadId for account: $hostEmail" -Color Yellow
        Start-Sleep -Milliseconds 200
    }

    Write-ThreadSafeLog "All $($runspaces.Count) threads started. Monitoring progress..." -Color Cyan

    # Monitor and start new runspaces as threads become available
    $progressTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $lastProgressUpdate = Get-Date
    while ($runspaces.Count -gt 0 -or $accountsQueue.Count -gt 0) {
        # Update progress every 30 seconds
        if ((Get-Date) - $lastProgressUpdate -gt [TimeSpan]::FromSeconds(30)) {
            $currentProgress = $sync.Progress
            Write-ThreadSafeLog "=== PROGRESS UPDATE ===" -Color Cyan
            Write-ThreadSafeLog "Processed: $($currentProgress.Processed) | Uploaded: $($currentProgress.Uploaded) | Skipped: $($currentProgress.Skipped) | Errors: $($currentProgress.Errors)" -Color Cyan
            Write-ThreadSafeLog "Active Threads: $(($runspaces | Where-Object { -not $_.AsyncResult.IsCompleted }).Count)/$MaxThreads" -Color Cyan
            Write-ThreadSafeLog "Runtime: $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" -Color Cyan
            $lastProgressUpdate = Get-Date
        }
        # Check for completed runspaces
        $completedRunspaces = $runspaces | Where-Object { $_.AsyncResult.IsCompleted -and $_.PowerShell }
        foreach ($completed in $completedRunspaces) {
            Write-ThreadSafeLog "Processing completed thread $($completed.ThreadId) for account: $($completed.HostEmail)" -Color Green
            try {
                $result = $completed.PowerShell.EndInvoke($completed.AsyncResult)
                $duration = (Get-Date) - $completed.StartTime
                Write-ThreadSafeLog "Thread $($completed.ThreadId) completed for account: $($completed.HostEmail) (Duration: $($duration.ToString('mm\:ss')))" -Color Green
            } catch {
                Write-ThreadSafeLog "Thread $($completed.ThreadId) error for account $($completed.HostEmail): $_" -Level "ERROR" -Color Red
            } finally {
                $completed.PowerShell.Dispose()
                $completed.PowerShell = $null
                # Return threadId to pool
                $availableThreadIds.Enqueue($completed.ThreadId)
                # Remove from runspaces
                $runspaces = $runspaces | Where-Object { $_ -ne $completed }
            }
        }
        # Start new runspaces if threads are available and accounts remain
        while ($availableThreadIds.Count -gt 0 -and $accountsQueue.Count -gt 0) {
            $threadId = $availableThreadIds.Dequeue()
            $hostEmail = $accountsQueue.Dequeue()
            Write-ThreadSafeLog "Starting thread $threadId for account: $hostEmail" -Color Yellow
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $runspacePool
            $null = $powershell.AddScript($workerScript)
            $null = $powershell.AddParameter("HostEmail", $hostEmail)
            $null = $powershell.AddParameter("ConnectionString", $config.database.connectionString)
            $null = $powershell.AddParameter("TableName", $config.database.tableName)
            $null = $powershell.AddParameter("UploadConfig", $config.upload)
            $null = $powershell.AddParameter("ThreadId", $threadId)
            $null = $powershell.AddParameter("MaxRecords", $MaxRecordsPerThread)
            $null = $powershell.AddParameter("BatchUpdateSize", $BatchUpdateSize)
            $null = $powershell.AddParameter("Sync", $sync)
            $asyncResult = $powershell.BeginInvoke()
            $runspaceInfo = [PSCustomObject]@{
                PowerShell = $powershell
                AsyncResult = $asyncResult
                HostEmail = $hostEmail
                ThreadId = $threadId
                StartTime = Get-Date
            }
            $runspaces += $runspaceInfo
            Write-ThreadSafeLog "Started thread $threadId for account: $hostEmail" -Color Yellow
            Start-Sleep -Milliseconds 200
        }
        Start-Sleep -Milliseconds 500
    }

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
    Write-ThreadSafeLog "Total files uploaded: $($finalProgress.Uploaded)" -Color Green
    Write-ThreadSafeLog "Total files skipped (already uploaded/not downloaded): $($finalProgress.Skipped)" -Color Yellow
    Write-ThreadSafeLog "Total errors: $($finalProgress.Errors)" -Color Red
    Write-ThreadSafeLog "Total runtime: $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" -Color Cyan
    Write-ThreadSafeLog "Upload completed successfully!" -Color Green
    Write-ThreadSafeLog "Upload provider: $($config.upload.provider)" -Color Cyan

    } catch {
        Write-ThreadSafeLog "Script execution failed: $_" -Level "ERROR" -Color Red
        exit 1
    } finally {
        # Cleanup is handled in the main execution block
    }