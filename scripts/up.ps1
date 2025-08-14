# Process records where Uploaded=false, set to true when completed
param(   
    [string]$ConnectionString = "Server=palomar.cj4cxnl2rpyc.us-west-2.rds.amazonaws.com,1433;Database=zda;User ID=zdauser;Password=YouAre#1;TrustServerCertificate=true",
    [string]$TableName = "recordings.ZoomRecordings",
    [string]$IdColumn = "GUID", 
    [int]$BatchSize = 2,
    [int]$MaxThreads = 1
)

# Function to get next batch of unprocessed records
# function Get-NextBatchToUpload {
#     param(
#         [string]$connString,
#         [int]$batchSize,
#         [string]$tableName,
#         [string]$idCol
#     )
#     Write-Output "Getting next batch of up to $batchSize records from $tableName where Uploaded = 0"
#     $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
    
#     try {
#         $connection.Open()
        
#         # Atomically claim a batch of records that haven't been uploaded
#         $query = @"
# WITH NextBatch AS (
#     SELECT TOP ($batchSize) $idCol
#     FROM $tableName WITH (READPAST)
#     WHERE Uploaded = 0
#     ORDER BY $idCol
# )
# UPDATE $tableName 
# SET ProcessingStarted = GETDATE(),
#     ProcessingThread = @ThreadId
# OUTPUT INSERTED.*
# FROM $tableName t
# INNER JOIN NextBatch nb ON t.$idCol = nb.$idCol
# WHERE t.Uploaded = 0  -- Double-check it hasn't been processed
# "@
        
#         $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
#         $command.Parameters.AddWithValue("@ThreadId", [System.Threading.Thread]::CurrentThread.ManagedThreadId)
        
#         $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
#         $dataTable = New-Object System.Data.DataTable
#         $adapter.Fill($dataTable)
        
#         return $dataTable
        
#     } finally {
#         $connection.Close()
#     }
# }

# Function to mark records as uploaded (or failed)
# function Set-UploadedStatus {
#     param(
#         [string]$connString,
#         [array]$recordIds,
#         [bool]$uploaded,
#         [string]$tableName,
#         [string]$idCol,
#         [string]$errorMessage = $null
#     )
    
#     if ($recordIds.Count -eq 0) { return }
    
#     $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
    
#     try {
#         $connection.Open()
        
#         $idList = ($recordIds -join ',')
#         $query = @"
# UPDATE $tableName 
# SET Uploaded = @Uploaded,
#     ProcessingCompleted = GETDATE(),
#     ErrorMessage = @ErrorMessage
# WHERE $idCol IN ($idList)
# "@
        
#         $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
#         $command.Parameters.AddWithValue("@Uploaded", $uploaded)
#         $command.Parameters.AddWithValue("@ErrorMessage", [System.DBNull]::Value)
        
#         if ($errorMessage) {
#             $command.Parameters["@ErrorMessage"].Value = $errorMessage
#         }
        
#         $command.ExecuteNonQuery()
        
#     } finally {
#         $connection.Close()
#     }
# }

# Worker script that processes batches
$workerScript = {
    param($connectionString, $batchSize, $tableName, $idCol, $workerNumber)
    
    # Create log file for this worker
    $logFile = "worker_$workerNumber.log"
    $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    
        # Simple logging function
    function Write-Log {
        param([string]$message, [string]$level = "INFO")
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logLine = "[$timestamp] [$level] Worker$workerNumber(T$threadId): $message"
        
        # Write to file (thread-safe with Out-File -Append)
        $logLine | Out-File -FilePath $logFile -Append -Encoding UTF8
    }
    
    # Load database functions in the runspace
    function Get-NextBatchToUpload {
        param($connString, $batchSize, $tableName, $idCol)
        $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
        try {
            $connection.Open()

            $query = @"
WITH NextBatch AS (
    SELECT TOP ($batchSize) $idCol
    FROM $tableName WITH (READPAST)
    where UPLOADED = 0 AND (UPLOAD_STARTED is NULL OR UPLOAD_MESSAGE is NOT NULL)
    ORDER BY $idCol
)
UPDATE $tableName 
SET UPLOAD_STARTED = GETDATE(),
    UPLOAD_THREAD = @ThreadId
OUTPUT INSERTED.*
FROM $tableName t
INNER JOIN NextBatch nb ON t.$idCol = nb.$idCol
"@
            
            $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
            $command.Parameters.AddWithValue("@ThreadId", [System.Threading.Thread]::CurrentThread.ManagedThreadId)
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable)

            $rows = @()
            foreach ($row in $dataTable.Rows) {
                $rowData = @{}
                foreach ($col in $dataTable.Columns) {
                    $rowData[$col.ColumnName] = $row[$col]
                }
                $rows += $rowData
            }

            Write-Log "Fetched $($rows.Count) records for upload"
           
            return $rows
        } finally {
            $connection.Close()
        }
    }
    
    function Set-UploadedStatus {
        param($connString, $recordIds, $uploaded, $tableName, $idCol, $errorMessage = $null)

        if ($recordIds.Count -eq 0) { return }
        
        $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
        try {
            $connection.Open()
            
        $idList = ($recordIds | ForEach-Object { "'$_'" }) -join ','
            Write-Log "Setting Uploaded=$uploaded for records: $idList"
            $query = @"
UPDATE $tableName 
SET Uploaded = @Uploaded,
    UPLOAD_COMPLETED = CASE WHEN @Uploaded = 1 THEN GETDATE() ELSE UPLOAD_COMPLETED END,
    UPLOAD_MESSAGE = @ErrorMessage
WHERE $idCol IN ($idList)           
"@
            $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
            $command.Parameters.AddWithValue("@Uploaded", $uploaded)
            $command.Parameters.AddWithValue("@ErrorMessage", [System.DBNull]::Value)
            
            if ($errorMessage) {
                $command.Parameters["@ErrorMessage"].Value = $errorMessage
            }
            
            $command.ExecuteNonQuery()
        } finally {
            $connection.Close()
        }
    }
    
    # Worker main processing loop
    $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    $totalProcessed = 0
    $totalUploaded = 0
    $totalFailed = 0
    $workerStartTime = Get-Date
    
    Write-Log  "Worker $workerNumber (Thread $threadId) started - processing records with Uploaded=false"
    
    while ($true) {
        try {
            # Get next batch of unuploaded records
            $batch = Get-NextBatchToUpload -connString $connectionString -batchSize $batchSize -tableName $tableName -idCol $idCol

            $uploadedIds = @()
            $failedIds = @()
            
            foreach ($row in $batch) {
                try {
                    $guid = $row[$idCol]
                    if ([string]::IsNullOrEmpty($guid)) {
                       # there is an issue with serializing the record from the Get-NextBatchToUpload.  It adds a couple
                       # records (or junk).  Skipping those
                       continue;
                    }
                    $processedValidRecord = $true
                    # *** YOUR UPLOAD/PROCESSING LOGIC HERE ***
                    # Example processing - replace with your actual logic:
                    Write-Log  "Worker $workerNumber - Processing guid:$guid"
                    # Get data from the record
                    # $data = $row["DataColumn"]
                    # $filename = $row["FileName"]
                    # $content = $row["Content"]
                    
                    # Example: Upload to API, FTP, cloud storage, etc.
                    # $uploadResult = Upload-ToDestination -Data $data -FileName $filename
                    
                    # For demo, simulate upload with random success/failure
                    $uploadSuccess = (Get-Random -Minimum 1 -Maximum 100) -gt 50  # 90% success rate
                    
                    if ($uploadSuccess) {
                        # Simulate upload time
                        Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 200)
                        
                        $uploadedIds += $guid
                        $totalUploaded++
                        
                        # Optional: Log successful upload details
                        Write-Log "Worker $workerNumber - Successfully processed record $guid"
                    } else {
                        throw "Simulated upload failure"
                    }
                    
                    $totalProcessed++
                    
                } catch {
                    Write-Log "Worker $workerNumber - Failed to process record: $($_.Exception.Message)"
                    #Write-Log "Worker $workerNumber - Failed to process record $($row[$idCol]): $($_.Exception.Message)"
                    $failedIds += $row[$idCol]
                    $totalFailed++
                }
            }
            
            if(($uploadedIds.Count + $failedIds.Count) -eq 0) {
                # NOTE: we can't just check for $batch.Count -eq 0 above
                # because of the odd serialization issue which adds 2 records that are invalid
                Write-Log "Worker $workerNumber - No more recordings to process, exiting."
                break
            }
            # Update successful records - set Uploaded = true
            if ($uploadedIds.Count -gt 0) {
                Set-UploadedStatus -connString $connectionString -recordIds $uploadedIds -uploaded $true -tableName $tableName -idCol $idCol
            }
            
            # Update failed records - keep Uploaded = false but log error
            if ($failedIds.Count -gt 0) {
                Set-UploadedStatus -connString $connectionString -recordIds $failedIds -uploaded $false -tableName $tableName -idCol $idCol -errorMessage "Upload failed during processing"
            }
            
            Write-Log "Worker $workerNumber - Batch completed. Uploaded: $($uploadedIds.Count), Failed: $($failedIds.Count), Total processed: $totalProcessed"
            
        } catch {
            Write-Log "Worker $workerNumber - Batch processing error: $($_.Exception.Message)"
            Start-Sleep -Seconds 5  # Brief pause before retrying
        }
    }
    
    $workerDuration = (Get-Date) - $workerStartTime
    Write-Log "Worker $workerNumber finished. Processed: $totalProcessed, Uploaded: $totalUploaded, Failed: $totalFailed in $([math]::Round($workerDuration.TotalMinutes, 2)) minutes"
    
    return @{
        WorkerNumber = $workerNumber
        ThreadId = $threadId
        TotalProcessed = $totalProcessed
        TotalUploaded = $totalUploaded
        TotalFailed = $totalFailed
        Duration = $workerDuration
    }
}

# Function to get count of remaining records
function Get-RemainingCount {
    param([string]$connString, [string]$tableName)
    
    $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
    try {
        $connection.Open()
        $command = New-Object System.Data.SqlClient.SqlCommand("SELECT COUNT(*) FROM $tableName WHERE Uploaded = 0", $connection)
        return $command.ExecuteScalar()
    } finally {
        $connection.Close()
    }
}

# Main execution function
function Start-UploadProcessing {
    Write-Host "Starting upload processing for records with Uploaded=false..." -ForegroundColor Green
    Write-Host "Connection: $ConnectionString" -ForegroundColor Yellow
    Write-Host "Table: $TableName" -ForegroundColor Yellow
    Write-Host "Batch Size: $BatchSize" -ForegroundColor Yellow
    Write-Host "Max Threads: $MaxThreads" -ForegroundColor Yellow
    
    # Check initial count
    $initialCount = Get-RemainingCount -connString $ConnectionString -tableName $TableName
    Write-Host "Records to process: $initialCount" -ForegroundColor Cyan
    
    if ($initialCount -eq 0) {
        Write-Host "No records found with Uploaded=false. Nothing to process." -ForegroundColor Yellow
        return
    }
    
    # Create RunspacePool
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $runspacePool.Open()
    
    $workers = @()
    $startTime = Get-Date
    
    # Start worker threads
    for ($i = 1; $i -le $MaxThreads; $i++) {
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
        $powershell.AddScript($workerScript).AddArgument($ConnectionString).AddArgument($BatchSize).AddArgument($TableName).AddArgument($IdColumn).AddArgument($i)
        
        $workers += @{
            WorkerNumber = $i
            PowerShell = $powershell
            AsyncResult = $powershell.BeginInvoke()
            StartTime = Get-Date
        }
        
        Write-Host "Started worker $i" -ForegroundColor Cyan
    }
    
    # Monitor progress
    $completedWorkers = 0
    $allResults = @()
    $lastRemainingCount = $initialCount
    $lastProgressTime = Get-Date
    
    Write-Host "`nMonitoring upload progress..." -ForegroundColor Yellow
    
    while ($completedWorkers -lt $MaxThreads) {
        # Check for completed workers
        foreach ($worker in $workers) {
            if ($worker.AsyncResult.IsCompleted -and $worker.PowerShell) {
                try {
                    $output = $worker.PowerShell.EndInvoke($worker.AsyncResult)
                    
                    foreach ($line in $output) {
                        if ($line -is [string]) {
                            Write-Host $line -ForegroundColor White
                        } else {
                            $allResults += $line
                        }
                    }
                    
                    $completedWorkers++
                    Write-Host "Worker $($worker.WorkerNumber) completed ($completedWorkers/$MaxThreads)" -ForegroundColor Green
                    
                    $worker.PowerShell.Dispose()
                    $worker.PowerShell = $null
                    
                } catch {
                    Write-Host "Error in worker $($worker.WorkerNumber): $($_.Exception.Message)" -ForegroundColor Red
                    $completedWorkers++
                }
            }
        }
        
        # Show progress every 30 seconds
        if (((Get-Date) - $lastProgressTime).TotalSeconds -gt 10) {
            $currentRemaining = Get-RemainingCount -connString $ConnectionString -tableName $TableName
            $processed = $initialCount - $currentRemaining
            $rate = if ($processed -gt 0) { [math]::Round($processed / ((Get-Date) - $startTime).TotalSeconds, 2) } else { 0 }
            
            Write-Host "Progress: $processed/$initialCount processed ($currentRemaining remaining) - Rate: $rate records/sec" -ForegroundColor Magenta
            $lastProgressTime = Get-Date
        }
        
        Start-Sleep -Milliseconds 1000
    }
    
    # Final statistics
    $totalDuration = (Get-Date) - $startTime
    $finalRemaining = Get-RemainingCount -connString $ConnectionString -tableName $TableName
    $totalProcessed = $initialCount - $finalRemaining
    $totalUploaded = ($allResults | Measure-Object TotalUploaded -Sum).Sum
    $totalFailed = ($allResults | Measure-Object TotalFailed -Sum).Sum
    
    Write-Host "`n" + "="*70 -ForegroundColor Magenta
    Write-Host "UPLOAD PROCESSING COMPLETE" -ForegroundColor Magenta
    Write-Host "="*70 -ForegroundColor Magenta
    Write-Host "Initial records with Uploaded=false: $initialCount" -ForegroundColor White
    Write-Host "Successfully uploaded (Uploaded=true): $totalUploaded" -ForegroundColor Green
    Write-Host "Failed uploads (Uploaded=false): $totalFailed" -ForegroundColor Red
    Write-Host "Remaining unprocessed: $finalRemaining" -ForegroundColor Yellow
    Write-Host "Total processing time: $([math]::Round($totalDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    Write-Host "Average rate: $([math]::Round($totalProcessed / $totalDuration.TotalSeconds, 2)) records/second" -ForegroundColor White
    
    # Show worker statistics
    if ($allResults.Count -gt 0) {
        Write-Host "`nWorker Statistics:" -ForegroundColor Yellow
        $allResults | Format-Table WorkerNumber, ThreadId, TotalProcessed, TotalUploaded, TotalFailed, @{Name="Duration(min)"; Expression={[math]::Round($_.Duration.TotalMinutes, 2)}} -AutoSize
    }
    
    # Clean up
    $runspacePool.Close()
    $runspacePool.Dispose()
    
    if ($finalRemaining -gt 0) {
        Write-Host "`nTo retry failed uploads, run this script again." -ForegroundColor Yellow
        Write-Host "Failed records still have Uploaded=false and can be reprocessed." -ForegroundColor Yellow
    }
}

# Optional: Add helpful columns if they don't exist
function Add-TrackingColumns {
    param([string]$connString, [string]$tableName)
    
    Write-Host "SQL to add optional tracking columns (run manually if desired):" -ForegroundColor Yellow
    Write-Host @"
-- Optional: Add tracking columns for better monitoring
ALTER TABLE $tableName ADD 
    UPLOAD_STARTED DATETIME2 NULL,
    UPLOAD_COMPLETED DATETIME2 NULL,
    UPLOAD_THREAD INT NULL,
    UPLOAD_MESSAGE NVARCHAR(MAX) NULL;

-- Add index for better performance
CREATE INDEX IX_${tableName}_Uploaded ON $tableName(Uploaded, $IdColumn);
"@ -ForegroundColor Green
}

# Display setup information
Write-Host "UPLOAD PROCESSING SETUP" -ForegroundColor Magenta
Write-Host "="*40 -ForegroundColor Magenta
Write-Host "Table: $TableName" -ForegroundColor White
Write-Host "Processing records where: Uploaded = false" -ForegroundColor White
Write-Host "Will set to: Uploaded = true (when successful)" -ForegroundColor White
Write-Host ""

Add-TrackingColumns -connString $ConnectionString -tableName $TableName

Write-Host "`nTo start processing:" -ForegroundColor Yellow
Write-Host "Start-UploadProcessing" -ForegroundColor Green

# Uncomment to start immediately:
Start-UploadProcessing