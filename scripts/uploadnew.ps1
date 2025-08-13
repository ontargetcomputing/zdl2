using module ../Modules/Database/Classes/SQLServerDatabase.psm1
using module ../Modules/FileStorage/Classes/OneDriveFileStorage.psm1
#using module ../Modules/FileStorage/Classes/S3FileStorage.psm1
using module ../Modules/Jobs/Classes/UploadJobs.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("upload")

$user_config = $configuration.ReadUserConfiguration()
$database = [SQLServerDatabase]::new($user_config)
$database.Connect()
$TO_UPLOAD = $database.SelectNotUploaded()


# Create RunspacePool with 3 threads
$runspacePool = [runspacefactory]::CreateRunspacePool(1, 3)
$runspacePool.Open()

Write-Host "Processing $($itemList.Count) items with 3 threads..." -ForegroundColor Green

# Create jobs array
$jobs = @()

# Script block
$scriptBlock = {
    param($item)
    
    $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    $startTime = Get-Date
    
    Write-Output "Thread $threadId - Starting: $item.GUID"
    
    # Simulate work
    $delay = Get-Random -Minimum 1 -Maximum 4
    Start-Sleep -Seconds $delay
    
    Write-Output "Thread $threadId - Completed: $item.GUID after $delay seconds"
    
    # Return structured result
    return [PSCustomObject]@{
        Item = $item
        ThreadId = $threadId
        StartTime = $startTime
        EndTime = Get-Date
        Duration = $delay
    }
}

# Start jobs
foreach ($recording in $TO_UPLOAD) {
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $runspacePool
    $powershell.AddScript($scriptBlock).AddArgument($recording)
    
    $jobs += [PSCustomObject]@{
        PowerShell = $powershell
        AsyncResult = $powershell.BeginInvoke()
        Item = $item
    }
}

Write-Host "All jobs started. Monitoring progress..." -ForegroundColor Yellow

# Collect results as jobs complete
$allResults = @()
$completedCount = 0

while ($completedCount -lt $jobs.Count) {
    foreach ($job in $jobs) {
        if ($job.AsyncResult.IsCompleted -and $job.PowerShell) {
            try {
                # Get all output from this job
                $jobOutput = $job.PowerShell.EndInvoke($job.AsyncResult)
                
                foreach ($output in $jobOutput) {
                    if ($output -is [string]) {
                        # Display console messages
                        Write-Host $output -ForegroundColor Cyan
                    } else {
                        # This is our structured result object
                        $allResults += $output
                    }
                }
                
                $completedCount++
                Write-Host "Completed: $($job.Item) ($completedCount/$($jobs.Count))" -ForegroundColor Green
                
                # Clean up
                $job.PowerShell.Dispose()
                $job.PowerShell = $null
                
            } catch {
                Write-Host "Error processing $($job.Item): $($_.Exception.Message)" -ForegroundColor Red
                $completedCount++
            }
        }
    }
    Start-Sleep -Milliseconds 100
}

# Display results summary
Write-Host ""
Write-Host "FINAL RESULTS SUMMARY" -ForegroundColor Magenta
Write-Host "=====================" -ForegroundColor Magenta

if ($allResults.Count -gt 0) {
    $allResults | Sort-Object EndTime | Format-Table Item, ThreadId, Duration, StartTime, EndTime -AutoSize
    
    Write-Host "Statistics:" -ForegroundColor Yellow
    Write-Host "Total items processed: $($allResults.Count)" -ForegroundColor White
    $avgDuration = ($allResults.Duration | Measure-Object -Average).Average
    Write-Host "Average duration: $([math]::Round($avgDuration, 2)) seconds" -ForegroundColor White
} else {
    Write-Host "No results were captured!" -ForegroundColor Red
}

# Clean up
$runspacePool.Close()
$runspacePool.Dispose()

Write-Host ""
Write-Host "RunspacePool completed and cleaned up." -ForegroundColor Green