# Create a list of items to process
$itemList = @("Server1", "Server2", "Server3", "Server4", "Server5", "Server6", "Server7", "Server8")

# Create RunspacePool with min/max threads
$minThreads = 1
$maxThreads = 4
$runspacePool = [runspacefactory]::CreateRunspacePool($minThreads, $maxThreads)
$runspacePool.Open()

Write-Host "Created RunspacePool with $maxThreads threads" -ForegroundColor Green
Write-Host "Processing $($itemList.Count) items..." -ForegroundColor Green

# Create jobs array to track all running jobs
$jobs = @()

# Script block to execute for each item
$scriptBlock = {
    param($item, $delay)
    
    $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    Write-Output "Thread $threadId - Starting: $item"
    
    # Simulate work with variable delay
    Start-Sleep -Seconds $delay
    
    Write-Output "Thread $threadId - Completed: $item after $delay seconds"
    
    # Return result
    return @{
        Item = $item
        ThreadId = $threadId
        ProcessedAt = Get-Date
    }
}

# Create and start a job for each item
foreach ($item in $itemList) {
    # Random delay between 1-3 seconds to simulate varying work
    $randomDelay = Get-Random -Minimum 1 -Maximum 4
    
    # Create PowerShell instance
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $runspacePool
    
    # Add script and parameters
    $powershell.AddScript($scriptBlock).AddArgument($item).AddArgument($randomDelay)
    
    # Start async execution
    $asyncResult = $powershell.BeginInvoke()
    
    # Store job info
    $jobs += [PSCustomObject]@{
        PowerShell = $powershell
        AsyncResult = $asyncResult
        Item = $item
        StartTime = Get-Date
    }
}

Write-Host "All jobs started. Waiting for completion..." -ForegroundColor Yellow

# Wait for all jobs to complete and collect results
$completedJobs = 0
$results = @()

while ($completedJobs -lt $jobs.Count) {
    foreach ($job in $jobs) {
        if ($job.AsyncResult.IsCompleted -and $job.PowerShell) {
            try {
                # Get the output
                $output = $job.PowerShell.EndInvoke($job.AsyncResult)
                
                # Display console output
                foreach ($line in $output) {
                    if ($line -is [string]) {
                        Write-Host $line -ForegroundColor Cyan
                    } else {
                        $results += $line
                    }
                }
                
                $completedJobs++
                Write-Host "Job completed for: $($job.Item) ($completedJobs/$($jobs.Count))" -ForegroundColor Green
                
                # Clean up this job
                $job.PowerShell.Dispose()
                $job.PowerShell = $null
                
            } catch {
                Write-Host "Error in job for $($job.Item): $($_.Exception.Message)" -ForegroundColor Red
                $completedJobs++
            }
        }
    }
    
    # Brief pause to avoid busy waiting
    Start-Sleep -Milliseconds 100
}

# Display final results
Write-Host "`nFinal Results:" -ForegroundColor Magenta
$results | Sort-Object ProcessedAt | Format-Table Item, ThreadId, ProcessedAt -AutoSize

# Clean up
$runspacePool.Close()
$runspacePool.Dispose()

Write-Host "RunspacePool completed and cleaned up." -ForegroundColor Green