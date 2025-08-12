using module ../Modules/Jobs/Classes/AbstractJobs.psm1

class DownloadJobs : AbstractJobs {
    $database

    DownloadJobs($database, [int]$max, [int]$sleep) : base($max, $sleep) {  
        $this.database = $database 
      }   
 
    [hashtable]ProcessCompleted() {
        $jobs = (Get-Job -State Completed)
        Write-Host "Processing Completed Downloads"
        $downloadCount = 0
        $NOTDOWNLOADED = 0
        $count = 1
        if ($jobs.Count -gt 0) { 
            foreach ($job in $jobs) {
                Write-Host("Processing completed job: $count")
                $count += 1
                $result = Receive-Job $job
                $results = $this.database.UpdateDownloadedRecording($result.GUID_ID, $result.TRYDLAGAIN,$result.FILEPATH, $result.CLOUDSIZE)
                $downloadCount += $results.downloadCount
                $NOTDOWNLOADED += $results.NOTDOWNLOADED 
                Write-Host("After processing job, downloadCount=$downloadCount, NOTDOWNLOADED=$NOTDOWNLOADED")
            }
            $jobs | Remove-Job 
        }
        return @{
            downloadCount = $downloadCount
            NOTDOWNLOADED   = $NOTDOWNLOADED
        }
    }
}


