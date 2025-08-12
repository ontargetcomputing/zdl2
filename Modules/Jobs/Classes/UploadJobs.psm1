using module ../Modules/Jobs/Classes/AbstractJobs.psm1

class UploadJobs : AbstractJobs {
    $database

    UploadJobs($database, [int]$max, [int]$sleep) : base($max, $sleep) {  
        $this.database = $database 
      }   

    [hashtable]ProcessCompleted() {
        Write-Host "Processing Completed Uploads"
        $uploadCount = 0
        $NOTUPLOADED = 0
        $jobs = (Get-Job -State Completed)
        if ($jobs.Count -gt 0) { 
            foreach ($job in $jobs) {
                $result = Receive-Job $job
                if( $result.uploadSuccess -eq $true ) {
                    Write-Host ("Processing Completed Successful Upload for GUID:$($result.GUID_ID) to $($result.ONEDRIVEPATH)")
                    $results = $this.database.UpdateUploadedRecording($result.GUID_ID, $result.uploadSuccess, $result.ONEDRIVEPATH)
                    $uploadCount += $results.uploadCount
                    $NOTUPLOADED += $results.NOTUPLOADED 
                } else {
                    Write-Host ("Processing Completed Failed Upload for GUID:$($result.GUID_ID) for $($result.HOST_EMAIL)")
                    $NOTUPLOADED += 1
                }
                
                $results = $this.database.UpdateUploadedRecording($result.GUID_ID, $result.uploadSuccess, $result.ONEDRIVEPATH)
                $uploadCount += $results.uploadCount
                $NOTUPLOADED += $results.NOTUPLOADED 
            }
            $jobs | Remove-Job 

            return @{
                uploadCount = $uploadCount
                NOTUPLOADED   = $NOTUPLOADED
            }
        } else {
            return @{
                uploadCount = 0
                NOTUPLOADED   = 0
            }
        }
    }
}


