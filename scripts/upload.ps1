using module ../Modules/Database/Classes/SQLServerDatabase.psm1
using module ../Modules/FileStorage/Classes/OneDriveFileStorage.psm1
using module ../Modules/FileStorage/Classes/S3FileStorage.psm1
using module ../Modules/Jobs/Classes/UploadJobs.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("upload")

$user_config = $configuration.ReadUserConfiguration()

# Select file storage class based on type
$fileStorage = $null
switch ($user_config.storage.type) {
    'OneDrive' {
        $office365 = $user_config.storage
        $fileStorage = [OneDriveFileStorage]::new($office365.appId, $office365.clientSecret, $office365.tenantName)
    }
    'S3' {
        $fileStorage = [S3FileStorage]::new(
            $user_config.storage.accessKey,
            $user_config.storage.secretAccessKey,
            $user_config.storage.bucketName,
            $user_config.storage.region
        )
    }
    default {
        Write-Host "Unsupported storage type: $($user_config.storage.type). Exiting upload script."
        exit
    }
}

$global:outputData = @()
$global:uploadCount = 0;
$global:NOTUPLOADED = 0;

$database = [SQLServerDatabase]::new($user_config)
$database.Connect()
$database.Backup("upload")
$uploadJobs = [UploadJobs]::new($database, 11, 5 )

#$TO_UPLOAD = QueryNotUploaded -connection $connection -table (GetDatabaseTable)
$TO_UPLOAD = $database.SelectNotUploaded()

foreach ($recording in $TO_UPLOAD) {
  Write-Host ("Current running jobs: " + (Get-Job -State Running).count);
  $jobsResult = $uploadJobs.ProcessCompleted()
  $global:NOTUPLOADED += $jobsResult.NOTUPLOADED
  $global:uploadCount += $jobsResult.uploadCount  
  $uploadJobs.Throttle() 
  try {
    Write-Host ("Attempting to upload : $($recording.GUID)");
    $fileStorage.Upload($recording)
  } catch {
    Write-Host "Failed to upload: $($recording.GUID), $($_.Exception.Message)"
  }
}

$uploadJobs.WaitForJobsToComplete()
$jobsResult = $uploadJobs.ProcessCompleted()
$global:NOTUPLOADED += $jobsResult.NOTUPLOADED
$global:uploadCount += $jobsResult.uploadCount
$database.Disconnect()

Write-Host "$global:uploadCount - Files Uploaded"
Write-Host "$Global:NOTUPLOADED - Failed to Upload"

Write-Host ("Finished at " + (Get-Date));
Stop-Transcript;



