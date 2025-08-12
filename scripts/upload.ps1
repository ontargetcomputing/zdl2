using module ../Modules/Database/Classes/SQLServerDatabase.psm1
using module ../Modules/FileStorage/Classes/OneDriveFileStorage.psm1
using module ../Modules/Jobs/Classes/UploadJobs.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("upload")

$user_config = $configuration.ReadUserConfiguration()

if ($user_config.storage.type -eq 'Local') {
  Write-Host "Storage type is Local. Nothing to do, exiting upload script."
  exit
}

$global:outputData = @()
$global:uploadCount = 0;
$global:NOTUPLOADED = 0;

# Create an instance of the Person class
$office365 = $user_config.office365
$oneDriveFileStorage = [OneDriveFileStorage]::new($office365.appId, $office365.clientSecret, $office365.tenantName)



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
    $oneDriveFileStorage.Upload($recording)
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



