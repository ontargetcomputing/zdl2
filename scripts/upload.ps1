using module ../Modules/Database/Classes/SQLServerDatabase.psm1
using module ../Modules/FileStorage/Classes/OneDriveFileStorage.psm1
#using module ../Modules/FileStorage/Classes/S3FileStorage.psm1
using module ../Modules/Jobs/Classes/UploadJobs.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("upload")

$user_config = $configuration.ReadUserConfiguration()

# Select file storage class based on type
# $fileStorage = $null
# switch ($user_config.storage.type) {
#     'OneDrive' {
#         $office365 = $user_config.storage
#         $fileStorage = [OneDriveFileStorage]::new($office365.appId, $office365.clientSecret, $office365.tenantName)
#     }
#     'S3' {
#         $fileStorage = [S3FileStorage]::new(
#             $user_config.storage.accessKey,
#             $user_config.storage.secretAccessKey,
#             $user_config.storage.bucketName,
#             $user_config.storage.region
#         )
#     }
#     default {
#         Write-Host "Unsupported storage type: $($user_config.storage.type). Exiting upload script."
#         exit
#     }
# }

$global:outputData = @()
$global:uploadCount = 0;
$global:NOTUPLOADED = 0;

$database = [SQLServerDatabase]::new($user_config)
$database.Connect()
$database.Backup("upload")
$uploadJobs = [UploadJobs]::new($database, 11, 5 )

#$TO_UPLOAD = QueryNotUploaded -connection $connection -table (GetDatabaseTable)
$TO_UPLOAD = $database.SelectNotUploaded()

$maxThreads = 12
$runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
$runspacePool.Open()

$runspaces = @()
foreach ($recording in $TO_UPLOAD) {
  $runspace = [powershell]::Create()
  $runspace.RunspacePool = $runspacePool
  $null = $runspace.AddScript({
      param($rec)
      Write-Output "Starting upload for GUID: $($rec.GUID)"
  }).AddArgument($rec)
  $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }  
}

foreach ($r in $runspaces) {
    $result = $r.Pipe.EndInvoke($r.Status)
    $r.Pipe.Dispose()
    Write-Host "Upload result: $($result.uploadSuccess) for GUID $($result.GUID_ID)"
}

$runspacePool.Close()
$runspacePool.Dispose()

Write-Host ("Finished at " + (Get-Date));
Stop-Transcript;



