using module ../Modules/Zoom/Classes/ZoomService.psm1
using module ../Modules/Database/Classes/SQLServerDatabase.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("identification")

$user_config = $configuration.ReadUserConfiguration()

$zoomService = [ZoomService]::new($user_config)
$zoomService.Connect()

$FromDate = ((Get-Date).AddDays(-60)).ToString("yyyy-MM-dd");
$ToDate = (Get-Date).AddDays(1).ToString("yyyy-MM-dd");

Write-Host("Determining date range for meeting identification...$($user_config.schedule.dateRange)");
if ( $user_config.schedule.dateRange -eq "2 Weeks" ) {
  $FromDate = ((Get-Date).AddDays(-14)).ToString("yyyy-MM-dd");
} elseif ($user_config.schedule.dateRange -eq "1 Month") {
  $FromDate = ((Get-Date).AddDays(-30)).ToString("yyyy-MM-dd");
} elseif ($user_config.schedule.dateRange -eq "Custom start date...") {
  $FromDate = $user_config.schedule.customFromDate;
} 

Write-Host ("Starting meeting load at " + (Get-Date -Format g));
Write-Host ("Using $($FromDate) to $($ToDate)");
 
$database = [SQLServerDatabase]::new($user_config)
$database.Connect()

$RECORDINGCOUNT = 0;
$VARERRORS = 0;

$accountsToIdentify = $database.SelectAccountsToDownload()

$startDate = [datetime]::ParseExact($FromDate, "yyyy-MM-dd", $null)
$endDate = [datetime]::ParseExact($ToDate, "yyyy-MM-dd", $null)

$currentStartDate = $startDate

while ($currentStartDate -lt $endDate) {
    $currentEndDate = $currentStartDate.AddDays(29)
    
    if ($currentEndDate -gt $endDate) {
        $currentEndDate = $endDate
    }

    $queryPageConfig = @{
      "from" = $currentStartDate.ToString("yyyy-MM-dd")
      "to" = $currentEndDate.ToString("yyyy-MM-dd")
      "pageToken" = ""
    }
    do {  
      #Write-Host("Querying for recordings with $($currentStartDate.ToString("yyyy-MM-dd")) to $($currentEndDate.ToString("yyyy-MM-dd")) and next page token of $($queryPageConfig.pageToken)")

      $dailyrecordings = $zoomService.GetPageOfZoomRecordings($queryPageConfig)
      $queryPageConfig.pageToken = $dailyrecordings.next_page_token;
      Write-Host("Zoom Response $($dailyrecordings)")
      foreach ($meeting in $dailyrecordings.meetings) {
        try {
          foreach ($recording in $meeting.recording_files) {
            if ($recording.status -eq "completed" -and $recording.file_type -ne "TIMELINE") {
              $HOST_EMAIL = $meeting.host_email;
              if ($accountsToIdentify.Count -eq 0 -or $accountsToIdentify -contains $HOST_EMAIL) {
                Write-Host "Identifing recording for $HOST_EMAIL"

                $GUID_ID = $recording.id
                $RECORDING_START = $recording.recording_start;
                $RECORDING_END = $recording.recording_end;
                
                $DOWNLOAD_URL = $recording.download_url;
                $MEETING_ID = $meeting.id;
                $FILE_SIZE = $recording.file_size;
                $TOPIC = $meeting.topic.Replace("'", "''").Replace("✨", '');
                $RECORDING_TYPE = $recording.recording_type;

                $guidExists = $database.SelectGuidExists($GUID_ID)
                #$existingRecording = $existingData | Where-Object { $_.GUID_ID -eq $GUID_ID }
                if ($guidExists -eq $true) {
                  Write-Host "Skipping existing recording:  $GUID_ID - $HOST_EMAIL - $MEETING_ID - $RECORDING_START - $TOPIC - $RECORDING_TYPE";
                }
                else {
                  Write-Host "Added recording:  $GUID_ID - $HOST_EMAIL - $MEETING_ID - $RECORDING_START - $TOPIC - $RECORDING_TYPE";

                  $recording = [PSCustomObject]@{ GUID_ID = "$GUID_ID"; HOST_EMAIL = "$HOST_EMAIL"; RECORDING_START = "$RECORDING_START"; RECORDING_END = "$RECORDING_END"; FILE_SIZE = "$FILE_SIZE"; DOWNLOAD_URL = "$DOWNLOAD_URL"; MEETING_ID = "$MEETING_ID"; TOPIC = "$TOPIC"; RECORDING_TYPE = "$RECORDING_TYPE" ; DOWNLOADED = $false ; TRYDLAGAIN = 0; DOWNLOAD_PATH = ''; UPLOADED = $false; UPLOAD_PATH = '' }
                  
                  #$outputData += $recording
                  $database.InsertRecording($recording)
                  $RECORDINGCOUNT += 1;
                }
              } else {
                Write-Host "Skipping recording for $HOST_EMAIL, email not in list of emails to download for"
              }
            }
          }
        }
        catch { Write-Host ("UNHANDLED ERROR - " + ($PSItem.Exception.Message)); $VARERRORS += 1; }
      }
      
    } while ($queryPageConfig.pageToken -ne "" -and $null -ne $dailyrecordings);



    $currentStartDate = $currentEndDate.AddDays(1)
}

$database.Disconnect()

if ($RECORDINGCOUNT -gt 0) {
  Write-Host "Added $RECORDINGCOUNT new recordings"
}
else {
  Write-Host "No new recordings found"
}


Write-Host ("Stopped meeting load at " + (Get-Date -Format g));
Stop-Transcript;


