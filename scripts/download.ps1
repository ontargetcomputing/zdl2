using module ../Modules/Database/Classes/SQLServerDatabase.psm1
using module ../Modules/Zoom/Classes/ZoomService.psm1
using module ../Modules/Jobs/Classes/DownloadJobs.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1

$configuration = [ZDAConfiguration]::new()
$configuration.StartTranscript("download")

$user_config = $configuration.ReadUserConfiguration()
$zoomService = [ZoomService]::new($user_config)
$zoomService.Connect()

####################################################
## Some utility functions below
####################################################
Function AddZeroIf($value) {
  if ($value -lt 10)
  { return ("0" + $value) }
  else
  { return $value };
}

Function DirectorySize($directory) {
  if (Test-Path -Path $directory -PathType Container) {
      if ($fileCount -gt 0) {
          ("{0}" -f ((Get-ChildItem $directory -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1GB)); 
      } else {
          0
      }
  } else {
      0
  }
}

Function GetLocalTime($dateString) {
  $timeFormat = "yyyy-MM-ddTHH:mm:ssZ"
  
  $utcTime = [datetime]::ParseExact($dateString, $timeFormat, [System.Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()
  $pacificTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Pacific Standard Time")

  # Convert UTC time to Pacific time
  $pacificTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcTime, $pacificTimeZone)
  
  # Output the converted time
  Write-Host "UTC Time: $($utcTime.ToString())"
  Write-Host "Pacific Time: $($pacificTime.ToString())"

  return $pacificTime
}

function RemoveInvalidCharacters($path) {

  try { 
    $UnsupportedChars = '[!&{}~#%$:/$¿]'

    filter Matches($UnsupportedChars) {
      $path | Select-String -AllMatches $UnsupportedChars |
      Select-Object -ExpandProperty Matches
      Select-Object -ExpandProperty Values
    }

    $newFileName = $path
    Matches $UnsupportedChars | ForEach-Object {
      if ($_.Value -match "&") { $newFileName = ($newFileName -replace "&", "and") }
      if ($_.Value -match "{") { $newFileName = ($newFileName -replace "{", "(") }
      if ($_.Value -match "}") { $newFileName = ($newFileName -replace "}", ")") }
      if ($_.Value -match "~") { $newFileName = ($newFileName -replace "~", "-") }
      if ($_.Value -match "#") { $newFileName = ($newFileName -replace "#", "") }
      if ($_.Value -match "%") { $newFileName = ($newFileName -replace "%", "") }
      if ($_.Value -match "!") { $newFileName = ($newFileName -replace "!", "") }
      if ($_.Value -match "/") { $newFileName = ($newFileName -replace "/", "-") }
      if ($_.Value -match '$') { $newFileName = ($newFileName -replace "$", "") }
      if ($_.Value -match ":") { $newFileName = ($newFileName -replace ":", "") }
      if ($_.Value -match "¿") { $newFileName = ($newFileName -replace "¿", "") }
    }

    return ($newfileName.Replace("`"", "").Replace("\", "").Replace("|", "-").Replace("?", "").Replace(";", "").Replace("<", "(").Replace(">", ")").Replace("*", "").Replace("`“", "").Replace("`”", "").Replace("[", "(").Replace("]", ")"));
  }
  catch {
    return "Unknown Meeting Subject";
  }
}

function AddStats {
  param (
    [Parameter(Mandatory = $true)]
    [int]$downloaded,
    [Parameter(Mandatory = $true)]
    [int]$not_downloaded,
    [Parameter(Mandatory = $true)]
    [int]$predownloadsize,
    [Parameter(Mandatory = $true)]
    [int]$postdownloadsize,
    [Parameter(Mandatory = $true)]
    [int]$increaseamount
    )

  $appName = $configuration.GetAppName()
  $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
  $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
  $statsFilepath = Join-Path -Path $appDataDirectory -ChildPath (($configuration.GetConfigurationItems())["stats-filename"])

  $today = ((Get-Date)).ToString("yyyy-MM-dd");
  $data = [pscustomobject]@{
    date             = $today
    downloaded       = $downloaded
    not_downloaded   = $not_downloaded
    predownloadsize  = $predownloadsize
    postdownloadsize = $postdownloadsize
    increaseamount   = $increaseamount
  }

  $data | Export-Csv -Path $statsFilepath -Append -NoTypeInformation
}

###########################################################################################

$global:recordings
$global:TZ = "";
$global:downloadCount = 0;
$global:NOTDOWNLOADED = 0;

#calculate how much is stored before we begin
$DOWNLOAD_DIRECTORY = $configuration.GetDownloadsDirectoryPath()
$PREDOWNLOADSIZE = DirectorySize -directory $DOWNLOAD_DIRECTORY


$database = [SQLServerDatabase]::new($user_config)
$database.Connect()
$database.Backup("download")
$downloadJobs = [DownloadJobs]::new($database, 7, 2 )

$global:recordings = $database.SelectNotDownloaded()

# $count = $global:recordings.Count

$access_token = $zoomService.GetAccessToken()
foreach ($recording in $global:recordings) {
  $GUID_ID = $recording.GUID
  if ( $recording.DOWNLOADED -eq $false ) {
    $jobsResult = $downloadJobs.ProcessCompleted()
    $global:NOTDOWNLOADED += $jobsResult.NOTDOWNLOADED
    $global:downloadCount += $jobsResult.downloadCount        
    $downloadJobs.Throttle()
    Write-Host "Attempting to download $GUID_ID"

    ##########################################
    $URL = $recording.DOWNLOAD_URL
    $MEETING_ID = $recording.MEETING_ID;
    $RECORDING_START = $recording.RECORDING_START;
    $TOPIC = $recording.TOPIC;
    $RECORDING_TYPE = $recording.RECORDING_TYPE;
    $HOST_EMAIL = $recording.HOST_EMAIL;
    #convert to PST from UTC
    $RECORDING_START = GetLocalTime($RECORDING_START);
    $FILENAME = ($RECORDING_START.Year.ToString() + (AddZeroIf($RECORDING_START.Month)) + (AddZeroIf($RECORDING_START.Day)));

    $FILETYPE = ".mp4";
    if ($RECORDING_TYPE -eq "chat_file") { $FILETYPE = ".txt"; };
    if ($RECORDING_TYPE -eq "audio_only") { $FILETYPE = ".m4a"; };
    if ($RECORDING_TYPE -eq "audio_transcript") { $FILETYPE = ".vtt"; };

    #sometimes the topic is too long; let's shorten it to 60 if needed for file pathing
    if ($TOPIC.Length -gt 60) { $TOPIC = $TOPIC.Substring(0, 59); }

    $FILENAME = ($FILENAME + "_" + ($RECORDING_START.ToShortTimeString().Replace(" ", "").Replace(":", "")) + "_" + $MEETING_ID + "_" + (RemoveInvalidCharacters($TOPIC)) + "_" + $RECORDING_TYPE + $FILETYPE);
    $FILENAME = $FILENAME -replace ' ', '_'
    
    WRITE-Host ("Transferring recording by " + $HOST_EMAIL + " for " + $FILENAME + " ...");
    $HOST_EMAIL_DIR = Join-Path -Path $DOWNLOAD_DIRECTORY -ChildPath $HOST_EMAIL
    New-Item -ItemType Directory -Force -Path $HOST_EMAIL_DIR | Out-Null;

    $FILENAME = Join-Path -Path $HOST_EMAIL_DIR -ChildPath $FILENAME

    #download_file

    $job = Start-Job -Name $GUID_ID -ArgumentList $GUID_ID, $URL, $FILENAME, ($recording.TRYDLAGAIN), ($recording.FILE_SIZE), ($access_token) -ScriptBlock {

      $J_ID = $args[0];
      $J_URL = $args[1];
      $J_FILEPATH = $args[2];
      $J_TRYDLAGAIN = $args[3];
      $J_FILESIZE = $args[4];
      $J_AUTHTOKEN = $args[5]

      #hide Invoke-WebRequest download progress bar
      $ProgressPreference = "SilentlyContinue"
      # Define the authorization header
      $headers = @{
        "Content-Type"  = "application/x-www-form-urlencoded"
        "Authorization" = "Bearer $J_AUTHTOKEN"
      }

      # Send the web request with the authorization header
      Invoke-WebRequest -Uri $J_URL -OutFile $J_FILEPATH -Headers $headers -ErrorVariable CatchVar
      #Not checking the error for now, the file either is on disk or it is not
      #if ($CatchVar[0] -like "*The server did not return the file size*")

      [pscustomobject]@{
        GUID_ID    = $J_ID
        FILEPATH   = $J_FILEPATH
        TRYDLAGAIN = $J_TRYDLAGAIN
        CATCHERROR = $CatchVar
        CLOUDSIZE  = $J_FILESIZE
      }
    }
  }
  else {
    Write-Host "Skipping $GUID_ID, already downloaded"
  }
}

$downloadJobs.WaitForJobsToComplete()
$jobsResult = $downloadJobs.ProcessCompleted()
$global:NOTDOWNLOADED += $jobsResult.NOTDOWNLOADED
$global:downloadCount += $jobsResult.downloadCount
$database.Disconnect()
Write-Host ("TOTAL DOWNLOADED ITEMS: " + $global:downloadCount);

$POSTDOWNLOADSIZE = DirectorySize -directory $DOWNLOAD_DIRECTORY
$INCREASEAMOUNT = $POSTDOWNLOADSIZE - $PREDOWNLOADSIZE

Write-Host "$global:downloadCount - Files Downloaded"
Write-Host "$Global:NOTDOWNLOADED - Failed to Download"
Write-Host "$PREDOWNLOADSIZE - Pre Download Directory Size"
Write-Host "$POSTDOWNLOADSIZE - Post Download Directory Size"
Write-Host "$INCREASEAMOUNT - Increase Amount"

AddStats -downloaded $global:downloadCount -not_downloaded $Global:NOTDOWNLOADED -predownloadsize $PREDOWNLOADSIZE -postdownloadsize $POSTDOWNLOADSIZE -increaseamount $INCREASEAMOUNT
Stop-Transcript;

####################################################
## Some utility functions below
####################################################
Function AddZeroIf($value) {
  if ($value -lt 10)
  { return ("0" + $value) }
  else
  { return $value };
}

Function DirectorySize($directory) {
  if (Test-Path -Path $directory -PathType Container) {
      if ($fileCount -gt 0) {
          ("{0}" -f ((Get-ChildItem $directory -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1GB)); 
      } else {
          0
      }
  } else {
      0
  }
}

Function GetLocalTime($dateString) {
  $timeFormat = "yyyy-MM-ddTHH:mm:ssZ"
  
  $utcTime = [datetime]::ParseExact($dateString, $timeFormat, [System.Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()
  $pacificTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Pacific Standard Time")

  # Convert UTC time to Pacific time
  $pacificTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcTime, $pacificTimeZone)
  return $pacificTime
}

function RemoveInvalidCharacters($path) {

  try { 
    $UnsupportedChars = '[!&{}~#%$:/$¿]'

    filter Matches($UnsupportedChars) {
      $path | Select-String -AllMatches $UnsupportedChars |
      Select-Object -ExpandProperty Matches
      Select-Object -ExpandProperty Values
    }

    $newFileName = $path
    Matches $UnsupportedChars | ForEach-Object {
      if ($_.Value -match "&") { $newFileName = ($newFileName -replace "&", "and") }
      if ($_.Value -match "{") { $newFileName = ($newFileName -replace "{", "(") }
      if ($_.Value -match "}") { $newFileName = ($newFileName -replace "}", ")") }
      if ($_.Value -match "~") { $newFileName = ($newFileName -replace "~", "-") }
      if ($_.Value -match "#") { $newFileName = ($newFileName -replace "#", "") }
      if ($_.Value -match "%") { $newFileName = ($newFileName -replace "%", "") }
      if ($_.Value -match "!") { $newFileName = ($newFileName -replace "!", "") }
      if ($_.Value -match "/") { $newFileName = ($newFileName -replace "/", "-") }
      if ($_.Value -match '$') { $newFileName = ($newFileName -replace "$", "") }
      if ($_.Value -match ":") { $newFileName = ($newFileName -replace ":", "") }
      if ($_.Value -match "¿") { $newFileName = ($newFileName -replace "¿", "") }
    }

    return ($newfileName.Replace("`"", "").Replace("\", "").Replace("|", "-").Replace("?", "").Replace(";", "").Replace("<", "(").Replace(">", ")").Replace("*", "").Replace("`“", "").Replace("`”", "").Replace("[", "(").Replace("]", ")"));
  }
  catch {
    return "Unknown Meeting Subject";
  }
}

function AddStats {
  param (
    [Parameter(Mandatory = $true)]
    [int]$downloaded,
    [Parameter(Mandatory = $true)]
    [int]$not_downloaded,
    [Parameter(Mandatory = $true)]
    [int]$predownloadsize,
    [Parameter(Mandatory = $true)]
    [int]$postdownloadsize,
    [Parameter(Mandatory = $true)]
    [int]$increaseamount
    )

  $appName = GetAppName
  $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
  $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
  $statsFilepath = Join-Path -Path $appDataDirectory -ChildPath (($configuration.GetConfigurationItems())["stats-filename"])

  $today = ((Get-Date)).ToString("yyyy-MM-dd");
  $data = [pscustomobject]@{
    date             = $today
    downloaded       = $downloaded
    not_downloaded   = $not_downloaded
    predownloadsize  = $predownloadsize
    postdownloadsize = $postdownloadsize
    increaseamount   = $increaseamount
  }

  $data | Export-Csv -Path $statsFilepath -Append -NoTypeInformation
}

###########################################################################################