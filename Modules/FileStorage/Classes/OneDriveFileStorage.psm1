using module ../Modules/FileStorage/Classes/AbstractFileStorage.psm1

class OneDriveFileStorage : AbstractFileStorage {
    [datetime] $LastRefresh = (Get-Date).AddMinutes(-1000)
    [string] $AppId
    [string] $AppSecret
    [string] $Scope
    [string] $TenantName
    [string] $AuthUrl
    [string] $UserAgent
    [string] $AccessToken
  
    [pscustomobject] $emailToUserIdHash = @{};
    [pscustomobject] $emailToRootFolderIdHash = @{};
    [pscustomobject] $emailToMeetingFolderIdHash = @{};
    #[pscustomobject] $sharepointBaseSessionUri = @{};
  
    OneDriveFileStorage([hashtable]$UserConfiguration) {   
      Write-Host("OneDriveFileStorage Constructor called.")
    }     

    OneDriveFileStorage([string] $appId, [string] $appSecret, [string] $tenantName) {        
      $this.AppId = $appId
      $this.AppSecret = $appSecret
      $this.Scope = "https://graph.microsoft.com/.default"
      $this.TenantName = $tenantName
      $this.AuthUrl = "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" 
      $this.UserAgent = "NONISV|Zoom Downloader|OneDrive Upload/1.0"
    }
  
    Authenticate() {
      if ($this.LastRefresh -lt ((Get-Date).AddMinutes( - 45))) { 
        Write-Host('OneDriveFileStorage:Refreshing Token')
        Add-Type -AssemblyName System.Web
  
        # Create body
        $Body = @{
          client_id     = $this.AppId
          client_secret = $this.AppSecret
          scope         = $this.Scope
          grant_type    = 'client_credentials'
        }
  
        # Splat the parameters for Invoke-Restmethod for cleaner code
        $PostSplat = @{
          ContentType = 'application/x-www-form-urlencoded'
          Method = 'POST'
          Body = $Body
          Uri = $this.AuthUrl
          UserAgent = $this.UserAgent
        }
  
        # Request the token!
        try {
          $Request = Invoke-RestMethod @PostSplat
          $this.AccessToken = $Request.access_token
          $this.LastRefresh = Get-Date;
          Write-Host "Authentication Successful"
        } catch {
          # Catch block to handle the exception
          Write-Host "Unable to Authenticate: $($_.Exception.Message)"
          $errorMessage = "Unable to authenticate"
          $exception = New-Object System.Exception($errorMessage)
          throw $exception
        }
  
      } else {
        Write-Host('OneDriveFileStorage:Token still fresh')
      }
    }
  
    Upload([pscustomobject] $recording) {
      $this.Authenticate()
      ##### BEGIN Start-Job
      # TODO - SEND THE HASH VALUES BACK
      Start-Job  -Name $recording.GUID -ArgumentList $recording, $this.AccessToken, $this.UserAgent, $this -ScriptBlock {
        param($recording, $token, $agent, $classInstance)
  
        class UploadJob {
          $recording
          [string] $token
          [string] $agent
          $authHeader
  
          $UserId
          $RootFolderId
          $MeetingFolderId
          $OneDrivepath
          $MeetingFolderName
  
          UploadJob($recording, [string] $token, [string] $agent, $userId, $rootFolderId, $meetingFolderId ) {
            $this.recording = $recording
            $this.token = $token
            $this.agent = $agent
            $this.UserId = $userId
            $this.RootFolderId = $rootFolderId
            $this.MeetingFolderId = $meetingFolderId
            $this.authHeader = @{
              'Content-Type'  = 'application/json'
              'Authorization' = "Bearer $($token)"
            }
          }
  
          UploadFile() {
            $FileToUpload = $this.recording.DOWNLOAD_PATH
            #$FILENAME = (Split-Path $FileToUpload -Leaf).Split('.')[0]
            $FILENAME = (Split-Path $FileToUpload -Leaf)
            Write-Host("Uploading $($FileToUpload) as $($FILENAME)")
            $BaseSessionUri = "https://graph.microsoft.com/v1.0/users/$($this.Userid)/drive/items";
            $Body = @{
              'item' = @{
                '@microsoft.graph.conflictBehavior' = "replace"
              }
            }
  
            $Body = $Body | ConvertTo-Json -Compress
            $UploadUri = ($BaseSessionUri + "/$($this.MeetingFolderID):/${FILENAME}:/createUploadSession");
  
            $Response = Invoke-RestMethod -Uri $UploadUri -Method Post -Headers $this.authHeader -Body $Body -ContentType 'application/json' -UserAgent $this.userAgent
            #Write-Host("The upload URL is $($Response.uploadUrl)")
            $fileInBytes = [System.IO.File]::ReadAllBytes($FileToUpload)
            $fileLength = $fileInBytes.Length
  
            $partSizeBytes = 320 * 1024 * 4  #Uploads 1.31MiB at a time.
            $index = 0
            $start = 0
            $end = 0
  
            $maxloops = [Math]::Round([Math]::Ceiling($fileLength / $partSizeBytes))
  
            while ($fileLength -gt ($end + 1)) {
              $start = $index * $partSizeBytes
              if (($start + $partSizeBytes - 1 ) -lt $fileLength) {
                $end = ($start + $partSizeBytes - 1)
              }
              else {
                $end = ($start + ($fileLength - ($index * $partSizeBytes)) - 1)
              }
              [byte[]]$body = $fileInBytes[$start..$end]
              $headers = @{    
                'Content-Range' = "bytes $start-$end/$fileLength"
              }
              Write-Host "bytes $start-$end/$fileLength | Index: $index and ChunkSize: $partSizeBytes"
              Invoke-WebRequest -Method Put -Uri $Response.uploadUrl -Body $body -Headers $headers | Out-Null
              $index++
              Write-Host "Percentage Complete: $([Math]::Ceiling($index/$maxloops*100)) %"
            }
            
            $this.OneDrivepath = "$($this.MeetingFolderName)\$($FILENAME)" 
          }
  
          [string] GetUserId() {
            $email = $this.recording.HOST_EMAIL
            try {
              Write-Host "Retrieving userid for $($this.recording.HOST_EMAIL) from graph api"
              $Response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$email" -Method Get -Headers $this.authHeader -ContentType 'application/json'
              Write-Host("ID for $($this.recording.HOST_EMAIL) is $($Response.Id)")
              return $Response.Id
            }
            catch {  
              Write-Host $_.Exception.Message
              if ($_.Exception.Message.Contains("404")) 
              { 
                Write-Host ("No OneDrive for $email.") 
                return "NA"
              }
              else
              { 
                Write-Host ("Unhandled error for $email - " + $_.Exception.Message)
                $errorMessage = "An error occurred for $email."
                $exception = New-Object System.Exception($errorMessage)
                throw $exception
              };
            }
          }
  
          [string] CreateRootFolder() {
            Write-Host("Getting root folder id for user id:$($this.UserId)")
            $BaseSessionUri = "https://graph.microsoft.com/v1.0/users/$($this.UserId)/drive/items";
            #$this.sharepointBaseSessionUri[$email] = $BaseSessionUri
            $BaseFolder = "Zoom Cloud Recordings";
            $CreateUri = $BaseSessionUri + "/root/children";
            #Create the folder if it does not exist
            $Body = @{
              name                                = "$BaseFolder"
              '@microsoft.graph.conflictBehavior' = "replace"
              folder                              = @{  }
            }
   
            $Body = $Body | ConvertTo-Json -Compress
            try {
              $Response = Invoke-RestMethod -Uri $CreateUri -Method Post -Headers $this.authHeader -Body $Body -ContentType 'application/json' -UserAgent $this.userAgent
              #$instance.sharepointRootFolderId[$email] = $Response.id;
              Write-Host("Root folder id is $($Response.Id)")
              return $Response.Id
            }
            catch {  
              write-host $PSItem.Exception.Message;
              if ($PSItem.Exception.Code.Contains("nameAlreadyExists")) { 
                Write-Host ("Folder Already Exists") 
              }
              else
              { Write-Host ("Unhandled error for $($this.recording.HOST_EMAIL) - " + $PSItem.Exception.Message) };
              return ""
            }
          }
  
          [string] CreateMeetingFolder() {
            write-host "creating meeting folder - $($this.recording.MEETING_ID) - $($this.recording.TOPIC)";
            #create meeting subfolder if needed
            #$RootFolderID = $this.sharepointRootFolderId[$email]
            #$CreateUri = $this.sharepointBaseSessionUri[$email] + "/$RootFolderID/children";
            $BaseSessionUri = "https://graph.microsoft.com/v1.0/users/USERIDGOESHERE/drive/items";
            $BaseSessionUri = $BaseSessionUri.Replace("USERIDGOESHERE", $this.Userid);
  
            $CreateUri = $BaseSessionUri + "/$($this.rootFolderId)/children";
            $this.MeetingFolderName = "$($this.recording.MEETING_ID) - $($this.recording.TOPIC)"
            $Body = @{
              name                                = $this.MeetingFolderName
              '@microsoft.graph.conflictBehavior' = "replace"
              folder                              = @{  }
            }
  
            $Body = $Body | ConvertTo-Json -Compress
            try {  
              $Response = Invoke-RestMethod -Uri $CreateUri -Method Post -Headers $this.authHeader -Body $Body -ContentType 'application/json' -UserAgent $this.userAgent
              Write-Host("Meeting folder id is $($Response.Id)")
              return $Response.Id
            }
            catch {  
              write-host $PSItem.Exception.Message;
              if ($PSItem.Exception.Code.Contains("nameAlreadyExists")) { 
                Write-Host ("Folder Already Exists") 
              }
              else
              { Write-Host ("Unhandled error for $($this.recording.HOST_EMAIL) - " + $PSItem.Exception.Message) };
              return ""
            }
          }
  
          [boolean] Upload() {
            Write-Host("Starting job for GUID:$($this.recording.GUID)")
            try {
              if ( "NA" -ne $this.UserId ) {
                if ( $null -eq $this.MeetingFolderId ) {
                  Write-Host("ID for Meeting Folder for $($this.recording.MEETING_ID) is unknown")
                  if ( $null -eq $this.RootFolderId ) {
                    Write-Host("ID for Root Folder for $($this.recording.HOST_EMAIL) is unknown")
                    if ( $null -eq $this.UserId ) {
                      $this.UserId = $this.GetUserId()
                    }
                    if ( "NA" -ne $this.UserId ) {
                      Write-Host "Creating folders"
                      $this.RootFolderId = $this.CreateRootFolder()
                      $this.MeetingFolderId = $this.CreateMeetingFolder()
                      $this.UploadFile()
                    } else {
                      Write-Host("$($this.recording.HOST_EMAIL) does not have OneDrive")
                      return $false                   
                    }
                  }
                  else {
                    Write-Host("ID for Root Folder for $($this.recording.HOST_EMAIL) is known")
                    $this.MeetingFolderId = $this.CreateMeetingFolder()
                    $this.UploadFile()
                    
                  }
                }
                else {
                  Write-Host("ID for Meeting Folder for $($this.MeetingFolderId) is known")
                  $this.UploadFile()
                }
                return $true
              } else {
                Write-Host("$($this.recording.HOST_EMAIL) does not have OneDrive")
                return $false
              }
            } catch {
              # Catch block to handle the exception
              Write-Host "Unable to Upload for $($this.recording.GUID): $($_.Exception.Message)"
              $errorMessage = "Unable to Upload"
              $exception = New-Object System.Exception($errorMessage)
              throw $exception            
            }
          } 
        }
        # #########################END CLASS
        
        $UserId = $classInstance.emailToUserIdHash[$recording.HOST_EMAIL]
        $RootFolderId = $classInstance.emailToRootFolderIdHash[$recording.HOST_EMAIL]
        $MeetingFolderId = $classInstance.emailToMeetingFolderIdHash[$recording.MEETING_ID]
        $uploadJob = [UploadJob]::new($recording, $token, $agent, $UserId, $RootFolderId, $MeetingFolderId)
        try {
          $uploadSuccess = $uploadJob.Upload()
        }
        catch {
          # Catch block to handle the exception
          Write-Host "Unable to upload $($recording.GUID): $($_.Exception.Message)"  
          $uploadSuccess = $false   
        }
        [pscustomobject]@{
          uploadSuccess   = $uploadSuccess
          GUID_ID         = $recording.GUID
          ONEDRIVEPATH    = $uploadJob.OneDrivepath
          USERID          = $uploadJob.UserID
          ROOTFOLDERID    = $uploadJob.RootFolderID
          MEETINGFOLDERID = $uploadJob.MeetingFolderID
          HOST_EMAIL      = $recording.HOST_EMAIL
          MEETING_ID      = $recoding.MEETING_ID
        }
      } 
      ##### END Start-Job
    }
  
}
