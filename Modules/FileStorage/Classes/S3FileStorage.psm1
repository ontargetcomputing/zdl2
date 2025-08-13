using module ../Modules/FileStorage/Classes/AbstractFileStorage.psm1

class S3FileStorage : AbstractFileStorage {
    [string] $RootFolder = "Zoom Cloud Recordings"

    [string] CreateRootFolder() {
        $folderKey = "$($this.RootFolder)/"
        try {
            Write-S3Object -BucketName $this.BucketName -Key $folderKey -Content "" -Region $this.Region -AccessKey $this.AccessKey -SecretKey $this.SecretAccessKey
            Write-Host "Created root folder in S3: $folderKey"
            return $folderKey
        } catch {
            Write-Host "Failed to create root folder: $($_.Exception.Message)"
            return $folderKey
        }
    }

    [string] CreateMeetingFolder($meetingId, $topic) {
        $meetingFolderName = "$meetingId - $topic"
        $meetingFolderKey = "$($this.RootFolder)/$meetingFolderName/"
        try {
            Write-S3Object -BucketName $this.BucketName -Key $meetingFolderKey -Content "" -Region $this.Region -AccessKey $this.AccessKey -SecretKey $this.SecretAccessKey
            Write-Host "Created meeting folder in S3: $meetingFolderKey"
            return $meetingFolderKey
        } catch {
            Write-Host "Failed to create meeting folder: $($_.Exception.Message)"
            return $meetingFolderKey
        }
    }
    [string] $AccessKey
    [string] $SecretAccessKey
    [string] $BucketName
    [string] $Region
    [string] $UserAgent
    [string] $SessionToken
    [datetime] $LastRefresh = (Get-Date).AddMinutes(-1000)

    S3FileStorage([hashtable]$UserConfiguration) {
        Write-Host("S3FileStorage Constructor called.")
        $this.AccessKey = $UserConfiguration.accessKey
        $this.SecretAccessKey = $UserConfiguration.secretAccessKey
        $this.BucketName = $UserConfiguration.bucketName
        $this.Region = $UserConfiguration.region
        $this.UserAgent = "NONISV|Zoom Downloader|S3 Upload/1.0"
    }

    S3FileStorage([string] $accessKey, [string] $secretAccessKey, [string] $bucketName, [string] $region) {
        Write-Host("S3FileStorage Constructor called.")
        $this.AccessKey = $accessKey
        $this.SecretAccessKey = $secretAccessKey
        $this.BucketName = $bucketName
        $this.Region = $region
        $this.UserAgent = "NONISV|Zoom Downloader|S3 Upload/1.0"
    }

    Authenticate() {
        Write-Host('S3FileStorage: Authenticating')
        try {
            # Set AWS credentials for the session (profile is temporary)
            Set-AWSCredential -AccessKey $this.AccessKey -SecretKey $this.SecretAccessKey -StoreAs 'ZoomDownloader'

            # Try to list buckets to validate credentials
            $buckets = Get-S3Bucket -ProfileName 'ZoomDownloader' -Region $this.Region
            $this.SessionToken = "VALIDATED"
            $this.LastRefresh = Get-Date
            Write-Host "Authentication Successful"
        } catch {
            Write-Host "Unable to Authenticate: $($_.Exception.Message)"
            $this.SessionToken = $null
            throw "Invalid AWS credentials or insufficient permissions."
        }
    }

    Upload([pscustomobject] $recording) {
        Write-Host("S3FileStorage: Starting upload job for recording GUID: $($recording.GUID)")
        Start-Job -Name $recording.GUID -ArgumentList $recording, $this.AccessKey, $this.SecretAccessKey, $this.BucketName, $this.Region, $this.UserAgent -ScriptBlock {
            param($recording, $accessKey, $secretAccessKey, $bucketName, $region, $userAgent)
            
            class UploadJob {
                $recording
                [string] $accessKey
                [string] $secretAccessKey
                [string] $bucketName
                [string] $region
                [string] $userAgent
                [string] $s3Path

                UploadJob($recording, [string] $accessKey, [string] $secretAccessKey, [string] $bucketName, [string] $region, [string] $userAgent) {
                    $this.recording = $recording
                    $this.accessKey = $accessKey
                    $this.secretAccessKey = $secretAccessKey
                    $this.bucketName = $bucketName
                    $this.region = $region
                    $this.userAgent = $userAgent
                }

                [bool] UploadFile() {
                    $FileToUpload = $this.recording.DOWNLOAD_PATH
                    $Key = $this.recording.GUID
                    Write-Host("Uploading $($FileToUpload) as $($Key) to S3 bucket $($this.bucketName)")
                    try {
                        Write-S3Object -BucketName $this.bucketName -File $FileToUpload -Key $Key -Region $this.region -AccessKey $this.accessKey -SecretKey $this.secretAccessKey
                        $this.s3Path = "$($this.bucketName)/$($Key)"
                        Write-Host("Upload complete for $($Key)")
                        return $true
                    } catch {
                        Write-Host "Failed to upload $($Key): $($_.Exception.Message)"
                        return $false
                    }
                }
            }
        }
    }
}
