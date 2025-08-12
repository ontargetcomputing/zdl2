using module ../Modules/FileStorage/Classes/AbstractFileStorage.psm1

class S3FileStorage : AbstractFileStorage {
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
            Write-Host("S3FileStorage: Upload job started for recording GUID: $($recording.GUID)")  
            $uploadJob = [UploadJob]::new($recording, $accessKey, $secretAccessKey, $bucketName, $region, $userAgent)
            try {
                Write-Host("Uploading")
                $uploadSuccess = $uploadJob.UploadFile()
                Write-Host("Uploaded")
            } catch {
                Write-Host "Unable to upload $($recording.GUID): $($_.Exception.Message)"
                $uploadSuccess = $false
            }
            [pscustomobject]@{
                uploadSuccess = $uploadSuccess
                GUID_ID       = $recording.GUID
                S3PATH        = $uploadJob.s3Path
                BUCKET        = $uploadJob.bucketName
                HOST_EMAIL    = $recording.HOST_EMAIL
                MEETING_ID    = $recording.MEETING_ID
            }
        }
    }
}
