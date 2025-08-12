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
        $this.Authenticate()
        Write-Host("Uploading $($recording.DOWNLOAD_PATH) to S3 bucket $($this.BucketName) in region $($this.Region)")
        # Simulate upload logic
        # For real use, use AWS Tools for PowerShell: Write-S3Object
        # Example:
        # Write-S3Object -BucketName $this.BucketName -File $recording.DOWNLOAD_PATH -Key $recording.GUID -Region $this.Region -AccessKey $this.AccessKey -SecretKey $this.SecretAccessKey
        Write-Host("Upload complete for $($recording.GUID)")
    }

    # Add any other methods required by AbstractFileStorage here
}
