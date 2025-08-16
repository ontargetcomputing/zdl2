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
    Upload([pscustomobject] $recording) {
        $this.Authenticate()
        $rootFolderKey = $this.CreateRootFolder()
        $meetingFolderKey = $this.CreateMeetingFolder($recording.MEETING_ID, $recording.TOPIC)
        $FileToUpload = $recording.DOWNLOAD_PATH
        $FileName = (Split-Path $FileToUpload -Leaf)
        $Key = "$meetingFolderKey$FileName"
        Write-Host("Uploading $($FileToUpload) as $($Key) to S3 bucket $($this.BucketName)")
        try {
            Write-S3Object -BucketName $this.BucketName -File $FileToUpload -Key $Key -Region $this.Region -AccessKey $this.AccessKey -SecretKey $this.SecretAccessKey
            Write-Host("Upload complete for $($Key)")
            [pscustomobject]@{
                uploadSuccess = $true
                GUID_ID       = $recording.GUID
                S3PATH        = $Key
                BUCKET        = $this.BucketName
                HOST_EMAIL    = $recording.HOST_EMAIL
                MEETING_ID    = $recording.MEETING_ID
            }
        } catch {
            Write-Host "Failed to upload $($Key): $($_.Exception.Message)"
            [pscustomobject]@{
                uploadSuccess = $false
                GUID_ID       = $recording.GUID
                S3PATH        = $Key
                BUCKET        = $this.BucketName
                HOST_EMAIL    = $recording.HOST_EMAIL
                MEETING_ID    = $recording.MEETING_ID
            }
        }
    }
}
