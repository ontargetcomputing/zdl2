# Example: Parallel OneDrive uploads using PowerShell runspaces
# This script demonstrates how to upload many files to OneDrive in parallel using runspaces
# Requires your OneDriveFileStorage class and a list of recordings

using module ../Modules/FileStorage/Classes/OneDriveFileStorage.psm1

# Example: create OneDriveFileStorage instance
$storage = [OneDriveFileStorage]::new($appId, $appSecret, $tenantName)

# Example: list of recordings to upload
$recordings = @(
    # Replace with your actual recording objects
    [PSCustomObject]@{ GUID = '1'; DOWNLOAD_PATH = 'C:\Temp\file1.mp4'; HOST_EMAIL = 'user1@example.com'; MEETING_ID = '1001'; TOPIC = 'Topic1' },
    [PSCustomObject]@{ GUID = '2'; DOWNLOAD_PATH = 'C:\Temp\file2.mp4'; HOST_EMAIL = 'user2@example.com'; MEETING_ID = '1002'; TOPIC = 'Topic2' }
    # ...
)

# Number of parallel uploads
$maxThreads = 8

# Create runspace pool
$runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
$runspacePool.Open()

$runspaces = @()
foreach ($rec in $recordings) {
    $runspace = [powershell]::Create()
    $runspace.RunspacePool = $runspacePool
    $null = $runspace.AddScript({
        param($storage, $rec)
        $storage.Upload($rec)
    }).AddArgument($storage).AddArgument($rec)
    $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
}

# Wait for all uploads to finish and collect results
foreach ($r in $runspaces) {
    $result = $r.Pipe.EndInvoke($r.Status)
    $r.Pipe.Dispose()
    Write-Host "Upload result: $($result.uploadSuccess) for GUID $($result.GUID_ID)"
}

$runspacePool.Close()
$runspacePool.Dispose()
