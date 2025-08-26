function Get-ConfigurationItems {
    $configHash = @{}
    $configHash["app-name"] = "ZoomDownloader"
    $configHash["user-config-filename"] = "conf.json"
    $configHash["downloads-directory"] = "ZoomRecordings"
    $configHash["stats-filename"] = "stats.csv"
    $configHash["database-name"] = "ZoomDownloader.sqbpro"
    return $configHash
}

function Get-AppName {
    return (Get-ConfigurationItems)["app-name"]
}

function Add-LocalAppDataFolder {
    $AppName = Get-AppName
    $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    $AppDataFolder = Join-Path -Path $localAppDataPath -ChildPath $AppName
    New-Item -ItemType Directory -Force -Path $AppDataFolder | Out-Null
}

function Get-SQLiteDatabasePath {
    $appName = Get-AppName
    $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
    $databasePath = Join-Path -Path $appDataDirectory -ChildPath (Get-ConfigurationItems)["database-name"]
    return $databasePath
}

function Get-DownloadsDirectoryPath {
    $appName = Get-AppName
    $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
    $downloadsDirectory = Join-Path -Path $appDataDirectory -ChildPath (Get-ConfigurationItems)["downloads-directory"]
    return $downloadsDirectory
}

function Get-UserConfigFilename {
    return (Get-ConfigurationItems)["user-config-filename"]
}

function Get-UserConfiguration {
    $configFilename = Get-UserConfigFilename
    $appName = Get-AppName
    $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
    $configFilePath = Join-Path -Path $appDataDirectory -ChildPath $configFilename
    if (Test-Path -Path $configFilePath) {
        Write-Host "Reading from $($configFilePath)"
        try {
            $conf = Get-Content -Path $configFilePath -Raw | ConvertFrom-Json
            return $conf
        } catch {
            Write-Host "Error reading file: $($_.Exception.Message)"
            throw "Unable to read $($configFilePath)"
        }
    } else {
        Write-Host "$configFilePath does not exist.  Loading empty dataset"
        return $null
    }
}

function Save-UserConfiguration {
    param([string]$jsonString)
    $configFilename = Get-UserConfigFilename
    $appName = Get-AppName
    $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
    $configFilePath = Join-Path -Path $appDataDirectory -ChildPath $configFilename
    $jsonString | Set-Content -Path $configFilePath
    Write-Host "Configuration has been written to $configFilePath"
}

function Start-TranscriptForApp {
    param([string]$name)
    $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    $appName = Get-AppName
    $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
    $LogFolder = Join-Path -Path $appDataDirectory -ChildPath 'logs'
    $LogPath = Join-Path -Path $LogFolder -ChildPath ($name + "_" + ((Get-Date).ToString("yyyy-MM-dd")) + ".txt")
    Start-Transcript -Append $LogPath
}
