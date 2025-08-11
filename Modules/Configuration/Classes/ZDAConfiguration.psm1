class ZDAConfiguration {
    ZDAConfiguration() {   
    }   

    [void]CreateLocalAppDataFolder() {
        $AppName = $this.GetAppName()
        $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        $AppDataFolder = Join-Path -Path $localAppDataPath -ChildPath $AppName
        New-Item -ItemType Directory -Force -Path $AppDataFolder
    }

    [string]GetAppName() {
        return ($this.GetConfigurationItems())["app-name"]
    }

    [hashtable]GetConfigurationItems() {
        $configHash = @{}
        $configHash["app-name"] = "ZoomDownloader"
        $configHash["user-config-filename"] = "conf.json"
        $configHash["downloads-directory"] = "ZoomRecordings"
        $configHash["stats-filename"] = "stats.csv"
        $configHash["database-name"] = "ZoomDownloader.sqbpro"
        return $configHash
    }

    [string]GetSQLiteDatabasePath() {
        $appName = $this.GetAppName()
        $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
        $databasePath = Join-Path -Path $appDataDirectory -ChildPath (($this.GetConfigurationItems())["database-name"])
        return $databasePath
    }

    [string]GetDownloadsDirectoryPath() {
        $appName = $this.GetAppName()
        $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
        $downloadsDirectory = Join-Path -Path $appDataDirectory -ChildPath (($this.GetConfigurationItems())["downloads-directory"])
        return $downloadsDirectory
    }

    [PSCustomObject]ReadUserConfiguration() {
        $configFilename = $this.GetUserConfigFilename()
        $appName = $this.GetAppName()
        $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
      
        $configFilePath = Join-Path -Path $appDataDirectory -ChildPath $configFilename
        if (Test-Path -Path $configFilePath) {
          Write-Host "Reading from $($configFilePath)"
          try {
            $conf = Get-Content -Path $configFilePath -Raw | ConvertFrom-Json
            return $conf
          }
          catch {
            Write-Host "Error reading file: $($_.Exception.Message)"
            throw "Unable to read $($configFilePath)"
          }
        }
        else {
          Write-Host "$configFilePath does not exist.  Loading empty dataset"
          #throw "$configFilePath does not exist."
          return $null
        }
    }

    [void]SaveUserConfiguration([string]$jsonString) {
        $configFilename = $this.GetUserConfigFilename()
        $appName = $this.GetAppName()
        $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
    
        $configFilePath = Join-Path -Path $appDataDirectory -ChildPath $configFilename
    
        # Write the JSON data to the file
        $jsonString | Set-Content -Path $configFilePath
        Write-Host "Configruation has been written to $configFilePath"
    }

    [void]StartTranscript([string]$name) {
        $localAppDataPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
        $appName = $this.GetAppName()
        $appDataDirectory = Join-Path -Path $localAppDataPath -ChildPath $appName
        $LogFolder = Join-Path -Path $appDataDirectory -ChildPath 'logs'
        $LogPath = Join-Path -Path $LogFolder -ChildPath ($name + "_" + ((Get-Date).ToString("yyyy-MM-dd")) + ".txt");
        Start-Transcript -Append $LogPath 
    }
    
    [string]GetUserConfigFilename() {
        return ($this.GetConfigurationItems())["user-config-filename"]
    }
}


