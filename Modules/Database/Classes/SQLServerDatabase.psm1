using module ../Modules/Database/Classes/AbstractDatabase.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1
if (-not (Get-Module -ListAvailable -Name SqlServer)) {
    Write-Host "SqlServer module not found. Installing..."
    
    # Install the module for the current user if not present
    try {
        Install-Module -Name SqlServer -Scope CurrentUser -Force -AllowClobber
        Write-Host "SqlServer module installed successfully."
    } catch {
        Write-Host "Failed to install SqlServer module. Exiting."
        exit 1
    }
} else {
    Write-Host "SqlServer module is already installed."
}
Import-Module SqlServer -ErrorAction Stop

# Concrete implementation of AbstractDatabase
class SQLServerDatabase : AbstractDatabase {
    static [string]$MainTableName = "ZoomRecordings"
    static [string]$AccountsToDownloadTable = "ZoomRecordingsAccounts"
    [string]$ConnectionString
    [string]$Schema
    $SQLServerConnection
    $configuration 

    SQLServerDatabase([object]$configuration) {
        Write-Host "SQLServerDatabase Constructor called"
        $this.configuration = [ZDAConfiguration]::new()
        $sqlserver = $configuration.sqlserver
        $this.ConnectionString = "Server=$($sqlserver.server),$($sqlserver.port);Database=$($sqlserver.database);User ID=$($sqlserver.userid);Password=$($sqlserver.password);TrustServerCertificate=true"
        $this.Schema = $sqlserver.schema
    }

    SQLServerDatabase([object]$configuration, [boolean]$create) {
        Write-Host "SQLServerDatabase Constructor called"
        $this.configuration = [ZDAConfiguration]::new()
        $sqlserver = $configuration.sqlserver
        $this.ConnectionString = "Server=$($sqlserver.server),$($sqlserver.port);Database=$($sqlserver.database);User ID=$($sqlserver.userid);Password=$($sqlserver.password);TrustServerCertificate=true"
        $this.Schema = $sqlserver.schema
        if($create -eq $true) {
            $this.CreateDatabase()
        }
        
    }

    [object]Connect() {
        if ($this.SQLServerConnection -eq $null) {
            $this.SQLServerConnection = New-Object System.Data.SqlClient.SqlConnection($this.ConnectionString)
            $this.SQLServerConnection.Open()
            Write-Host("Connected to SQL Server")
        }
        return $this.SQLServerConnection
    }

    [void]Disconnect() {
        if ($this.SQLServerConnection -ne $null) {
            $this.SQLServerConnection.Close()
            Write-Host("Disconnected from SQL Server")
        }
    }

    [void]InsertIntoAccountsToDownloadTable([String]$accountsList) {
        $table = $this.Schema + "." + [SQLServerDatabase]::AccountsToDownloadTable
        Write-Host "Inserting accounts to download into: $table"
        $connection = $this.Connect()
        try {
            $sql = "DELETE FROM $table"
            $command = $connection.CreateCommand()
            $command.CommandText = $sql
            $command.ExecuteNonQuery()
            Write-Host "Cleared table $table"

            $emails = $accountsList -split "`r`n" | Where-Object { $_.Trim() -ne "" }
            if ($emails.Count -gt 0) {
                foreach ($email in $emails) {
                    $sql = "INSERT INTO $table (HOST_EMAIL) VALUES (@Email)"
                    $command = $connection.CreateCommand()
                    $command.CommandText = $sql
                    $parameter = $command.Parameters.Add("@Email", [System.Data.SqlDbType]::NVarChar)
                    $parameter.Value = $email
                    $command.ExecuteNonQuery()
                    Write-Host "Inserted email: $email"
                }
            } else {
                Write-Host "No emails to insert into accounts list"
            }
        } catch {
            Write-Host "An error occurred: $($_.Exception.Message)"
        }
    }

    [void]Backup([string]$context) {
        Write-Host "Backing up SQL Server database is beyond this script's scope. Use SQL Server Management Studio or `BACKUP DATABASE` T-SQL commands."
    }

    [object[]]SelectAccountsToDownload() {
        $table = $this.Schema + "." + [SQLServerDatabase]::AccountsToDownloadTable
        Write-Host "Selecting data from $table"
        $sql = "SELECT HOST_EMAIL FROM $table"
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $reader = $command.ExecuteReader()

        $emailAddresses = @()
        while ($reader.Read()) {
            $emailAddresses += $reader["HOST_EMAIL"]
        }

        $reader.Close()
        Write-Host "Retrieved accounts to download"
        return $emailAddresses
    }

    [boolean]SelectGuidExists([String]$guid) {
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
        Write-Host "Checking if GUID exists in $table"
        $sql = "SELECT COUNT(1) FROM $table WHERE GUID = @GUID"
        
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $parameter = $command.Parameters.Add("@GUID", [System.Data.SqlDbType]::NVarChar)
        $parameter.Value = $guid
    
        $count = $command.ExecuteScalar()
        return $count -gt 0
    }    

    [object]SelectNotDownloaded() {
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
        Write-Host "Selecting not downloaded recordings from $table"
        $sql = "SELECT * FROM $table WHERE DOWNLOADED = 0"
    
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
    
        $reader = $command.ExecuteReader()
        $outputData = @()
    
        while ($reader.Read()) {
            $recording = [PSCustomObject]@{
                GUID            = $reader["GUID"]
                HOST_EMAIL      = $reader["HOST_EMAIL"]
                RECORDING_START = $reader["RECORDING_START"]
                RECORDING_END   = $reader["RECORDING_END"]
                FILE_SIZE       = $reader["FILE_SIZE"]
                DOWNLOAD_URL    = $reader["DOWNLOAD_URL"]
                MEETING_ID      = $reader["MEETING_ID"]
                TOPIC           = $reader["TOPIC"]
                RECORDING_TYPE  = $reader["RECORDING_TYPE"]
                DOWNLOADED      = $reader["DOWNLOADED"]
                TRYDLAGAIN      = $reader["TRYDLAGAIN"]
                DOWNLOAD_PATH   = $reader["DOWNLOAD_PATH"]
                UPLOADED        = $reader["UPLOADED"]
                UPLOAD_PATH     = $reader["UPLOAD_PATH"]
            }
            $outputData += $recording
            Write-Host "Found GUID: $($recording.GUID)"
        }
    
        $reader.Close()
        return $outputData
    }

    [hashtable]UpdateDownloadedRecording($GUID_ID, $TRYDLAGAIN, $FILEPATH, $CLOUDSIZE) {
        $connection = $this.Connect()
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
        $downloadCount = 0
        $NOTDOWNLOADED = 0
        $command = $connection.CreateCommand()
        try {
            if ((Test-Path -LiteralPath $FILEPATH) -and ((Get-Item $FILEPATH).Length -eq $CLOUDSIZE)) {
                $sql = "UPDATE $table SET DOWNLOADED = 1, DOWNLOAD_PATH = @FILE_PATH WHERE GUID = @GUID"
                $command.CommandText = $sql
                $command.Parameters.Add("@GUID", [System.Data.SqlDbType]::NVarChar).Value = $GUID_ID
                $command.Parameters.Add("@FILE_PATH", [System.Data.SqlDbType]::NVarChar).Value = $FILEPATH
                $command.ExecuteNonQuery()
                $downloadCount = 1
                Write-Host "Successfully updated downloaded recording for GUID: $GUID_ID"
            } else {
                $sql = "UPDATE $table SET DOWNLOADED = 0, DOWNLOAD_PATH = 'DELETED', TRYDLAGAIN = @TRYDLAGAIN WHERE GUID = @GUID"
                $command.CommandText = $sql
                $command.Parameters.Add("@GUID", [System.Data.SqlDbType]::NVarChar).Value = $GUID_ID
                $command.Parameters.Add("@TRYDLAGAIN", [System.Data.SqlDbType]::Int).Value = $TRYDLAGAIN
                $command.ExecuteNonQuery()
                $NOTDOWNLOADED = 1
                Write-Host "Failed to update downloaded recording for GUID: $GUID_ID"
            }
        } finally {
            $command.Dispose()
        }
    
        return @{
            downloadCount = $downloadCount
            NOTDOWNLOADED = $NOTDOWNLOADED
        }
    }

    [hashtable]UpdateUploadedRecording($GUID_ID, $UPLOADSUCCESS, $ONEDRIVEPATH) {
        $connection = $this.Connect()
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
        $uploadCount = 0
        $NOTUPLOADED = 0
        $command = $connection.CreateCommand()
        try {
            if ($UPLOADSUCCESS -eq $true) {
                $sql = "UPDATE $table SET UPLOADED = 1, UPLOAD_PATH = @UPLOAD_PATH WHERE GUID = @GUID"
                $command.CommandText = $sql
                $command.Parameters.Add("@GUID", [System.Data.SqlDbType]::NVarChar).Value = $GUID_ID
                $command.Parameters.Add("@UPLOAD_PATH", [System.Data.SqlDbType]::NVarChar).Value = $ONEDRIVEPATH
                $command.ExecuteNonQuery()
                $uploadCount = 1
                Write-Host "Successfully updated uploaded recording for GUID: $GUID_ID"
            } else {
                $NOTUPLOADED = 1
                Write-Host "Failed to update uploaded recording for GUID: $GUID_ID"
            }
        } finally {
            $command.Dispose()
        }
    
        return @{
            uploadCount = $uploadCount
            NOTUPLOADED = $NOTUPLOADED
        }
    }

    [void]InsertRecording([PSCustomObject]$record) {
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
    
        $sql = @"
        INSERT INTO $table (
            GUID,
            HOST_EMAIL,
            RECORDING_START,
            RECORDING_END,
            FILE_SIZE,
            DOWNLOAD_URL,
            MEETING_ID,
            TOPIC,
            RECORDING_TYPE,
            DOWNLOADED,
            TRYDLAGAIN,
            DOWNLOAD_PATH,
            UPLOADED,
            UPLOAD_PATH
        ) VALUES (
            @GUID,
            @HOST_EMAIL,
            @RECORDING_START,
            @RECORDING_END,
            @FILE_SIZE,
            @DOWNLOAD_URL,
            @MEETING_ID,
            @TOPIC,
            @RECORDING_TYPE,
            @DOWNLOADED,
            @TRYDLAGAIN,
            @DOWNLOAD_PATH,
            @UPLOADED,
            @UPLOAD_PATH
        )
"@
    
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
    
        $command.Parameters.Add("@GUID", [System.Data.SqlDbType]::NVarChar).Value = $record.GUID_ID
        $command.Parameters.Add("@HOST_EMAIL", [System.Data.SqlDbType]::NVarChar).Value = $record.HOST_EMAIL
        $command.Parameters.Add("@RECORDING_START", [System.Data.SqlDbType]::NVarChar).Value = $record.RECORDING_START
        $command.Parameters.Add("@RECORDING_END", [System.Data.SqlDbType]::NVarChar).Value = $record.RECORDING_END
        $command.Parameters.Add("@FILE_SIZE", [System.Data.SqlDbType]::NVarChar).Value = $record.FILE_SIZE
        $command.Parameters.Add("@DOWNLOAD_URL", [System.Data.SqlDbType]::NVarChar).Value = $record.DOWNLOAD_URL
        $command.Parameters.Add("@MEETING_ID", [System.Data.SqlDbType]::NVarChar).Value = $record.MEETING_ID
        $command.Parameters.Add("@TOPIC", [System.Data.SqlDbType]::NVarChar).Value = $record.TOPIC
        $command.Parameters.Add("@RECORDING_TYPE", [System.Data.SqlDbType]::NVarChar).Value = $record.RECORDING_TYPE
        $command.Parameters.Add("@DOWNLOADED", [System.Data.SqlDbType]::Bit).Value = $record.DOWNLOADED
        $command.Parameters.Add("@TRYDLAGAIN", [System.Data.SqlDbType]::Int).Value = $record.TRYDLAGAIN
        $command.Parameters.Add("@DOWNLOAD_PATH", [System.Data.SqlDbType]::NVarChar).Value = $record.DOWNLOAD_PATH
        $command.Parameters.Add("@UPLOADED", [System.Data.SqlDbType]::Bit).Value = $record.UPLOADED
        $command.Parameters.Add("@UPLOAD_PATH", [System.Data.SqlDbType]::NVarChar).Value = $record.UPLOAD_PATH
    
        $command.ExecuteNonQuery()
        Write-Host "Recording inserted successfully"
    }

    hidden CreateDatabase() {
        Write-Host "Ensuring tables exist in SQL Server database"
        $connection = $this.Connect()
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
        $mainTableSQL = @"
        IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '$($this.Schema)' AND TABLE_NAME = '$([SQLServerDatabase]::MainTableName)')
        CREATE TABLE $table (
            GUID NVARCHAR(255) PRIMARY KEY,
            HOST_EMAIL NVARCHAR(255),
            RECORDING_START NVARCHAR(255),
            RECORDING_END NVARCHAR(255),
            FILE_SIZE NVARCHAR(255),
            DOWNLOAD_URL NVARCHAR(255),
            MEETING_ID NVARCHAR(255),
            TOPIC NVARCHAR(255),
            RECORDING_TYPE NVARCHAR(255),
            DOWNLOADED BIT,
            TRYDLAGAIN INT,
            DOWNLOAD_PATH NVARCHAR(255),
            UPLOADED BIT,
            UPLOAD_PATH NVARCHAR(255)
        )
"@
        #Write-Host("Query:$mainTableSQL")
        $command = $connection.CreateCommand()
        $command.CommandText = $mainTableSQL
        $command.ExecuteNonQuery()

        $table = $this.Schema + "." + [SQLServerDatabase]::AccountsToDownloadTable
        $accountsTableSQL = @"
        IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '$($this.Schema)' AND TABLE_NAME = '$([SQLServerDatabase]::AccountsToDownloadTable)')
        CREATE TABLE $table (
            HOST_EMAIL NVARCHAR(255) PRIMARY KEY
        )
"@
        #Write-Host("Query:$accountsTableSQL")
        $command.CommandText = $accountsTableSQL
        $command.ExecuteNonQuery()
        $command.Dispose()
    }

    [object]SelectNotUploaded() {
        $table = $this.Schema + "." + [SQLServerDatabase]::MainTableName
        Write-Host "Selecting not uploaded recordings from $($table)"
        $sql = "SELECT * FROM $($table) WHERE UPLOADED = 0"        
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $command.CommandType = [System.Data.CommandType]::Text
        $reader = $command.ExecuteReader()

        $outputData = @()
        
        while ($reader.Read()) {
            $recording = [PSCustomObject]@{
                GUID            = $reader['GUID'] 
                HOST_EMAIL      = $reader['HOST_EMAIL'] 
                RECORDING_START = $reader['RECORDING_START'] 
                RECORDING_END   = $reader['RECORDING_END'] 
                FILE_SIZE       = $reader['FILE_SIZE'] 
                DOWNLOAD_URL    = $reader['DOWNLOAD_URL'] 
                MEETING_ID      = $reader['MEETING_ID'] 
                TOPIC           = $reader['TOPIC'] 
                RECORDING_TYPE  = $reader['RECORDING_TYPE']
                DOWNLOADED      = $reader['DOWNLOADED']
                TRYDLAGAIN      = $reader['TRYDLAGAIN']
                DOWNLOAD_PATH   = $reader['DOWNLOAD_PATH'] 
                UPLOADED        = $reader['UPLOADED']
                UPLOAD_PATH     = $reader['UPLOAD_PATH']
            }
            $outputData += $recording    
            Write-Host "Found GUID: $($recording.GUID)"
        }
        
        # Clean up
        $reader.Close()
        $command.Dispose()
        
        return $outputData
    }
}