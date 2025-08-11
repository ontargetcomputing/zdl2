using module ../Modules/Database/Classes/AbstractDatabase.psm1
using module ../Modules/Configuration/Classes/ZDAConfiguration.psm1
try {
    [Reflection.Assembly]::LoadFile("C:\Users\Administrator\Downloads\sqlite-netFx45-binary-x64-2012-1.0.118.0\System.Data.SQLite.dll") 2> $null
  } 
  catch {
  
  }
if (-not (Get-Module -ListAvailable -Name SQLite)) {
    Write-Host "SQLite module not found. Installing..."
    
    # Install the module for the current user if not present
    try {
        Install-Module -Name SQLite -Scope CurrentUser -Force -AllowClobber
        Write-Host "SQLite module installed successfully."
    } catch {
        Write-Host "Failed to install SQLite module. Exiting."
        exit 1
    }
} else {
    Write-Host "SQLite module is already installed."
}
Import-Module SQLite -ErrorAction Stop

# Concrete implementation of AbstractDatabase
class SQLiteDatabase : AbstractDatabase {
    static [string]$MainTableName = "ZoomRecordings"
    static [string]$AccountsToDownloadTable = "ZoomRecordingsAccounts"
    [string]$DatabasePath
    $SQLiteConnection
    $configuration 

    SQLiteDatabase([string]$path ) {
        Write-Host "SQLiteDatabase Constructor called"
        $this.configuration = [ZDAConfiguration]::new()
        $this.databasePath = $this.configuration.GetSQLiteDatabasePath()
        Write-Host($this.databasePath)
        $this.CreateDatabase()
    }

    [object]Connect() {
        if( $this.SQLiteConnection -eq $null) {
            $connectionString = "Data Source=$($this.configuration.GetSQLiteDatabasePath())"
            Write-Host("Connection string = $connectionString")
            $this.SQLiteConnection = New-Object System.Data.SQLite.SQLiteConnection
            $this.SQLiteConnection.ConnectionString = $connectionString
            $this.SQLiteConnection.Open() 
            Write-Host("Connected to $($this.databasePath)")
        }
        return $this.SQLiteConnection
    }

    [void]Disconnect() {
        if( $this.SQLiteConnection -ne $null) {
            $this.SQLiteConnection.Close()
            Write-Host("Disconnected from $($this.databasePath)")
        }
    }

    [void]InsertIntoAccountsToDownloadTable([String]$accountsList) {
        $table = [SQLiteDatabase]::AccountsToDownloadTable
        Write-Host "Inserting accounts to download into : $table"
        $sql = "delete from $table"
        $connection = $this.Connect()
        try {
            Write-Host "Truncating $($table)"
            $command = $connection.CreateCommand()
            $command.CommandText = $sql
            $command.ExecuteNonQuery() 
                    
            $emails = $accountsList -split "`r`n" | Where-Object { $_.Trim() -ne "" }
    
        if ($emails.Count -gt 0) {
            foreach ($email in $emails) {
                # Define the query to insert emails
                $query = "INSERT INTO $table (HOST_EMAIL) VALUES (@Email)"
                
                # Create the command and parameter
                $command = $connection.CreateCommand()
                $command.CommandText = $query
                $command.Parameters.AddWithValue("@Email", $email)
    
                # Execute the query
                $command.ExecuteNonQuery()
    
                Write-Host "Inserted email: $email"
            }
        } else {
            # Print a message if there are no emails
            Write-Host "No emails to download added to list, all accounts will be processed."
        }        
        } catch { 
            Write-Host "An error occurred: $($PSItem.Exception.Message)"
        } finally {
        }  
    }

   
    [void]Backup([string]$context) {
        Copy-Item -Path $this.DatabasePath -Destination ($this.DatabasePath + "." + $context + ".bak") -Force
    }

    [object[]]SelectAccountsToDownload() {
        $table = [SQLiteDatabase]::AccountsToDownloadTable
        Write-Host "Selecting data from $($table)"
        $sql = "SELECT HOST_EMAIL FROM $($table)"
        
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql

        $reader = $command.ExecuteReader()

        $emailAddresses = @()

        while ($reader.Read()) {
            $emailAddresses += $reader["HOST_EMAIL"]
        }
        
        if ($emailAddresses.Count -eq 0) {
            Write-Host "No email addresses found.  All accounts will be downloaded"
        } else {
            Write-Host "Email addresses retrieved:"
            $emailAddresses | ForEach-Object { Write-Host $_ }
        }

        # Close the reader and the database connection
        $reader.Close()

        return $emailAddresses
    }

    [boolean]SelectGuidExists([String]$guid) {
        Write-Host "Selecting data from $([SQLiteDatabase]::MainTableName)"
        $sql = "SELECT count(1) FROM $($([SQLiteDatabase]::MainTableName)) where GUID = '$($guid)'"
        
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $command.CommandType = [System.Data.CommandType]::Text
        #$command.Parameters.AddWithValue("@guid", $guid) 
        $count = $command.ExecuteScalar()
        $command.Dispose()
        if ($count -gt 0) {
            return $true
        } else {
            return $false
        }
    }


    [object]SelectNotDownloaded() {
        $table = $([SQLiteDatabase]::MainTableName)
        Write-Host "Selecting not downloaded recordings from $($table)"
        $sql = "SELECT * from $($table) where DOWNLOADED = 0"
        
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $command.CommandType = [System.Data.CommandType]::Text
        $reader = $command.ExecuteReader()
        $outputData = @()
        # Loop through the result set
        while ($reader.Read()) {
            $GUID = $reader['GUID'] 
            $HOST_EMAIL = $reader['HOST_EMAIL'] 
            $RECORDING_START = $reader['RECORDING_START'] 
            $RECORDING_END = $reader['RECORDING_END'] 
            $FILE_SIZE = $reader['FILE_SIZE'] 
            $DOWNLOAD_URL = $reader['DOWNLOAD_URL'] 
            $MEETING_ID = $reader['MEETING_ID'] 
            $TOPIC = $reader['TOPIC'] 
            $RECORDING_TYPE = $reader['RECORDING_TYPE'] 
            # $DOWNLOADED = $reader['DOWNLOADED'] 
            # $TRYDLAGAIN = $reader['TRYDLAGAIN'] 
            # $DOWNLOAD_PATH = $reader['DOWNLOAD_PATH']  
            # $UPLOADED = $reader['UPLOADED'] 
            # $UPLOAD_PATH = $reader['UPLOAD_PATH'] 
    
            $recording = [PSCustomObject]@{ GUID = "$GUID"; HOST_EMAIL = "$HOST_EMAIL"; RECORDING_START = "$RECORDING_START"; RECORDING_END = "$RECORDING_END"; FILE_SIZE = "$FILE_SIZE"; DOWNLOAD_URL = "$DOWNLOAD_URL"; MEETING_ID = "$MEETING_ID"; TOPIC = "$TOPIC"; RECORDING_TYPE = "$RECORDING_TYPE" ; DOWNLOADED = $false ; TRYDLAGAIN = 0; DOWNLOAD_PATH = ''; UPLOADED = $false; UPLOAD_PATH = '' }
            $outputData += $recording    
            Write-Host "Found GUID: $GUID"
        }
    
        $reader.Close()
        $command.Dispose()
        return $outputData
    }

    [object]SelectNotUploaded() {
        $table = $([SQLiteDatabase]::MainTableName)
        Write-Host "Selecting not uploaded recordings from $($table)"
        $sql = "SELECT * from $($table) where UPLOADED = 0"
    
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $command.CommandType = [System.Data.CommandType]::Text
        $reader = $command.ExecuteReader()
        $outputData = @()
        # Loop through the result set
        while ($reader.Read()) {
            $GUID = $reader['GUID'] 
            $HOST_EMAIL = $reader['HOST_EMAIL'] 
            $RECORDING_START = $reader['RECORDING_START'] 
            $RECORDING_END = $reader['RECORDING_END'] 
            $FILE_SIZE = $reader['FILE_SIZE'] 
            $DOWNLOAD_URL = $reader['DOWNLOAD_URL'] 
            $MEETING_ID = $reader['MEETING_ID'] 
            $TOPIC = $reader['TOPIC'] 
            $RECORDING_TYPE = $reader['RECORDING_TYPE'] 
            $DOWNLOADED = $reader['DOWNLOADED'] 
            $TRYDLAGAIN = $reader['TRYDLAGAIN'] 
            $DOWNLOAD_PATH = $reader['DOWNLOAD_PATH']  
            $UPLOADED = $reader['UPLOADED'] 
            $UPLOAD_PATH = $reader['UPLOAD_PATH'] 
    
            $recording = [PSCustomObject]@{ GUID = "$GUID"; HOST_EMAIL = "$HOST_EMAIL"; RECORDING_START = "$RECORDING_START"; RECORDING_END = "$RECORDING_END"; FILE_SIZE = "$FILE_SIZE"; DOWNLOAD_URL = "$DOWNLOAD_URL"; MEETING_ID = "$MEETING_ID"; TOPIC = "$TOPIC"; RECORDING_TYPE = "$RECORDING_TYPE" ; DOWNLOADED = $DOWNLOADED ; TRYDLAGAIN = $TRYDLAGAIN; DOWNLOAD_PATH = "$DOWNLOAD_PATH"; UPLOADED = $UPLOADED; UPLOAD_PATH = $UPLOAD_PATH }
            $outputData += $recording    
            Write-Host "Found GUID: $GUID"
        }
    
        $reader.Close()
        $command.Dispose()
        return $outputData
    }

    [hashtable]UpdateDownloadedRecording($GUID_ID, $TRYDLAGAIN, $FILEPATH, $CLOUDSIZE) {
        $downloadCount = 0
        $NOTDOWNLOADED = 0

        $connection = $this.Connect()
        $table = $([SQLiteDatabase]::MainTableName)
        try {
            $NEWTRYDLAGAIN = "0";
            #$existingRecording = $global:recordings | Where-Object { $_.GUID_ID -eq $GUID_ID }
            #if the file doesn't exist on the file system, try again one time on the next script run
            if ((Test-Path -LiteralPath $FILEPATH) -and ((Get-Item $FILEPATH).length -eq $CLOUDSIZE)) { 
                $downloadCount = 1;        
                $sql = "Update $($table) set DOWNLOADED =  1, DOWNLOAD_PATH = @FILE_PATH WHERE GUID = @GUID;" 
                Write-Host "Update $($table) set DOWNLOADED =  1, DOWNLOAD_PATH = '$FILEPATH' WHERE GUID = '$GUID_ID';" 
                $command = $connection.CreateCommand()
                $command.CommandText = $sql
                $command.Parameters.AddWithValue("@GUID", $GUID_ID) 
                $command.Parameters.AddWithValue("@FILE_PATH", $FILEPATH)     
                $command.ExecuteNonQuery()
            }
            else {
                $FILEPATH = "DELETED";
                $NOTDOWNLOADED = 1;
                Write-Host "File was DELETED or ISN'T FINISHED!";
                If ($TRYDLAGAIN -eq 0) { 
                    $NEWTRYDLAGAIN = "1"; 
                }
                else { 
                    $NEWTRYDLAGAIN = "0"; 
                } 
            
                $sql = "Update $($table) set DOWNLOADED =  0, DOWNLOAD_PATH = 'DELETED', TRYDLAGAIN = @TRYLDAGAIN WHERE GUID = @GUID;" 
                Write-Host "Update $($table) set DOWNLOADED =  0, DOWNLOAD_PATH = 'DELETED', TRYDLAGAIN = $TRYDLAGAIN WHERE GUID = '$GUID_ID';" 
                $command = $connection.CreateCommand()
                $command.CommandText = $sql
                $command.Parameters.AddWithValue("@GUID", $GUID_ID) 
                $command.Parameters.AddWithValue("@FILE_PATH", $FILEPATH)  
                $command.Parameters.AddWithValue("@TRYDLAGAIN", $TRYDLAGAIN)     
                $command.ExecuteNonQuery()
        
            }  
        }  finally {
            
        }
        
        return @{
            downloadCount = $downloadCount
            NOTDOWNLOADED   = $NOTDOWNLOADED
        }
    }
          


    [hashtable]UpdateUploadedRecording($GUID_ID, $UPLOADSUCESS, $ONEDRIVEPATH) {
        $connection = $this.Connect()
        $table = $([SQLiteDatabase]::MainTableName)
        $uploadCount = 0
        $NOTUPLOADED = 0
        #$existingRecording = $global:recordings | Where-Object { $_.GUID_ID -eq $GUID_ID }
        #if the file doesn't exist on the file system, try again one time on the next script run
        #$existingRecording.UPLOADED = $UPLOADSUCESS
        try {
          if( $UPLOADSUCESS -eq $true ) {
            $uploadCount += 1; 
        
            Write-Host "Writing to $($table)"
            $sql = "Update $($table) set UPLOADED =  1, UPLOAD_PATH = @FILE_PATH WHERE GUID = @GUID;" 
            Write-Host $sql
            $command = $connection.CreateCommand()
            $command.CommandText = $sql
            $command.Parameters.AddWithValue("@GUID", $GUID_ID) 
            $command.Parameters.AddWithValue("@FILE_PATH", $ONEDRIVEPATH)     
            $command.ExecuteNonQuery()
          } else {
            $NOTUPLOADED +=1
          }
 
        }  finally {
        } 
        return @{
            uploadCount = $uploadCount
            NOTUPLOADED   = $NOTUPLOADED
        } 
    }

    [void]InsertRecording([PSCustomObject]$record) {
        $table = $([SQLiteDatabase]::MainTableName)
        Write-Host "Writing to $($table)"
        $sql = @"
        Insert into $($table)(
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
        Write-Host $sql
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $command.Parameters.AddWithValue("@GUID", $record.GUID_ID) 
        $command.Parameters.AddWithValue("@HOST_EMAIL", $record.HOST_EMAIL) 
        $command.Parameters.AddWithValue("@RECORDING_START", $record.RECORDING_START) 
        $command.Parameters.AddWithValue("@RECORDING_END", $record.RECORDING_END) 
        $command.Parameters.AddWithValue("@FILE_SIZE", $record.FILE_SIZE) 
        $command.Parameters.AddWithValue("@DOWNLOAD_URL", $record.DOWNLOAD_URL) 
        $command.Parameters.AddWithValue("@MEETING_ID", $record.MEETING_ID) 
        $command.Parameters.AddWithValue("@TOPIC", $record.TOPIC) 
        $command.Parameters.AddWithValue("@RECORDING_TYPE", $record.RECORDING_TYPE) 
        $command.Parameters.AddWithValue("@DOWNLOADED", $record.DOWNLOADED) 
        $command.Parameters.AddWithValue("@TRYDLAGAIN", $record.TRYDLAGAIN) 
        $command.Parameters.AddWithValue("@DOWNLOAD_PATH", $record.DOWNLOAD_PATH) 
        $command.Parameters.AddWithValue("@UPLOADED", $record.UPLOADED) 
        $command.Parameters.AddWithValue("@UPLOAD_PATH", $record.UPLOAD_PATH) 
        
        $command.ExecuteNonQuery()
    }

    hidden CreateDatabase() {
        #Check if the database file already exists
        if (-not (Test-Path -Path $this.databasePath)) {            
            [System.Data.SQLite.SQLiteConnection]::CreateFile($this.databasePath)  
        }
        else {
            Write-Host "Database file already exists."
        }
    
        $connection = $this.Connect()
        $command = $connection.CreateCommand()
    
        $sql = @"
        CREATE TABLE IF NOT EXISTS $([SQLiteDatabase]::MainTableName) (
            GUID TEXT PRIMARY KEY,
            HOST_EMAIL TEXT,
            RECORDING_START TEXT,
            RECORDING_END TEXT,
            FILE_SIZE TEXT,
            DOWNLOAD_URL TEXT,
            MEETING_ID TEXT,
            TOPIC TEXT,
            RECORDING_TYPE TEXT,
            DOWNLOADED TEXT,
            TRYDLAGAIN TEXT,
            DOWNLOAD_PATH TEXT,
            UPLOADED TEXT,
            UPLOAD_PATH TEXT
        );
"@

        Write-Host "Creating table $($sql)"
        try {
            $command = $connection.CreateCommand()
            $command.CommandText = $sql
            $command.ExecuteNonQuery()  
        } catch {  
            Write-Host "An error occurred: $($PSItem.Exception.Message)"
        } finally {
        }
    
        $sql = @"
        CREATE TABLE IF NOT EXISTS $([SQLiteDatabase]::AccountsToDownloadTable) (
            HOST_EMAIL TEXT PRIMARY KEY
        );
"@
        Write-Host "Creating table $($sql)"
        try {
            $command = $connection.CreateCommand()
            $command.CommandText = $sql
            $command.ExecuteNonQuery()  
        } catch { 
            Write-Host "An error occurred: $($PSItem.Exception.Message)"
        } finally {
        }        
    
        Write-Host "SQLite database created successfully."
    }
}
