# Define a base class to simulate an abstract class
class AbstractDatabase {
    # Constructor
    AbstractDatabase() {
        Write-Output "AbstractDatabase Constructor Called"
    }

    # Pseudo-abstract methods: Throw exceptions if not overridden
    [void]Backup([string]$context) {
        throw "The 'Backup' method must be implemented in a derived class."
    }

    [object[]]SelectAccountsToDownload() {
        throw "The 'SelectAccountsToDownload' method must be implemented in a derived class."
    }

    [void]Connect() {
        throw "The 'GetConnection' method must be implemented in a derived class."
    }

    [void]Disconnect() {
        throw "The 'GetConnection' method must be implemented in a derived class."
    }

    [void]GetDatabaseAccountsToDownloadTable() {
        throw "The 'GetDatabaseAccountsToDownloadTable' method must be implemented in a derived class."
    }


    [boolean]SelectGuidExists() {
        throw "The 'SelectGuidExists' method must be implemented in a derived class."
    }

    [object]SelectNotDownloaded() {
        throw "The 'SelectNotDownloaded' method must be implemented in a derived class."
    }

    [object]SelectNotUploaded() {
        throw "The 'SelectNotUploaded' method must be implemented in a derived class."
    }

    [hashtable]UpdateDownloadedRecording($GUID_ID, $TRYDLAGAIN, $FILEPATH, $CLOUDSIZE) {
        throw "The 'UpdateDatabaseInDownload' method must be implemented in a derived class."
    }

    [hashtable]UpdateUploadedRecording($GUID_ID, $UPLOADSUCESS, $ONEDRIVEPATH) {
        throw "The 'UpdateUploadedRecording' method must be implemented in a derived class."
    }

    [void]InsertRecording() {
        throw "The 'InsertRecording' method must be implemented in a derived class."
    }
}