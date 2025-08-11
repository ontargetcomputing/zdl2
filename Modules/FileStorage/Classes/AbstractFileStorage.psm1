class AbstractFileStorage {
    [hashtable]$UserConfiguration

    AbstractFileStorage() {   
        Write-Host("AbstractFileStorage Empty Constructor called.")
    }   

    AbstractFileStorage([hashtable]$UserConfiguration) {
        $this.UserConfiguration = $UserConfiguration
    }

    [void]Authenticate() {
        throw "The method 'Authenticate' must be implemented in a derived class."
    }

    [boolean]Upload([pscustomobject] $recording) {
        throw "The method 'Upload' must be implemented in a derived class."
    }
}


