if (-not (Get-Module -ListAvailable -Name PSZoom)) {
    Write-Host "PSZoom module not found. Installing..."
    
    # Install the module for the current user if not present
    try {
        Install-Module -Name PSZoom -Scope CurrentUser -Force -AllowClobber
        Write-Host "PSZoom module installed successfully."
    } catch {
        Write-Host "Failed to install PSZoom module. Exiting."
        exit 1
    }
} else {
    Write-Host "PSZoom module is already installed."
}
Import-Module PSZoom -ErrorAction Stop

class ZoomService {

    [PSObject]$configuration

    ZoomService([PSObject]$configuration) {    
        $this.configuration = $configuration    
    }
  
    Connect() {  
        $AccountID = $this.configuration.zoom.accountId
        $ClientID = $this.configuration.zoom.clientId
        $ClientSecret = $this.configuration.zoom.clientSecret
        
        Connect-PSZoom -AccountID $AccountID -ClientID $ClientID -ClientSecret $ClientSecret
    }

    [object]GetPageOfZoomRecordings([PSObject]$query_configuration) {
        Write-Host "Retrieving From:$($query_configuration.from), To:$($query_configuration.to), NextPageToken:$($query_configuration.pageToken)"
        return get-zoomaccountrecordings -AccountID me -PageSize 300 -From $query_configuration.from -To $query_configuration.to -NextPageToken $query_configuration.pageToken
    }

    [object]GetAccessToken() {
        $AccountID = $this.configuration.zoom.accountId
        $ClientID = $this.configuration.zoom.clientId
        $ClientSecret = $this.configuration.zoom.clientSecret
        $stringToEncode = $ClientID + ":" + $ClientSecret
        $base64Encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($stringToEncode))
        try {
            $access = Invoke-WebRequest -Method POST -Uri https://zoom.us/oauth/token -ContentType 'application/x-www-form-urlencoded' -Body @{ grant_type = 'account_credentials'; account_id = "$AccountID" } -Headers @{ Host = 'zoom.us'; Authorization = "Basic $base64Encoded" } | ConvertFrom-Json
            return $access.access_token
        }
        catch {
            Write-Host "Unable to authenticate $($_.Exception.Message)"
            throw "Unable to authenticate"
        }
    }
}
