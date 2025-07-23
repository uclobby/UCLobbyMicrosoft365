function Test-UcServiceConnection {
    <#
        .SYNOPSIS
        Test connection to a Service

        .DESCRIPTION
        This function will validate if the there is an active connection to a service and also if the required module is installed.

        Requirements:   MsGraph, TeamsDeviceTAC - EntraAuth PowerShell module (Install-Module EntraAuth)
                        TeamsModule - MicrosoftTeams PowerShell module (Install-Module MicrosoftTeams)

        .PARAMETER Type
        Specifies a Type of Service, valid options:
            MSGraph - Microsoft Graph
            TeamsModule - Microsoft Teams PowerShell module
            TeamsDeviceTAC - Teams Admin Center (TAC) API for Teams Devices

        .PARAMETER Scopes
        When present it will check if the require permissions are in the current Scope, only applicable to Microsoft Graph API.
        
        .PARAMETER AltScopes
        Allows checking for alternative permissions to the ones specified in AltScopes, only applicable to Microsoft Graph API.

        .PARAMETER AuthType
        Some Ms Graph APIs can require specific AuthType, Application or Delegated (User)
    #>
    param(
        [Parameter(mandatory = $true)]
        [ValidateSet("MSGraph", "TeamsPowerShell", "TeamsDeviceTAC")]
        [string]$Type,
        [string[]]$Scopes,
        [string[]]$AltScopes,
        [ValidateSet("Application", "Delegated")]
        [string]$AuthType
    )
    switch ($Type) {
        "MSGraph" {
            #UCLobbyTeams is moving to use EntraAuth instead of Microsoft.Graph.Authentication, both will be supported for now.
            $script:GraphEntraAuth = $false
            $EntraAuthModuleAvailable = Get-Module EntraAuth -ListAvailable
            $MSGraphAuthAvailable = Get-Module Microsoft.Graph.Authentication -ListAvailable
      
            if ($EntraAuthModuleAvailable) {
                $AuthToken = Get-EntraToken -Service Graph
                if ($AuthToken) {
                    $script:GraphEntraAuth = $true
                    $currentScopes = $AuthToken.Scopes
                    $AuthTokenType = $AuthToken.tokendata.idtyp.replace('app', 'Application').replace('user', 'Delegated')
                }
            }

            #EntraAuth has priority if already connected.
            if ($MSGraphAuthAvailable -and !$script:GraphEntraAuth) {
                $MgGraphContext = Get-MgContext
                $currentScopes = $MgGraphContext.Scopes
                $AuthTokenType = (""+$MgGraphContext.AuthType).replace('AppOnly', 'Application')
            }

            if(!$EntraAuthModuleAvailable -and !$MSGraphAuthAvailable) {
                Write-Warning ("Missing EntraAuth PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module EntraAuth") 
                return $false
            }

            if (!($currentScopes)) {
                Write-Warning  ("Not Connected to Microsoft Graph" + `
                        [Environment]::NewLine + "Please connect to Microsoft Graph before running this cmdlet." + `
                        [Environment]::NewLine + "Commercial Tenant: Connect-EntraService -ClientID Graph -Scopes " + ($Scopes -join ",") + `
                        [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-EntraService -ClientID Graph " + ($Scopes -join ",") + " -Environment USGov")
                return $false
            }

            if ($AuthType -and $AuthTokenType -ne $AuthType) {
                Write-Warning "Wrong Permission Type: $AuthTokenType, this PowerShell cmdlet requires: $AuthType"
                return $false
            }
            $strScope = ""
            $strAltScope = ""
            $missingScopes = ""
            $missingAltScopes = ""
            $missingScope = $false
            $missingAltScope = $false
            foreach ($scope in $Scopes) {
                $strScope += "`"" + $scope + "`","
                if ($scope -notin $currentScopes) {
                    $missingScope = $true
                    $missingScopes += $scope + ","
                }
            }
            if ($missingScope -and $AltScopes) {
                foreach ($altScope in $AltScopes) {
                    $strAltScope += "`"" + $altScope + "`","
                    if ($altScope -notin $currentScopes) {
                        $missingAltScope = $true
                        $missingAltScopes += $altScope + ","
                    }
                }
            }
            else {
                $missingAltScope = $true
            }
            #If scopes are missing we need to connect using the required scopes
            if ($missingScope -and $missingAltScope) {
                if ($Scopes -and $AltScopes) {
                    Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0, $missingScopes.Length - 1) + " and missing alternative Scope(s): " + $missingAltScopes.Substring(0, $missingAltScopes.Length - 1) + `
                            [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
                            [Environment]::NewLine + "Commercial Tenant: Connect-EntraService -ClientID Graph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " or Connect-EntraService -ClientID Graph -Scopes " + $strAltScope.Substring(0, $strAltScope.Length - 1) + `
                            [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-EntraService -ClientID Graph -Environment USGov -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " or Connect-EntraService -ClientID Graph -Environment USGov -Scopes " + $strAltScope.Substring(0, $strAltScope.Length - 1) )
                }
                else {
                    Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0, $missingScopes.Length - 1) + `
                            [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
                            [Environment]::NewLine + "Commercial Tenant: Connect-EntraService -ClientID Graph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + `
                            [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-EntraService -ClientID Graph -Scopes " + $strScope.Substring(0, $strScope.Length - 1))
                }
                return $false
            }
            return $true
        }
        "TeamsPowerShell" { 
            #Checking if MicrosoftTeams module is installed
            if (!(Get-Module MicrosoftTeams -ListAvailable)) {
                Write-Warning ("Missing MicrosoftTeams PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module MicrosoftTeams") 
                return $false
            }
            #We need to use a cmdlet to know if we are connected to MicrosoftTeams PowerShell
            try {
                Get-CsTenant -ErrorAction SilentlyContinue | Out-Null
                return $true
            }
            catch [System.UnauthorizedAccessException] {
                Write-Warning ("Please connect to Microsoft Teams PowerShell with Connect-MicrosoftTeams before running this cmdlet")
                return $false
            }
        }
        "TeamsDeviceTAC" {
            #Checking if EntraAuth module is installed
            if (!(Get-Module EntraAuth -ListAvailable)) {
                Write-Warning ("Missing EntraAuth PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module EntraAuth") 
                return $false
            }
            if (Get-EntraToken TeamsDeviceTAC) {
                return $true
            }
            else {
                Write-Warning "Please connect to Teams TAC API with Connect-UcTeamsDeviceTAC before running this cmdlet"
            }
        }
        Default {
            return $false
        }
    }
}