function Test-UcPowerShellModule {
    <#
        .SYNOPSIS
        Test if PowerShell module is installed and updated

        .DESCRIPTION
        This function returns FALSE if PowerShell module is not installed.

        .PARAMETER ModuleName
        Specifies PowerShell module name

        .EXAMPLE
        PS> Test-UcPowerShellModule -ModuleName UCLobbyTeams
    #>
    param(
        [string]$ModuleName
    )
    
    try { 
        #Region 2025-07-23: We can use the current module name, this will make the code simpler in the other functions.
        $ModuleName = $MyInvocation.MyCommand.Module.Name
        if (!($ModuleName)) {
            Write-Warning "Please specify a module name using the ModuleName parameter."
            return
        }
        $ModuleNameCheck = Get-Variable -Scope Global -Name ($ModuleName + "ModuleCheck") -ErrorAction SilentlyContinue
        if ($ModuleNameCheck.Value) {
            return $true
        }
        if ($ModuleNameCheck) {
            Set-Variable -Scope Global -Name ($ModuleName + "ModuleCheck") -Value $true
        }
        else {
            New-Variable -Scope Global -Name ($ModuleName + "ModuleCheck") -Value $true
        }
        #endRegion
        
        #Get all installed versions
        $installedVersions = (Get-Module $ModuleName -ListAvailable | Sort-Object Version -Descending).Version

        #Get the lastest version available
        $availableVersion = (Find-Module -Name $ModuleName -Repository PSGallery -ErrorAction SilentlyContinue).Version

        if (!($installedVersions)) {
            if ($availableVersion ) {
                #Module not installed and there is an available version to install.
                Write-Warning ("The PowerShell Module $ModuleName is not installed, please install the latest available version ($availableVersion) with:" + [Environment]::NewLine + "Install-Module $ModuleName")
            }
            else {
                #Wrong name or not found in the registered PS Repository.
                Write-Warning ("The PowerShell Module $ModuleName not found in the registered PS Repository, please check the module name and try again.")
            }
            return $false
        }

        #Get the current loaded version
        $tmpCurrentVersion = (Get-Module $ModuleName | Sort-Object Version -Descending)
        if ($tmpCurrentVersion) {
            $currentVersion = $tmpCurrentVersion[0].Version.ToString()
        }

        if (!($currentVersion)) {
            #Module is installed but not imported, in this case we check if there is a newer version available.
            if ($availableVersion -in $installedVersions) {
                Write-Warning ("The lastest available version of $ModuleName module is installed, however the module is not imported." + [Environment]::NewLine + "Please make sure you import it with:" + [Environment]::NewLine + "Import-Module $ModuleName -RequiredVersion $availableVersion")
                return $false
            }
            else {
                Write-Warning ("There is a new version available $availableVersion, the lastest installed version is " + $installedVersions[0] + "." + [Environment]::NewLine + "Please update the module with:" + [Environment]::NewLine + "Update-Module $ModuleName")
            }
        }

        if ($currentVersion -ne $availableVersion ) {
            if ($availableVersion -in $installedVersions) {
                Write-Warning ("The lastest available version of $ModuleName module is installed, however version $currentVersion is imported." + [Environment]::NewLine + "Please make sure you import it with:" + [Environment]::NewLine + "Import-Module $ModuleName -RequiredVersion $availableVersion")
            }
            else {
                Write-Warning ("There is a new version available $availableVersion, current version $currentVersion." + [Environment]::NewLine + "Please update the module with:" + [Environment]::NewLine + "Update-Module $ModuleName")
            }
        }
        return $true
    }
    catch {
    }
    return $false
}