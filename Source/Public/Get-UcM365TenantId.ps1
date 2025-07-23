function Get-UcM365TenantId {
    <#
        .SYNOPSIS
        Get Microsoft 365 Tenant Id 

        .DESCRIPTION
        This function returns the Tenant ID associated with a domain that is part of a Microsoft 365 Tenant.

        .PARAMETER Domain
        Specifies a domain registered with Microsoft 365

        .EXAMPLE
        PS> Get-UcM365TenantId -Domain uclobby.com
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )
    
    $regexTenantID = "^(.*@)(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})$"
    $regexOnMicrosoftDomain = "^(.*@)(?!.*mail)(.*.onmicrosoft.com)$"

    try {
        #2025-07-23: All logic to check if we run this and getting the module name moved to the Test-UcPowerShellModule.
        Test-UcPowerShellModule | Out-Null

        $AllowedAudiences = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/" + $Domain + "/metadata/json/1") -UseBasicParsing | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
    }
    catch [System.Net.Http.HttpRequestException] {
        if ($PSItem.Exception.Response.StatusCode -eq "BadRequest") {
            Write-Error "The domain $Domain is not part of a Microsoft 365 Tenant."
        }
        else {
            Write-Error $PSItem.Exception.Message
        }
    }
    catch {
        Write-Error "Unknown error while checking domain: $Domain"
    }
    $output = [System.Collections.ArrayList]::new()
    $OnMicrosoftDomains = [System.Collections.ArrayList]::new()
    $TenantID = ""
    foreach ($AllowedAudience in $AllowedAudiences) {
        $tempTID = [regex]::Match($AllowedAudience , $regexTenantID).captures.groups
        $tempID = [regex]::Match($AllowedAudience , $regexOnMicrosoftDomain).captures.groups
        if ($tempTID.count -ge 2) {
            $TenantID = $tempTID[2].value 
        }
        if ($tempID.count -ge 2) {
            [void]$OnMicrosoftDomains.Add($tempID[2].value)
        }
    }
    #Multi Geo will have multiple OnMicrosoft Domains
    foreach ($OnMicrosoftDomain in $OnMicrosoftDomains) {
        if ($TenantID -and $OnMicrosoftDomain) {
            $M365TidPSObj = [PSCustomObject]@{ TenantID = $TenantID
                OnMicrosoftDomain                       = $OnMicrosoftDomain
            }
            $M365TidPSObj.PSObject.TypeNames.Insert(0, 'M365TenantId')
            [void]$output.Add($M365TidPSObj)
        }
    }
    return $output 
}