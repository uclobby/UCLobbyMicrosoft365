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
    try {
        #CD20250723: All logic to check if we run this and getting the module name moved to the Test-UcPowerShellModule.
        Test-UcPowerShellModule | Out-Null

        $TenantId = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/" + $Domain + "/metadata/json/1") -UseBasicParsing | ConvertFrom-Json | Select-Object -ExpandProperty Realm
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

    #CD20260414: Return tenant ID only after the recent Microsoft 365 change.
    return $TenantId 
}