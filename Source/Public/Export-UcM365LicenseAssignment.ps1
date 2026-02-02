function Export-UcM365LicenseAssignment {
    <#
        .SYNOPSIS
        Generate a report of the User assigned licenses either direct or assigned by group (Inherited)

        .DESCRIPTION
        This script will get a report of all Service Plans assigned to users and how the license is assigned to the user (Direct, Inherited)

        Contributors: David Paulino, Freydem Fernandez Lopez, Gal Naor

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)
                        
                        Microsoft Graph Scopes:
                            "Directory.Read.All"
        
        .PARAMETER UseFriendlyNames
        When present will download a csv file containing the License/ServicePlans friendly names

        Product names and service plan identifiers for licensing
        https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

        .PARAMETER SkipServicePlan
        When present will just check the licenses and not the service plans assigned to the user.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results. By default, it will save on current user Download.

        .PARAMETER DuplicateServicePlansOnly
        When present the report will be the users that have the same service plan from different assigned licenses.

        .EXAMPLE 
        PS> Export-UcM365LicenseAssignment

        .EXAMPLE 
        PS> Export-UcM365LicenseAssignment -UseFriendlyNames
    #>
    [cmdletbinding()]
    param(
        [string]$SKU,    
        [switch]$UseFriendlyNames,
        [switch]$SkipServicePlan,
        [string]$OutputPath,
        [switch]$DuplicateServicePlansOnly
    )

    $startTime = Get-Date
    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "Directory.Read.All" -AltScopes ("User.Read.All", "Organization.Read.All"))) {
        return
    }

    #CD20250723: All logic to check if we run this and getting the module name moved to the Test-UcPowerShellModule.
    Test-UcPowerShellModule | Out-Null
    #endregion
    
    $outFile = "M365LicenseAssigment_" 
    #region CD20240905: Users with Duplicate Service Plans
    if ($DuplicateServicePlansOnly) {
        $outFile += "DuplicateServicePlansOnly_"
    }
    #endregion
    $outFile += (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"

    #Verify if the Output Path exists
    if ($OutputPath) {
        if (!(Test-Path $OutputPath -PathType Container)) {
            Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
            return
        }
        $OutputFilePath = [System.IO.Path]::Combine($OutputPath, $outFile)
    }
    else {                
        $OutputFilePath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFile)
    }
        
    if ($UseFriendlyNames) {
        #CD20231019: OutputPath will be for both report and Product names and service plan identifiers for licensing.csv
        $SKUnSPFilePath = [System.IO.Path]::Combine($OutputPath, "Product names and service plan identifiers for licensing.csv")
        if (!(Test-Path -Path $SKUnSPFilePath)) {
            try {
                Write-Warning "M365 Product Names and Service Plans file not found, attempting to download it."
                Invoke-WebRequest -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" -OutFile $SKUnSPFilePath
            }
            catch {
                Write-Warning "Could not download M365 Product Names and Service Plans."
            }
        }
        try {
            $SKUnSP = Import-Csv -Path $SKUnSPFilePath
        }
        catch {
            Write-Warning "Could not import Service Plan ID file."
            $UseFriendlyNames = $false
        }
    }

    #region CD20240905: Combined Graph calls for SKUs, Licensed Groups
    #Tenant SKUs - All Licenses that exist in the tenant
    $graphRequests = [System.Collections.ArrayList]::new()
    $gRequestTmp = New-Object -TypeName PSObject -Property @{
        id     = "TenantSKUs"
        method = "GET"
        url    = "/subscribedSkus?`$select=skuID,skuPartNumber,servicePlans,appliesTo,consumedUnits"
    }
    [void]$graphRequests.Add($gRequestTmp)

    #Groups with Licenses Assignment.
    $GraphRequestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
    $GraphRequestHeader.Add("ConsistencyLevel", "eventual")
    $gRequestTmp = New-Object -TypeName PSObject -Property @{
        id      = "GroupsWithLicenses"
        method  = "GET"
        headers = $GraphRequestHeader
        url     = "/groups?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=id,displayName,assignedLicenses&`$top=999"
    }
    [void]$graphRequests.Add($gRequestTmp)
    $BatchResponse = Invoke-UcGraphRequest -Requests $graphRequests -Beta -IncludeBody -Activity "Export-UcM365LicenseAssignment, Step 1: Getting Tenant License details"

    $tmpGraphResponse = $BatchResponse | Where-Object { $_.id -eq ("TenantSKUs") }
    if ($tmpGraphResponse.status -eq 200) {
        $TenantSKUs = $tmpGraphResponse.body.value
    }
    $tmpGraphResponse = $BatchResponse | Where-Object { $_.id -eq ("GroupsWithLicenses") }
    if ($tmpGraphResponse.status -eq 200) {
        $GroupsWithLicenses = $tmpGraphResponse.body.value
    }
    #endregion

    #region CD20231019: Adding filter to SKU
    if ($SKU) {
        if ($UseFriendlyNames) {
            #CD20241022: Change to allow to search using SKU parameter instead of exact match.
            $SKUGUID = ($SKUnSP | Where-Object { $_.String_Id -match $SKU -or $_.Product_Display_Name -match $SKU } | Sort-Object GUID -Unique ).GUID
            $TenantSKUs = $TenantSKUs | Where-Object { $_.skuId -in $SKUGUID -or $_.skuPartNumber -match $SKU }
            if ($TenantSKUs.count -eq 0) {
                Write-Warning "Could not find `"$SKU`" (SKU Name/Part Number) subscription associated with the tenant."
                return 
            }
        }
        else {
            #CD20241022: Change to allow to search using SKU parameter instead of exact match.
            $TenantSKUs = $TenantSKUs | Where-Object { $_.skuPartNumber -match $SKU }
            if ($TenantSKUs.count -eq 0) {
                Write-Warning "Could not find `"$SKU`" (SKU Part Number) subscription associated with the tenant."
                return 
            }
        }
    }
    else {
        $TenantSKUs = $TenantSKUs | Where-Object -Property consumedUnits -GT -Value 0 | Sort-Object skuPartNumber
    }
    #endregion

    #region CD20231019: Getting all Service Plans for new matrix style report
    $allServicePlans = [System.Collections.ArrayList]::new()
    foreach ($TenantSKU in $TenantSKUs) {
        $tmpUserServicePlans = $TenantSKU.ServicePlans | Where-Object -Property appliesTo -EQ -Value "User" 
        foreach ($ServicePlan in $tmpUserServicePlans) {
            if (!($ServicePlan.ServicePlanId -in $allServicePlans.ServicePlanId)) {
                if ($UseFriendlyNames) {
                    $servicePlanName = ($SKUnSP | Where-Object { $_.Service_Plan_Id -eq $ServicePlan.ServicePlanId -and $_.GUID -eq $TenantSKU.skuID } | Sort-Object Service_Plans_Included_Friendly_Names -Unique).Service_Plans_Included_Friendly_Names
                    if ([string]::IsNullOrEmpty($servicePlanName)) {
                        $servicePlanName = $ServicePlan.servicePlanName    
                    }
                }
                else {
                    $servicePlanName = $ServicePlan.ServicePlanName
                }
                $tmpSP = New-Object -TypeName PSObject -Property @{
                    servicePlanId   = $ServicePlan.ServicePlanId
                    servicePlanName = $servicePlanName
                }
                [void]$allServicePlans.Add($tmpSP)
            }
        }
    }
    #Sorting service plans by name and creating the file header
    $allServicePlans = $allServicePlans | Sort-Object ServicePlanName
    $row = "UserPrincipalName,LicenseAssigned,LicenseState,LicenseAssignment,LicenseAssignmentGroup"
    if (!($SkipServicePlan)) {
        foreach ($ServicePlan in $allServicePlans) {
            $row += "," + $ServicePlan.servicePlanName
        }
    }
    $row += [Environment]::NewLine
    #endregion

    Write-Progress -Id 2 -Activity "Export-UcM365LicenseAssignment, Step 2: Reading users assigned licenses/service plans"
    #region CD20240905: Users with Duplicate Service Plans
    if ($DuplicateServicePlansOnly) {
        #We need to check license per user, this is slower than check per SKU like in Licensing Assignment but required in order to detect duplicates.
        $usersProcessed = 1
        $UserLicenseAssignmentURI = "/users?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=userPrincipalName,licenseAssignmentStates&`$top=999"
        #region CD20251017: On larger tenants is better to process each graph response that can have up to 999 results.
        try {
            do {
                $GraphResponse = Invoke-UcGraphRequest -Path $UserLicenseAssignmentURI -Header $GraphRequestHeader -Raw
                $UserLicenseAssignmentURI = $GraphResponse.'@odata.nextLink'
                #We only want to run this once so we can use the usersProcessed.
                if ($usersProcessed -eq 1 ) {
                    $TotalUsers = $GraphResponse.'@odata.count'
                    #In case we dont have @odata.count it means only one entry is returned.
                    if (!($TotalUsers)) {
                        $TotalUsers = 1
                    }
                }
                if ($GraphResponse.value) {
                    $UsersWithLicenses = $GraphResponse.value
                }
                else {
                    $UsersWithLicenses = $GraphResponse
                }
                foreach ($LicensedUser in $UsersWithLicenses) {
                    #Update the status every 100 users
                    if (($usersProcessed % 100 -eq 0) -or ($usersProcessed -eq $TotalUsers) -or ($usersProcessed -eq 1) ) {
                        Write-Progress -Id 2 -Activity "Export-UcM365LicenseAssignment, Step 2: Reading users assigned licenses/service plans" -Status "$usersProcessed of $TotalUsers"
                    }
                    $tmpUserServicePlans = [System.Collections.ArrayList]::new()
                    #We only need to process users that have 2 or more licenses assigned.
                    if ($LicensedUser.licenseAssignmentStates.count -gt 1) {
                        foreach ($licenseState in $LicensedUser.licenseAssignmentStates) {
                            #If not a in the Tenant SKUs we can skip it
                            if ($licenseState.skuId -in $TenantSKUs.skuId) {
                                $tmpLicenseInfo = ($TenantSKUs | Where-Object { $_.skuId -eq $licenseState.skuId })
                                $LicenseDisplayName = $tmpLicenseInfo.skuPartNumber
                                if ($UseFriendlyNames) {
                                    $LicenseDisplayName = ($SKUnSP | Where-Object { $_.GUID -eq $licenseState.skuId } | Sort-Object Product_Display_Name -Unique).Product_Display_Name
                                }
                                if ([string]::IsNullOrEmpty($LicenseDisplayName)) {
                                    $LicenseDisplayName = $licenseState.skuId
                                }
                                $licenseAssignment = "Direct"
                                $licenseAssignmentGroup = "NA"
                                if (!([string]::IsNullOrEmpty($licenseState.assignedByGroup))) {
                                    $licenseAssignment = "Inherited"
                                    $licenseAssignmentGroup = ($GroupsWithLicenses | Where-Object -Property "id" -EQ -Value $licenseState.assignedByGroup).displayName
                                    if ([string]::IsNullOrEmpty($licenseAssignmentGroup)) {
                                        $licenseAssignmentGroup = $licenseState.assignedByGroup
                                    }
                                }
                                if ($SkipServicePlan) {
                                    $ObjUserServicePlans = [PSCustomObject]@{
                                        LicenseSkuId           = $licenseState.skuId
                                        LicenseDisplayName     = $LicenseDisplayName
                                        LicenseState           = $licenseState.state
                                        LicenseAssignment      = $licenseAssignment
                                        LicenseAssignmentGroup = $licenseAssignmentGroup
                                    }
                                    [void]$tmpUserServicePlans.Add($ObjUserServicePlans)
                                }
                                else {
                                    $SKUUserServicePlans = $tmpLicenseInfo.servicePlans | Where-Object -Property appliesTo -EQ -Value "User" | Sort-Object servicePlanName
                                    foreach ($SKUUserServicePlan in $SKUUserServicePlans) {
                                        $SPStatus = "Off"
                                        if ($SKUUserServicePlan.servicePlanId -notin $licenseState.disabledPlans) {
                                            $SPStatus = "On"
                                        }
                                        $ObjUserServicePlans = [PSCustomObject]@{
                                            LicenseSkuId           = $licenseState.skuId
                                            LicenseDisplayName     = $LicenseDisplayName
                                            LicenseState           = $licenseState.state
                                            LicenseAssignment      = $licenseAssignment
                                            LicenseAssignmentGroup = $licenseAssignmentGroup
                                            ServicePlanId          = $SKUUserServicePlan.servicePlanId
                                            ServicePlanName        = $SKUUserServicePlan.servicePlanName
                                            Status                 = $SPStatus
                                        }
                                        [void]$tmpUserServicePlans.Add($ObjUserServicePlans)
                                    }
                                }
                            }
                        }
                        #Checking if we have more then one Service Plan
                        #In the future we can add filters, like only if both are ON or Ignore Direct/Inherited
                        $userServicePlans = ""
                        if ($SkipServicePlan) {
                            $skuWithDupServicePlans = $tmpUserServicePlans | Group-Object -Property LicenseSkuId | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group | Select-Object LicenseSkuId, LicenseDisplayName, LicenseState , LicenseAssignment, LicenseAssignmentGroup  | Sort-Object -Property LicenseDisplayName, LicenseAssignment, LicenseAssignmentGroup -Unique 
                        }
                        else {
                            $skuWithDupServicePlans = $tmpUserServicePlans | Group-Object -Property ServicePlanId | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group | Select-Object LicenseSkuId, LicenseDisplayName, LicenseState , LicenseAssignment, LicenseAssignmentGroup | Sort-Object -Property LicenseDisplayName, LicenseAssignment, LicenseAssignmentGroup -Unique 
                        }
                        foreach ($UserLicenseState in  $skuWithDupServicePlans) {
                            if (($skuWithDupServicePlans.Count -gt 0)) {
                                if (!($SkipServicePlan)) {
                                    foreach ($ServicePlan in $allServicePlans) {
                                        $tmpSPStatus = $tmpUserServicePlans | Where-Object { $UserLicenseState.LicenseSkuId -eq $_.LicenseSkuId -and $UserLicenseState.LicenseAssignment -eq $_.LicenseAssignment -and $UserLicenseState.LicenseAssignmentGroup -eq $_.LicenseAssignmentGroup -and $_.servicePlanId -eq $servicePlan.servicePlanId }
                                        $userServicePlans += "," + $tmpSPStatus.Status
                                    }                        
                                }
                                $row += $LicensedUser.userPrincipalName + "," + $UserLicenseState.LicenseDisplayName + "," + $UserLicenseState.LicenseState + "," + $UserLicenseState.LicenseAssignment + "," + $UserLicenseState.LicenseAssignmentGroup + $userServicePlans + [Environment]::NewLine
                                $userServicePlans = ""
                            }
                        }
                    }
                    $usersProcessed++
                }
                if ($row) {
                    Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -Append -NoNewline
                    $row = ""
                }
            } while ($UserLicenseAssignmentURI)
        }
        catch {}
        #endregion
    }
    #endregion
    else {
        #region License Assignment
        foreach ($TenantSKU in $TenantSKUs) {
            $LicenseDisplayName = $TenantSKU.skuPartNumber
            if ($UseFriendlyNames) {
                $tmpFriendlyName = ($SKUnSP | Where-Object { $_.GUID -eq $TenantSKU.skuID } | Sort-Object Product_Display_Name -Unique).Product_Display_Name
                #CD20241022: To prevent empty name when a license exists in the tenant but the data is not available in "Products names and Services Identifiers" file.
                if ($tmpFriendlyName) {
                    $LicenseDisplayName = $tmpFriendlyName
                }  
            }
            $SKUUserServicePlans = $TenantSKU.servicePlans | Where-Object -Property appliesTo -EQ -Value "User" | Sort-Object servicePlanName
            $usersProcessed = 1       
            $GraphRequestURI = "/users?`$filter=assignedLicenses/any(u:u/skuId eq " + $TenantSKU.skuId + " )&`$select=userPrincipalName,licenseAssignmentStates&`$orderby=userPrincipalName&`$count=true&`$top=999"
            Write-Verbose ("Next Request URI: $GraphRequestURI - " + $GraphRequestURI.count)
            #region CD20251017: On larger tenants is better to process each graph response that can have up to 999 results.
            try {
                do {
                    $GraphResponse = Invoke-UcGraphRequest -Path $GraphRequestURI -Header $GraphRequestHeader -Raw
                    $GraphRequestURI = $GraphResponse.'@odata.nextLink'
                    #We only want to run this once so we can use the usersProcessed.
                    if ($usersProcessed -eq 1 ) {
                        
                        $TotalUsers = $GraphResponse.'@odata.count'
                        #In case we dont have @odata.count it means only one entry is returned.
                        if (!($TotalUsers)) {
                            $TotalUsers = 1
                        }
                    }
                    if ($GraphResponse.value) {
                        $UsersWithLicenses = $GraphResponse.value
                    }
                    else {
                        $UsersWithLicenses = $GraphResponse
                    }
                    foreach ($UserWithLicense in $UsersWithLicenses) {
                        if (($usersProcessed % 100 -eq 0) -or ($usersProcessed -eq $TotalUsers)) {
                            Write-Progress -ParentId 2 -Activity "Checking license assignments for $LicenseDisplayName" -Status "$usersProcessed of $TotalUsers"
                        }
                        $tmpLicenseAssignmentStates = $UserWithLicense.licenseAssignmentStates | Where-Object -Property skuId -EQ -Value $TenantSKU.skuId | Sort-Object assignedByGroup
                        foreach ($licenseState in $tmpLicenseAssignmentStates) {
                            $licenseAssignment = "Direct"
                            $licenseAssignmentGroup = ""
                            if (!([string]::IsNullOrEmpty($licenseState.assignedByGroup))) {
                                $licenseAssignment = "Inherited"
                                $licenseAssignmentGroup = ($GroupsWithLicenses | Where-Object -Property "id" -EQ -Value $licenseState.assignedByGroup).displayName
                                if ([string]::IsNullOrEmpty($licenseAssignmentGroup)) {
                                    $licenseAssignmentGroup = $licenseState.assignedByGroup
                                }
                            }
                            $userServicePlans = ""
                            if (!($SkipServicePlan)) {
                                foreach ($ServicePlan in $allServicePlans) {
                                    if ($servicePlan.servicePlanId -in $SKUUserServicePlans.servicePlanId) {
                                        if ($servicePlan.servicePlanId -notin $licenseState.disabledPlans) {
                                            $userServicePlans += ",On"
                                        }
                                        else {
                                            $userServicePlans += ",Off"
                                        }
                                    }
                                    else {
                                        $userServicePlans += ","
                                    }
                                }
                            }
                            $row += $UserWithLicense.userPrincipalName + "," + $LicenseDisplayName + "," + $licenseState.state + "," + $LicenseAssignment + "," + $LicenseAssignmentGroup + $userServicePlans + [Environment]::NewLine 
                        }
                        $usersProcessed++
                    }
                    #CD20251017: Improving performance by only writing to the output file once per request instead of every line.
                    if ($row) {
                        Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -Append -NoNewline
                        $row = ""
                    }
                } while ($GraphRequestURI)
            }
            catch {
                Write-Warning ("Failed to get Users with assigned SKU Id: " + $TenantSKU.skuID)
                $GraphRequestURI = ""
            }
            #endregion
        }
        #endregion
    }

    if ($usersProcessed -gt 0) {
        Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
        #region CD20231019: Added execution time to the output.
        $endTime = Get-Date
        $totalSeconds = [math]::round(($endTime - $startTime).TotalSeconds, 2)
        $totalTime = New-TimeSpan -Seconds $totalSeconds
        Write-Host "Execution time:" $totalTime.Hours "Hours" $totalTime.Minutes "Minutes" $totalTime.Seconds "Seconds" -ForegroundColor Green
        #endregion
    }
}
