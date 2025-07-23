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

    #2025-07-23: All logic to check if we run this and getting the module name moved to the Test-UcPowerShellModule.
    Test-UcPowerShellModule | Out-Null
    #endregion
    
    $outFile = "M365LicenseAssigment_" 
    #region 2024-09-05: Users with Duplicate Service Plans
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
        #2023-10-19: Change: OutputPath will be for both report and Product names and service plan identifiers for licensing.csv
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
            $SKUnSP = import-CSV -Path $SKUnSPFilePath
        }
        catch {
            Write-Warning "Could not import Service Plan ID file."
            $UseFriendlyNames = $false
        }
    }

    #region 2024-09-05: Combined Graph calls for SKUs, Licensed Groups
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

    #region 2023-10-19: Adding filter to SKU
    if ($SKU) {
        if ($UseFriendlyNames) {
            #2024-10-22: Change to allow to search using SKU parameter instead of exact match.
            $SKUGUID = ($SKUnSP | Where-Object { $_.String_Id -match $SKU -or $_.Product_Display_Name -match $SKU } | Sort-Object GUID -Unique ).GUID
            $TenantSKUs = $TenantSKUs | Where-Object { $_.skuId -in $SKUGUID -or $_.skuPartNumber -match $SKU }
            if ($TenantSKUs.count -eq 0) {
                Write-Warning "Could not find `"$SKU`" (SKU Name/Part Number) subscription associated with the tenant."
                return 
            }
        }
        else {
            #2024-10-22: Change to allow to search using SKU parameter instead of exact match.
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

    #region 2023-10-19: Getting all Service Plans for new matrix style report
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
    $row = "UserPrincipalName,LicenseAssigned,LicenseAssignment,LicenseAssignmentGroup"
    if (!($SkipServicePlan)) {
        foreach ($ServicePlan in $allServicePlans) {
            $row += "," + $ServicePlan.servicePlanName
        }
    }
    $row += [Environment]::NewLine
    #endregion

    Write-Progress -Id 2 -Activity "Export-UcM365LicenseAssignment, Step 2: Reading users assigned licenses/service plans"
    if ($DuplicateServicePlansOnly) {
        #region 2024-09-05: Users with Duplicate Service Plans
        #We need to check license per user, this is slower than check per SKU like in Licensing Assignment but required in order to detect duplicates.
        $TotalUsers = 0
        $usersProcessed = 0
        $GraphNextPage = "/users?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=userPrincipalName,licenseAssignmentStates&`$top=999"
        do {
            $GraphResponse = Invoke-UcGraphRequest -Path $GraphNextPage -Headers $GraphRequestHeader
            $GraphNextPage = $GraphResponse.'@odata.nextLink'
            $UsersWithLicenses = $GraphResponse.value
            if (![string]::IsNullOrEmpty($GraphResponse.'@odata.count')) {
                $TotalUsers = $GraphResponse.'@odata.count'
            }
            
            foreach ($LicensedUser in $UsersWithLicenses) {
                $usersProcessed++
                #Update the status every 100 users
                if (($usersProcessed % 100 -eq 0) -or ($usersProcessed -eq $TotalUsers) -or ($usersProcessed -eq 1) ) {
                    Write-Progress -Id 2 -Activity "Export-UcM365LicenseAssignment, Step 2: Reading users assigned licenses/service plans" -Status "$usersProcessed of $TotalUsers"
                }
                #We only need to process users that have 2 or more licenses assigned.
                if ($LicensedUser.licenseAssignmentStates.count -gt 1) {
                    $tmpUserServicePlans = [System.Collections.ArrayList]::new()
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
                            
                            $SKUUserServicePlans = $tmpLicenseInfo.servicePlans | Where-Object -Property appliesTo -EQ -Value "User" | Sort-Object servicePlanName
                            foreach ($SKUUserServicePlan in $SKUUserServicePlans) {
                                $SPStatus = "Off"
                                if ($SKUUserServicePlan.servicePlanId -notin $licenseState.disabledPlans) {
                                    $SPStatus = "On"
                                }
                                $ObjUserServicePlans = [PSCustomObject]@{
                                    LicenseSkuId           = $licenseState.skuId
                                    LicenseDisplayName     = $LicenseDisplayName
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
                    
                    #Checking if we have more then one Service Plan
                    #In the future we can add filters, like only if both are ON or Ignore Direct/Inherited
                    $skuWithDupServicePlans = $tmpUserServicePlans | Group-Object -Property ServicePlanId | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group | Select-Object LicenseSkuId, LicenseDisplayName, LicenseAssignment, LicenseAssignmentGroup  | Sort-Object -Property LicenseDisplayName, LicenseAssignment, LicenseAssignmentGroup -Unique 
                    if (($skuWithDupServicePlans.Count -gt 0)) {
                        foreach ($UserLicenseState in  $skuWithDupServicePlans) {
                            foreach ($ServicePlan in $allServicePlans) {
                                $tmpSPStatus = $tmpUserServicePlans | Where-Object { $UserLicenseState.LicenseSkuId -eq $_.LicenseSkuId -and $UserLicenseState.LicenseAssignment -eq $_.LicenseAssignment -and $UserLicenseState.LicenseAssignmentGroup -eq $_.LicenseAssignmentGroup -and $_.servicePlanId -eq $servicePlan.servicePlanId }
                                if ($tmpSPStatus.Status -in ("On", "Off")) {
                                    $userServicePlans += "," + $tmpSPStatus.Status
                                }
                                else {
                                    $userServicePlans += ","
                                }
                            }
                            $row += $LicensedUser.userPrincipalName + "," + $UserLicenseState.LicenseDisplayName + "," + $UserLicenseState.LicenseAssignment + "," + $UserLicenseState.LicenseAssignmentGroup + $userServicePlans
                            Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
                            $row = ""
                            $userServicePlans = ""
                        }
                    }
                }
            }
        } while (!([string]::IsNullOrEmpty($GraphNextPage)))
        #endregion
    }
    else {
        #region License Assignment
        foreach ($TenantSKU in $TenantSKUs) {
            $LicenseDisplayName = $TenantSKU.skuPartNumber
            if ($UseFriendlyNames) {
                $tmpFriendlyName = ($SKUnSP | Where-Object { $_.GUID -eq $TenantSKU.skuID } | Sort-Object Product_Display_Name -Unique).Product_Display_Name
                #2024-10-22: To prevent empty name when a license exists in the tenant but the data is not available in "Products names and Services Identifiers" file.
                if ($tmpFriendlyName) {
                    $LicenseDisplayName = $tmpFriendlyName
                }  
            }
            $SKUUserServicePlans = $TenantSKU.servicePlans | Where-Object -Property appliesTo -EQ -Value "User" | Sort-Object servicePlanName
            $usersProcessed = 0       
            $GraphRequestURI = "https://graph.microsoft.com/v1.0/users?`$filter=assignedLicenses/any(u:u/skuId eq " + $TenantSKU.skuId + " )&`$select=userPrincipalName,licenseAssignmentStates&`$orderby=userPrincipalName&`$count=true&`$top=999"
            do {
                try {
                    $UsersWithLicenses = Invoke-MgGraphRequest -Method Get -Uri $GraphRequestURI -Headers $GraphRequestHeader
                    if (![string]::IsNullOrEmpty($UsersWithLicenses.'@odata.count')) {
                        $TotalUsers = $UsersWithLicenses.'@odata.count'
                    }
                    $GraphRequestURI = $UsersWithLicenses.'@odata.nextLink'
                    foreach ($UserWithLicense in $UsersWithLicenses.value) {
                        if (($usersProcessed % 1000 -eq 0) -or ($usersProcessed -eq $TotalUsers)) {
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
                            $row += $UserWithLicense.userPrincipalName + "," + $LicenseDisplayName + "," + $LicenseAssignment + "," + $LicenseAssignmentGroup + $userServicePlans
                            Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
                            $row = ""
                        }
                        $usersProcessed++
                    }
                }
                catch {
                    Write-Warning ("Failed to get Users with assigned SKU Id: " + $TenantSKU.skuID)
                    $GraphRequestURI = ""
                }
            } while (![string]::IsNullOrEmpty($GraphRequestURI))
        }
        #endregion
    }

    if ($usersProcessed -gt 0) {
        Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
        #region 2023-10-19: Added execution time to the output.
        $endTime = Get-Date
        $totalSeconds = [math]::round(($endTime - $startTime).TotalSeconds, 2)
        $totalTime = New-TimeSpan -Seconds $totalSeconds
        Write-Host "Execution time:" $totalTime.Hours "Hours" $totalTime.Minutes "Minutes" $totalTime.Seconds "Seconds" -ForegroundColor Green
        #endregion
    }
}