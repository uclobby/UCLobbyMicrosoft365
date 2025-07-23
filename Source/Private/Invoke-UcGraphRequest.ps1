function Invoke-UcGraphRequest {
    <#
        .SYNOPSIS
        Invoke a Microsoft Graph Request using Entra Auth or Microsoft.Graph.Authentication 

        .DESCRIPTION
        This function will send a Microsoft Graph request to an available connections, "Test-UcServiceConnection -Type MsGraph" will have to be executed first to determine if we have a session with EntraAuth or Microsoft.Graph.Authentication.

        Requirements:   EntraAuth PowerShell module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

        .PARAMETER Path
        Specifies Microsoft Graph Path that we want to send the request.

        .PARAMETER Header
        Specify the header for cases we need to have a custom header.

        .PARAMETER Requests
        If wwe want to send a batch request.

        .PARAMETER Beta
        When present, it will use the Microsoft Graph Beta API.

        .PARAMETER IncludeBody
        Some Ms Graph APIs can require specific AuthType, Application or Delegated (User).

        .PARAMETER Activity
        For Batch requests we have use this for Activity Progress.
    #>
    param(
        [string]$Path = "/`$batch",
        [object]$Header,
        [object]$Requests,
        [switch]$Beta,
        [switch]$IncludeBody,
        [string]$Activity
    )
    #This is an easy way to switch between v1.0 and beta.
    $BatchPath = "`$batch"
    if ($Beta) {
        $Path = "../beta" + $Path
        $BatchPath = "../beta/`$batch"
    }

    #If requests then we need to do a batch request to Graph.
    if (!$Requests) {
        if ($script:GraphEntraAuth) {
            if ($Header) {
                return Invoke-EntraRequest -Path $Path -NoPaging -Header $Header
            } 
            return Invoke-EntraRequest -Path $Path -NoPaging
        }
        else {
            if ($Header) {
                $GraphResponse = Invoke-MgGraphRequest -Uri ("/v1.0/" + $Path) -Headers $Header
            }
            else {
                $GraphResponse = Invoke-MgGraphRequest -Uri ("/v1.0/" + $Path)
            }
            #When it's more than one result Invoke-MgGraphRequest returns "value", we need to remove it to match EntraAuth behaviour.
            if ($GraphResponse.value) {
                return $GraphResponse.value
            }
            else {
                return $GraphResponse
            }
        }
    }
    else {
        $outBatchResponses = [System.Collections.ArrayList]::new()
        $tmpGraphRequests = [System.Collections.ArrayList]::new()
        $g = 1
        $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
        $requestHeader.Add("Content-Type", "application/json")
        #If activity is null then we can use this to get the function that call this function.
        if (!($Activity)) {
            $Activity = [string]$(Get-PSCallStack)[1].FunctionName
        }
        $batchCount = [int][Math]::Ceiling(($Requests.count / 20))
        foreach ($GraphRequest in $Requests) {
            Write-Progress -Activity $Activity -Status "Running batch $g of $batchCount"
            [void]$tmpGraphRequests.Add($GraphRequest) 
            if ($tmpGraphRequests.Count -ge 20) {
                $g++
                $grapRequestBody = ' { "requests":  ' + ($tmpGraphRequests  | ConvertTo-Json) + ' }' 
                if ($script:GraphEntraAuth) {
                    #TODO: Add support for Graph Batch with EntraAuth
                    $GraphResponses += (Invoke-EntraRequest -Path $BatchPath -Body $grapRequestBody -Method Post -Header $requestHeader).responses
                }
                else {
                    $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri ("/v1.0/" + $BatchPath) -Body $grapRequestBody).responses
                }
                $tmpGraphRequests = [System.Collections.ArrayList]::new()
            }
        }
        
        if ($tmpGraphRequests.Count -gt 0) {
            Write-Progress -Activity $Activity -Status "Running batch $g of $batchCount"
            #TO DO: Look for alternatives instead of doing this.
            if ($tmpGraphRequests.Count -gt 1) {
                $grapRequestBody = ' { "requests":  ' + ($tmpGraphRequests | ConvertTo-Json) + ' }' 
            }
            else {
                $grapRequestBody = ' { "requests": [' + ($tmpGraphRequests | ConvertTo-Json) + '] }' 
            }
            try {
                if ($script:GraphEntraAuth) {
                    #TODO: Add support for Graph Batch with EntraAuth
                    $GraphResponses += (Invoke-EntraRequest -Path $BatchPath -Body $grapRequestBody -Method Post -Header $requestHeader).responses
                }
                else {
                    $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri  ("/v1.0/" + $BatchPath) -Body $grapRequestBody).responses
                }
            }
            catch {
                Write-Warning "Error while getting the Graph Request."
            }
        }
        
        #In some cases we will need the complete graph response, in that case the calling function will have to process pending pages.
        $attempts = 1
        for ($j = 0; $j -lt $GraphResponses.length; $j++) {
            $ResponseCount = 0
            if ($IncludeBody) {
                $outBatchResponses += $GraphResponses[$j]
            }
            else {
                $outBatchResponses += $GraphResponses[$j].body
                if ($GraphResponses[$j].status -eq "200") {
                    #Checking if there are more pages available    
                    $GraphURI_NextPage = $GraphResponses[$j].body.'@odata.nextLink'
                    $GraphTotalCount = $GraphResponses[$j].body.'@odata.count'
                    $ResponseCount += $GraphResponses[$j].body.value.count
                    while (![string]::IsNullOrEmpty($GraphURI_NextPage)) {
                        try {
                            if ($script:GraphEntraAuth) {
                                #TODO: Add support for Graph Batch with EntraAuth, for now we need to use NoPaging to have the same behaviour as Invoke-MgGraphRequest
                                $graphNextPageResponse = Invoke-EntraRequest -Path $GraphURI_NextPage -NoPaging
                            }
                            else {
                                $graphNextPageResponse = Invoke-MgGraphRequest -Method Get -Uri $GraphURI_NextPage
                            }
                            $outBatchResponses += $graphNextPageResponse
                            $GraphURI_NextPage = $graphNextPageResponse.'@odata.nextLink'
                            $ResponseCount += $graphNextPageResponse.value.count
                            Write-Progress -Activity $Activity -Status "$ResponseCount of $GraphTotalCount"
                        }
                        catch {
                            Write-Warning "Failed to get the next batch page, retrying..."
                            $attempts--
                        }
                        if ($attempts -eq 0) {
                            Write-Warning "Could not get next batch page, skiping it."
                            break
                        }
                    }
                }
                else {
                    Write-Warning ("Failed to get Graph Response" + [Environment]::NewLine + `
                            "Error Code: " + $GraphResponses[$j].status + " " + $GraphResponses[$j].body.error.code + [Environment]::NewLine + `
                            "Error Message: " + $GraphResponses[$j].body.error.message + [Environment]::NewLine + `
                            "Request Date: " + $GraphResponses[$j].body.error.innerError.date + [Environment]::NewLine + `
                            "Request ID: " + $GraphResponses[$j].body.error.innerError.'request-id' + [Environment]::NewLine + `
                            "Client Request Id: " + $GraphResponses[$j].body.error.innerError.'client-request-id')
                } 
            }
        }
        return $outBatchResponses
    }
}