<# 

.SYNOPSIS
    
    This function lists all security groups in Microsoft 365, excluding dynamic groups and those with administrative rights.

.DESCRIPTION
    
    This function lists all security groups in Microsoft 365, excluding dynamic groups and those with administrative rights.
    
    The function requires the following environment variables to be set:
    
    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
    SecurityKey - Optional, use this as an additional step to secure the function
 
    The function requires the following modules to be installed:
    
    Microsoft.Graph
    
.INPUTS

    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    TicketId - optional - string value of the ticket id used for transaction tracking
    SecurityKey - Optional, use this as an additional step to secure the function

.OUTPUTS

    JSON response with the following fields:

    Name - The display name of the security group
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"

#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "ListSecurityGroups function triggered."

$resultCode = 200
$message = ""

$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $TenantId) {
    $TenantId = $env:Ms365_TenantId
}
else {
    $TenantId = $TenantId.Trim()
}

if (-Not $TicketId) {
    $TicketId = ""
}

Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

if ($resultCode -Eq 200) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId

    $groups = Get-MgGroup -Filter "groupTypes/Any(x:x eq 'Unified') and securityEnabled eq true"

    $filteredGroups = $groups | Where-Object {
        $_.GroupTypes -notcontains 'DynamicMembership' -and
        $_.DisplayName -notmatch 'Admin|Administrator|Owner|Root'
    }

    if (-Not $filteredGroups) {
        $message = "No security groups found."
        $resultCode = 500
    }

    $groupsList = $filteredGroups | ForEach-Object {
        $_.DisplayName
    }

    if ($resultCode -Eq 200) {
        $message = "Request completed. Security groups listed successfully."
    }
}

$body = @{
    Message      = $message + $groupsList
    Groups       = $groupsList
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
