<# 

.SYNOPSIS
    
    This function is used to add a user to multiple distribution groups in Microsoft 365.

.DESCRIPTION
             
    This function is used to add a user to multiple distribution groups in Microsoft 365.
    
    The function requires the following environment variables to be set:
        
    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
        
    The function requires the following modules to be installed:
        
    Microsoft.Graph

.INPUTS

    UserPrincipalName - user principal name that exists in the tenant
    GroupNames - array of group names that exist in the tenant
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    TicketId - optional - string value of the ticket id used for transaction tracking
    SecurityKey - Optional, use this as an additional step to secure the function

    JSON Structure
    {
        "UserPrincipalName": "user@domain.com",
        "GroupNames": ["Group Name 1", "Group Name 2"],
        "TenantId": "12345678-1234-1234-123456789012",
        "TicketId": "123456",
        "SecurityKey": "optional"
    }

.OUTPUTS 

    JSON response with the following fields:

    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"
    Internal - Boolean value indicating if the operation is internal

#>
using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Add User to Groups function triggered."

$resultCode = 200
$message = ""

$UserPrincipalName = $Request.Body.UserPrincipalName
$GroupNames = $Request.Body.GroupNames
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $UserPrincipalName) {
    $message = "UserPrincipalName cannot be blank."
    $resultCode = 500
}
else {
    $UserPrincipalName = $UserPrincipalName.Trim()
}

if (-Not $GroupNames -or $GroupNames.Count -eq 0 -or $GroupNames -eq "No groups available at this time.") {
    $message = "No groups specified on the form."
    $resultCode = 500
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

Write-Host "User Principal Name: $UserPrincipalName"
Write-Host "Group Names: $($GroupNames -join ', ')"
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

if ($resultCode -Eq 200)
{
    $secure365Password = ConvertTo-SecureString -String $env:
