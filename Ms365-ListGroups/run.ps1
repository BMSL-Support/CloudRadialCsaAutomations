<# 

.SYNOPSIS
    
    This function lists all MS Teams in Microsoft 365.

.DESCRIPTION
    
    This function lists all MS Teams in Microsoft 365.
    
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
    TicketNotes - optional - string value of the ticket notes
    SecurityKey - Optional, use this as an additional step to secure the function

.OUTPUTS

    JSON response with the following fields:

    Name - The display name of the MS Team
    TicketId - TicketId passed in Parameters
    TicketNotes - TicketNotes passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"

#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "ListTeams function triggered."

$resultCode = 200
$message = ""

$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$TicketNotes = $Request.Body.TicketNotes
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

if (-Not $TicketNotes) {
    $TicketNotes = ""
}

Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"
Write-Host "Ticket Notes: $TicketNotes"

if ($resultCode -Eq 200) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId

    $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')"

    if (-Not $teams) {
        $message = "No MS Teams found."
        $resultCode = 500
    }

    $teamsList = $teams | ForEach-Object {
        $_.DisplayName
    }

    if ($resultCode -Eq 200) {
        $message = "Request completed. MS Teams listed successfully."
    }
}

$body = @{
    Message      = $message
    Teams        = $teamsList
    TicketId     = $TicketId
    TicketNotes  = $TicketNotes
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
