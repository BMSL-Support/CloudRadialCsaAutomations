<# 

.SYNOPSIS
    
    This function is used to add a note to a ConnectWise ticket.

.DESCRIPTION

    This function creates a new user in the tenant with the same licenses and group memberships as an existing user.

    The function requires the following environment variables to be set:

    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
    SecurityKey - Optional, use this as an additional step to secure the function

    The function requires the following modules to be installed:
    
    Microsoft.Graph

.INPUTS

    UserEmail - user email address that exists in the tenant
    GroupName - group name that exists in the tenant
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    TicketId - optional - string value of the ticket id used for transaction tracking
    SecurityKey - Optional, use this as an additional step to secure the function

    JSON Structure

    {
        "UserEmail": "email@address.com",
        "GroupName": "Group Name",
        "TenantId": "12345678-1234-1234-123456789012",
        "TicketId": "123456,
        "SecurityKey", "optional"
    }

.OUTPUTS

    JSON response with the following fields:

    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"

#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Create User Like Another User function triggered."

$resultCode = 200
$message = ""

$NewUserEmail = $Request.Body.NewUserEmail
$ExistingUserEmail = $Request.Body.ExistingUserEmail
$NewUserFirstName = $Request.Body.NewUserFirstName
$NewUserLastName = $Request.Body.NewUserLastName
$NewUserDisplayName = "$NewUserFirstName $NewUserLastName"
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

# Optional parameters
$JobTitle = $Request.Body.JobTitle
$AddJobTitle = $Request.Body.AddJobTitle
$Dept = $Request.Body.Dept
$OfficePhone = $Request.Body.OfficePhone
$MobilePhone = $Request.Body.MobilePhone

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $NewUserEmail) {
    $message = "NewUserEmail cannot be blank."
    $resultCode = 500
}
else {
    $NewUserEmail = $NewUserEmail.Trim()
}

if (-Not $ExistingUserEmail) {
    $message = "ExistingUserEmail cannot be blank."
    $resultCode = 500
}
else {
    $ExistingUserEmail = $ExistingUserEmail.Trim()
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

Write-Host "New User Email: $NewUserEmail"
Write-Host "Existing User Email: $ExistingUserEmail"
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

if ($resultCode -Eq 200) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId

    # Define the existing user's UserPrincipalName (UPN) and the new user's UPN
    $existingUserUpn = $ExistingUserEmail
    $newUserUpn = $NewUserEmail

    # Retrieve the existing user's details
    $existingUser = Get-MgUser -UserPrincipalName $existingUserUpn

    if (-Not $existingUser) {
        $message = "Request failed. User `"$ExistingUserEmail`" could not be found."
        $resultCode = 500
    }

    if ($resultCode -eq 200) {
        # Create the new user
        $newUser = New-MgUser -UserPrincipalName $newUserUpn -DisplayName $NewUserDisplayName -GivenName $NewUserFirstName -Surname $NewUserLastName -JobTitle $JobTitle -Department $Dept -OfficePhone $OfficePhone -MobilePhone $MobilePhone

        $message = "New user `"$NewUserEmail`" created successfully like user `"$ExistingUserEmail`"."
    }
}

# Prepare the output JSON
$outputJson = @{
    TenantId = $TenantId
    UserPrincipalName = $NewUserEmail
    RequestedLicense = @($existingUser.AssignedLicenses.SkuId)
    TicketId = $TicketId
}

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
    TenantId     = $TenantId
    UserPrincipalName = $NewUserEmail
    RequestedLicense = @($existingUser.AssignedLicenses.SkuId)
    TicketId     = $TicketId
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
