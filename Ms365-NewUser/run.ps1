<# 
.SYNOPSIS
    This function creates a new user in the tenant.
.DESCRIPTION
    This function creates a new user in the tenant.
    The function requires the following environment variables to be set:
    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
    SecurityKey - Optional, use this as an additional step to secure the function
    The function requires the following modules to be installed:
    Microsoft.Graph
.INPUTS
    JSON Structure
    {
        "TicketId": "123456",
        "TenantId": "12345678-1234-1234-123456789012",
        "NewUserFirstName": "John",
        "NewUserLastName": "Doe",
        "NewUserEmail": "john.doe@example.com",
        "JobTitle": "Software Engineer",
        "OfficePhone": "+1234567890",
        "MobilePhone": "+0987654321",
        "LicenseTypes": ["ENTERPRISEPACK", "STANDARDPACK"]
    }
.OUTPUTS
    JSON response with the following fields:
    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"
    UserPrincipalName - UPN of the new user created
    TenantId - Tenant Id of the Microsoft 365 tenant
    LicenseTypes - Array of license types
#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Create New User function triggered."

$resultCode = 200
$message = ""
$UserPrincipalName = ""

$TicketId = $Request.Body.TicketId
$TenantId = $Request.Body.TenantId
$NewUserFirstName = $Request.Body.NewUserFirstName
$NewUserLastName = $Request.Body.NewUserLastName
$NewUserEmail = $Request.Body.NewUserEmail
$JobTitle = $Request.Body.JobTitle
$OfficePhone = $Request.Body.OfficePhone
$MobilePhone = $Request.Body.MobilePhone
$LicenseTypes = $Request.Body.LicenseTypes
$SecurityKey = $env:SecurityKey

Write-Host "Received inputs:"
Write-Host "TicketId: $TicketId"
Write-Host "TenantId: $TenantId"
Write-Host "NewUserFirstName: $NewUserFirstName"
Write-Host "NewUserLastName: $NewUserLastName"
Write-Host "NewUserEmail: $NewUserEmail"
Write-Host "JobTitle: $JobTitle"
Write-Host "OfficePhone: $OfficePhone"
Write-Host "MobilePhone: $MobilePhone"
Write-Host "LicenseTypes: $LicenseTypes"

# Function to generate a random password
function New-RandomPassword {
    param (
        [int]$length = 12
    )

    $characters = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!#%&*()'
    $password = -join ((1..$length) | ForEach-Object { $characters[(Get-Random -Minimum 0 -Maximum $characters.Length)] })
    return $password
}

$password = New-RandomPassword -length 16
Write-Host "Generated Password: $Password"

try {
    if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
        throw "Invalid security key"
    }

    if (-Not $NewUserEmail) {
        throw "NewUserEmail cannot be blank."
    }
    else {
        $NewUserEmail = $NewUserEmail.Trim()
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
    Write-Host "Tenant Id: $TenantId"
    Write-Host "Ticket Id: $TicketId"

    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId
    Write-Host "Connected to Microsoft Graph."

    # Generate the display name
    $NewUserDisplayName = "$NewUserFirstName $NewUserLastName"

    # Extract the mailNickname from the NewUserEmail
    $mailNickname = $NewUserEmail.Split("@")[0]

    Write-Host "Creating new user..."
    # Create the new user
    $newUser = New-MgUser -UserPrincipalName $NewUserEmail -DisplayName $NewUserDisplayName -GivenName $NewUserFirstName -Surname $NewUserLastName -MailNickname $mailNickname -JobTitle $JobTitle -BusinessPhones @($OfficePhone) -MobilePhone $MobilePhone -PasswordProfile @{ Password = $Password; ForceChangePasswordNextSignIn = $true } -AccountEnabled

    if ($newUser) {
        Write-Host "New user created: $($newUser.Id)"
        $message = "New user $NewUserEmail created successfully. `rUsername: $NewUserEmail `rPassword: $Password"
        $UserPrincipalName = $newUser.UserPrincipalName
    } else {
        throw "Failed to create new user."
    }
}
catch {
    $message = "Error: $_"
    $resultCode = 500
    Write-Host "Error: $_"
}

$body = @{
    Message           = $message
    TicketId          = $TicketId
    ResultCode        = $resultCode
    ResultStatus      = if ($resultCode -eq 200) { "Success" } else { "Failure" }
    UserPrincipalName = $UserPrincipalName
    TenantId          = $TenantId
    LicenseTypes      = $LicenseTypes
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::InternalServerError }
        Body        = $body
        ContentType = "application/json"
    })
