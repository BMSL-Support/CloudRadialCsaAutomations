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
    RequestedLicense - Array of license types
#>

using namespace System.Net

param($Request, $TriggerMetadata)

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

    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

    # Generate the display name
    $NewUserDisplayName = "$NewUserFirstName $NewUserLastName"

    # Extract the mailNickname from the NewUserEmail
    $mailNickname = $NewUserEmail.Split("@")[0]

    # Check and ignore fields if not specified or contain certain values
    if ($JobTitle -eq "@NUsersJobTitle" -or -not $JobTitle) {
        $JobTitle = $null
    }
    if ($OfficePhone -eq "@NUsersOfficePhone" -or -not $OfficePhone) {
        $OfficePhone = $null
    }
    if ($MobilePhone -eq "@NUsersMobilePhone" -or -not $MobilePhone) {
        $MobilePhone = $null
    }

    # Create the new user
    $newUserParams = @{
        UserPrincipalName = $NewUserEmail
        DisplayName       = $NewUserDisplayName
        GivenName         = $NewUserFirstName
        Surname           = $NewUserLastName
        MailNickname      = $mailNickname
        PasswordProfile   = @{ Password = $Password; ForceChangePasswordNextSignIn = $true }
        UsageLocation     = "GB"
        AccountEnabled    = $true
    }

    if ($JobTitle) {
        $newUserParams.JobTitle = $JobTitle
    }
    if ($OfficePhone) {
        $newUserParams.BusinessPhones = $OfficePhone
    }
    if ($MobilePhone) {
        $newUserParams.MobilePhone = $MobilePhone
    }

    $newUser = New-MgUser @newUserParams

    if ($newUser) {
        $message = "New user $NewUserDisplayName created successfully.`r `rUsername: $NewUserEmail `rPassword: $Password"
        $UserPrincipalName = $newUser.UserPrincipalName
    } else {
        throw "Failed to create new user."
    }
}
catch {
    $message = "Error: $_"
    $resultCode = 500
}

$body = @{
    Message           = $message
    TicketId          = $TicketId
    ResultCode        = $resultCode
    ResultStatus      = if ($resultCode -eq 200) { "Success" } else { "Failure" }
    UserPrincipalName = $UserPrincipalName
    TenantId          = $TenantId
    RequestedLicense  = $LicenseTypes
    Internal          = $true
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::InternalServerError }
        Body        = $body
        ContentType = "application/json"
    })
