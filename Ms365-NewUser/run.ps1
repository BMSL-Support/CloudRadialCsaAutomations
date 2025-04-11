<# 
.SYNOPSIS
    This script creates a new user in the Microsoft 365 tenant using details provided in a JSON payload.

.DESCRIPTION
    This script automates the creation of a new user in the Microsoft 365 tenant. 
    It reads user details from a JSON payload, including the tenant ID, ticket ID, account details, 
    and license types. The script validates the input, generates a random password for the new user, 
    and creates the user in the tenant using the Microsoft Graph API.

    The script requires the following environment variables to be set:
    - Ms365_AuthAppId: Application ID of the service principal
    - Ms365_AuthSecretId: Secret ID of the service principal
    - Ms365_TenantId: Tenant ID of the Microsoft 365 tenant
    - SecurityKey: Optional, used for additional security validation

    The script requires the Microsoft.Graph module to be installed.

.INPUTS
    JSON Structure:
    {
        "TenantId": "@CompanyTenantId",
        "TicketId": "@TicketId",
        "AccountDetails": {
            "GivenName": "@NUsersFirstName",
            "Surname": "@NUsersLastName",
            "UserPrincipalName": "@NUsersEmail",
            "AdditionalAccountDetails": {
                "JobTitle": "@NUsersJobTitle",
                "City": "@NUsersAddJobTitle",
                "Department": "@NUsersDept",
                "BusinessPhones": "@NUsersOfficePhone",
                "MobilePhone": "@NUsersMobilePhone"
            }
        },
        "LicenseTypes": ["@LicenseType"]
    }

.OUTPUTS
    JSON response with the following fields:
    - Message: Descriptive string of the result
    - TicketId: Ticket ID passed in the parameters
    - ResultCode: 200 for success, 500 for failure
    - ResultStatus: "Success" or "Failure"
    - UserPrincipalName: UPN of the new user created
    - TenantId: Tenant ID of the Microsoft 365 tenant
#>

using namespace System.Net

param($Request, $TriggerMetadata)

$resultCode = 200
$message = ""
$UserPrincipalName = ""

# Parse the JSON structure
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$AccountDetails = $Request.Body.AccountDetails
$GivenName = $AccountDetails.GivenName
$Surname = $AccountDetails.Surname
$UserPrincipalName = $AccountDetails.UserPrincipalName
$AdditionalAccountDetails = $AccountDetails.AdditionalAccountDetails
$JobTitle = $AdditionalAccountDetails.JobTitle
$City = $AdditionalAccountDetails.City
$Department = $AdditionalAccountDetails.Department
$BusinessPhones = $AdditionalAccountDetails.BusinessPhones
$MobilePhone = $AdditionalAccountDetails.MobilePhone
$LicenseTypes = $Request.Body.LicenseTypes
$SecurityKey = $env:SecurityKey

# Treat values starting with '@' as null
if ($TenantId -like "@*") { $TenantId = $null }
if ($TicketId -like "@*") { $TicketId = $null }
if ($GivenName -like "@*") { $GivenName = $null }
if ($Surname -like "@*") { $Surname = $null }
if ($UserPrincipalName -like "@*") { $UserPrincipalName = $null }
if ($JobTitle -like "@*") { $JobTitle = $null }
if ($City -like "@*") { $City = $null }
if ($Department -like "@*") { $Department = $null }
if ($BusinessPhones -like "@*") { $BusinessPhones = $null }
if ($MobilePhone -like "@*") { $MobilePhone = $null }
if ($LicenseTypes -is [Array]) {
    $LicenseTypes = $LicenseTypes | Where-Object { $_ -notlike "@*" }
}

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

    if (-Not $UserPrincipalName) {
        throw "UserPrincipalName cannot be blank."
    }
    else {
        $UserPrincipalName = $UserPrincipalName.Trim()
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
    $DisplayName = "$GivenName $Surname"

    # Extract the mailNickname from the UserPrincipalName
    $mailNickname = $UserPrincipalName.Split("@")[0]

    # Create the new user
    $newUserParams = @{
        UserPrincipalName = $UserPrincipalName
        DisplayName       = $DisplayName
        GivenName         = $GivenName
        Surname           = $Surname
        MailNickname      = $mailNickname
        PasswordProfile   = @{ Password = $Password; ForceChangePasswordNextSignIn = $true }
        UsageLocation     = "GB"
        AccountEnabled    = $true
    }

    if ($JobTitle) {
        $newUserParams.JobTitle = $JobTitle
    }
    if ($City) {
        $newUserParams.City = $City
    }
    if ($Department) {
        $newUserParams.Department = $Department
    }
    if ($BusinessPhones) {
        $newUserParams.BusinessPhones = $BusinessPhones
    }
    if ($MobilePhone) {
        $newUserParams.MobilePhone = $MobilePhone
    }

    $newUser = New-MgUser @newUserParams

    if ($newUser) {
        $message = "New user $DisplayName created successfully.`r `rUsername: $UserPrincipalName `rPassword: $Password"
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
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::InternalServerError }
        Body        = $body
        ContentType = "application/json"
    })
