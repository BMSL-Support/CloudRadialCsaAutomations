<# 
.SYNOPSIS
    This function creates a new user in the tenant.
.DESCRIPTION
    This function creates a new user in the tenant and assigns a license if available.
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
        "NewUserDisplayName": "John Doe",
        "NewUserEmail": "john.doe@example.com",
        "LicenseType": "ENTERPRISEPACK"
    }
.OUTPUTS
    JSON response with the following fields:
    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success with license, 201 for success without license, 500 for failure
    ResultStatus - "Success" or "Failure"
#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Create New User function triggered."

$resultCode = 201
$message = ""

$TicketId = $Request.Body.TicketId
$TenantId = $Request.Body.TenantId
$NewUserFirstName = $Request.Body.NewUserFirstName
$NewUserLastName = $Request.Body.NewUserLastName
$NewUserDisplayName = $Request.Body.NewUserDisplayName
$NewUserEmail = $Request.Body.NewUserEmail
$LicenseType = $Request.Body.LicenseType
$SecurityKey = $env:SecurityKey

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

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId

    # Create the new user
    $newUser = New-MgUser -UserPrincipalName $NewUserEmail -DisplayName $NewUserDisplayName -GivenName $NewUserFirstName -Surname $NewUserLastName

    if ($LicenseType) {
        # Retrieve all available licenses
        $availableLicenses = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $LicenseType }

        if ($availableLicenses) {
            # Assign the license to the new user
            Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $availableLicenses.SkuId }
            $message = "New user `$NewUserEmail` created successfully with license."
            $resultCode = 200
        }
        else {
            $message = "License type `$LicenseType` not available. New user `$NewUserEmail` created without license."
        }
    }
    else {
        $message = "No license type specified. New user `$NewUserEmail` created without license."
    }
}
catch {
    $message = "Error: $_"
    $resultCode = 500
}

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200 -or $resultCode -eq 201) { "Success" } else { "Failure" }
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::Created }
        Body        = $body
        ContentType = "application/json"
    })
