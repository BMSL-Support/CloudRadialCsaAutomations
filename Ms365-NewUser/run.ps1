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
        "NewUserEmail": "john.doe@example.com",
        "LicenseType": "ENTERPRISEPACK",
        "JobTitle": "Software Engineer",
        "OfficePhone": "+1234567890",
        "MobilePhone": "+0987654321"
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
$NewUserEmail = $Request.Body.NewUserEmail
$LicenseType = $Request.Body.LicenseType
$JobTitle = $Request.Body.JobTitle
$OfficePhone = $Request.Body.OfficePhone
$MobilePhone = $Request.Body.MobilePhone
$SecurityKey = $env:SecurityKey

Write-Host "Received inputs:"
Write-Host "TicketId: $TicketId"
Write-Host "TenantId: $TenantId"
Write-Host "NewUserFirstName: $NewUserFirstName"
Write-Host "NewUserLastName: $NewUserLastName"
Write-Host "NewUserEmail: $NewUserEmail"
Write-Host "LicenseType: $LicenseType"
Write-Host "JobTitle: $JobTitle"
Write-Host "OfficePhone: $OfficePhone"
Write-Host "MobilePhone: $MobilePhone"

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

    Write-Host "Creating new user..."
    # Create the new user
    $newUser = New-MgUser -UserPrincipalName $NewUserEmail -DisplayName $NewUserDisplayName -GivenName $NewUserFirstName -Surname $NewUserLastName -JobTitle $JobTitle -BusinessPhones @($OfficePhone) -MobilePhone $MobilePhone

    if ($newUser) {
        Write-Host "New user created: $($newUser.Id)"
    } else {
        throw "Failed to create new user."
    }

    if ($LicenseType) {
        Write-Host "Retrieving available licenses..."
        # Retrieve all available licenses
        $availableLicenses = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $LicenseType }

        if ($availableLicenses) {
            Write-Host "Assigning license to new user..."
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
    Write-Host "Error: $_"
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
