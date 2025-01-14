<# 

.SYNOPSIS

    This function lists all available Microsoft 365 licenses.

.DESCRIPTION

    This function lists all available Microsoft 365 licenses.
    
    The function requires the following environment variables to be set:
    
    Ms365_AuthAppId - Application Id of the Azure AD application
    Ms365_AuthSecretId - Secret Id of the Azure AD application
    Ms365_TenantId - Tenant Id of the Azure AD application
    SecurityKey - Optional, use this as an additional step to secure the function

    The function requires the following modules to be installed:
    
    Microsoft.Graph     

.INPUTS

    CompanyId - numeric company id
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    SecurityKey - Optional, use this as an additional step to secure the function

    JSON Structure

    {
        "CompanyId": "12"
        "TenantId": "12345678-1234-1234-1234-123456789012",
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

Write-Host "Create New User function triggered."

$resultCode = 201
$message = ""

# Log the raw request body for debugging
Write-Host "Raw Request Body: $($Request.Body | ConvertTo-Json -Depth 10)"

$TicketId = $Request.Body.TicketId
$TenantId = $Request.Body.TenantId
$NewUserFirstName = $Request.Body.NewUserFirstName
$NewUserLastName = $Request.Body.NewUserLastName
$NewUserEmail = $Request.Body.NewUserEmail
$LicenseTypes = $Request.Body.LicenseType -split ","
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
Write-Host "LicenseTypes: $LicenseTypes"
Write-Host "JobTitle: $JobTitle"
Write-Host "OfficePhone: $OfficePhone"
Write-Host "MobilePhone: $MobilePhone"

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

# Download the CSV once and store it in a variable
$csvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$csvData = Invoke-WebRequest -Uri $csvUrl -UseBasicParsing | ConvertFrom-Csv

# Function to convert pretty license names to SKU IDs
function Get-SkuId {
    param (
        [string]$licenseName,
        [array]$csvData
    )

    $skuId = ($csvData | Where-Object { $_.'Product_Display_Name' -eq $licenseName }).'GUID'
    Write-Host "License Name: $licenseName, SKU ID: $skuId"
    return $skuId
}

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
    } else {
        throw "Failed to create new user."
    }

    $licensesAssigned = @()
    foreach ($licenseType in $LicenseTypes) {
        $licenseType = $licenseType.Trim()
        $skuId = Get-SkuId -licenseName $licenseType -csvData $csvData

        if ($skuId) {
            Write-Host "Assigning license $licenseType to new user..."
            Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $skuId }
            $licensesAssigned += $licenseType
        }
        else {
            Write-Host "The license type $licenseType was not available."
        }
    }

    if ($licensesAssigned.Count -gt 0) {
        $message = "New user $NewUserEmail created successfully with licenses: $($licensesAssigned -join ', '). `rUsername: $NewUserEmail `rPassword: $Password"
        $resultCode = 200
    }
    else {
        $message = "No valid licenses were assigned. New user $NewUserEmail created without license. `rUsername: $NewUserEmail `rPassword: $Password"
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
