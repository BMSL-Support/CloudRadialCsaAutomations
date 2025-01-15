<# 

.SYNOPSIS

    This function adds Microsoft 365 licenses to an existing user based on the "RequestedLicense".

.DESCRIPTION

    This function adds Microsoft 365 licenses to an existing user based on the "RequestedLicense".
    
    The function requires the following environment variables to be set:
    
    Ms365_AuthAppId - Application Id of the Azure AD application
    Ms365_AuthSecretId - Secret Id of the Azure AD application
    Ms365_TenantId - Tenant Id of the Azure AD application
    SecurityKey - Optional, use this as an additional step to secure the function

    The function requires the following modules to be installed:
    
    Microsoft.Graph     

.INPUTS

    UserPrincipalName - string value of the user's principal name
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    RequestedLicense - array of license types to be assigned to the user
    TicketId - string value of the ticket id

    JSON Structure

    {
        "UserPrincipalName": "user@example.com",
        "TenantId": "12345678-1234-1234-1234-123456789012",
        "RequestedLicense": [
            "Office 365 E3",
            "Microsoft 365 Business Standard"
        ],
        "TicketId": "TICKET12345"
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

function Add-UserLicenses {
    param (
        [string]$UserPrincipalName,
        [string]$AppId,
        [string]$SecretId,
        [string]$TenantId,
        [array]$RequestedLicense,
        [string]$TicketId
    )

    # Check for required inputs
    if (-not $UserPrincipalName) {
        Write-Host "ERROR: UserPrincipalName is required."
        return
    }
    if (-not $TenantId) {
        Write-Host "ERROR: TenantId is required."
        return
    }
    if (-not $RequestedLicense) {
        Write-Host "ERROR: RequestedLicense is required."
        return
    }
    if (-not $TicketId) {
        Write-Host "ERROR: TicketId is required."
        return
    }

    # Construct the basic authentication header
    $securePassword = ConvertTo-SecureString -String $SecretId -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($AppId, $securePassword)
    Connect-MgGraph -ClientSecretCredential $credential -TenantId $TenantId -NoWelcome

    # Get all licenses in the tenant
    $licenses = Get-MgSubscribedSku

    # Output available licenses
    Write-Host "Available licenses in the tenant:"
    foreach ($license in $licenses) {
        Write-Host "SKU ID: $($license.SkuId), SKU Part Number: $($license.SkuPartNumber), Enabled: $($license.PrepaidUnits.Enabled), Consumed: $($license.ConsumedUnits)"
    }

    # Get license types
    $licenseTypes = Get-LicenseTypes -CsvUri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"

    $licensesToAdd = @()
    $licensesNotAvailable = @()
    Write-Host "RequestedLicense input: $($RequestedLicense -join ', ')"
    foreach ($licenseType in $RequestedLicense) {
        Write-Host "Processing license type: $licenseType"
        $skuId = $licenseTypes.GetEnumerator() | Where-Object { $_.Value -eq $licenseType } | Select-Object -ExpandProperty Key
        Write-Host "Checking license type: $licenseType, SKU ID: $skuId"
        $license = $licenses | Where-Object { $_.SkuId -eq $skuId }
        if ($license) {
            Write-Host "License found: $($license.SkuPartNumber), Enabled: $($license.PrepaidUnits.Enabled), Consumed: $($license.ConsumedUnits)"
            if ($license.PrepaidUnits.Enabled -gt $license.ConsumedUnits) {
                $licensesToAdd += $skuId
            } else {
                $licensesNotAvailable += $licenseType
            }
        } else {
            Write-Host "License not found for SKU ID: $skuId"
            $licensesNotAvailable += $licenseType
        }
    }

    try {
        if ($licensesToAdd.Count -gt 0) {
            $user = Get-MgUser -UserPrincipalName $UserPrincipalName
            foreach ($skuId in $licensesToAdd) {
                Write-Host "Adding license SKU ID: $skuId to user: $UserPrincipalName"
                Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = $skuId}
            }
            $message = "Added licenses: $($licensesToAdd -join ', ')."
        } else {
            $message = "No licenses available to add."
        }

        if ($licensesNotAvailable.Count -gt 0) {
            $message += " Licenses not available: $($licensesNotAvailable -join ', ')."
        }
    } catch {
        Write-Host "ERROR: $_"
        $message = "An error occurred while adding licenses: $_"
    }

    return $message
}

function Get-LicenseTypes {
    param (
        [Parameter (Mandatory=$true)] [String]$CsvUri
    )

    $csvData = Invoke-RestMethod -Method Get -Uri $CsvUri | ConvertFrom-Csv
    $licenseTypes = @{ }
    foreach ($row in $csvData) {
        $licenseTypes[$row.'GUID'] = $row.'Product_Display_Name'
    }

    return $licenseTypes
}

# Debugging: Output the entire body to check structure
Write-Host "Request Body: $($Request.Body | ConvertTo-Json -Depth 10)"

# Ensure the UserPrincipalName exists in the body
if ($Request.Body.UserPrincipalName) {
    $UserPrincipalName = $Request.Body.UserPrincipalName
} else {
    Write-Host "ERROR: UserPrincipalName is missing from the request body."
    exit
}

$TenantId = $Request.Body.TenantId
$RequestedLicense = $Request.Body.RequestedLicense
$TicketId = $Request.Body.TicketId
$AppId = $env:Ms365_AuthAppId
$SecretId = $env:Ms365_AuthSecretId

Write-Host "UserPrincipalName: $UserPrincipalName"
Write-Host "TenantId: $TenantId"
Write-Host "RequestedLicense: $($RequestedLicense -join ', ')"
Write-Host "TicketId: $TicketId"

# Debug the structure of the input JSON
Write-Host "Request Body: $($Request.Body | ConvertTo-Json -Depth 10)"

# Run the function to add licenses
$message = Add-UserLicenses -UserPrincipalName $UserPrincipalName -AppId $AppId -SecretId $SecretId -TenantId $TenantId -RequestedLicense $RequestedLicense -TicketId $TicketId

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = 200
    ResultStatus = "Success"
}

# Associate values to output bindings by calling 'Push-OutputBinding' to send the response
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
