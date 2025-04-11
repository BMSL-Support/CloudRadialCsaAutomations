<# 
.SYNOPSIS
    This function adds Microsoft 365 licenses to an existing user based on the "LicenseTypes" field in the JSON file.

.DESCRIPTION
    This function adds Microsoft 365 licenses to an existing user based on the "LicenseTypes" field in the JSON file.
    The function requires the following environment variables to be set:
    - Ms365_AuthAppId: Application Id of the Azure AD application
    - Ms365_AuthSecretId: Secret Id of the Azure AD application
    - Ms365_TenantId: Tenant Id of the Azure AD application
    - SecurityKey: Optional, used for additional security validation

    The function requires the following modules to be installed:
    - Microsoft.Graph

.INPUTS
    JSON Structure:
    {
        "TenantId": "@CompanyTenantId",
        "TicketId": "@TicketId",
        "AccountDetails": {
            "UserPrincipalName": "@NUsersEmail"
        },
        "LicenseTypes": ["@LicenseType"]
    }

.OUTPUTS
    JSON response with the following fields:
    - Message: Descriptive string of result
    - TicketId: TicketId passed in Parameters
    - ResultCode: 200 for success, 500 for failure
    - ResultStatus: "Success" or "Failure"
#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Add Licenses function triggered."

$resultCode = 200
$message = ""

# Parse the JSON structure
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$AccountDetails = $Request.Body.AccountDetails
$UserPrincipalName = $AccountDetails.UserPrincipalName
$LicenseTypes = $Request.Body.LicenseTypes
$SecurityKey = $env:SecurityKey

# Treat values starting with '@' as null
if ($TenantId -like "@*") { $TenantId = $null }
if ($TicketId -like "@*") { $TicketId = $null }
if ($UserPrincipalName -like "@*") { $UserPrincipalName = $null }
if ($LicenseTypes -is [Array]) {
    $LicenseTypes = $LicenseTypes | Where-Object { $_ -notlike "@*" }
} else {
    $LicenseTypes = @()
}

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break
}

if (-Not $UserPrincipalName) {
    $message = "UserPrincipalName cannot be blank."
    $resultCode = 500
} else {
    $UserPrincipalName = $UserPrincipalName.Trim()
}

if (-Not $TenantId) {
    $TenantId = $env:Ms365_TenantId
} else {
    $TenantId = $TenantId.Trim()
}

if (-Not $TicketId) {
    $TicketId = ""
}

Write-Host "User Principal Name: $UserPrincipalName"
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"
Write-Host "License Types: $($LicenseTypes -join ', ')"

if ($resultCode -eq 200) {
    try {
        $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
        $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

        Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

        if (-not $LicenseTypes) {
            $message = "No licenses specified for assignment."
        } else {
            $user = Get-MgUser -UserId $UserPrincipalName
            $assignedLicenses = $user.AssignedLicenses | ForEach-Object { $_.SkuId }

            $licenses = Get-MgSubscribedSku
            $licenseTypes = Get-LicenseTypes -CsvUri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"

            $licensesToAdd = @()
            $licensesNotAvailable = @()
            $licensesPrettyNamesToAdd = @()
            $licensesPrettyNamesNotAvailable = @()

            foreach ($licenseType in $LicenseTypes) {
                $skuId = $licenseTypes.GetEnumerator() | Where-Object { $_.Value -eq $licenseType } | Select-Object -ExpandProperty Key
                if ($assignedLicenses -contains $skuId) {
                    continue
                } else {
                    $license = $licenses | Where-Object { $_.SkuId -eq $skuId }
                    if ($license) {
                        if ($license.PrepaidUnits.Enabled -gt $license.ConsumedUnits) {
                            $licensesToAdd += $skuId
                            $licensesPrettyNamesToAdd += $licenseType
                        } else {
                            $licensesNotAvailable += $skuId
                            $licensesPrettyNamesNotAvailable += $licenseType
                        }
                    } else {
                        $licensesNotAvailable += $skuId
                        $licensesPrettyNamesNotAvailable += $licenseType
                    }
                }
            }

            if ($licensesToAdd.Count -gt 0) {
                Set-MgUserLicense -UserId $UserPrincipalName -AddLicenses @{ SkuId = $licensesToAdd } -RemoveLicenses @()
                $message = "The following licenses were successfully assigned: $($licensesPrettyNamesToAdd -join ', ')."
            }

            if ($licensesNotAvailable.Count -gt 0) {
                $message += " The following licenses are not available: $($licensesPrettyNamesNotAvailable -join ', ')."
            }
        }
    } catch {
        $message = "An error occurred while assigning licenses: $($_ | Out-String)"
        $resultCode = 500
    }
}

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode  = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::InternalServerError }
    Body        = $body
    ContentType = "application/json"
})

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
