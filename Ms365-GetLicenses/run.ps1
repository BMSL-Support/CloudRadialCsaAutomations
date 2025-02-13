<#
.SYNOPSIS
    This function retrieves Microsoft 365 licenses based on either the license type or the UPN of the user.
.DESCRIPTION
    This function retrieves Microsoft 365 licenses based on either the license type or the UPN of the user.
    The function requires the following environment variables to be set:
    Ms365_AuthAppId - Application Id of the Azure AD application
    Ms365_AuthSecretId - Secret Id of the Azure AD application
    Ms365_TenantId - Tenant Id of the Azure AD application
    The function requires the following modules to be installed:
    Microsoft.Graph
.INPUTS
    UserPrincipalName - string value of the user's principal name (optional)
    LicenseType - string value of the license type (optional)
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    JSON Structure
    {
        "UserPrincipalName": "user@example.com",
        "LicenseType": "Microsoft Power Automate Free",
        "TenantId": "12345678-1234-1234-1234-123456789012"
    }
.OUTPUTS
    JSON response with the following fields:
    Message - Descriptive string of result
    Result - Array of users or licenses
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"
#>

using namespace System.Net

param($Request, $TriggerMetadata)

function Get-Ms365Licenses {
    param (
        [Parameter(Mandatory=$false)][string]$UserPrincipalName,
        [Parameter(Mandatory=$false)][string]$LicenseType,
        [Parameter(Mandatory=$true)][string]$AppId,
        [Parameter(Mandatory=$true)][string]$SecretId,
        [Parameter(Mandatory=$true)][string]$TenantId
    )

    try {
        # Construct the basic authentication header
        $securePassword = ConvertTo-SecureString -String $SecretId -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($AppId, $securePassword)
        Connect-MgGraph -ClientSecretCredential $credential -TenantId $TenantId -NoWelcome

        if ($UserPrincipalName) {
            # Get user licenses
            $user = Get-MgUser -UserId $UserPrincipalName
            $licenses = $user.AssignedLicenses | ForEach-Object { $_.SkuId }
            $result = $licenses
            $message = "Licenses assigned to user $UserPrincipalName: $($licenses -join ', ')."
        } elseif ($LicenseType) {
            # Get all users with the specified license
            $licenseTypes = Get-LicenseTypes -CsvUri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
            $skuId = $licenseTypes.GetEnumerator() | Where-Object { $_.Value -eq $LicenseType } | Select-Object -ExpandProperty Key
            $users = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $skuId)"
            $result = $users | ForEach-Object { $_.UserPrincipalName }
            $message = "Users with license $LicenseType: $($result -join ', ')."
        } else {
            $message = "Either UserPrincipalName or LicenseType must be specified."
            $result = @()
        }

    } catch {
        $message = "An error occurred while retrieving licenses: $($_ | Out-String)"
        $result = @()
    }

    return @{
        Message = $message
        Result = $result
        ResultCode = 200
        ResultStatus = "Success"
    }
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

$UserPrincipalName = $Request.Body.UserPrincipalName
$LicenseType = $Request.Body.LicenseType
$TenantId = $Request.Body.TenantId
$AppId = $env:Ms365_AuthAppId
$SecretId = $env:Ms365_AuthSecretId

$response = Get-Ms365Licenses -UserPrincipalName $UserPrincipalName -LicenseType $LicenseType -AppId $AppId -SecretId $SecretId -TenantId $TenantId

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $response
        ContentType = "application/json"
    })
