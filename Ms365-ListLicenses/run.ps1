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

function Set-CompanyM365Licenses {
    param (
        [string]$Token,
        [string]$AppId,
        [string]$SecretId,
        [int]$CompanyId,
        [string]$LicenseList
    )

    # Construct the basic authentication header
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${AppId}:${SecretId}"))
    $headers = @{
        "Authorization" = "Basic $base64AuthInfo"
        "Content-Type" = "application/json"
    }

    $body = @{
        "companyId" = $CompanyId
        "token" = "$Token"
        "value" = "$LicenseList"
    }

    $bodyJson = $body | ConvertTo-Json

    # Replace the following URL with the actual REST API endpoint
    $apiUrl = "https://api.us.cloudradial.com/api/beta/token"

    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Body $bodyJson -Method Post

    Write-Host "API response: $($response | ConvertTo-Json -Depth 4)"
}

function Get-PrettyLicenseNames {
    param (
        [Parameter (Mandatory=$true)] [String]$CsvUri
    )

    $csvData = Invoke-RestMethod -Method Get -Uri $CsvUri | ConvertFrom-Csv
    $prettyNames = @{}
    foreach ($row in $csvData) {
        $prettyNames[$row.'GUID'] = $row.'Product_Display_Name'
    }

    return $prettyNames
}

$companyId = $Request.Body.CompanyId
$tenantId = $Request.Body.TenantId
$SecurityKey = $env:SecurityKey

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $companyId) {
    $companyId = 1
}

if (-Not $tenantId) {
    $tenantId = $env:Ms365_TenantId
}

$resultCode = 200
$message = ""

$secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
$credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $tenantId

# Get all licenses in the tenant
$licenses = Get-MgSubscribedSku

# Get pretty names for licenses
$prettyNames = Get-PrettyLicenseNames -CsvUri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"

# Extract license product names
$licenseNames = $licenses | ForEach-Object { 
    if ($prettyNames.ContainsKey($_.SkuId)) {
        $prettyNames[$_.SkuId]
    } else {
        Write-Host "Warning: No pretty name found for SKU ID $($_.SkuId)"
        $_.SkuPartNumber
    }
}
$licenseNames = $licenseNames | Sort-Object

# Convert the array of license names to a comma-separated string
$licenseNamesString = $licenseNames -join ","

Set-CompanyM365Licenses -Token "CompanyM365Licenses" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -LicenseList $licenseNamesString

Write-Host "Updated CompanyM365Licenses for Company Id: $companyId."

$message = "Company licenses for $companyId have been updated."

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
