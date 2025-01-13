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

function Set-CloudRadialToken {
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

# Fetch licenses for the tenant
$licenses = Get-MgSubscribedSku

# Debugging: Output the fetched licenses to see what we have
Write-Host "Fetched Licenses:"
$licenses | ForEach-Object { Write-Host "$($_.SkuPartNumber) - $($_.SkuId)" }

# Path to the CSV containing Service Plan identifiers and friendly names
$csvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"

# Download the CSV content
$csvContent = Invoke-WebRequest -Uri $csvUrl -UseBasicPipelines

# Convert the CSV content into an object
$servicePlans = $csvContent.Content | ConvertFrom-Csv

# Debugging: Output the CSV data to check what is loaded
Write-Host "Service Plans CSV Data:"
$servicePlans | ForEach-Object { Write-Host "$($_.ServicePlanId) - $($_.Service_Plans_Included_Friendly_Name)" }

# Initialize an array to store the license names
$licenseNames = @()

foreach ($license in $licenses) {
    $skuPartNumber = $license.SkuPartNumber
    # Debugging: Output the SKU part number for each license
    Write-Host "Processing License: $skuPartNumber"

    $servicePlan = $servicePlans | Where-Object { $_.ServicePlanId -eq $skuPartNumber }
    
    if ($servicePlan) {
        Write-Host "Found Matching Service Plan: $($servicePlan.Service_Plans_Included_Friendly_Name)"
        $licenseNames += $servicePlan.Service_Plans_Included_Friendly_Name
    } else {
        Write-Host "No matching service plan found for SKU: $skuPartNumber"
    }
}

# Check if we found any license names
if ($licenseNames.Count -eq 0) {
    Write-Host "No licenses were found or matched."
} else {
    Write-Host "Found the following license names: $($licenseNames -join ', ')"
}

# Join all license names into a comma-separated string
$licenseNamesString = $licenseNames -join ","

# Send the list of licenses to CloudRadial
Set-CloudRadialToken -Token "CompanyM365Licenses" -AppId $env:CloudRadialCsa_ApiPublicKey -SecretId $env:CloudRadialCsa_ApiPrivateKey -CompanyId $companyId -LicenseList $licenseNamesString

Write-Host "Updated CompanyM365Licenses for Company Id: $companyId."

$message = "Company tokens for $companyId have been updated with the M365 license types."

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
} 

# Associate values to output bindings by calling 'Push-OutputBinding' 
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
