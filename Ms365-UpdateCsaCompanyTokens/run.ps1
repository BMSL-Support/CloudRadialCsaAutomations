<#

.SYNOPSIS

    This function is used to update the company tokens in CloudRadial from a Microsoft 365 tenant.

.DESCRIPTION
    
    This function is used to update the company tokens in CloudRadial from a Microsoft 365 tenant.
    
    The function requires the following environment variables to be set:
    
    Ms365_AuthAppId - Application Id of the Azure AD application
    Ms365_AuthSecretId - Secret Id of the Azure AD application
    Ms365_TenantId - Tenant Id of the Azure AD application
    CloudRadialCsa_ApiPublicKey - Public Key of the CloudRadial API
    CloudRadialCsa_ApiPrivateKey - Private Key of the CloudRadial API
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
        [string]$GroupList
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
        "value" = "$GroupList"
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

# Output the list of M365 Licenses types
$CompanyM365Licenses = Get-MsolAccountSku

# Output the list of M365 Security groups
$CompanyM365SecGroups = Get-MsolGroup -All | Where-Object {$_.GroupType -eq 'Security'}

# Output the list of M365 Team names
$CompanyM365Teams = Get-Team

# Output the list of Exchange Distribution Groups
$CompanyM365EOLDG = Get-DistributionGroup

# Output the list of Exchange Shared Mailboxes
$CompanyM365EOLSMB = Get-Mailbox -RecipientTypeDetails SharedMailbox

# Output the list of Exchange Shared Calendars
$CompanyM365EOLSC = Get-MailboxFolderStatistics -FolderScope Calendar

Set-CloudRadialToken -Token "CompanyM365Licenses" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -GroupList ($CompanyM365Licenses | ConvertTo-Json)
Set-CloudRadialToken -Token "CompanyM365SecGroups" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -GroupList ($CompanyM365SecGroups | ConvertTo-Json)
Set-CloudRadialToken -Token "CompanyM365Teams" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -GroupList ($CompanyM365Teams | ConvertTo-Json)
Set-CloudRadialToken -Token "CompanyM365EOLDG" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -GroupList ($CompanyM365EOLDG | ConvertTo-Json)
Set-CloudRadialToken -Token "CompanyM365EOLSMB" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -GroupList ($CompanyM365EOLSMB | ConvertTo-Json)
Set-CloudRadialToken -Token "CompanyM365EOLSC" -AppId ${env:CloudRadialCsa_ApiPublicKey} -SecretId ${env:CloudRadialCsa_ApiPrivateKey} -CompanyId $companyId -GroupList ($CompanyM365EOLSC | ConvertTo-Json)

$message = "Company tokens for $companyId have been updated."

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
