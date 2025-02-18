<#
.SYNOPSIS
    This function collects products from ConnectWise and updates tokens in CloudRadial.

.DESCRIPTION
    This function fetches product data from ConnectWise using specified product identifiers, extracts the description and price, converts them to basic HTML, and saves them as strings in company-wide tokens in CloudRadial. The function requires the following environment variables to be set:
    
    - ConnectWise API Token
    - CloudRadial API Public Key
    - CloudRadial API Private Key

    The function updates tokens for specified products.

.INPUTS
    CompanyId - numeric company id (optional, defaults to 1)
    Identifiers - array of product identifiers

.OUTPUTS
    JSON response with the following fields:
    - Message: Descriptive string of result
    - ResultCode: 200 for success, 500 for failure
    - ResultStatus: "Success" or "Failure"

.EXPECTED INPUT
    {
        "CompanyId": 12,
        "Identifiers": ["TypicalLaptop", "PerformanceLaptop"]
    }

.EXPECTED OUTPUT
    {
        "Message": "Company tokens for 12 have been updated.",
        "ResultCode": 200,
        "ResultStatus": "Success"
    }
#>

using namespace System.Net

param($Request, $TriggerMetadata)

function Get-ConnectWiseProduct {
    param (
        [string]$Identifier
    )

    Import-Module "C:\home\site\wwwroot\Modules\ConnectWiseManageAPI\ConnectWiseManageAPI.psm1"
    
    # Create the CWConnection
    $Connection = @{
        Server = $env:ConnectWisePsa_ApiBaseUrl
        Company = $env:ConnectWisePsa_ApiCompanyId
        PubKey = $env:ConnectWisePsa_ApiPublicKey
        PrivateKey = $env:ConnectWisePsa_ApiPrivateKey
        ClientID = $env:ConnectWisePsa_ApiClientId
    }
    Connect-CWM @Connection

    # Fetch product data
    $product = Get-CWMProduct -Where "identifier='$Identifier'"
    return $product
}

function Set-CloudRadialToken {
    param (
        [string]$Token,
        [string]$AppId,
        [string]$SecretId,
        [int]$CompanyId,
        [string]$Value
    )

    # Construct the basic authentication header
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${AppId}:${SecretId}"))
    $headers = @{
        "Authorization" = "Basic $base64AuthInfo"
        "Content-Type"  = "application/json"
    }

    $body = @{
        "companyId" = $CompanyId
        "token"     = $Token
        "value"     = $Value
    }

    $bodyJson = $body | ConvertTo-Json

    # Replace with actual CloudRadial API endpoint
    $apiUrl = "https://api.us.cloudradial.com/api/beta/token"

    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Body $bodyJson -Method Post

    Write-Host "API response: $($response | ConvertTo-Json -Depth 4)"
}

function Update-CloudRadialTokens {
    param (
        [int]$CompanyId,
        [string[]]$Identifiers
    )

    foreach ($identifier in $Identifiers) {
        $productData = Get-ConnectWiseProduct -Identifier $identifier
        $description = $productData.description
        $price = $productData.price

        $htmlValue = "<p>Description: $description</p><p>Price: $price</p>"

        Set-CloudRadialToken -Token "@$identifier" -AppId $env:CloudRadialCsa_ApiPublicKey -SecretId $env:CloudRadialCsa_ApiPrivateKey -CompanyId $CompanyId -Value $htmlValue
    }

    Write-Host "Updated tokens for Company Id: $CompanyId."
}

$companyId = $Request.Body.CompanyId
$identifiers = $Request.Body.Identifiers

if (-Not $companyId) {
    $companyId = 1
}

if (-Not $identifiers) {
    Write-Host "No product identifiers specified."
    break;
}

Update-CloudRadialTokens -CompanyId $companyId -Identifiers $identifiers

$body = @{
    Message      = "Company tokens for $companyId have been updated."
    ResultCode   = 200
    ResultStatus = "Success"
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode  = [HttpStatusCode]::OK
    Body        = $body
    ContentType = "application/json"
})
