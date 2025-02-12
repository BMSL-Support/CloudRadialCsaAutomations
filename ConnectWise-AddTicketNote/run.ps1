<# 

.SYNOPSIS
    
    This function is used to add a note to a ConnectWise ticket.

.DESCRIPTION
                    
    This function is used to add a note to a ConnectWise ticket.
                    
    The function requires the following environment variables to be set:
                    
    ConnectWisePsa_ApiBaseUrl - Base URL of the ConnectWise API
    ConnectWisePsa_ApiCompanyId - Company Id of the ConnectWise API
    ConnectWisePsa_ApiPublicKey - Public Key of the ConnectWise API
    ConnectWisePsa_ApiPrivateKey - Private Key of the ConnectWise API
    ConnectWisePsa_ApiClientId - Client Id of the ConnectWise API
    SecurityKey - Optional, use this as an additional step to secure the function
                    
    The function requires the following modules to be installed:
                   
    None        

.INPUTS

    TicketId - string value of numeric ticket number
    Message - text of note to add
    Internal - boolean indicating whether not should be internal only
    SecurityKey - optional security key to secure the function

    JSON Structure

    {
        "TicketId": "123456",
        "Message": "This is a note",
        "Internal": true,
        "SecurityKey", "optional"
    }

.OUTPUTS
    
    JSON structure of the response from the ConnectWise API

#>
using namespace System.Net

param($Request, $TriggerMetadata)

function Add-ConnectWiseTicketNote {
    param (
        [string]$ConnectWiseUrl,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId,
        [string]$TicketId,
        [string]$Text,
        [boolean]$Internal = $false
    )

    # Construct the API endpoint for adding a note
    $apiUrl = "$ConnectWiseUrl/v4_6_release/apis/3.0/service/tickets/$TicketId/notes"
    Write-Host "API URL: $apiUrl"

    # Create the note payload
    $notePayload = @{
        ticketId = $TicketId
        text = $Text
        detailDescriptionFlag = $true
        internalAnalysisFlag = $Internal
    } | ConvertTo-Json
    Write-Host "Note Payload: $notePayload"

    # Set up the authentication headers
    $headers = @{
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${PublicKey}:${PrivateKey}"))
        "Content-Type" = "application/json"
        "Accept" = "application/vnd.connectwise.com+json; version=v2024.1"
        "clientId" = $ClientId
    }
    Write-Host "Headers: $($headers | ConvertTo-Json)"

    # Make the API request to add the note
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $notePayload -ErrorAction Stop
        $responseContent = $response | ConvertTo-Json
        Write-Host "API Response: $responseContent"
        return $response
    } catch {
        Write-Host "Error: $($_.Exception.Message)"
        if ($_.Exception.Response -ne $null) {
            $statusCode = $_.Exception.Response.StatusCode
            $responseContent = $_.Exception.Response.Content.ReadAsStringAsync().Result
            Write-Host "StatusCode: $statusCode"
            Write-Host "Response: $responseContent"
        } else {
            Write-Host "No response received from server."
        }
        throw $_.Exception
    }
}

$TicketId = $Request.Body.TicketId
$Text = $Request.Body.Message
$Internal = $Request.Body.Internal
$SecurityKey = $env:SecurityKey

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $TicketId) {
    Write-Host "Missing ticket number"
    break;
}
if (-Not $Text) {
    Write-Host "Missing ticket text"
    break;
}
if (-Not $Internal) {
    $Internal = $false
}

Write-Host "TicketId: $TicketId"
Write-Host "Text: $Text"
Write-Host "Internal: $Internal"

$result = Add-ConnectWiseTicketNote -ConnectWiseUrl $env:ConnectWisePsa_ApiBaseUrl `
    -PublicKey "$env:ConnectWisePsa_ApiCompanyId+$env:ConnectWisePsa_ApiPublicKey" `
    -PrivateKey $env:ConnectWisePsa_ApiPrivateKey `
    -ClientId $env:ConnectWisePsa_ApiClientId `
    -TicketId $TicketId `
    -Text $Text `
    -Internal $Internal

Write-Host "Result: $($result | ConvertTo-Json)"

$body = @{
    response = ($result | ConvertTo-Json);
} 

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})
