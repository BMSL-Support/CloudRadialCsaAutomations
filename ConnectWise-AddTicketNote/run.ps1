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
        "SecurityKey": "optional"
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

    # Create the note serviceObject
    $notePayload = @{
        ticketId = $TicketId
        text = $Text
        detailDescriptionFlag = $true
        internalAnalysisFlag = $Internal
        #resolutionFlag = $false
        #customerUpdatedFlag = $false 
    } | ConvertTo-Json
    
    # Set up the authentication headers
    $headers = @{
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${PublicKey}:${PrivateKey}"))
        "Content-Type" = "application/json"
        "clientId" = $ClientId
    }

    # Make the API request to add the note
    $result = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $notePayload
    Write-Host $result
    return $result
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
    $internal = $false
}

Write-Host "TicketId: $TicketId"
Write-Host "Text: $Text"
Write-Host "Internal: $Internal"

# Prepare the payload and headers
$notePayload = @{
    text = $Text
    internal = $Internal
} | ConvertTo-Json

$headers = @{
    clientId = $env:ConnectWisePsa_ApiClientId
    Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$env:ConnectWisePsa_ApiCompanyId+$env:ConnectWisePsa_ApiPublicKey:$env:ConnectWisePsa_ApiPrivateKey"))
}

# Debugging output
Write-Host "Payload: $notePayload"
Write-Host "Headers: $headers"

# Make the API request
$result = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $notePayload

Write-Host $result.Message

$body = @{
    response = ($result | ConvertTo-Json);
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})
