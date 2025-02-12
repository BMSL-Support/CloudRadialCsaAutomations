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
                   
    ConnectWiseManageAPI     

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

# Extract data from the request body
$TicketId = $Request.Body.TicketId
$Text = $Request.Body.Message
$Internal = [bool]$Request.Body.Internal
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

Write-Host "TicketId: $TicketId"
Write-Host "Text: $Text"
Write-Host "Internal: $Internal"

# Create the note serviceObject
$notePayload = @{
    ticketId = $TicketId
    text = $Text
    detailDescriptionFlag = $true
    internalAnalysisFlag = $Internal
}

$result = New-CWMTicketNote -ticketId $TicketId -text $Text -detailDescriptionFlag $true -internalFlag $Internal -resolutionFlag $false

Write-Host $result.Message

$body = @{
    response = ($result | ConvertTo-Json);
} 

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})

# Disconnect from ConnectWise
Disconnect-CWM
