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

# Check if $Request.Body is already a string or object and use it directly
$requestBody = $Request.Body

# If it's a Stream, read it and convert to a string
if ($requestBody -is [System.IO.Stream]) {
    $requestBody = [System.Text.Encoding]::UTF8.GetString($Request.Body.ReadToEnd())
}

# If $requestBody is already an object, no need to decode
if ($requestBody -is [string]) {
    $data = $requestBody | ConvertFrom-Json
} elseif ($requestBody -is [hashtable]) {
    $data = $requestBody
}

# Extract data from the request
$ticketId = $data.TicketId
$message = $data.Message
$internalNote = $data.Internal

if (-not $ticketId -or -not $message) {
    return @{
        statusCode = [HttpStatusCode]::BadRequest
        body = "Please pass a valid TicketId and Message in the request body"
    }
}

# Prepare the note object
$note = @{
    ticketId = $ticketId
    text = $message
    internalAnalysisFlag = $internalNote
}

# Convert the note object to JSON
$json = $note | ConvertTo-Json

# Set up the API request
#$apiUrl = "${env:ConnectWisePsa_ApiBaseUrl}/v4_6_release/apis/3.0/service/tickets/$ticketId/notes"
$apiUrl = "${env:webhook_test}"
$authHeader = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${env:ConnectWisePsa_ApiCompanyId}+${env:ConnectWisePsa_ApiPublicKey}:${env:ConnectWisePsa_ApiPrivateKey}"))
$headers = @{
    "Authorization" = "Basic $authHeader"
    "Accept" = "application/vnd.connectwise.com+json; version=v2024_1"
}

# Send the API request
$response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $json -ContentType "application/json"

# Check for successful response
if ($response -and $response.StatusCode -eq 200) {
    return @{
        statusCode = [HttpStatusCode]::OK
        body = "Note created successfully"
    }
} else {
    return @{
        statusCode = [HttpStatusCode]::BadRequest
        body = "Failed to create note"
    }
}
