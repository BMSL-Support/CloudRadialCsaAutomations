<# 
.SYNOPSIS
    This function is used to add a note to a ConnectWise ticket.
.DESCRIPTION
    This function adds a note to a ConnectWise ticket using the ConnectWise Manage API.
    The function requires the following environment variables to be set:
    ConnectWisePsa_ApiBaseUrl - Base URL of the ConnectWise API
    ConnectWisePsa_ApiCompanyId - Company Id of the ConnectWise API
    ConnectWisePsa_ApiPublicKey - Public Key of the ConnectWise API
    ConnectWisePsa_ApiPrivateKey - Private Key of the ConnectWise API
    ConnectWisePsa_ApiClientId - Client Id of the ConnectWise API
    SecurityKey - Optional, use this as an additional step to secure the function
.INPUTS
    TicketId - string value of numeric ticket number
    Message - text of note to add
    Internal - boolean indicating whether the note should be internal only
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

# Get environment variables
$ApiBaseUrl = ${env:ConnectWisePsa_ApiBaseUrl}
$CompanyId = ${env:ConnectWisePsa_ApiCompanyId}
$PublicKey = ${env:ConnectWisePsa_ApiPublicKey}
$PrivateKey = ${env:ConnectWisePsa_ApiPrivateKey}
$ClientId = ${env:ConnectWisePsa_ApiClientId}
$SecretKey = ${env:SecurityKey}

# Ensure the HTTP request body is correctly parsed
$body = $Request.Body | ConvertFrom-Json

$TicketId = $body.TicketId
$Message = $body.Message
$Internal = $body.Internal
$SecurityKey = $body.SecurityKey

# Check if the SecurityKey is valid (optional)
if ($SecurityKey -and $SecurityKey -ne $SecretKey) {
    $Response = @{
        status = "error"
        message = "Invalid SecurityKey"
    }
    $Response | ConvertTo-Json
    exit
}

# Prepare the note payload
$note = @{
    "ticketId" = $TicketId
    "note"     = $Message
    "internal" = $Internal
}

# Convert the note to JSON
$noteJson = $note | ConvertTo-Json

# Create the API endpoint
$endpoint = "$ApiBaseUrl/v4_6_release/apis/3.0/$CompanyId/service/tickets/$TicketId/notes"

# Set up the authorization headers
$authorizationValue = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${PublicKey}:${PrivateKey}"))
$headers = @{
    "Authorization" = $authorizationValue
    "Content-Type"  = "application/json"
    "clientid"      = $ClientId
}

# Make the API call to ConnectWise
$response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers -Body $noteJson

# Return the API response as JSON
$Response = @{
    status  = "success"
    message = "Note added successfully"
    data    = $response
}

$Response | ConvertTo-Json
