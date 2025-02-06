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

param(
    [Parameter(Mandatory=$true)]
    [string]$TicketId,

    [Parameter(Mandatory=$true)]
    [string]$Message,

    [Parameter(Mandatory=$true)]
    [bool]$Internal,

    [string]$SecurityKey
)

# Get environment variables
$ApiBaseUrl = $env:ConnectWisePsa_ApiBaseUrl
$CompanyId = $env:ConnectWisePsa_ApiCompanyId
$PublicKey = $env:ConnectWisePsa_ApiPublicKey
$PrivateKey = $env:ConnectWisePsa_ApiPrivateKey
$ClientId = $env:ConnectWisePsa_ApiClientId
$SecretKey = $env:SecurityKey

# Validate if the SecurityKey matches, if provided
if ($SecurityKey -and $SecurityKey -ne $SecretKey) {
    Write-Host "Invalid SecurityKey"
    exit
}

# Create the note payload
$note = @{
    "ticketId" = $TicketId
    "note"     = $Message
    "internal" = $Internal
}

# Convert the note to JSON
$noteJson = $note | ConvertTo-Json

# Create the API endpoint
$endpoint = "$ApiBaseUrl/v4_6_release/apis/3.0/$CompanyId/service/tickets/$TicketId/notes"

# Define the headers for authentication
$authorizationValue = "Bearer " + $PublicKey + ":" + $PrivateKey
$headers = @{
    "Authorization" = $authorizationValue
    "Content-Type"  = "application/json"
    "clientid"      = $ClientId
}

# Make the API call
$response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers -Body $noteJson

# Return the response as JSON
$response | ConvertTo-Json
