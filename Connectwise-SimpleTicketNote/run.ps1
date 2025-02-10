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

# ConnectWise Ticket Note Function

param(
    [Parameter(Mandatory=$true)]
    [string]$TicketId,

    [Parameter(Mandatory=$true)]
    [string]$Message,

    [Parameter(Mandatory=$true)]
    [bool]$Internal,

    [string]$SecurityKey
)

# Check for security key (if required)
$EnvSecurityKey = $env:SecurityKey
if ($SecurityKey -and $SecurityKey -ne $EnvSecurityKey) {
    return @{ message = "Invalid Security Key." } | ConvertTo-Json
}

# Read the environment variables
$ApiBaseUrl = $env:ConnectWisePsa_ApiBaseUrl
$ApiCompanyId = $env:ConnectWisePsa_ApiCompanyId
$ApiPublicKey = $env:ConnectWisePsa_ApiPublicKey
$ApiPrivateKey = $env:ConnectWisePsa_ApiPrivateKey
$ApiClientId = $env:ConnectWisePsa_ApiClientId

# Check if all required environment variables are set
if (-not $ApiBaseUrl -or -not $ApiCompanyId -or -not $ApiPublicKey -or -not $ApiPrivateKey -or -not $ApiClientId) {
    return @{ message = "Missing required environment variables." } | ConvertTo-Json
}

# Initialize ConnectWise API Client
$modulePath = "C:\path\to\ConnectWiseManageAPI\ConnectWiseManageAPI.psm1" # Adjust this path
Import-Module $modulePath

# Build the ticket note object
$ticketNote = @{
    "internalNote" = $Internal
    "text"         = $Message
    "ticket"       = @{
        "id" = $TicketId
    }
}

# Use the ConnectWise API to add the note to the ticket
try {
    $response = Add-CWTicketNote -CompanyId $ApiCompanyId `
                                  -ApiPublicKey $ApiPublicKey `
                                  -ApiPrivateKey $ApiPrivateKey `
                                  -ApiClientId $ApiClientId `
                                  -ApiBaseUrl $ApiBaseUrl `
                                  -TicketNote $ticketNote

    # Return the response from the API
    return $response | ConvertTo-Json
}
catch {
    return @{ message = "Error occurred: $_" } | ConvertTo-Json
}
