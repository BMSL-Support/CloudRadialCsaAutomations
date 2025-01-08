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
    }
    $Body = ConvertTo-Json $notePayload
    
    # Set up the authentication headers
    $AuthString  = "$($CientId)+$($PublicKey):$($PrivateKey)"
    $EncodedAuth  = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($AuthString));
    $headers = @{
        Authorization = "Basic $EncodedAuth"
        ClientID = $ClientID
        'Cache-Control'= 'no-cache'
        ConnectionMethod = 'Key'
        Accept = "application/vnd.connectwise.com+json; version=v2020_2"
    }

    # Make the API request to add the note
    $result = Invoke-WebRequest -Uri $apiUrl -Method 'Post' -Headers $headers -Body $Body
    Write-Host $result
    return $result.content

}
