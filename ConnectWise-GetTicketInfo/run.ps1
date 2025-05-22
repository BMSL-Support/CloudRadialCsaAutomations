<# 
.SYNOPSIS
    Retrieves ConnectWise tickets with notes and resolution details.

.DESCRIPTION
    This function queries ConnectWise for tickets and includes their notes and resolution info.

.INPUTS
    JSON Structure:
    {
        "SummaryContains": "printer",
        "Status": "New",
        "CreatedAfter": "2024-01-01",
        "SecurityKey": "optional"
    }

.OUTPUTS
    JSON array of tickets with notes and resolution details
#>

using namespace System.Net

param($Request, $TriggerMetadata)

Import-Module "C:\home\site\wwwroot\Modules\ConnectWiseManageAPI\ConnectWiseManageAPI.psm1"

# Connect to ConnectWise
$Connection = @{
    Server     = $env:ConnectWisePsa_ApiBaseUrl
    Company    = $env:ConnectWisePsa_ApiCompanyId
    PubKey     = $env:ConnectWisePsa_ApiPublicKey
    PrivateKey = $env:ConnectWisePsa_ApiPrivateKey
    ClientID   = $env:ConnectWisePsa_ApiClientId
}
Connect-CWM @Connection

# Extract filters
$Summary        = $Request.Body.SummaryContains
$Status         = $Request.Body.Status
$CreatedAfter   = $Request.Body.CreatedAfter
$CreatedBefore  = $Request.Body.CreatedBefore
$SecurityKey    = $env:SecurityKey

if ($SecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break
}

# Build query conditions
$conditions = @()
if ($Summary)       { $conditions += "summary contains '$Summary'" }
if ($Status)        { $conditions += "status/name='$Status'" }
if ($CreatedAfter)  { $conditions += "dateEntered>[$CreatedAfter]" }
if ($CreatedBefore) { $conditions += "dateEntered<[$CreatedBefore]" }

$filter = $conditions -join " and "

# Fetch tickets
$tickets = Get-CWMTickets -conditions $filter -pageSize 50

# Enrich each ticket with notes and resolution
$enrichedTickets = @()
foreach ($ticket in $tickets) {
    $ticketId = $ticket.id
    $notes = Get-CWMTicketNotes -ticketId $ticketId
    # Try to extract resolution from notes or description
    $resolutionNote = $notes | Where-Object { $_.internalAnalysisFlag -eq $true -or $_.resolutionFlag -eq $true } | Select-Object -First 1
    $resolutionText = if ($resolutionNote) { $resolutionNote.text } else { $ticket.resolution }

    $enrichedTickets += @{
        id          = $ticket.id
        summary     = $ticket.summary
        status      = $ticket.status.name
        priority    = $ticket.priority.name
        company     = $ticket.company.name
        dateEntered = $ticket.dateEntered
        notes       = $notes
        resolution  = $resolutionText
    }
}

$body = @{
    tickets = ($enrichedTickets | ConvertTo-Json -Depth 6)
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})

Disconnect-CWM
