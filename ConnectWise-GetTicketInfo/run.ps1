<# 
.SYNOPSIS
    Retrieves ConnectWise tickets with filters, notes, and resolution details.

.DESCRIPTION
    This function queries ConnectWise for tickets using advanced filters and enriches each ticket with notes and resolution information.

.INPUTS
    JSON Structure:
    {
        "TicketId": "123456",
        "SummaryContains": "printer",
        "Status": "New",
        "Priority": "High",
        "Company": "Contoso Ltd",
        "Contact": "John Doe",
        "Board": "Service Desk",
        "ConfigItem": "Printer-01",
        "CreatedAfter": "2024-01-01",
        "CreatedBefore": "2024-12-31",
        "SecurityKey": "optional"
    }

.OUTPUTS
    JSON array of enriched tickets
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
$TicketId       = $Request.Body.TicketId
$Summary        = $Request.Body.SummaryContains
$Status         = $Request.Body.Status
$Priority       = $Request.Body.Priority
$Company        = $Request.Body.Company
$Contact        = $Request.Body.Contact
$Board          = $Request.Body.Board
$ConfigItem     = $Request.Body.ConfigItem
$CreatedAfter   = $Request.Body.CreatedAfter
$CreatedBefore  = $Request.Body.CreatedBefore
$SecurityKey    = $env:SecurityKey

if ($SecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break
}

# Build query conditions
$conditions = @()
if ($TicketId)      { $conditions += "id=$TicketId" }
if ($Summary)       { $conditions += "summary contains '$Summary'" }
if ($Status)        { $conditions += "status/name='$Status'" }
if ($Priority)      { $conditions += "priority/name='$Priority'" }
if ($Company)       { $conditions += "company/name='$Company'" }
if ($Contact)       { $conditions += "contact/name='$Contact'" }
if ($Board)         { $conditions += "board/name='$Board'" }
if ($ConfigItem)    { $conditions += "configurationItems/identifier='$ConfigItem'" }
if ($CreatedAfter)  { $conditions += "dateEntered>[$CreatedAfter]" }
if ($CreatedBefore) { $conditions += "dateEntered<[$CreatedBefore]" }

$filter = $conditions -join " and "

# Fetch tickets using the existing Get-CWMTicket function
$tickets = Get-CWMTicket -condition $filter -pageSize 50 -all

# Enrich each ticket with notes and resolution
$enrichedTickets = @()
foreach ($ticket in $tickets) {
    $ticketId = $ticket.id
    $notes = Get-CWMTicketNote -ticketId $ticketId
    # Try to extract resolution from notes or ticket field
    $resolutionNote = $notes | Where-Object { $_.internalAnalysisFlag -eq $true -or $_.resolutionFlag -eq $true } | Select-Object -First 1
    $resolutionText = if ($resolutionNote) { $resolutionNote.text } else { $ticket.resolution }

    $enrichedTickets += @{
        id          = $ticket.id
        summary     = $ticket.summary
        status      = $ticket.status.name
        priority    = $ticket.priority.name
        company     = $ticket.company.name
        contact     = $ticket.contact.name
        board       = $ticket.board.name
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
