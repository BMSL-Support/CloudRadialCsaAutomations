<# 
.SYNOPSIS
    Retrieves ConnectWise tickets with filters, notes, and resolution details.

.DESCRIPTION
    This function queries ConnectWise for tickets using advanced filters and enriches each ticket with notes and resolution information.
    It also supports post-filtering based on a keyword found in the ticket notes.

.INPUTS
    JSON Structure:
    {
        "TicketId": "123456",
        "SummaryContains": "printer",
        "Status": "New",
        "Priority": "High",
        "Company": "Fabrikam Ltd",
        "Contact": "Joe Bloggs",
        "Board": "Service Desk",
        "ConfigItem": "Printer-01",
        "CreatedAfter": "2023-01-01",
        "CreatedBefore": "2024-12-31",
        "Keyword": "AI",
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

# Security key check
$SecurityKey = $env:SecurityKey
if ($SecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    Disconnect-CWM
    return
}

# Extract request body
$Body = $Request.Body

# Apply defaults (last 12 months)
$CreatedAfter  = if ($Body.CreatedAfter) { $Body.CreatedAfter } else { (Get-Date).AddMonths(-12).ToString("yyyy-MM-dd") }
$CreatedBefore = if ($Body.CreatedBefore) { $Body.CreatedBefore } else { (Get-Date).ToString("yyyy-MM-dd") }

# Build filter conditions
$conditions = @()
if ($Body.TicketId)      { $conditions += "id=$($Body.TicketId)" }
if ($Body.SummaryContains) { $conditions += "summary contains '$($Body.SummaryContains)'" }
if ($Body.Status)        { $conditions += "status/name='$($Body.Status)'" }
if ($Body.Priority)      { $conditions += "priority/name='$($Body.Priority)'" }
if ($Body.Company)       { $conditions += "company/name contains '$($Body.Company)'" }
if ($Body.Contact)       { $conditions += "contact/name='$($Body.Contact)'" }
if ($Body.Board)         { $conditions += "board/name='$($Body.Board)'" }
if ($Body.ConfigItem)    { $conditions += "configurationItems/identifier='$($Body.ConfigItem)'" }
$conditions += "dateEntered>[$CreatedAfter]"
$conditions += "dateEntered<[$CreatedBefore]"

$filter = $conditions -join " and "
$pageSize = if ($Body.MaxResults) { [int]$Body.MaxResults } else { 50 }

# Fetch tickets
$tickets = Get-CWMTicket -condition $filter -pageSize $pageSize -all:$false

# Enrich tickets (with optional keyword filtering)
$enrichedTickets = foreach ($ticket in $tickets) {
    $notes = @()
    $includeTicket = $true

    if ($Body.Keyword) {
        $notes = Get-CWMTicketNote -ticketId $ticket.id
        $includeTicket = $notes | Where-Object { $_.text -like "*$($Body.Keyword)*" } | Select-Object -First 1
    } else {
        $notes = Get-CWMTicketNote -ticketId $ticket.id
    }

    if ($includeTicket) {
        $resolutionNote = $notes | Where-Object { $_.internalAnalysisFlag -or $_.resolutionFlag } | Select-Object -First 1
        [PSCustomObject]@{
            id          = $ticket.id
            summary     = $ticket.summary
            status      = $ticket.status.name
            priority    = $ticket.priority.name
            company     = $ticket.company.name
            contact     = $ticket.contact.name
            board       = $ticket.board.name
            dateEntered = $ticket.dateEntered
            notes       = $notes
            resolution  = if ($null -ne $resolutionNote -and $null -ne $resolutionNote.text -and $resolutionNote.text -ne "") { $resolutionNote.text } else { $ticket.resolution }
        }
    }
}

# Convert and respond
$body = $enrichedTickets | ConvertTo-Json -Depth 10
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body       = $body
    ContentType = "application/json"
})

Disconnect-CWM
