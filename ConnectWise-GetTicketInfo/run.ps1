
<# .SYNOPSIS
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
$Keyword        = $Request.Body.Keyword
$SecurityKey    = $env:SecurityKey

if ($SecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break
}

# Default to recent tickets (last 6 months) if CreatedAfter is not provided
$defaultStartDate = (Get-Date).AddMonths(-6).ToString("yyyy-MM-dd")
if (-not $CreatedAfter) {
    $CreatedAfter = $defaultStartDate
}

# Build query conditions
function Build-Conditions {
    param($TicketId, $Summary, $Status, $Priority, $Company, $Contact, $Board, $ConfigItem, $CreatedAfter, $CreatedBefore)
    $conditions = @()
    if ($TicketId)      { $conditions += "id=$TicketId" }
    if ($Summary)       { $conditions += "summary contains '$Summary'" }
    if ($Status)        { $conditions += "status/name='$Status'" }
    if ($Priority)      { $conditions += "priority/name='$Priority'" }
    if ($Company)       { $conditions += "company/name contains '$Company'" }
    if ($Contact)       { $conditions += "contact/name='$Contact'" }
    if ($Board)         { $conditions += "board/name='$Board'" }
    if ($ConfigItem)    { $conditions += "configurationItems/identifier='$ConfigItem'" }
    if ($CreatedAfter)  { $conditions += "dateEntered>[$CreatedAfter]" }
    if ($CreatedBefore) { $conditions += "dateEntered<[$CreatedBefore]" }
    return $conditions -join " and "
}

$filter = Build-Conditions -TicketId $TicketId -Summary $Summary -Status $Status -Priority $Priority -Company $Company -Contact $Contact -Board $Board -ConfigItem $ConfigItem -CreatedAfter $CreatedAfter -CreatedBefore $CreatedBefore

# Fetch tickets using the existing Get-CWMTicket function
$tickets = Get-CWMTicket -condition $filter -pageSize 50 -all:$false

# If no tickets found and CreatedAfter was not explicitly provided, retry without CreatedAfter
if ($tickets.Count -eq 0 -and -not $Request.Body.CreatedAfter) {
    Write-Host "No tickets found in last 6 months. Expanding search..."
    $CreatedAfter = $null
    $filter = Build-Conditions -TicketId $TicketId -Summary $Summary -Status $Status -Priority $Priority -Company $Company -Contact $Contact -Board $Board -ConfigItem $ConfigItem -CreatedAfter $CreatedAfter -CreatedBefore $CreatedBefore
    $tickets = Get-CWMTicket -condition $filter -pageSize 50 -all:$false
}

# Enrich and filter tickets
$enrichedTickets = @()
foreach ($ticket in $tickets) {
    $ticketId = $ticket.id
    $notes = Get-CWMTicketNote -ticketId $ticketId

    # If a keyword is provided, filter tickets based on note content
    if ($Keyword) {
        $keywords = $Keyword -split ',\s*'
        $matchFound = $false
        foreach ($note in $notes) {
            foreach ($kw in $keywords) {
                if ($note.text -like "*$kw*") {
                    $matchFound = $true
                    break
                }
            }
            if ($matchFound) { break }
        }
        if (-not $matchFound) {
            continue
        }
    }

    # Extract resolution if available
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
    tickets = ($enrichedTickets | ConvertTo-Json -Depth 10)
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})

Disconnect-CWM
