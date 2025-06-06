using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# === LOAD MODULES ===
. "$PSScriptRoot\modules\utils.ps1"
$InformationPreference = 'Continue'

# === STEP 0: Read and Clean JSON ===
try {
    Write-Host "📥 Reading JSON input..."
    $rawJson = ($Request.Body | Out-String).Trim()
    Write-Host "🧼 Cleaning placeholders..."

    $CleanedJson = Clear-Placeholders -JsonString $rawJson
    $JsonObject = $CleanedJson | ConvertFrom-Json -Depth 10
    $JsonObject = Update-Placeholders -JsonObject $JsonObject

    Write-Host "✅ Placeholders removed and JSON parsed."
}
catch {
    Write-Host "❌ Failed to clean or parse JSON: $($_.Exception.Message)"
    return @{
        error = $_.Exception.Message
        message = "Dispatcher failed at JSON parsing"
    } | ConvertTo-Json -Depth 10
}

# === STEP 1: Initialize Metadata ===
try {
    Initialize-Metadata -Json $JsonObject
}
catch {
    Write-Host "⚠ Failed to initialize metadata: $($_.Exception.Message)"
}

# === MODULE EXECUTION LOGGING ===
$AllOutputs = @{}

# === STEP 2: Validate JSON ===
try {
    Write-Host "🔍 Running JSON validation..."
    $validationResult = Test-NewUserJson -Data $JsonObject
    $AllOutputs["Validation"] = $validationResult

    # Update metadata status after validation
    $JsonObject.metadata.status.validation = "successful"
}
catch {
    $errorMsg = "❌ Exception during JSON validation: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
    $JsonObject.metadata.status.validation = "failed"
}

# === STEP 3: Group Mirroring ===
if (($JsonObject.Groups.MirroredUsers.MirroredUserEmail -and $JsonObject.Groups.MirroredUsers.MirroredUserEmail.Trim()) -or 
    ($JsonObject.Groups.MirroredUsers.MirroredUserGroups -and $JsonObject.Groups.MirroredUsers.MirroredUserGroups.Trim())) {
    try {
        Write-Host "➡ Fetching mirrored group memberships..."
        & "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1" -Json $JsonObject

        # No patching needed; module modifies $JsonObject in-place
        $AllOutputs["MirroredGroups"] = @{
            Groups   = $JsonObject.Groups
            Metadata = $JsonObject.metadata
        }

        # Update metadata status
        $JsonObject.metadata.status.groupAssignment = "successful"
    }
    catch {
        $errorMsg = "❌ Exception during mirrored group fetch: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
        $JsonObject.metadata.status.groupAssignment = "failed"
    }
}

# === STEP 4: Create User ===
try {
    Write-Host "👤 Creating user..."
    $userCreationOutput = & "$PSScriptRoot\modules\Invoke-CreateNewUser.ps1" -Json $JsonObject
    $AllOutputs["CreateUser"] = $userCreationOutput
    if ($userCreationOutput.ResultStatus -eq 'failed') {
        $userCreationFailed = $true
        Write-Host "❌ User creation reported failure. Skipping groups and licenses."
    }

    # Update metadata status after user creation
    if ($userCreationFailed) {
        $JsonObject.metadata.status.userCreation = "failed"
    } else {
        $JsonObject.metadata.status.userCreation = "successful"
    }
}
catch {
    $errorMsg = "❌ Exception during user creation: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
    $userCreationFailed = $true
    $JsonObject.metadata.status.userCreation = "failed"
}
Write-Host "📤 CreateUser module output: $($userCreationOutput | ConvertTo-Json -Depth 5)"
# === STEP 5: Add to Groups ===
if (-not ($JsonObject.metadata.status.userCreation -eq "failed")) {
    try {
        Write-Host "👥 Adding user to groups..."
        $groupAssignmentOutput = & "$PSScriptRoot\modules\Add-UserGroups.ps1" -Json $JsonObject
        $AllOutputs["Groups"] = $groupAssignmentOutput

        # Propagate returned groupAssignment status
        if ($groupAssignmentOutput.metadata.status.groupAssignment) {
            $JsonObject.metadata.status.groupAssignment = $groupAssignmentOutput.metadata.status.groupAssignment
        } else {
            $JsonObject.metadata.status.groupAssignment = "unknown"
        }

        # Merge any returned errors
        if ($groupAssignmentOutput.metadata.errors) {
            $JsonObject.metadata.errors += $groupAssignmentOutput.metadata.errors
        }

        # Optionally, mirror manualFallback and mirroredFrom for Format-TicketNote compatibility
        if ($groupAssignmentOutput.metadata.manualFallback) {
            $JsonObject.metadata.manualFallback = $groupAssignmentOutput.metadata.manualFallback
        }
        if ($groupAssignmentOutput.metadata.mirroredFrom) {
            $JsonObject.metadata.mirroredFrom = $groupAssignmentOutput.metadata.mirroredFrom
        }

        # If GroupsAssigned returned, append to root Json for consistency
        if ($groupAssignmentOutput.GroupsAssigned) {
            $JsonObject.GroupsAssigned = $groupAssignmentOutput.GroupsAssigned
        }
    }
    catch {
        $errorMsg = "❌ Exception during group assignment: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
        $JsonObject.metadata.status.groupAssignment = "failed"
    }
}

Write-Host "📤 GroupAssignment module output: $($groupAssignmentOutput | ConvertTo-Json -Depth 5)"

# === STEP 6: Licensing (optional) ===
$licenseModule = "$PSScriptRoot\modules\Assign-License.ps1"
if ((-not $userCreationFailed) -and (Test-Path $licenseModule)) {
    try {
        Write-Host "🎫 Assigning licenses..."
        $licenseOutput = & $licenseModule -Json $JsonObject
        $AllOutputs["Licenses"] = $licenseOutput

        # Update metadata status after licensing
        $JsonObject.metadata.status.licensing = "successful"
    }
    catch {
        $errorMsg = "❌ Exception during licensing: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
        $JsonObject.metadata.status.licensing = "failed"
    }
}

# === STEP 7: Format Final Ticket Note ===
try {
    Write-Host "📝 Formatting ConnectWise ticket note..."
    
    # Debug: Show the full outputs structure
    Write-Host "=== ALL OUTPUTS STRUCTURE ==="
    Write-Host ($AllOutputs | ConvertTo-Json -Depth 5)

    $ticketNoteObject = & "$PSScriptRoot\modules\Format-TicketNote.ps1" -AllOutputs $AllOutputs
    
    # Validate output
    if (-not $ticketNoteObject.Message) {
        throw "Formatted note content is empty"
    }

    $TicketId = if ($ticketNoteObject.TicketId) { 
        $ticketNoteObject.TicketId 
    } else { 
        # Fallback to finding ticket ID in any output
        $AllOutputs | ForEach-Object {
            if ($_.TicketId) { $_.TicketId }
            elseif ($_.metadata.ticket.id) { $_.metadata.ticket.id }
        } | Select-Object -First 1
    }

    if (-not $TicketId) {
        throw "Could not determine TicketId from any source"
    }

    $ticketNote = $ticketNoteObject.Message

    Write-Host "✅ Final Note Content:"
    Write-Host $ticketNote
}
catch {
    $errorMsg = "❌ Exception formatting ticket note: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        status  = "failed"
        error   = $_.Exception.Message
        message = "Dispatcher failed during final formatting"
    } | ConvertTo-Json
}
# === STEP 8: Create ConnectWise Ticket Note ===
try {
    Write-Host "📬 Adding note to ConnectWise ticket $TicketId..."
    
    . "$PSScriptRoot\modules\Update-ConnectWiseTicketNote.ps1"
    
    # Pass the clean note content
    $ticketNoteResult = Update-ConnectWiseTicketNote -TicketId $TicketId -Message $ticketNote

    if ($ticketNoteResult.Status -ne "Success") {
        Write-Warning "⚠️ ConnectWise ticket note failed to add: $($ticketNoteResult.Message)"
        if ($ticketNoteResult.Error) {
            Write-Error "Details: $($ticketNoteResult.Error)"
        }
    }

    return @{
        status     = "completed"
        ticketId   = $TicketId
        noteStatus = $ticketNoteResult.Status
        success    = ($ticketNoteResult.Status -eq "Success")
    } | ConvertTo-Json
}
catch {
    $errorMsg = "❌ Exception while adding ConnectWise note: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        status    = "failed"
        error     = $_.Exception.Message
        ticketId  = $TicketId
    } | ConvertTo-Json
}
