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
if ($JsonObject.Groups.MirroredUsers.MirroredUserEmail -or $JsonObject.Groups.MirroredUsers.MirroredUserGroups) {
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

        # Update metadata status after group assignment
        $JsonObject.metadata.status.groupAssignment = "successful"
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
    $ticketNoteObject = & "$PSScriptRoot\modules\Format-TicketNote.ps1" -Json $JsonObject -ModuleOutputs $AllOutputs
    $TicketId = $ticketNoteObject.TicketId
    $ticketNote = $ticketNoteObject.Message
    Write-Host "✅ Ticket note formatted."
}
catch {
    $errorMsg = "❌ Exception formatting ticket note: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        error   = $_.Exception.Message
        message = "Dispatcher failed during final formatting"
        errors  = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
Write-Host "📤 Final Ticket Note module output: $($ticketNoteObject | ConvertTo-Json -Depth 5)"
# === STEP 8: Create ConnectWise Ticket Note ===
try {
    Write-Host "📬 Adding note to ConnectWise ticket $TicketId..."

    . "$PSScriptRoot\modules\Update-ConnectWiseTicketNote.ps1"
    $ticketNoteResult = Update-ConnectWiseTicketNote -TicketId $TicketId -Message $ticketNote

    if ($ticketNoteResult.Status -ne "Success") {
        Write-Warning "⚠️ ConnectWise ticket note failed to add"
        Write-Warning "Message: $($ticketNoteResult.Message)"

        if ($ticketNoteResult.Error) {
            Write-Error "Error: $($ticketNoteResult.Error)"
        }

        if ($ticketNoteResult.Stack) {
            Write-Verbose "Stack Trace: $($ticketNoteResult.Stack)"
        }
    }
    else {
        Write-Information $ticketNoteResult.Message
    }

    return @{
        result       = $ticketNoteResult.Status
        message      = $ticketNote
        noteStatus   = $ticketNoteResult.Status
        noteMessage  = $ticketNoteResult.Message
        errors       = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
catch {
    $errorMsg = "❌ Exception while adding ConnectWise note: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        error   = $_.Exception.Message
        message = "Dispatcher failed during ConnectWise note creation"
        errors  = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
