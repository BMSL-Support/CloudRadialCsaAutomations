using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# === ENHANCED LOGGING SETUP ===
$InformationPreference = 'Continue'
$DebugPreference = 'Continue'
$ErrorActionPreference = 'Stop'

# === LOAD MODULES ===
. "$PSScriptRoot\modules\utils.ps1"

# === STEP 0: Read and Clean JSON ===
try {
    Write-Host "📥 Reading JSON input..."
    $rawJson = ($Request.Body | Out-String).Trim()
    Write-Debug "Raw JSON input:`n$rawJson"

    Write-Host "🧼 Cleaning placeholders..."
    $CleanedJson = Clear-Placeholders -JsonString $rawJson
    Write-Debug "Cleaned JSON:`n$CleanedJson"
    
    $JsonObject = $CleanedJson | ConvertFrom-Json -Depth 10
    $JsonObject = Update-Placeholders -JsonObject $JsonObject
    
    # Validate critical fields
    if (-not $JsonObject.TicketId -or $JsonObject.TicketId -match '^@') {
        throw "Invalid or missing TicketId in JSON input"
    }
    if (-not $JsonObject.AccountDetails.UserPrincipalName) {
        throw "Missing UserPrincipalName in AccountDetails"
    }
    
    Write-Host "✅ JSON parsed and validated. Ticket ID: $($JsonObject.TicketId)"
}
catch {
    $errorMsg = "❌ JSON processing failed: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher failed at JSON parsing"
        details = @{
            inputJson = $rawJson
            error     = $_.Exception | Select-Object *
        }
    } | ConvertTo-Json -Depth 10
}

# === STEP 1: Initialize Metadata ===
try {
    Initialize-Metadata -Json $JsonObject
    Write-Debug "Metadata initialized"
}
catch {
    Write-Host "⚠ Metadata initialization warning: $($_.Exception.Message)"
}

# === MODULE EXECUTION LOGGING ===
$AllOutputs = @{
    Timestamp = [DateTime]::UtcNow.ToString('o')
    Steps     = @{}
}

# === STEP 2: Validate JSON ===
try {
    Write-Host "🔍 Running JSON validation..."
    $validationResult = Test-NewUserJson -Data $JsonObject
    $AllOutputs.Steps["Validation"] = $validationResult
    
    if ($validationResult.Valid -eq $false) {
        throw "JSON validation failed: $($validationResult.Message)"
    }
    
    $JsonObject.metadata.status.validation = "successful"
    Write-Host "✅ JSON validation passed"
}
catch {
    $errorMsg = "❌ JSON validation failed: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
    $JsonObject.metadata.status.validation = "failed"
    
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher failed during JSON validation"
        errors  = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}

# === STEP 3: Group Mirroring ===
if ($JsonObject.Groups.MirroredUsers.MirroredUserEmail -or $JsonObject.Groups.MirroredUsers.MirroredUserGroups) {
    try {
        Write-Host "➡ Fetching mirrored group memberships..."
        $mirrorResult = & "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1" -Json $JsonObject
        
        $AllOutputs.Steps["MirroredGroups"] = @{
            Groups   = $JsonObject.Groups
            Metadata = $JsonObject.metadata
            Result   = $mirrorResult
        }
        
        $JsonObject.metadata.status.groupAssignment = "successful"
        Write-Host "✅ Mirrored groups processed"
    }
    catch {
        $errorMsg = "❌ Mirrored group fetch failed: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
        $JsonObject.metadata.status.groupAssignment = "failed"
    }
}

# === STEP 4: Create User ===
$userCreationFailed = $false
try {
    Write-Host "👤 Creating user $($JsonObject.AccountDetails.UserPrincipalName)..."
    $userCreationOutput = & "$PSScriptRoot\modules\Invoke-CreateNewUser.ps1" -Json $JsonObject
    
    Write-Host "🔍 User Creation Output:"
    $userCreationOutput | Format-List | Out-Host
    
    $AllOutputs.Steps["CreateUser"] = $userCreationOutput
    
    if ($userCreationOutput.ResultStatus -ne 'success') {
        $userCreationFailed = $true
        throw "User creation reported failure: $($userCreationOutput.Message)"
    }
    
    # Verify user exists in Azure AD
    try {
        $user = Get-MgUser -UserId $JsonObject.AccountDetails.UserPrincipalName -ErrorAction Stop
        if (-not $user) {
            throw "User verification failed - user not found in Azure AD"
        }
        Write-Host "✅ User verified in Azure AD"
    }
    catch {
        $userCreationFailed = $true
        throw "User verification failed: $($_.Exception.Message)"
    }
    
    $JsonObject.metadata.status.userCreation = "successful"
}
catch {
    $userCreationFailed = $true
    $errorMsg = "❌ User creation failed: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
    $JsonObject.metadata.status.userCreation = "failed"
    
    # Skip remaining steps if user creation failed
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher failed during user creation"
        errors  = $JsonObject.metadata.errors
        outputs = $AllOutputs
    } | ConvertTo-Json -Depth 10
}

# === STEP 5: Add to Groups ===
if (-not $userCreationFailed) {
    try {
        Write-Host "👥 Adding user to groups..."
        $groupAssignmentOutput = & "$PSScriptRoot\modules\Add-UserGroups.ps1" -Json $JsonObject
        $AllOutputs.Steps["Groups"] = $groupAssignmentOutput
        
        $JsonObject.metadata.status.groupAssignment = "successful"
        Write-Host "✅ Group assignment completed"
    }
    catch {
        $errorMsg = "❌ Group assignment failed: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
        $JsonObject.metadata.status.groupAssignment = "failed"
    }
}

# === STEP 6: Licensing ===
$licenseModule = "$PSScriptRoot\modules\Assign-License.ps1"
if ((-not $userCreationFailed) -and (Test-Path $licenseModule)) {
    try {
        Write-Host "🎫 Assigning licenses..."
        $licenseOutput = & $licenseModule -Json $JsonObject
        $AllOutputs.Steps["Licenses"] = $licenseOutput
        
        $JsonObject.metadata.status.licensing = "successful"
        Write-Host "✅ License assignment completed"
    }
    catch {
        $errorMsg = "❌ Licensing failed: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
        $JsonObject.metadata.status.licensing = "failed"
    }
}

# === STEP 7: Format Final Ticket Note ===
try {
    Write-Host "📝 Formatting ConnectWise ticket note..."
    $ticketNoteObject = & "$PSScriptRoot\modules\Format-TicketNote.ps1" -Json $JsonObject -ModuleOutputs $AllOutputs
    
    if (-not $ticketNoteObject.TicketId) {
        throw "TicketId not returned from Format-TicketNote"
    }
    if (-not $ticketNoteObject.Message) {
        throw "Empty message returned from Format-TicketNote"
    }
    
    $TicketId = $ticketNoteObject.TicketId
    $ticketNote = $ticketNoteObject.Message
    
    Write-Host "✅ Ticket note formatted for ticket $TicketId"
}
catch {
    $errorMsg = "❌ Ticket note formatting failed: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher failed during final formatting"
        errors  = $JsonObject.metadata.errors
        outputs = $AllOutputs
    } | ConvertTo-Json -Depth 10
}

# === STEP 8: Create ConnectWise Ticket Note ===
try {
    Write-Host "📬 Adding note to ConnectWise ticket $TicketId..."
    
    . "$PSScriptRoot\modules\Update-ConnectWiseTicketNote.ps1"
    $ticketNoteResult = Update-ConnectWiseTicketNote -TicketId $TicketId -Message $ticketNote
    
    if ($ticketNoteResult.Status -ne "Success") {
        throw "ConnectWise ticket note failed: $($ticketNoteResult.Message)"
    }
    
    Write-Host "✅ Ticket note added successfully"
    
    return @{
        status      = "success"
        message     = "User provisioning completed"
        ticketId    = $TicketId
        userCreated = (-not $userCreationFailed)
        outputs     = $AllOutputs
        errors      = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
catch {
    $errorMsg = "❌ ConnectWise note creation failed: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        status  = "partial"
        error   = $errorMsg
        message = "User provisioning completed but ticket update failed"
        outputs = $AllOutputs
        errors  = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
