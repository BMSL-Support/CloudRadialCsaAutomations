using namespace System.Net

# Enable strict error handling
$ErrorActionPreference = 'Stop'
$DebugPreference = 'Continue'
$VerbosePreference = 'Continue'
$InformationPreference = 'Continue'

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# === INITIALIZATION ===
$global:FunctionStartTime = [DateTime]::UtcNow
$AllOutputs = @{
    Timestamp = $global:FunctionStartTime.ToString('o')
    Steps     = @{}
    Errors    = @()
}

function Write-Log {
    param($Message, $Level = 'Information')
    $timestamp = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    Write-Host "[$timestamp][$Level] $Message"
}

# === LOAD MODULES ===
try {
    Write-Log "Loading modules..."
    $moduleRoot = "$PSScriptRoot\modules"
    
    # Load utils first
    $utilsPath = "$moduleRoot\utils.ps1"
    if (-not (Test-Path $utilsPath)) {
        throw "Utils.ps1 not found at $utilsPath"
    }
    . $utilsPath
    
    # Verify all required modules exist
    $requiredModules = @(
        'Get-MirroredUserGroupMemberships.ps1',
        'Invoke-CreateNewUser.ps1',
        'Add-UserGroups.ps1',
        'Format-TicketNote.ps1',
        'Update-ConnectWiseTicketNote.ps1'
    )
    
    foreach ($module in $requiredModules) {
        $modulePath = "$moduleRoot\$module"
        if (-not (Test-Path $modulePath)) {
            throw "Required module $module not found at $modulePath"
        }
    }
}
catch {
    $errorMsg = "❌ Module loading failed: $($_.Exception.Message)"
    Write-Log $errorMsg -Level 'Error'
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher initialization failed"
        details = @{
            error = $_.Exception | Select-Object *
            availableModules = (Get-ChildItem "$PSScriptRoot\modules" | Select-Object Name)
        }
    } | ConvertTo-Json -Depth 5
}

# === STEP 0: PROCESS INPUT ===
try {
    Write-Log "📥 Reading JSON input..."
    $rawJson = ($Request.Body | Out-String).Trim()
    Write-Debug "Raw input JSON:`n$rawJson"

    # Validate basic JSON structure
    if ([string]::IsNullOrWhiteSpace($rawJson)) {
        throw "Empty request body received"
    }

    # Clean and parse JSON
    $CleanedJson = Clear-Placeholders -JsonString $rawJson
    $JsonObject = $CleanedJson | ConvertFrom-Json -Depth 10
    $JsonObject = Update-Placeholders -JsonObject $JsonObject

    # Validate required fields
    $requiredFields = @('TenantId', 'TicketId', 'AccountDetails')
    foreach ($field in $requiredFields) {
        if (-not $JsonObject.$field) {
            throw "Missing required field: $field"
        }
    }

    if (-not $JsonObject.AccountDetails.UserPrincipalName) {
        throw "Missing UserPrincipalName in AccountDetails"
    }

    Write-Log "✅ JSON parsed and validated. Ticket ID: $($JsonObject.TicketId)"
}
catch {
    $errorMsg = "❌ JSON processing failed: $($_.Exception.Message)"
    Write-Log $errorMsg -Level 'Error'
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher failed at JSON parsing"
        input   = $rawJson
        errorDetails = $_.Exception | Select-Object *
    } | ConvertTo-Json -Depth 5
}

# === MAIN EXECUTION FLOW ===
try {
    # Initialize metadata
    Initialize-Metadata -Json $JsonObject

    # STEP 1: JSON Validation
    try {
        Write-Log "🔍 Validating JSON structure..."
        $validationResult = Test-NewUserJson -Data $JsonObject
        $AllOutputs.Steps["Validation"] = $validationResult
        
        if (-not $validationResult.Valid) {
            throw $validationResult.Message
        }
        $JsonObject.metadata.status.validation = "successful"
    }
    catch {
        $errorMsg = "❌ Validation failed: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        throw $errorMsg
    }

    # STEP 2: Group Mirroring
    if ($JsonObject.Groups.MirroredUsers.MirroredUserEmail) {
        try {
            Write-Log "🔄 Processing mirrored groups..."
            $mirrorResult = & "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1" -Json $JsonObject
            $AllOutputs.Steps["MirroredGroups"] = $mirrorResult
            $JsonObject.metadata.status.groupAssignment = "successful"
        }
        catch {
            $errorMsg = "⚠️ Group mirroring failed: $($_.Exception.Message)"
            $JsonObject.metadata.errors += $errorMsg
            $JsonObject.metadata.status.groupAssignment = "partial"
            Write-Log $errorMsg -Level 'Warning'
        }
    }

    # STEP 3: User Creation
    try {
        Write-Log "👤 Creating user $($JsonObject.AccountDetails.UserPrincipalName)..."
        $userCreationOutput = & "$PSScriptRoot\modules\Invoke-CreateNewUser.ps1" -Json $JsonObject
        
        # Debug output
        Write-Debug "User Creation Module Output:"
        $userCreationOutput | Format-List | Out-Default
        
        if ($userCreationOutput.ResultStatus -ne 'success') {
            throw $userCreationOutput.Message
        }
        
        $AllOutputs.Steps["CreateUser"] = $userCreationOutput
        $JsonObject.metadata.status.userCreation = "successful"
    }
    catch {
        $errorMsg = "❌ User creation failed: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        $JsonObject.metadata.status.userCreation = "failed"
        
        return @{
            status  = "failed"
            error   = $errorMsg
            message = "User creation failed"
            outputs = $AllOutputs
            debug   = $userCreationOutput.Debug
        } | ConvertTo-Json -Depth 5
    }

    # STEP 4: Group Assignment
    if (-not $userCreationFailed) {
        try {
            Write-Log "👥 Assigning groups..."
            $groupResult = & "$PSScriptRoot\modules\Add-UserGroups.ps1" -Json $JsonObject
            $AllOutputs.Steps["Groups"] = $groupResult
            $JsonObject.metadata.status.groupAssignment = "successful"
        }
        catch {
            $errorMsg = "⚠️ Group assignment failed: $($_.Exception.Message)"
            $JsonObject.metadata.errors += $errorMsg
            $JsonObject.metadata.status.groupAssignment = "partial"
            Write-Log $errorMsg -Level 'Warning'
        }
    }

    # STEP 5: License Assignment
    $licenseModule = "$PSScriptRoot\modules\Assign-License.ps1"
    if ((Test-Path $licenseModule) -and (-not $userCreationFailed)) {
        try {
            Write-Log "🎫 Assigning licenses..."
            $licenseResult = & $licenseModule -Json $JsonObject
            $AllOutputs.Steps["Licenses"] = $licenseResult
            $JsonObject.metadata.status.licensing = "successful"
        }
        catch {
            $errorMsg = "⚠️ License assignment failed: $($_.Exception.Message)"
            $JsonObject.metadata.errors += $errorMsg
            $JsonObject.metadata.status.licensing = "partial"
            Write-Log $errorMsg -Level 'Warning'
        }
    }

    # STEP 6: Ticket Note
    try {
        Write-Log "📝 Generating ticket note..."
        $ticketNoteObject = & "$PSScriptRoot\modules\Format-TicketNote.ps1" -Json $JsonObject -ModuleOutputs $AllOutputs
        
        if (-not $ticketNoteObject.TicketId) {
            throw "Missing TicketId in note object"
        }
        
        $TicketId = $ticketNoteObject.TicketId
        $ticketNote = $ticketNoteObject.Message
    }
    catch {
        $errorMsg = "❌ Ticket note generation failed: $($_.Exception.Message)"
        Write-Log $errorMsg -Level 'Error'
        return @{
            status  = "partial"
            error   = $errorMsg
            message = "Processing completed but ticket note failed"
            outputs = $AllOutputs
        } | ConvertTo-Json -Depth 5
    }

    # STEP 7: Update Ticket
    try {
        Write-Log "📬 Updating ticket $TicketId..."
        $ticketResult = & "$PSScriptRoot\modules\Update-ConnectWiseTicketNote.ps1" -TicketId $TicketId -Message $ticketNote
        
        if ($ticketResult.Status -ne "Success") {
            throw $ticketResult.Message
        }
        
        return @{
            status      = "success"
            message     = "User provisioning completed"
            ticketId    = $TicketId
            outputs     = $AllOutputs
            errors      = $JsonObject.metadata.errors
            duration    = ([DateTime]::UtcNow - $global:FunctionStartTime).TotalSeconds
        } | ConvertTo-Json -Depth 5
    }
    catch {
        $errorMsg = "⚠️ Ticket update failed: $($_.Exception.Message)"
        Write-Log $errorMsg -Level 'Warning'
        return @{
            status  = "partial"
            error   = $errorMsg
            message = "Processing completed but ticket update failed"
            outputs = $AllOutputs
            ticketNote = $ticketNote
        } | ConvertTo-Json -Depth 5
    }
}
catch {
    $errorMsg = "❌ Dispatcher fatal error: $($_.Exception.Message)"
    Write-Log $errorMsg -Level 'Error'
    return @{
        status  = "failed"
        error   = $errorMsg
        message = "Dispatcher encountered a fatal error"
        outputs = $AllOutputs
        stackTrace = $_.ScriptStackTrace
    } | ConvertTo-Json -Depth 5
}
