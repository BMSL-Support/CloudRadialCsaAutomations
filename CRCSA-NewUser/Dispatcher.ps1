using namespace System.Net

# REQUIRED: Azure Functions parameter declaration
param($Request, $TriggerMetadata)

# Initialize execution state tracking
$ExecutionState = @{
    CurrentStep = "initialization"
    Steps = @(
        @{Name = "initialization"; Status = "started"},
        @{Name = "validation"; Status = "pending"},
        @{Name = "user_creation"; Status = "pending"},
        @{Name = "group_assignment"; Status = "pending"},
        @{Name = "licensing"; Status = "pending"},
        @{Name = "ticket_update"; Status = "pending"}
    )
}

# Ensure metadata exists before any processing
Initialize-Metadata -Json $JsonObject

# Enable strict error handling
$ErrorActionPreference = 'Stop'
$DebugPreference = 'Continue'
$VerbosePreference = 'Continue'
$InformationPreference = 'Continue'

# === INITIALIZATION ===
if (-not $JsonObject.metadata) {
    Initialize-Metadata -Json $JsonObject
}
else {
    # Ensure legacy format compatibility
    if (-not $JsonObject.metadata.PSObject.Properties['status']) {
        $JsonObject.metadata | Add-Member -NotePropertyName 'status' -NotePropertyValue @{}
    }
}

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
    $errorMsg = "❌ Error in step '$($ExecutionState.CurrentStep)': $($_.Exception.Message)"
    
    # Ensure metadata is properly initialized
    Initialize-Metadata -Json $JsonObject
    
    # Safely add error to metadata
    $JsonObject.metadata.errors += $errorMsg
    
    # Update step status
    if ($JsonObject.metadata.status.PSObject.Properties[$ExecutionState.CurrentStep]) {
        $JsonObject.metadata.status.$($ExecutionState.CurrentStep) = "failed"
    }
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host "Error details: $($_.Exception | Out-String)" -ForegroundColor DarkRed
    
    return @{
        status   = "failed"
        error    = $errorMsg
        message  = "Processing stopped due to errors"
        metadata = $JsonObject.metadata
        debug    = @{
            timestamp    = [DateTime]::UtcNow.ToString('o')
            currentStep  = $ExecutionState.CurrentStep
            errorDetails = $_.Exception | Select-Object *
        }
    } | ConvertTo-Json -Depth 5
}
# === STEP 0: PROCESS INPUT ===
try {
    Write-Host "🔍 Running JSON validation..."
    $validationResult = Test-NewUserJson -Data $JsonObject
    
    if (-not $validationResult.IsValid) {
        throw "Validation failed: $($validationResult.Errors -join ', ')"
    }
    
    $AllOutputs["Validation"] = @{
        Status = if ($validationResult.Warnings) {"completed_with_warnings"} else {"success"}
        Errors = $validationResult.Errors
        Warnings = $validationResult.Warnings
    }
    
    if ($validationResult.Warnings) {
        Write-Warning "Validation completed with warnings: $($validationResult.Warnings -join ', ')"
    }
}
catch {
    $errorMsg = "❌ Error in step '$($ExecutionState.CurrentStep)': $($_.Exception.Message)"
    
    # Ensure metadata is properly initialized
    Initialize-Metadata -Json $JsonObject
    
    # Safely add error to metadata
    $JsonObject.metadata.errors += $errorMsg
    
    # Update step status
    if ($JsonObject.metadata.status.PSObject.Properties[$ExecutionState.CurrentStep]) {
        $JsonObject.metadata.status.$($ExecutionState.CurrentStep) = "failed"
    }
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host "Error details: $($_.Exception | Out-String)" -ForegroundColor DarkRed
    
    return @{
        status   = "failed"
        error    = $errorMsg
        message  = "Processing stopped due to errors"
        metadata = $JsonObject.metadata
        debug    = @{
            timestamp    = [DateTime]::UtcNow.ToString('o')
            currentStep  = $ExecutionState.CurrentStep
            errorDetails = $_.Exception | Select-Object *
        }
    } | ConvertTo-Json -Depth 5
}
# === MAIN EXECUTION FLOW ===
try {
    # Initialize metadata
    Initialize-Metadata -Json $JsonObject

    # STEP 1: JSON Validation
try {
    Write-Log "🔍 Running JSON validation..."
    
    # Run validation with non-strict mode (groups are optional)
    $validationResult = Test-NewUserJson -Data $JsonObject -StrictValidation:$false

    if (-not $validationResult.Valid) {
        $errorMsg = "❌ JSON validation failed: $($validationResult.Message)`n" +
                    ($validationResult.Errors -join "`n")
        throw $errorMsg
    }

    $AllOutputs.Steps["Validation"] = $validationResult
    $JsonObject.metadata.status.validation = if ($validationResult.Warnings.Count -gt 0) {
        "completed_with_warnings"
    } else {
        "successful"
    }

    # Add warnings to metadata if any
    if ($validationResult.Warnings.Count -gt 0) {
        $JsonObject.metadata.warnings = @($JsonObject.metadata.warnings) + $validationResult.Warnings
    }

    Write-Log "✅ Validation passed with $($validationResult.Warnings.Count) warnings"
}
catch {
    $errorMsg = "❌ Error in step '$($ExecutionState.CurrentStep)': $($_.Exception.Message)"
    
    # Ensure metadata is properly initialized
    Initialize-Metadata -Json $JsonObject
    
    # Safely add error to metadata
    $JsonObject.metadata.errors += $errorMsg
    
    # Update step status
    if ($JsonObject.metadata.status.PSObject.Properties[$ExecutionState.CurrentStep]) {
        $JsonObject.metadata.status.$($ExecutionState.CurrentStep) = "failed"
    }
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host "Error details: $($_.Exception | Out-String)" -ForegroundColor DarkRed
    
    return @{
        status   = "failed"
        error    = $errorMsg
        message  = "Processing stopped due to errors"
        metadata = $JsonObject.metadata
        debug    = @{
            timestamp    = [DateTime]::UtcNow.ToString('o')
            currentStep  = $ExecutionState.CurrentStep
            errorDetails = $_.Exception | Select-Object *
        }
    } | ConvertTo-Json -Depth 5
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
    $errorMsg = "❌ Error in step '$($ExecutionState.CurrentStep)': $($_.Exception.Message)"
    
    # Ensure metadata is properly initialized
    Initialize-Metadata -Json $JsonObject
    
    # Safely add error to metadata
    $JsonObject.metadata.errors += $errorMsg
    
    # Update step status
    if ($JsonObject.metadata.status.PSObject.Properties[$ExecutionState.CurrentStep]) {
        $JsonObject.metadata.status.$($ExecutionState.CurrentStep) = "failed"
    }
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host "Error details: $($_.Exception | Out-String)" -ForegroundColor DarkRed
    
    return @{
        status   = "failed"
        error    = $errorMsg
        message  = "Processing stopped due to errors"
        metadata = $JsonObject.metadata
        debug    = @{
            timestamp    = [DateTime]::UtcNow.ToString('o')
            currentStep  = $ExecutionState.CurrentStep
            errorDetails = $_.Exception | Select-Object *
        }
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
    $errorMsg = "❌ Error in step '$($ExecutionState.CurrentStep)': $($_.Exception.Message)"
    
    # Ensure metadata is properly initialized
    Initialize-Metadata -Json $JsonObject
    
    # Safely add error to metadata
    $JsonObject.metadata.errors += $errorMsg
    
    # Update step status
    if ($JsonObject.metadata.status.PSObject.Properties[$ExecutionState.CurrentStep]) {
        $JsonObject.metadata.status.$($ExecutionState.CurrentStep) = "failed"
    }
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host "Error details: $($_.Exception | Out-String)" -ForegroundColor DarkRed
    
    return @{
        status   = "failed"
        error    = $errorMsg
        message  = "Processing stopped due to errors"
        metadata = $JsonObject.metadata
        debug    = @{
            timestamp    = [DateTime]::UtcNow.ToString('o')
            currentStep  = $ExecutionState.CurrentStep
            errorDetails = $_.Exception | Select-Object *
        }
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
    $errorMsg = "❌ Error in step '$($ExecutionState.CurrentStep)': $($_.Exception.Message)"
    
    # Ensure metadata is properly initialized
    Initialize-Metadata -Json $JsonObject
    
    # Safely add error to metadata
    $JsonObject.metadata.errors += $errorMsg
    
    # Update step status
    if ($JsonObject.metadata.status.PSObject.Properties[$ExecutionState.CurrentStep]) {
        $JsonObject.metadata.status.$($ExecutionState.CurrentStep) = "failed"
    }
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host "Error details: $($_.Exception | Out-String)" -ForegroundColor DarkRed
    
    return @{
        status   = "failed"
        error    = $errorMsg
        message  = "Processing stopped due to errors"
        metadata = $JsonObject.metadata
        debug    = @{
            timestamp    = [DateTime]::UtcNow.ToString('o')
            currentStep  = $ExecutionState.CurrentStep
            errorDetails = $_.Exception | Select-Object *
        }
    } | ConvertTo-Json -Depth 5
}
