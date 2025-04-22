using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# === LOAD MODULES ===
. "$PSScriptRoot\modules\utils.ps1"

# === STEP 0: Read and Clean JSON ===
try {
    Write-Host "📥 Reading JSON input..."
    $rawJson = ($Request.Body | Out-String).Trim()     # ✅ Correct source
    Write-Host "🧼 Cleaning placeholders..."

    $CleanedJson = Clear-Placeholders -JsonString $rawJson   # ✅ Fixed reference
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
$AllOutputs = @()

# === STEP 2: Validate JSON ===
try {
    Write-Host "🔍 Running JSON validation..."
    $validationResult = Test-NewUserJson -Json $JsonObject
    $AllOutputs += $validationResult
}
catch {
    $errorMsg = "❌ Exception during JSON validation: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
}

# === STEP 3: Group Mirroring ===
if ($JsonObject.Groups.MirroredUsers.MirroredUserEmail -or $JsonObject.Groups.MirroredUsers.MirroredUserGroups) {
    try {
        Write-Host "➡ Fetching mirrored group memberships..."
        $groupResult = & "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1" -Json $JsonObject
        $AllOutputs += "✅ Mirrored groups fetched successfully."
    }
    catch {
        $errorMsg = "❌ Exception during mirrored group fetch: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
    }
}

# === STEP 4: Create User ===
try {
    Write-Host "👤 Creating user..."
    $userCreationOutput = & "$PSScriptRoot\modules\Invoke-CreateNewUser.ps1" -Json $JsonObject
    $AllOutputs += $userCreationOutput
}
catch {
    $errorMsg = "❌ Exception during user creation: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
}

# === STEP 5: Add to Groups ===
try {
    Write-Host "👥 Adding user to groups..."
    $groupAssignmentOutput = & "$PSScriptRoot\modules\Add-UserGroups.ps1" -Json $JsonObject
    $AllOutputs += $groupAssignmentOutput
}
catch {
    $errorMsg = "❌ Exception during group assignment: $($_.Exception.Message)"
    $JsonObject.metadata.errors += $errorMsg
    Write-Host $errorMsg
}

# === STEP 6: Licensing (optional) ===
$licenseModule = "$PSScriptRoot\modules\Assign-License.ps1"
if (Test-Path $licenseModule) {
    try {
        Write-Host "🎫 Assigning licenses..."
        $licenseOutput = & $licenseModule -Json $JsonObject
        $AllOutputs += $licenseOutput
    }
    catch {
        $errorMsg = "❌ Exception during licensing: $($_.Exception.Message)"
        $JsonObject.metadata.errors += $errorMsg
        Write-Host $errorMsg
    }
}

# === STEP 7: Format Final Ticket Note ===
try {
    Write-Host "📝 Formatting ConnectWise ticket note..."
    $ticketNote = & "$PSScriptRoot\modules\Format-TicketNote.ps1" -Json $JsonObject -ModuleOutputs $AllOutputs
    Write-Host "✅ Ticket note formatted."

    return @{
        result  = "success"
        message = $ticketNote
        errors  = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
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
