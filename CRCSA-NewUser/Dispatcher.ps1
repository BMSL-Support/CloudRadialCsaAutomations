using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# Read JSON body from HTTP request
$rawJson = ($Request.Body | Out-String).Trim()

# Load utilities and modules
. "$PSScriptRoot\modules\utils.ps1"

# Clean up placeholder values (e.g., "@User", [@Groups])
try {
    Write-Host "🧼 Cleaning placeholders..."
    $CleanedJson = Clear-Placeholders -JsonString $JsonInput
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

# Initialize metadata
try {
    Initialize-Metadata -Json $JsonObject
}
catch {
    Write-Host "⚠ Failed to initialize metadata: $($_.Exception.Message)"
}

# Store all output logs from modules
$AllOutputs = @()

# === STEP 1: Validate JSON ===
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

# === STEP 2: Group Mirroring ===
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

# === STEP 3: Create User ===
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

# === STEP 4: Add to Groups ===
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

# === STEP 5: Licensing (if module exists) ===
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

# === FINAL STEP: Format ticket note ===
try {
    Write-Host "📝 Formatting ConnectWise ticket note..."
    $ticketNote = & "$PSScriptRoot\modules\Format-TicketNote.ps1" -Json $JsonObject -ModuleOutputs $AllOutputs
    Write-Host "✅ Ticket note formatted."

    return @{
        result = "success"
        message = $ticketNote
        errors = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
catch {
    $errorMsg = "❌ Exception formatting ticket note: $($_.Exception.Message)"
    Write-Host $errorMsg
    return @{
        error = $_.Exception.Message
        message = "Dispatcher failed during final formatting"
        errors = $JsonObject.metadata.errors
    } | ConvertTo-Json -Depth 10
}
