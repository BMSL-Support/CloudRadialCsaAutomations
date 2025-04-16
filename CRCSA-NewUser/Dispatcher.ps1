using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# Load and import the Create-NewUser module
. "$PSScriptRoot\modules\Create-NewUser.ps1"
. "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1"

# Utilities
function Update-Placeholders {
    param (
        [string]$JsonInput
    )

    # Replace placeholders with null or empty arrays
    $JsonInput = $JsonInput -replace '"([^"]+)":\s?"@[^"]+"', '"$1": null'
    $JsonInput = $JsonInput -replace '"([^"]+)":\s?\[@[^"]+\]', '"$1": []'
    return $JsonInput
}

function Test-NewUserJson {
    param (
        [psobject]$Data
    )

    $errors = @()

    if (-not $Data.TenantId) { $errors += "Missing: TenantId" }
    if (-not $Data.TicketId) { $errors += "Missing: TicketId" }

    $acc = $Data.AccountDetails
    if (-not $acc) {
        $errors += "Missing: AccountDetails block"
    } else {
        if (-not $acc.GivenName) { $errors += "Missing: AccountDetails.GivenName" }
        if (-not $acc.Surname) { $errors += "Missing: AccountDetails.Surname" }
        if (-not $acc.UserPrincipalName -or $acc.UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $errors += "Missing or invalid: AccountDetails.UserPrincipalName"
        }
    }

    if (-not $Data.LicenseTypes -or -not ($Data.LicenseTypes -is [array])) {
        $errors += "Missing or invalid: LicenseTypes (must be an array)"
    }

    return $errors
}

# Main Logic
try {
    $raw = $Request.Body | ConvertTo-Json -Depth 10
    $rawClean = Update-Placeholders -JsonInput $raw
    $json = $rawClean | ConvertFrom-Json -ErrorAction Stop

    # Add metadata
    if (-not $json.metadata) {
        $json | Add-Member -MemberType NoteProperty -Name "metadata" -Value @{
            createdTimestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            status = @{
                userCreation = "pending"
                groupAssignment = "pending"
                licensing = "pending"
            }
            errors = @()
        }
    }

    $validationErrors = Test-NewUserJson -Data $json

    if ($validationErrors.Count -eq 0) {
        Write-Host "✅ JSON is valid. Proceeding with user creation..."

    # ✅ Handle Mirrored Users (augment JSON with inherited groups)
    if ($json.Groups.MirroredUsers) {
        Write-Host "➡ Fetching mirrored group memberships..."
        $mirroredGroups = Get-MirroredUserGroupMemberships -MirroredUsers $json.Groups.MirroredUsers

        # Merge groups only if not already defined manually
        if (-not $json.Groups.Teams -or $json.Groups.Teams.Count -eq 0) {
            $json.Groups.Teams = $mirroredGroups.Teams
        }
        if (-not $json.Groups.Security -or $json.Groups.Security.Count -eq 0) {
            $json.Groups.Security = $mirroredGroups.Security
        }
        if (-not $json.Groups.Distribution -or $json.Groups.Distribution.Count -eq 0) {
            $json.Groups.Distribution = $mirroredGroups.Distribution
        }
        if (-not $json.Groups.SharedMailboxes -or $json.Groups.SharedMailboxes.Count -eq 0) {
            $json.Groups.SharedMailboxes = $mirroredGroups.SharedMailboxes
        }
    }

        # Call the user creation script and capture result
        $result = Invoke-CreateNewUser -Json $json

        # You could add calls here for groups/licensing too
        # $groupResult = Invoke-AssignGroups -Json $result.Json
        # $licenseResult = Invoke-AssignLicenses -Json $result.Json

        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = [HttpStatusCode]::OK
            Body        = $result
            ContentType = "application/json"
        })
    } else {
        Write-Host "❌ JSON is invalid."
        $json.metadata.status.userCreation = "failed"
        $json.metadata.errors += $validationErrors

        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = @{
                message = "Validation failed"
                errors = $validationErrors
                metadata = $json.metadata
            }
            ContentType = "application/json"
        })
    }
}
catch {
    Write-Host "❌ Exception during dispatch: $_"

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = @{
            message = "Dispatcher failed"
            error   = "$_"
        }
        ContentType = "application/json"
    })
}
