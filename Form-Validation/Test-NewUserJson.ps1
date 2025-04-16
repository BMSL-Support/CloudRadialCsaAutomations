using namespace System.Net

param(
    [Parameter(Mandatory = $true)]
    [object]$Request  # Input from HTTP request (JSON body)
)

function Update-Placeholders {
    param (
        [Parameter(Mandatory = $true)]
        [string]$JsonInput
    )

    # Replace placeholders starting with '@' for both strings and arrays
    # Replace placeholders in strings with null
    $JsonInput = $JsonInput -replace '"([^"]+)":\s?"@[^"]+"', '"$1": null'

    # Replace placeholders in arrays with empty arrays
    $JsonInput = $JsonInput -replace '"([^"]+)":\s?\[@[^"]+\]', '"$1": []'

    return $JsonInput
}

function Test-NewUserJson {
    param (
        [Parameter(Mandatory = $true)]
        [psobject]$Data
    )

    $errors = @()

    # Root-level fields validation
    if (-not $Data.TenantId) { $errors += "Missing: TenantId" }
    if (-not $Data.TicketId) { $errors += "Missing: TicketId" }

    # AccountDetails
    $acc = $Data.AccountDetails
    if (-not $acc) {
        $errors += "Missing: AccountDetails block"
    } else {
        if (-not $acc.GivenName) { $errors += "Missing: AccountDetails.GivenName" }
        if (-not $acc.Surname) { $errors += "Missing: AccountDetails.Surname" }
        if (-not $acc.UserPrincipalName) {
            $errors += "Missing: AccountDetails.UserPrincipalName"
        } elseif ($acc.UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $errors += "Invalid format: AccountDetails.UserPrincipalName"
        }

        if ($acc.AdditionalAccountDetails) {
            $aad = $acc.AdditionalAccountDetails
            if ($aad.BusinessPhones -and (-not ($aad.BusinessPhones -is [array]))) {
                $errors += "AccountDetails.AdditionalAccountDetails.BusinessPhones must be an array"
            }
        }
    }

    # LicenseTypes
    if (-not $Data.LicenseTypes) {
        $errors += "Missing: LicenseTypes"
    } elseif (-not ($Data.LicenseTypes -is [array])) {
        $errors += "LicenseTypes must be an array"
    }

    # Groups validation
    $groups = $Data.Groups
    if (-not $groups) {
        $errors += "Missing: Groups block"
    } else {
        $mirrored = $groups.MirroredUsers
        if ($mirrored) {
            $hasEmail = [bool]$mirrored.MirroredUserEmail
            $hasGroups = [bool]$mirrored.MirroredUserGroups

            # Validate conditions based on MirroredUserEmail
            if ($hasEmail) {
                if ($groups.Distribution -eq $null -or $groups.SharedMailboxes -eq $null) {
                    $errors += "MirroredUserEmail requires at least one of Distribution or SharedMailboxes."
                }
                if ($groups.Teams) {
                    $errors += "Cannot specify Teams Groups when using MirroredUserEmail."
                }
                if ($groups.Security) {
                    $errors += "Cannot specify Security Groups when using MirroredUserEmail."
                }
            }

            # Validate conditions based on MirroredUserGroups
            if ($hasGroups) {
                if ($groups.Teams -eq $null -or $groups.Security -eq $null) {
                    $errors += "MirroredUserGroups requires at least one of Teams or Security."
                }
                if ($groups.Distribution) {
                    $errors += "Cannot specify Distribution Groups when using MirroredUserGroups."
                }
                if ($groups.SharedMailboxes) {
                    $errors += "Cannot specify Shared Mailboxes when using MirroredUserGroups."
                }
            }
        }

        # Validate group-type arrays
        foreach ($key in @("Software", "Teams", "Security", "Distribution", "SharedMailboxes")) {
            if ($groups.$key -and (-not ($groups.$key -is [array]))) {
                $errors += "Groups.$key must be an array"
            }
        }
    }

    return $errors
}

# ----------- MAIN EXECUTION -----------
try {
    # Load JSON input (file or raw JSON)
    if (Test-Path $JsonInput) {
        $raw = Get-Content -Path $JsonInput -Raw
    } else {
        $Request = $JsonInput
    }

    # Pre-process and replace placeholders
    $processedJson = Update-Placeholders -JsonInput $raw

    # Convert JSON to object
    $json = $processedJson | ConvertFrom-Json -ErrorAction Stop

    $validationErrors = Test-NewUserJson -Data $json

    if ($validationErrors.Count -eq 0) {
        Write-Host "✅ JSON is valid."
        exit 0
    } else {
        Write-Host "❌ JSON is invalid:"
        $validationErrors | ForEach-Object { Write-Host " - $_" }
        exit 1
    }
}
catch {
    Write-Host "❌ Failed to parse JSON: $_"
    exit 2
}
