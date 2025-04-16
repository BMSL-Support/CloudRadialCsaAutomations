using namespace System.Net

param(
    [Parameter(Mandatory = $true)]
    [object]$Request  # Input from HTTP request (JSON body)
)

# Function to update placeholders in the JSON
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

# Function to validate the JSON structure
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
    # Get the raw JSON input from the request
    $JsonInput = $Request.Body

    # Log the raw JSON input to diagnose issues
    Write-Host "Raw Input JSON: $JsonInput"

    # Validate if the input is not null or empty
    if (-not $JsonInput) {
        throw "❌ Error: The provided JSON input is null or empty."
    }

    # Pre-process and replace placeholders in the JSON
    $processedJson = Update-Placeholders -JsonInput $JsonInput

    # Log processed JSON to ensure it's valid before parsing
    Write-Host "Processed JSON: $processedJson"

    # Convert the processed JSON string to a PowerShell object
    $json = $processedJson | ConvertFrom-Json -ErrorAction Stop

    # Validate the JSON object
    $validationErrors = Test-NewUserJson -Data $json

    # Add metadata to the JSON structure for tracking progress
    $metadata = @{
        createdTimestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        status = @{
            userCreation   = "pending"
            groupAssignment = "pending"
            licensing       = "pending"
        }
        errors = $validationErrors
    }

    # Add the metadata to the JSON object
    $json | Add-Member -MemberType NoteProperty -Name "metadata" -Value $metadata

    # Return the updated JSON object (with metadata) as the response
    if ($validationErrors.Count -eq 0) {
        Write-Host "✅ JSON is valid."
        return $json | ConvertTo-Json -Depth 3
    } else {
        Write-Host "❌ JSON is invalid:"
        $validationErrors | ForEach-Object { Write-Host " - $_" }
        return @{
            status  = "error"
            message = "Invalid JSON"
            errors  = $validationErrors
        } | ConvertTo-Json
    }
}
catch {
    Write-Host "❌ Failed to parse JSON: $_"
    return @{
        status  = "error"
        message = "Failed to parse JSON"
        error   = $_.Exception.Message
    } | ConvertTo-Json
}
