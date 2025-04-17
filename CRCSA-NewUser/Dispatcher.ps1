using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# Load modules
. "$PSScriptRoot\modules\Create-NewUser.ps1"
. "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1"
. "$PSScriptRoot\modules\Add-UserGroups.ps1"
. "$PSScriptRoot\modules\utils.ps1"

# Utilities
function Update-Placeholders {
    param (
        [string]$JsonInput
    )
    $JsonInput = $JsonInput -replace '"([^\"]+)":\s?"@[^\"]+"', '"$1": null'
    $JsonInput = $JsonInput -replace '"([^\"]+)":\s?\[@[^\"]+\]', '"$1": []'
    return $JsonInput
}

function Test-NewUserJson {
    param (
        [psobject]$Data
    )

    $errors = @()

    if (-not $Data.TenantId)    { $errors += "Missing: TenantId" }
    if (-not $Data.TicketId)    { $errors += "Missing: TicketId" }

    $acc = $Data.AccountDetails
    if (-not $acc) {
        $errors += "Missing: AccountDetails block"
    } else {
        if (-not $acc.GivenName)         { $errors += "Missing: AccountDetails.GivenName" }
        if (-not $acc.Surname)           { $errors += "Missing: AccountDetails.Surname" }
        if (-not $acc.UserPrincipalName -or $acc.UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $errors += "Missing or invalid: AccountDetails.UserPrincipalName"
        }
    }

    if (-not $Data.Groups) {
        $errors += "Missing: Groups block"
    } else {
        foreach ($required in @("Teams", "Security", "Distribution", "SharedMailboxes", "Software", "MirroredUsers")) {
            if (-not $Data.Groups.PSObject.Properties[$required]) {
                $errors += "Missing: Groups.$required"
            }
        }
    }

    # LicenseTypes is optional but must be an array if present
    if ($Data.PSObject.Properties.Match('LicenseTypes')) {
        if ($Data.LicenseTypes -and -not ($Data.LicenseTypes -is [array])) {
            $errors += "Invalid format: LicenseTypes must be an array"
        }
    }

    return $errors
}

# Main Logic
try {
    $raw = $Request.Body | ConvertTo-Json -Depth 10
    $rawClean = Update-Placeholders -JsonInput $raw
    $json = $rawClean | ConvertFrom-Json -ErrorAction Stop

    # Use shared metadata initializer
    Initialize-Metadata -Json $json

    $validationErrors = Test-NewUserJson -Data $json

    if ($validationErrors.Count -eq 0) {
        Write-Host "✅ JSON is valid. Proceeding..."

        # Handle mirrored user groups if defined
        if ($json.Groups.MirroredUsers) {
            Write-Host "➡ Fetching mirrored group memberships..."
            $mirroredGroups = Get-MirroredUserGroupMemberships -MirroredUsers $json.Groups.MirroredUsers -TenantId $json.TenantId

            foreach ($groupType in @("Teams", "Security", "Distribution", "SharedMailboxes")) {
                if (-not $json.Groups.$groupType -or $json.Groups.$groupType.Count -eq 0) {
                    $json.Groups.$groupType = $mirroredGroups[$groupType]
                }
            }
        }

        # Create user
        $result = Invoke-CreateNewUser -Json $json
        $userUpn = $json.AccountDetails.UserPrincipalName

        $dispatcherMessage = $result.Message
        $dispatcherErrors = @()

        if (-not $result.Success) {
            $json.metadata.status.userCreation = "failed"
            $json.metadata.errors += $result.Message

            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body = @{
                    message  = "User creation failed"
                    errors   = @($result.Message)
                    metadata = $json.metadata
                }
                ContentType = "application/json"
            })
            return
        }

        # Add to groups if provided
        if ($json.Groups) {
            $groupResult = Add-UserGroups -UserPrincipalName $userUpn -Groups $json.Groups -TenantId $json.TenantId -TicketId $json.TicketId
            $dispatcherMessage += "`n`n" + $groupResult.Message
            $dispatcherErrors += $groupResult.Errors
        }

        # Success response
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode  = [HttpStatusCode]::OK
            Body        = @{
                message  = $dispatcherMessage
                upn      = $userUpn
                metadata = $json.metadata
                result   = $result.Result
                errors   = $dispatcherErrors
            }
            ContentType = "application/json"
        })
    }
    else {
        Write-Host "❌ JSON validation failed."
        $json.metadata.status.userCreation = "failed"
        $json.metadata.errors += $validationErrors

        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = @{
                message  = "Validation failed"
                errors   = $validationErrors
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
            error   = $_.Exception.Message
        }
        ContentType = "application/json"
    })
}
