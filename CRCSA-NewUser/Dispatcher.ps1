using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request
)

# Load modules
. "$PSScriptRoot\modules\Create-NewUser.ps1"
. "$PSScriptRoot\modules\Get-MirroredUserGroupMemberships.ps1"
. "$PSScriptRoot\modules\Add-UserGroups.ps1"
. "$PSScriptRoot\modules\Update-ConnectWiseTicketNote.ps1"
. "$PSScriptRoot\modules\utils.ps1"

# Main Logic
try {
    # Parse raw JSON and sanitize placeholders
    $rawJson = $Request.Body | Out-String
    $cleanJson = Clear-Placeholders -JsonString $rawJson

    try {
        $json = $cleanJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        return @{
            message = "Dispatcher failed"
            error   = "Conversion from JSON failed with error: $($_.Exception.Message)"
        }
    }

    # Clean up any remaining @placeholders post-parsing
    $json = Update-Placeholders -JsonObject $json

    # Initialize metadata
    Initialize-Metadata -Json $json -Step "userCreation"

    # Validate JSON structure
    $validationErrors = Test-NewUserJson -Data $json

    if ($validationErrors.Count -eq 0) {
        Write-Host "✅ JSON is valid. Proceeding..."

        # Handle mirrored group logic
        $mirroredInfo = $json.Groups.MirroredUsers
        if ($mirroredInfo.MirroredUserEmail -or $mirroredInfo.MirroredUserGroups) {
            Write-Host "➡ Fetching mirrored group memberships..."

            $mirroredGroups = Get-MirroredUserGroupMemberships `
                -MirroredUserEmail $mirroredInfo.MirroredUserEmail `
                -MirroredUserGroups $mirroredInfo.MirroredUserGroups `
                -TenantId $json.TenantId

            foreach ($groupType in @("Teams", "Security", "Distribution", "SharedMailboxes")) {
                if (-not $json.Groups.$groupType -or $json.Groups.$groupType.Count -eq 0) {
                    $json.Groups.$groupType = $mirroredGroups[$groupType]
                }
            }
        }

        # Create user
        $result = Invoke-CreateNewUser -Json $json
        $userUpn = $json.AccountDetails.UserPrincipalName

        $dispatcherMessage = ""
        $dispatcherErrors = @()

        if ($result.ResultStatus -ne "Success") {
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

        # Assign groups
        if ($json.Groups) {
            $groupResult = Add-UserGroups -Json $json
            $dispatcherMessage += "`n`n" + $groupResult.Message
            $dispatcherErrors += $groupResult.Errors
        }

        # Update ConnectWise ticket
        $ticketNoteResponse = Update-ConnectWiseTicketNote -TicketId $json.TicketId -Message $dispatcherMessage -Internal $true

        if ($ticketNoteResponse.Status -ne "Success") {
            $dispatcherErrors += "Failed to update ConnectWise ticket note: $($ticketNoteResponse.Message)"
            $dispatcherMessage += "`n❌ Failed to update ConnectWise ticket note: $($ticketNoteResponse.Message)"
        }

        # Final response
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
