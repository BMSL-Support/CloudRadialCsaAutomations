<# 

.SYNOPSIS
    
    This function is used to add a user to multiple distribution groups in Microsoft 365.

.DESCRIPTION
             
    This function is used to add a user to multiple distribution groups in Microsoft 365.
    
    The function requires the following environment variables to be set:
        
    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
        
    The function requires the following modules to be installed:
        
    Microsoft.Graph

.INPUTS

    UserPrincipalName - user principal name that exists in the tenant
    GroupNames - array of group names that exist in the tenant
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    TicketId - optional - string value of the ticket id used for transaction tracking
    SecurityKey - Optional, use this as an additional step to secure the function
    LikeUserGroup - optional - email address or UPN of a user to list similar groups
    LikeUserEmail - optional - email address or UPN of a user to list similar email groups

    JSON Structure

    {
        "UserPrincipalName": "user@domain.com",
        "GroupNames": ["Group Name 1", "Group Name 2"],
        "TenantId": "12345678-1234-1234-123456789012",
        "TicketId": "123456",
        "SecurityKey": "optional",
        "LikeUserGroup": "likeuser@domain.com",
        "LikeUserEmail": "likeuser@domain.com"
    }

.OUTPUTS 

    JSON response with the following fields:

    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"
    Internal - Boolean value indicating if the operation is internal

#>
using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Add User to Groups function triggered."

$resultCode = 200
$message = ""

$UserPrincipalName = $Request.Body.UserPrincipalName
$GroupNames = $Request.Body.GroupNames
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey
$LikeUserGroup = $Request.Body.LikeUserGroup
$LikeUserEmail = $Request.Body.LikeUserEmail

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $UserPrincipalName) {
    $message = "UserPrincipalName cannot be blank."
    $resultCode = 500
}
else {
    $UserPrincipalName = $UserPrincipalName.Trim()
}

if (-Not $GroupNames -or $GroupNames.Count -eq 0 -or $GroupNames -eq "No groups available at this time.") {
    $message = "No groups specified on the form."
    $resultCode = 500
}

if (-Not $TenantId) {
    $TenantId = $env:Ms365_TenantId
}
else {
    $TenantId = $TenantId.Trim()
}

if (-Not $TicketId) {
    $TicketId = ""
}

Write-Host "User Principal Name: $UserPrincipalName"
Write-Host "Group Names: $($GroupNames -join ', ')"
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

if ($resultCode -Eq 200)
{
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

    $UserObject = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'"

    Write-Host $UserObject.userPrincipalName
    Write-Host $UserObject.Id

    if (-Not $UserObject) {
        $message = "Request failed. User `"$UserPrincipalName`" could not be found."
        $resultCode = 500
    }

    $addedGroups = @()
    $likeUserGroups = @()
    $likeUserEmails = @()

    if ($LikeUserGroup) {
        $likeUserObject = Get-MgUser -Filter "userPrincipalName eq '$LikeUserGroup'"
        if ($likeUserObject) {
            $userGroups = Get-MgUserMemberOf -UserId $likeUserObject.Id | Where-Object { $_.DisplayName -like "Security -*" -or $_.DisplayName -like "Data -*" -or $_.DisplayName -like "SP Data -*" }
            foreach ($group in $userGroups) {
                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $UserObject.Id
                $likeUserGroups += $group.DisplayName
            }
        } else {
            $message += "Request failed. LikeUserGroup `"$LikeUserGroup`" could not be found.`n"
            $resultCode = 500
        }
    }

    if ($LikeUserEmail) {
        $likeUserObject = Get-MgUser -Filter "userPrincipalName eq '$LikeUserEmail'"
        if ($likeUserObject) {
            $distributionGroups = Get-MgGroup -Filter "mailEnabled eq true" -All
            $filteredGroups = $distributionGroups | Where-Object { $_.groupTypes -notcontains 'Unified' -and $_.Mail -notlike "*.onmicrosoft.com" }
            foreach ($group in $filteredGroups) {
                $likeUserEmails += $group.DisplayName
            }
        } else {
            $message += "Request failed. LikeUserEmail `"$LikeUserEmail`" could not be found.`n"
            $resultCode = 500
        }
    }

    foreach ($GroupName in $GroupNames) {
        $GroupObject = Get-MgGroup -Filter "displayName eq '$GroupName'"

        Write-Host $GroupObject.DisplayName
        Write-Host $GroupObject.Id

        if (-Not $GroupObject) {
            $message += "Request failed. Group `"$GroupName`" could not be found to add user `"$UserPrincipalName`" to.`n"
            $resultCode = 500
            continue
        }

        if ($GroupObject.mailEnabled -eq $true -and $GroupObject.groupTypes -notcontains 'Unified') {
            $message += "Group `"$GroupName`" is a mail distribution group or shared mailbox and needs to be added manually in the Exchange Online Management portal.`n"
            continue
        }

        $GroupMembers = Get-MgGroupMember -GroupId $GroupObject.Id

        if ($GroupMembers.Id -Contains $UserObject.Id) {
            $message += "Request failed. User `"$UserPrincipalName`" is already a member of group `"$GroupName`".`n"
            $resultCode = 500
            continue
        } 

        if ($resultCode -Eq 200) {
            New-MgGroupMember -GroupId $GroupObject.Id -DirectoryObjectId $UserObject.Id
            $addedGroups += $GroupName
        }
    }

    if ($addedGroups.Count -gt 0) {
        $message = "The following Teams and Security Groups were selected and successfully added:`n`n" + ($addedGroups -join "`n")
    }

    if ($likeUserGroups.Count -gt 0) {
        $message += "`nThe following groups were added based on mirroring the user `"$LikeUserGroup`":`n`n" + ($likeUserGroups -join "`n")
    }

    if ($likeUserEmails.Count -gt 0) {
        $message += "`nThe following are mail distribution group or shared mailbox mirrored from `"$LikeUserEmail`" need to be added manually in the Exchange Online Management portal:`n`n" + ($likeUserEmails -join "`n")
    }
}

$body = @{
    Message = $message
    TicketId = $TicketId
    ResultCode = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
    Internal     = $true
} 

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})
