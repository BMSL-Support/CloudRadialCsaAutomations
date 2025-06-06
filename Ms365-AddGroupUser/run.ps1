
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

    JSON Structure
    {
        "UserPrincipalName": "user@domain.com",
        "GroupNames": ["Group Name 1", "Group Name 2"],
        "TenantId": "12345678-1234-1234-123456789012",
        "TicketId": "123456",
        "SecurityKey": "optional"
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

$UserPrincipalName = $Request.Body.AccountDetails.UserPrincipalName
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

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
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

$Groups = $Request.Body.Groups

# Handle MirroredUsers
$MirroredUserEmail = $Groups.MirroredUsers.MirroredUserEmail
$MirroredUserGroups = $Groups.MirroredUsers.MirroredUserGroups

if ($MirroredUserEmail -match "^@") {
    $MirroredUserEmail = $null
}

if ($MirroredUserGroups -match "^@") {
    $MirroredUserGroups = $null
}

if ($MirroredUserGroups) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

    $MirroredUserObject = Get-MgUser -Filter "userPrincipalName eq '$MirroredUserGroups'"

    if ($MirroredUserObject) {
        $TeamsGroups = Get-MgUserMemberOf -UserId $MirroredUserObject.Id | Where-Object { $_.ODataType -eq '#microsoft.graph.group' -and $_.GroupTypes -contains 'Unified' }
        $SecurityGroups = Get-MgUserMemberOf -UserId $MirroredUserObject.Id | Where-Object { $_.ODataType -eq '#microsoft.graph.group' -and $_.GroupTypes -notcontains 'Unified' }

        
        Write-Output "Teams Groups: $TeamsGroups"
        Write-Output "Security Groups: $SecurityGroups"


        $addedTeamsGroups = @()
        $addedSecurityGroups = @()

        foreach ($Group in $TeamsGroups) {
            if ($Group.Id -ne "") {
                # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserPrincipalName
                $addedTeamsGroups += $Group.DisplayName
            }
        }

        foreach ($Group in $SecurityGroups) {
            if ($Group.Id -ne "") {
                # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserPrincipalName
                $addedSecurityGroups += $Group.DisplayName
            }
        }

        $message += "$UserPrincipalName was added to the following Teams based on ${MirroredUserGroups}:`n" + ($addedTeamsGroups -join "`n") + "`n`n"
        $message += "$UserPrincipalName was added to the following Security Groups based on ${MirroredUserGroups}:`n" + ($addedSecurityGroups -join "`n") + "`n`n"
    }
}

if ($MirroredUserEmail) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

    $MirroredUserObject = Get-MgUser -Filter "userPrincipalName eq '$MirroredUserEmail'"

    if ($MirroredUserObject) {
        $DistributionGroups = Get-MgUserMemberOf -UserId $MirroredUserObject.Id | Where-Object { $_.ODataType -eq '#microsoft.graph.group' -and $_.MailEnabled -eq $true }
        $SharedMailboxes = Get-MgUserMemberOf -UserId $MirroredUserObject.Id | Where-Object { $_.ODataType -eq '#microsoft.graph.group' -and $_.MailEnabled -eq $false }

        
        Write-Output "Distribution Groups: $DistributionGroups"
        Write-Output "Shared Mailboxes: $SharedMailboxes"

        $message += "The following actions will need to be completed manually in the Exchange Online Admin Centre -`n`n"
        $message += "$UserPrincipalName will need to be added to the following Exchange Groups based on ${MirroredUserEmail}:`n" + ($DistributionGroups.DisplayName -join "`n") + "`n`n"
        $message += "$UserPrincipalName will need to be given access to the following Shared Mailboxes based on  ${MirroredUserEmail}:`n" + ($SharedMailboxes.DisplayName -join "`n") + "`n`n"
    }
}

# Handle Software groups
$SoftwareGroups = $Groups.Software | Where-Object { $_ -notmatch "^@" -and $_ -ne "No groups available at this time." }

if ($SoftwareGroups.Count -eq 0) {
    $message += "No software groups were defined in the request.`n`n"
}
else {
    $addedSoftwareGroups = @()
    foreach ($Group in $SoftwareGroups) {
        if ($Group.Id -ne "") {
            # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserPrincipalName
            $addedSoftwareGroups += $Group
        }
    }
    $message += "$UserPrincipalName was added to the following software groups:`n" + ($addedSoftwareGroups -join "`n") + "`n`n"
}

# Handle Teams groups
$TeamsGroups = $Groups.Teams | Where-Object { $_ -notmatch "^@" -and $_ -ne "No groups available at this time." }

if ($TeamsGroups.Count -eq 0) {
    $message += "No Teams were defined in the request.`n`n"
}
else {
    $addedTeamsGroups = @()
    foreach ($Group in $TeamsGroups) {
        if ($Group.Id -ne "") {
            # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserPrincipalName
            $addedTeamsGroups += $Group
        }
    }
    $message += "$UserPrincipalName was added to the following Teams:`n" + ($addedTeamsGroups -join "`n") + "`n`n"
}

# Handle Security groups
$SecurityGroups = $Groups.Security | Where-Object { $_ -notmatch "^@" -and $_ -ne "No groups available at this time." }

if ($SecurityGroups.Count -eq 0) {
    $message += "No security groups were defined in the request.`n`n"
}
else {
    $addedSecurityGroups = @()
    foreach ($Group in $SecurityGroups) {
        if ($Group.Id -ne "") {
            # New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserPrincipalName
            $addedSecurityGroups += $Group
        }
    }
    $message += "$UserPrincipalName was added to the following security groups:`n" + ($addedSecurityGroups -join "`n") + "`n`n"
}


# Handle Distribution groups
$DistributionGroups = $Groups.Distribution | Where-Object { $_.MailEnabled -eq $true -and $_.GroupTypes -notcontains 'Unified' -and $_.Mail -notlike "*.onmicrosoft.com" }

if ($DistributionGroups.Count -eq 0) {
    $message += "No distribution groups were defined in the request.`n`n"
}
else {
    $addedDistributionGroups = @()
    foreach ($Group in $DistributionGroups) {
        $GroupObject = Get-MgGroup -Filter "displayName eq '$Group'"
        if ($GroupObject.Id -ne "") {
            # New-MgGroupMember -GroupId $GroupObject.Id -DirectoryObjectId $UserPrincipalName
            $addedDistributionGroups += $Group
        }
    }
    $message += "$UserPrincipalName was added to the following distribution groups:`n" + ($addedDistributionGroups -join "`n") + "`n`n"
}

$DistributionGroups = $Groups.Distribution

if ($DistributionGroups.Count -gt 0) {
    $message += "The following actions will need to be completed manually in the Exchange Online Admin Centre -`n`n"
    $message += "$UserPrincipalName will need to be added to the following Exchange Groups:`n" + ($DistributionGroups -join "`n") + "`n`n"
}


# Handle Shared Mailboxes
$SharedMailboxes = $Groups.SharedMailboxes

if ($SharedMailboxes.Count -eq 0) {
    $message += "No shared mailboxes were defined in the request.`n`n"
}
else {
    $addedSharedMailboxes = @()
    foreach ($Mailbox in $SharedMailboxes) {
        $MailboxObject = Get-MgUser -Filter "displayName eq '$Mailbox'"
        if ($MailboxObject.Id -ne "") {
            # New-MgUserMailboxSetting -UserId $MailboxObject.Id -AccessRights "FullAccess" -UserPrincipalName $UserPrincipalName
            $addedSharedMailboxes += $Mailbox
        }
    }
    $message += "$UserPrincipalName was given access to the following shared mailboxes:`n" + ($addedSharedMailboxes -join "`n") + "`n`n"
}

$SharedMailboxes = $Groups.SharedMailboxes

if ($SharedMailboxes.Count -gt 0) {
    $message += "$UserPrincipalName will need to be given access to the following Shared Mailboxes:`n" + ($SharedMailboxes -join "`n") + "`n`n"
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
