<# 

.SYNOPSIS
    This function is used to add a user to multiple distribution groups in Microsoft 365.

.DESCRIPTION
    This function is used to add a user to multiple distribution groups in Microsoft 365.
    
    The function requires the following environment variables to be set:
    - Ms365_AuthAppId: Application Id of the service principal
    - Ms365_AuthSecretId: Secret Id of the service principal
    - Ms365_TenantId: Tenant Id of the Microsoft 365 tenant
    
    The function requires the following modules to be installed:
    - Microsoft.Graph

.INPUTS
    JSON Structure:
    {
        "TenantId": "@CompanyTenantId",
        "TicketId": "@TicketId",
        "AccountDetails": {
            "GivenName": "@NUsersFirstName",
            "Surname": "@NUsersLastName",
            "UserPrincipalName": "@NUsersEmail",
            "AdditionalAccountDetails": {
                "JobTitle": "@NUsersJobTitle",
                "City": "@NUsersAddJobTitle",
                "Department": "@NUsersDept",
                "BusinessPhones": "@NUsersOfficePhone",
                "MobilePhone": "@NUsersMobilePhone"
            }
        },
        "LicenseTypes": ["@LicenseType"],
        "Groups": {
            "MirroredUsers": {
                "MirroredUserEmail": "mirroreduser@domain.com",
                "MirroredUserGroups": "mirroreduser@domain.com"
            },
            "Software": ["@NUSoftwareGroups"],
            "Teams": ["@NUTeamNames"],
            "Security": ["@NUSecGroupNames"],
            "Distribution": ["@NUDistributionGroups"],
            "SharedMailboxes": ["@NUSharedMailboxes"]
        }
    }

.OUTPUTS
    JSON response with the following fields:
    - Message: Descriptive string of result
    - TicketId: TicketId passed in Parameters
    - ResultCode: 200 for success, 500 for failure
    - ResultStatus: "Success" or "Failure"
    - Internal: Boolean value indicating if the operation is internal

#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Add User to Groups function triggered."

$resultCode = 200
$message = ""

# Parse the JSON structure
$UserPrincipalName = $Request.Body.AccountDetails.UserPrincipalName
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey
$MirroredUserEmail = $Request.Body.Groups.MirroredUsers.MirroredUserEmail
$MirroredUserGroups = $Request.Body.Groups.MirroredUsers.MirroredUserGroups

# Treat values starting with '@' as null
if ($UserPrincipalName -like "@*") { $UserPrincipalName = $null }
if ($TenantId -like "@*") { $TenantId = $null }
if ($TicketId -like "@*") { $TicketId = $null }
if ($MirroredUserEmail -like "@*") { $MirroredUserEmail = $null }
if ($MirroredUserGroups -like "@*") { $MirroredUserGroups = $null }

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break
}

if (-Not $UserPrincipalName) {
    $message = "UserPrincipalName cannot be blank."
    $resultCode = 500
} else {
    $UserPrincipalName = $UserPrincipalName.Trim()
}

if (-Not $TenantId) {
    $TenantId = $env:Ms365_TenantId
} else {
    $TenantId = $TenantId.Trim()
}

if (-Not $TicketId) {
    $TicketId = ""
}

Write-Host "User Principal Name: $UserPrincipalName"
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

if ($resultCode -eq 200) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

    $UserObject = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'"

    if (-Not $UserObject) {
        $message = "Request failed. User `"$UserPrincipalName`" could not be found."
        $resultCode = 500
    }

    $addedGroups = @()
    $manualGroups = @()

    # Handle MirroredUserEmail and MirroredUserGroups
    if ($MirroredUserEmail) {
        $MirroredUserObject = Get-MgUser -Filter "userPrincipalName eq '$MirroredUserEmail'"
        if ($MirroredUserObject) {
            $MirroredGroups = Get-MgUserMemberGroup -UserId $MirroredUserObject.Id -SecurityEnabledOnly $false
            foreach ($GroupId in $MirroredGroups) {
                $GroupObject = Get-MgGroup -GroupId $GroupId
                if ($GroupObject.mailEnabled -eq $true -and $GroupObject.groupTypes -notcontains 'Unified') {
                    $manualGroups += $GroupObject.DisplayName
                } else {
                    New-MgGroupMember -GroupId $GroupObject.Id -DirectoryObjectId $UserObject.Id
                    $addedGroups += $GroupObject.DisplayName
                }
            }
        } else {
            $message += "Mirrored user `"$MirroredUserEmail`" could not be found.`n"
        }
    }

    if ($MirroredUserGroups) {
        $GroupObject = Get-MgGroup -Filter "mail eq '$MirroredUserGroups'"
        if ($GroupObject) {
            if ($GroupObject.mailEnabled -eq $true -and $GroupObject.groupTypes -notcontains 'Unified') {
                $manualGroups += $GroupObject.DisplayName
            } else {
                New-MgGroupMember -GroupId $GroupObject.Id -DirectoryObjectId $UserObject.Id
                $addedGroups += $GroupObject.DisplayName
            }
        } else {
            $message += "Mirrored group `"$MirroredUserGroups`" could not be found.`n"
        }
    }

    if ($addedGroups.Count -gt 0) {
        $message += "The following groups were successfully added:`n`n" + ($addedGroups -join "`n")
    }

    if ($manualGroups.Count -gt 0) {
        $message += "`nThe following groups need to be added manually in the Exchange Online Management portal:`n`n" + ($manualGroups -join "`n")
    }
}

$body = @{
    Message = $message
    TicketId = $TicketId
    ResultCode = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
    Internal = $true
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::InternalServerError }
    Body = $body
    ContentType = "application/json"
})
