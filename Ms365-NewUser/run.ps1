using namespace System.Net

param (
    [Parameter(Mandatory = $true)]
    [object]$Request,
    [object]$TriggerMetadata
)

# Default output values
$resultCode = 200
$resultStatus = "Success"
$message = ""
$UserPrincipalName = ""

# Pull JSON body
$json = $Request.Body

# Setup metadata if not present
if (-not $json.metadata) {
    $json | Add-Member -MemberType NoteProperty -Name "metadata" -Value @{
        createdTimestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        status = @{
            userCreation = "pending"
        }
        errors = @()
    }
}

$json.metadata.status.userCreation = "in_progress"

# Extract inputs
$TenantId = $json.TenantId
$TicketId = $json.TicketId
$AccountDetails = $json.AccountDetails
$GivenName = $AccountDetails.GivenName
$Surname = $AccountDetails.Surname
$UserPrincipalName = $AccountDetails.UserPrincipalName
$aad = $AccountDetails.AdditionalAccountDetails
$JobTitle = $aad.JobTitle
$City = $aad.City
$Department = $aad.Department
$BusinessPhones = $aad.BusinessPhones
$MobilePhone = $aad.MobilePhone
$LicenseTypes = $json.LicenseTypes
$SecurityKey = $env:SecurityKey

# Function to generate a random password
function New-RandomPassword {
    param ([int]$length = 16)
    $characters = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!#%&*()'
    -join ((1..$length) | ForEach-Object { $characters[(Get-Random -Minimum 0 -Maximum $characters.Length)] })
}

$password = New-RandomPassword

try {
    if ($SecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
        throw "Invalid security key"
    }

    if (-not $UserPrincipalName) {
        throw "UserPrincipalName cannot be blank."
    } else {
        $UserPrincipalName = $UserPrincipalName.Trim()
    }

    if (-not $TenantId) {
        $TenantId = $env:Ms365_TenantId
    } else {
        $TenantId = $TenantId.Trim()
    }

    if (-not $TicketId) {
        $TicketId = ""
    }

    # Connect to Microsoft Graph
    $secureSecret = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secureSecret)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId -NoWelcome

    # Display name & mail nickname
    $DisplayName = "$GivenName $Surname"
    $mailNickname = $UserPrincipalName.Split("@")[0]

    $newUserParams = @{
        UserPrincipalName = $UserPrincipalName
        DisplayName       = $DisplayName
        GivenName         = $GivenName
        Surname           = $Surname
        MailNickname      = $mailNickname
        PasswordProfile   = @{ Password = $password; ForceChangePasswordNextSignIn = $true }
        UsageLocation     = "GB"
        AccountEnabled    = $true
    }

    if ($JobTitle)       { $newUserParams.JobTitle = $JobTitle }
    if ($City)           { $newUserParams.City = $City }
    if ($Department)     { $newUserParams.Department = $Department }
    if ($BusinessPhones) { $newUserParams.BusinessPhones = $BusinessPhones }
    if ($MobilePhone)    { $newUserParams.MobilePhone = $MobilePhone }

    # Create the user
    $newUser = New-MgUser @newUserParams

    if ($newUser) {
        $message = "✅ New user $DisplayName created successfully.`r`nUsername: $UserPrincipalName `r`nPassword: $password"
        $UserPrincipalName = $newUser.UserPrincipalName
        $json.metadata.status.userCreation = "completed"
    } else {
        throw "User creation failed."
    }
}
catch {
    $resultCode = 500
    $resultStatus = "Failure"
    $message = "❌ Error: $_"
    $json.metadata.status.userCreation = "failed"
    $json.metadata.errors += "User creation error: $_"
}

# Prepare response
$body = @{
    Message           = $message
    TicketId          = $TicketId
    ResultCode        = $resultCode
    ResultStatus      = $resultStatus
    UserPrincipalName = $UserPrincipalName
    TenantId          = $TenantId
    RequestedLicense  = $LicenseTypes
    Metadata          = $json.metadata
}

# Return HTTP response
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode  = if ($resultCode -eq 200) { [HttpStatusCode]::OK } else { [HttpStatusCode]::InternalServerError }
    Body        = $body
    ContentType = "application/json"
})
