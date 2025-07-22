using namespace System.Net

param($Request, $TriggerMetadata)

# Helper function to recursively remove placeholder values
function Clear-ObjectPlaceholders {
    param (
        [psobject]$obj,
        [hashtable]$visited = $null
    )

    if ($null -eq $visited) {
        $visited = @{}
    }

    # Only guard recursion for objects
    if ($obj -is [System.Collections.IDictionary] -or $obj -is [PSCustomObject]) {
        $objHash = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)
        if ($visited.ContainsKey($objHash)) {
            return $null
        }
        $visited[$objHash] = $true

        $cleaned = @{}
        foreach ($property in $obj.PSObject.Properties) {
            $value = Clear-ObjectPlaceholders -obj $property.Value -visited $visited
            if ($null -ne $value -and `
                -not ($value -is [System.Collections.IEnumerable] -and $value.Count -eq 0)) {
                $cleaned[$property.Name] = $value
            }
        }
        if ($cleaned.Count -eq 0) {
            return $null
        } else {
            return $cleaned
        }
    }
    elseif ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
        if ($obj.Count -eq 1 -and $obj[0] -is [string] -and $obj[0].Trim() -match '^@') {
            return @()
        }
        $newArray = @()
        foreach ($item in $obj) {
            $cleaned = Clear-ObjectPlaceholders -obj $item -visited $visited
            if ($null -ne $cleaned) { $newArray += ,$cleaned }
        }
        if ($newArray.Count -eq 0) {
            return $null
        } else {
            return $newArray
        }
    }
    elseif ($obj -is [string] -and $obj.Trim() -match '^@') {
        return $null
    }
    return $obj
}

# Validate required fields and check for placeholders
function Validate-UserJson {
    param([psobject]$jsonObj)
    $missingFields = @()

    if (-not $jsonObj.TenantId -or $jsonObj.TenantId -match '^@') {
        $missingFields += "TenantId"
    }
    if (-not $jsonObj.TicketId -or $jsonObj.TicketId -match '^@') {
        $missingFields += "TicketId"
    }
    if (-not $jsonObj.AccountDetails) {
        $missingFields += "AccountDetails"
    } else {
        if (-not $jsonObj.AccountDetails.GivenName -or $jsonObj.AccountDetails.GivenName -match '^@') {
            $missingFields += "AccountDetails.GivenName"
        }
        if (-not $jsonObj.AccountDetails.Surname -or $jsonObj.AccountDetails.Surname -match '^@') {
            $missingFields += "AccountDetails.Surname"
        }
        if (-not $jsonObj.AccountDetails.UserPrincipalName -or $jsonObj.AccountDetails.UserPrincipalName -match '^@') {
            $missingFields += "AccountDetails.UserPrincipalName"
        }
    }
    return $missingFields
}

# Main logic
try {
    $jsonObj = $Request.Body
    $missingFields = Validate-UserJson -jsonObj $jsonObj

    if ($missingFields.Count -gt 0) {
        return Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = 400
            Body = "Validation failed: Missing or placeholder values for: $($missingFields -join ', ')"
            Headers = @{ "Content-Type" = "text/plain" }
        })
    }

    $cleaned = Clear-ObjectPlaceholders -obj $jsonObj
    return Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 200
        Body = ($cleaned | ConvertTo-Json -Depth 10)
        Headers = @{ "Content-Type" = "application/json" }
    })
}
catch {
    return Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 400
        Body = "Invalid JSON or unexpected error."
        Headers = @{ "Content-Type" = "text/plain" }
    })
}
