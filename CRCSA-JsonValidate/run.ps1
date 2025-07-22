using namespace System.Net

param($Request, $TriggerMetadata)

function Clear-ObjectPlaceholders {
    param (
        [psobject]$obj,
        [hashtable]$visited = $null
    )

    if ($null -eq $visited) {
        $visited = @{}
    }

    # Prevent infinite recursion for circular references
    if ($obj -is [object] -and $visited.ContainsKey($obj)) {
        return $null
    }
    if ($obj -is [object]) {
        $visited[$obj] = $true
    }

    if ($obj -is [System.Collections.IDictionary] -or $obj -is [PSCustomObject]) {
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

# Use the deserialized request body directly
$jsonObj = $Request.Body
Write-Host "Parsed body type: $($jsonObj.GetType().FullName)"

# Validate required fields
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

if ($missingFields.Count -gt 0) {
    Write-Host "Validation failed for fields: $($missingFields -join ', ')"
    return Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 400
        Body = "Validation failed: Missing or placeholder values for: $($missingFields -join ', ')"
        Headers = @{ "Content-Type" = "text/plain" }
    })
}

# Clean the JSON
$cleaned = Clear-ObjectPlaceholders -obj $jsonObj
Write-Host "Validation passed. Returning cleaned JSON."

# Return the cleaned JSON
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = ($cleaned | ConvertTo-Json -Depth 10)
    Headers = @{ "Content-Type" = "application/json" }
})
