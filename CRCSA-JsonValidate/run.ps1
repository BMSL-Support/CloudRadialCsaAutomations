using namespace System.Net

param($Request, $TriggerMetadata)

function Clear-ObjectPlaceholders {
    param ([psobject]$obj)

    if ($obj -is [System.Collections.IDictionary]) {
        $hasNull = $false
        foreach ($key in $obj.Keys) {
            $obj[$key] = Clear-ObjectPlaceholders -obj $obj[$key]
            if ($null -eq $obj[$key]) { $hasNull = $true }
        }
        if ($hasNull) { return $null }
    }
    elseif ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
        if ($obj.Count -eq 1 -and $obj[0] -is [string] -and $obj[0] -match '^@') {
            return @()
        }
        $newArray = @()
        foreach ($item in $obj) {
            $cleaned = Clear-ObjectPlaceholders -obj $item
            if ($null -ne $cleaned) { $newArray += ,$cleaned }
        }
        return $newArray
    }
    elseif ($obj -is [string] -and $obj -match '^@') {
        return $null
    }
    return $obj
}

# Read and parse the incoming JSON
$body = $Request.Body
try {
    $jsonObj = $body | ConvertFrom-Json
} catch {
    return Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = 400
        Body = "Invalid JSON"
        Headers = @{ "Content-Type" = "text/plain" }
    })
}
# Clean the JSON
$cleaned = Clear-ObjectPlaceholders -obj $jsonObj

# Return the cleaned JSON
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = ($cleaned | ConvertTo-Json -Depth 10)
    Headers = @{ "Content-Type" = "application/json" }
})
