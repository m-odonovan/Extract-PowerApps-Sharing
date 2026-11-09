#Requires -Version 5.1

function Get-DataverseEnvironmentUrl {
    param(
        [string]$Id,
        [bool]$IncludePause = $false
    )

    $lookupErrors = @()
    if ($IncludePause) {
        Start-Sleep -Milliseconds (Get-Random -Minimum 0 -Maximum 2001)
    }
    $environment = Get-AdminPowerAppEnvironment -EnvironmentName $Id -ErrorAction Stop -ErrorVariable lookupErrors
    # Some versions of the admin module catch HTTP errors without rethrowing.
    if ($lookupErrors.Count -gt 0) {
        throw "Environment lookup failed: $($lookupErrors[0])"
    }
    $url = [string]$environment.Internal.properties.linkedEnvironmentMetadata.instanceUrl
    $uri = $null
    if (-not [uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') {
        throw "No valid Dataverse URL was returned for environment '$Id'."
    }
    return $url.TrimEnd('/')
}

function Invoke-DataverseRequest {
    param(
        [hashtable]$Context,
        [string]$RequestUri,
        [bool]$IncludePause = $false
    )

    $uri = [uri]$RequestUri
    $baseUri = [uri]$Context.Url
    if ($uri.Scheme -ne 'https' -or $uri.Authority -ne $baseUri.Authority) {
        throw 'Dataverse returned a paging URL outside the expected environment.'
    }

    $accessToken = Get-AzAccessToken -ResourceUrl "$($Context.Url)/" -ErrorAction Stop
    if ($accessToken.Token -is [System.Security.SecureString]) {
        $token = [System.Net.NetworkCredential]::new('', $accessToken.Token).Password
    }
    else {
        $token = [string]$accessToken.Token
    }
    if ([string]::IsNullOrWhiteSpace($token)) {
        throw 'Get-AzAccessToken returned an empty Dataverse token.'
    }

    $headers = @{
        Authorization      = "Bearer $token"
        Accept             = 'application/json'
        'OData-MaxVersion' = '4.0'
        'OData-Version'    = '4.0'
        Prefer             = 'odata.maxpagesize=5000'
    }
    try {
        if ($IncludePause) {
            Start-Sleep -Milliseconds (Get-Random -Minimum 0 -Maximum 2001)
        }
        $response = Invoke-RestMethod -Uri $uri.AbsoluteUri -Method Get -Headers $headers -ErrorAction Stop
    }
    finally {
        $headers.Remove('Authorization')
        $token = $null
    }
    return $response
}

function Invoke-DataverseQuery {
    param(
        [hashtable]$Context,
        [string]$Query,
        [bool]$IncludePause = $false
    )

    $nextLink = "$($Context.Url)/api/data/v9.2/$Query"
    while ($nextLink) {
        $response = Invoke-DataverseRequest -Context $Context -RequestUri $nextLink -IncludePause $IncludePause
        if ($null -eq $response -or $null -eq $response.PSObject.Properties['value']) {
            throw 'Dataverse did not return the expected collection response.'
        }
        foreach ($record in $response.value) {
            $record
        }
        $nextLink = $response.'@odata.nextLink'
    }
}
