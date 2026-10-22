$ShowSensitiveInformation = $false

if($ShowSensitiveInformation) {
    Write-Host -ForegroundColor Yellow "WARNING: API and CSRF tokens will be visible. Anyone viewing your screen or terminal recordings may capture them."
    if((Read-Host -Prompt "Type ACCEPT to allow sensitive information to be displayed") -cne "ACCEPT") {
        throw "Sensitive information disclaimer was not accepted."
    }
}

function Invoke-ApiPost {
<#
.SYNOPSIS
Sends an HTTP POST request to an API endpoint.

.DESCRIPTION
Constructs a request URL from a base URL and a relative URI, then sends
an HTTP POST request using the specified web session. The request body
is serialized to JSON before being sent.

.PARAMETER RelativeUri
The relative path of the API endpoint.

.PARAMETER WebSession
The WebRequestSession to use for the request. This session can contain
cookies, authentication information, and other session state.

.PARAMETER Body
The object to send as the request payload. The object is converted to
JSON using ConvertTo-Json before being transmitted.

.OUTPUTS
System.Object

Returns the deserialized response from the API.

Sends a POST request to the users endpoint and returns the response.

.NOTES
The request is sent with the Content-Type header set to
'application/json'.
#>
param(
    [Parameter(Mandatory)][string]$RelativeUri,
    [Parameter(Mandatory)][object]$Body
)

    $FullUri = "$Uri/$RelativeUri"
    $JsonData = ($Body | ConvertTo-Json -Depth 10)

    Write-Host "[POST] $FullUri"

    if($RelativeUri -ne "sessions/") {
        # Don't display the password.
        Write-Host $JsonData
    }

    $Result = Invoke-RestMethod `
        -Uri "$FullUri" `
        -Method Post `
        -WebSession $Session `
        -Body $JsonData `
        -ContentType "application/json" `
        -StatusCodeVariable "StatusCode"

    Write-Host $StatusCode
    
    if($StatusCode -lt 200 -or $StatusCode -gt 299) {
        $Result
        throw "Invalid status code: $($StatusCode)"
    }
    
    $FileName = $RelativeUri -Replace "/", "_"
    $FileName = $FileName.Trim("_")

    Set-Content -Path "data/$FileName.json" -Value $($Result | ConvertTo-Json -Depth 10)


    if($ShowSensitiveInformation -and $RelativeUri -eq "sessions") {
        
        $CsrfToken = $Session.Cookies.GetCookies($FullUri)['csrftoken'].Value
        Write-Host -ForegroundColor Green "CSRF Token: $CsrfToken"
    }

    return $Result

}


function Invoke-ApiGet {
<#
.SYNOPSIS
Sends an HTTP GET request to an API endpoint.

.DESCRIPTION
Constructs a request URL from a base URL and a relative URI, then sends
an HTTP GET request using the configured web session. List requests use
limit=1000 and follow next links, combining all pages in the results wrapper.

.PARAMETER RelativeUri
The relative path of the API endpoint.

.PARAMETER List
Request up to 1000 records per page for a list endpoint.

.OUTPUTS
System.Object

Returns the deserialized response, with all pages combined when paginated.

.NOTES
The request is sent with the Content-Type header set to
'application/json'.
#>
param(
    [Parameter(Mandatory)][string]$RelativeUri,
    [switch]$List
)

    $FullUri = "$Uri/$RelativeUri"
    if($List) {
        $Separator = if($FullUri.Contains("?")) { "&" } else { "?" }
        $FullUri += "${Separator}limit=1000"
    }

    $ApiUri = [uri]"$Uri/"
    $VisitedUris = [System.Collections.Generic.HashSet[string]]::new()
    $AllResults = [System.Collections.Generic.List[object]]::new()
    $Result = $null
    $IsPaginated = $false

    do {
        if(!$VisitedUris.Add($FullUri)) {
            throw "Repeated pagination URI: $FullUri"
        }

        Write-Host "[GET] $FullUri"

        $Page = Invoke-RestMethod `
            -Uri "$FullUri" `
            -Method Get `
            -WebSession $Session `
            -ContentType "application/json" `
            -StatusCodeVariable "StatusCode"

        Write-Host $StatusCode

        if($StatusCode -lt 200 -or $StatusCode -gt 299) {
            throw "Invalid status code: $($StatusCode)"
        }

        if($VisitedUris.Count -eq 1) {
            $Result = $Page
            $IsPaginated = $null -ne $Page.PSObject.Properties['results'] -and
                $null -ne $Page.PSObject.Properties['next']
        }

        if(!$IsPaginated) {
            break
        }

        if($null -eq $Page.PSObject.Properties['next'] -or
            $Page.results -isnot [array]) {
            throw "Invalid paginated response from: $FullUri"
        }

        foreach($Item in $Page.results) {
            $AllResults.Add($Item)
        }

        if(!$Page.next) {
            break
        }

        $NextUri = [uri]::new([uri]$FullUri, [string]$Page.next)
        if($NextUri.Scheme -ne $ApiUri.Scheme -or
            $NextUri.Authority -ne $ApiUri.Authority -or
            !$NextUri.AbsolutePath.StartsWith($ApiUri.AbsolutePath, [StringComparison]::Ordinal)) {
            throw "Pagination link is outside the configured API: $NextUri"
        }
        $FullUri = $NextUri.AbsoluteUri
    } while($true)

    if($IsPaginated) {
        $Result.results = $AllResults.ToArray()
        $Result.next = $null
        if($null -ne $Result.PSObject.Properties['previous']) {
            $Result.previous = $null
        }
    }
    
    $FileName = $RelativeUri -Replace "/", "_"
    $FileName = $FileName.Trim("_")

    Set-Content -Path "data/$FileName.json" -Value $($Result | ConvertTo-Json -Depth 10)

    return $Result

}



if(!(Test-Path "data")) {
    New-Item -ItemType Directory "data"
}


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ProductName = "Parallels Secure Workspace"
$Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

Write-Host -ForegroundColor Yellow "To export all data from $ProductName, an API token or credentials from a Global Admin are required."

$Uri = Read-Host -Prompt "Enter the base URI of the $ProductName API. For example: https://your-own.org/api/v2"

if($Uri -notmatch "/api/v2$") {
    throw "The URI does not appear to be valid. Expected format: https://your-own.org/api/v2"
}

$Option = Read-Host -Prompt @"
Select authentication method:
1. Use credentials.
2. Use API token.

Enter number:
"@

switch ($Option) {

    "1" {

        $Domain = $(Read-Host -Prompt "Enter the administrative Workspace domain to authenticate to").ToUpper()
        $Username = Read-Host -Prompt "Enter the username (format: simple username, e.g. jsmith)"
        $Password = Read-Host -Prompt "Enter the password" -MaskInput

        $Params = @{
            "domain" = $Domain;
            "username" = $Username;
            "password" = $Password;
            logout_other_sessions = $true;
        }
        Invoke-ApiPost -RelativeUri "sessions/" -Body $Params


        # Extract csrftoken from cookie and set the x-csrftoken header for future requests
        $CsrfToken = $Session.Cookies.GetCookies("$Uri/sessions/")['csrftoken'].Value
        $Session.Headers.Add('x-csrftoken', $CsrfToken)
        if($ShowSensitiveInformation) {
            Write-Host "Adding x-csrftoken: $CsrfToken"
        }


    }
    
    "2" {

        $Token = Read-Host -Prompt "Enter the API authentication token." -MaskInput:(-not $ShowSensitiveInformation)

        if($Token -notmatch "^[a-f0-9]{40}") {

            throw "This does not appear to be a valid token."



        }

        $Session.Headers.Add("Authorization", "Token $($Token)")

    }

    Default {

        throw "An invalid option was chosen: $Option"

    }
}

# Export generic config.
Invoke-ApiGet -RelativeUri "configuration/1"

Invoke-ApiGet -RelativeUri "apps" -List
Invoke-ApiGet -RelativeUri "app-servers" -List
Invoke-ApiGet -RelativeUri "auth-providers" -List
Invoke-ApiGet -RelativeUri "branding" -List
Invoke-ApiGet -RelativeUri "branding-images" -List
Invoke-ApiGet -RelativeUri "categories" -List
Invoke-ApiGet -RelativeUri "domains" -List
Invoke-ApiGet -RelativeUri "features" -List
Invoke-ApiGet -RelativeUri "labels" -List
Invoke-ApiGet -RelativeUri "ssl-offloader-certificates" -List
Invoke-ApiGet -RelativeUri "twofactor-providers" -List
Invoke-ApiGet -RelativeUri "user-file-types" -List

Write-Host -ForegroundColor Green "Export complete."