
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

    Set-Content -Path "data/$FileName.json" -Value $($Result | ConvertTo-Json)


    if($RelativeUri -eq "sessions") {
        
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
an HTTP GET request using the specified web session. The request body
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
    [Parameter(Mandatory)][string]$RelativeUri
)

    $FullUri = "$Uri/$RelativeUri"

    Write-Host "[GET] $FullUri"
    Write-Host $JsonData

    $Result = Invoke-RestMethod `
        -Uri "$FullUri" `
        -Method Get `
        -WebSession $Session `
        -ContentType "application/json" `
        -StatusCodeVariable "StatusCode"

    Write-Host $StatusCode
    
    if($StatusCode -lt 200 -or $StatusCode -gt 299) {
        $Result
        throw "Invalid status code: $($StatusCode)"
    }
    
    $FileName = $RelativeUri -Replace "/", "_"
    $FileName = $FileName.Trim("_")

    Set-Content -Path "data/$FileName.json" -Value $($Result | ConvertTo-Json)

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

$Option = Read-Host -Prompt `
    "Select authentication method:`n`n" `
    "1. Use credentials.`n" `
    "2. Use API token.`n`n" `
    "Enter number"

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
        Write-Host "Adding x-csrftoken: $CsrfToken"


    }
    
    "2" {

        $Token = Read-Host -Prompt "Enter the API authentication token."

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

Invoke-ApiGet -RelativeUri "apps"
Invoke-ApiGet -RelativeUri "app-servers"
Invoke-ApiGet -RelativeUri "auth-providers"
Invoke-ApiGet -RelativeUri "branding"
Invoke-ApiGet -RelativeUri "branding-images"
Invoke-ApiGet -RelativeUri "categories"
Invoke-ApiGet -RelativeUri "domains"
Invoke-ApiGet -RelativeUri "features"
Invoke-ApiGet -RelativeUri "labels"
Invoke-ApiGet -RelativeUri "ssl-offloader-certificates"
Invoke-ApiGet -RelativeUri "twofactor-providers"
Invoke-ApiGet -RelativeUri "user-file-types"

Write-Host -ForegroundColor Green "Export complete."