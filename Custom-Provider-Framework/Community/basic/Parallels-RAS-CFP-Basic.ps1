# Sample request json
#  
# {"method": "provider/connect", "params": {"settings": {"token" : "7819cf5ca94a30ad154"}}}
# {"method": "guests/list"}
# {"method":"guests/control","params":{"control":"start","id":"67c58327384f5b51"}}


if($Host.Name -notmatch "ISE")
{
	[Console]::InputEncoding = [System.Text.Encoding]::UTF8
	[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$stdout = [Console]::OpenStandardOutput()
$writer = New-Object System.IO.StreamWriter($stdout, [System.Text.Encoding]::UTF8)
$writer.AutoFlush = $true  # Ensure immediate flushing

$providerNamePrefix = "Sample:"

$ErrorCodes = @{
    ParseError     = -32700
    MethodNotFound = -32601
    InvalidParams  = -32602
    InternalError  = -32603
}

# Predefined method registry with required parameters
$MethodRegistry = @{

    "provider/initialize"  = @{ Handler = { param($data) Handle-Initialize }; RequiredFields = @() }

    "provider/connect"     = @{ Handler = { param($data) Handle-Connect $data.params }; RequiredFields = @("params.settings") }

    "provider/disconnect"  = @{ Handler = { param($data) Handle-Disconnect}; RequiredFields = @() }

    "guests/list"          = @{ Handler = { param($data) Handle-GuestList }; RequiredFields = @() }
	
    "guests/get"           = @{ Handler = { param($data) Handle-GuestInfo $data.params }; RequiredFields = @("params.id") }

    "guests/control"       = @{ Handler = { param($data) Handle-GuestControl $data.params}; RequiredFields = @("params.id", "params.control") }
}

# Send a structured JSON response
function Send-Response {

    param([object]$responseObj)

    $responseJson = $responseObj | ConvertTo-Json -Compress -Depth 10

    $writer.WriteLine($responseJson)
}

# Safely parse JSON from input
function ConvertTo-JsonSafe {
   param([string]$inputLine)

    try {
		return $inputLine | ConvertFrom-Json -ErrorAction Stop
    }

    catch {
        return $null
    }

}

# Validate required fields in JSON input
function Validate-MethodInput {

    param([object]$data, [array]$requiredFields)

    foreach ($field in $requiredFields) {

        $keys = $field -split '\.'

        $value = $data

        foreach ($key in $keys) {

            if ($value -and $value.PSObject.Properties[$key]) {

                $value = $value.$key

            } else {

                return @{ error = @{ code = $ErrorCodes.InvalidParams; message = "$providerNamePrefix Missing field: $field" } }
            }

        }

    }

    return $null

}

# Process methods
function Process-Method {

    param([string]$inputLine)

    $methodData = ConvertTo-JsonSafe $inputLine

    if (-not $methodData) {

        return @{ error = @{ code = $ErrorCodes.ParseError; message = "$providerNamePrefix Invalid JSON format" } }

    }

    $methodName = $methodData.method

    if (-not $methodName) {

        return @{ error = @{ code = $ErrorCodes.MethodNotFound; message = "$providerNamePrefix Missing method name" } }

    }

    $methodEntry = $MethodRegistry[$methodName.ToLower()]

    if (-not $methodEntry) {

        return @{ error = @{ code = $ErrorCodes.MethodNotFound; message = "$providerNamePrefix Unknown method: $methodName" } }

    }

    # Validate required fields
    $validationError = Validate-MethodInput -data $methodData -requiredFields $methodEntry.RequiredFields

    if ($validationError) {

        return $validationError

    }

    try {

        return & $methodEntry.Handler -data $methodData
    }

    catch {
        return @{ error = @{ code = $ErrorCodes.InternalError; message = "$providerNamePrefix Method execution message" } }
    }

}

# Method Handlers
function Handle-Initialize {

       $capabilities = [PSCustomObject]@{
           suspend = $true
           guests_polling_rate = 30
       }
	   
    return @{
        result = @{
            version = "1.0.0";
            capabilities = $capabilities
        }
    }
}

function Handle-Connect {

    param($params)

    $settings = $params.settings

    if (-not $settings.token) {

		return @{ error = @{ code = $ErrorCodes.InvalidParams; message = "$providerNamePrefix Invalid token" } }
    }

   return [PSCustomObject]@{
	   result = @{
       }
   }
}

function Handle-Disconnect {
   return [PSCustomObject]@{
	   result = @{
       }
   }
}

function Handle-GuestInfo {

    param($params)

    if (-not $params.id) {

		return @{ error = @{ code = $ErrorCodes.InvalidParams; message = "$providerNamePrefix Invalid guest id" } }
    }

    $guests = @{
        "123456" = @{ name = "vm1"; state = "powered_off" }
        "7890abc" = @{ name = "vm2"; state = "powering_off" }
	    "5623456" = @{ name = "vm3"; state = "suspended" }
        "7890xys" = @{ name = "vm4"; state = "suspending" }
    }

   return [PSCustomObject]@{
	   result = $guests[$params.id]
   }
}

function Handle-GuestList {

    $guests = @(
        "123456",
        "7890abc",
	    "5623456",
        "7890xys"
    )
	
   $result = [ordered]@{
	   result = @{
           guests = $guests
       }
   }
   return $result
}

function Handle-GuestControl {

    param($params)

    if (-not $params.id) {

		return @{ error = @{ code = $ErrorCodes.InvalidParams; message = "$providerNamePrefix Invalid guest id" } }
    }
	
	if (-not $params.control) {

		return @{ error = @{ code = $ErrorCodes.InvalidParams; message = "$providerNamePrefix Invalid guest control" } }
    }


   return [PSCustomObject]@{
	   result = @{
       }
   }
}


# Main processing loop
while ($true) {

    $inputLine = [Console]::In.ReadLine()

    if (-not $inputLine) { break }
	
    $response = Process-Method $inputLine.Trim()

	#"Response: $response" | Tee-Object -FilePath C:\Temp\ps_debug_log.txt -Append
    Send-Response $response
} 
