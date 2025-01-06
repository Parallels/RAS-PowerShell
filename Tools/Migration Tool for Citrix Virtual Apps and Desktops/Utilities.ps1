<#
.SYNOPSIS
Utilities used in PrepareImport.ps1

.DESCRIPTION
This is a utility script which contains all the OU/AD resolution functions,
Table creation, and Logs. Without this script PrepareImport
will not work

.EXAMPLE
. "./Utilities.ps1"
#>
function CreateTable ($TableName, $Columns) {
	Log -type "INFO" -message "Creating table '$($TableName)' ... "

	$table = New-Object System.Data.DataTable
	$table.TableName = $TableName
	for ($i = 0; $i -lt $Columns.Length; $i++) {
		[void]$table.Columns.Add($Columns[$i])
	}
	return , $table
}

function GetCmdletCallDetails () {
	$stack = Get-PSCallStack
	[string]$callerName = (Get-PSCallStack).Command
	$linenumber = @((Get-PSCallStack).ScriptLineNumber)

	$callerName = $callerName -replace '\s+', ' '
	$callerName = [regex]::Replace($callerName, "\s+", " ")
	[System.Collections.ArrayList]$callerList = $callerName.Split(' ')
	$callerList.RemoveAt($callerList.Count - 1)
	$callerList.RemoveAt(0)
	$callerList.Reverse()

	$callerName = [string]::Join('::', $callerList.ToArray())

	return @{
		LineNumber = $linenumber
		CallerName = $callerName
		CallerList = $callerList
	}
}

function Log ([string] $type, [string] $message, $testInfo = $null, $testStatus = $null, [Exception]$exception = $null, [int] $tabSize) {
	$status_color = "DarkGrey"

	if ($type -eq "INFO") {
		$status_color = "Gray"
	}
	elseif ($type -eq "WARNING") {
		$status_color = "Yellow"
	}
	elseif ($type -eq "SUCCESS") {
		$status_color = "Green"
	}
	else {
		$status_color = "Red"
	}

	$date = Get-Date -Format g
	Write-Host "[$($date)]" -ForegroundColor DarkCyan -NoNewline
	Write-Host "[$($type)]" -ForegroundColor $status_color -NoNewline

	$res = GetCmdletCallDetails
	$callerName = $res.CallerList[$res.CallerList.Count - 2]
	Write-Host "[Func:$($callerName)][Line: $($res.LineNumber[2])]" -ForegroundColor Magenta -NoNewline
	Write-Host " ===> " -ForegroundColor DarkGray -NoNewline
	$message += "$("`t"*$tabSize)"
	if ($exception) {
		$message += " Error was: $($exception.Message)"
	}

	if ($testInfo -ne $null) {
		$message += "[TestInfo: $testInfo]"
	}

	if ($testStatus -eq $null) {
		Write-Host "$($message)" -ForegroundColor Green
	}
	else {
		$status = ""
		if ($testStatus) {
			$status = "PASSED"
			$color = "Green"
		}
		else {
			$status = "FAILED"
			$color = "Red"
		}
		Write-Host "[$($status)]" -ForegroundColor $color -NoNewline
		Write-Host "$($message)" -ForegroundColor Gray
	}
}

function LoadXML ([string] $path) {
	[xml] $data = New-Object xml
    $xmlReader = $null
	try {
		$path = Resolve-Path $path
		Log -type "INFO" -message "Reading XML from $($path) ..."
		[System.Xml.XmlTextReader] $xmlReader = New-Object System.Xml.XmlTextReader ($path)
		$xmlReader.Namespaces = $false

		Log -type "INFO" -message "Marshalling into XML type ..."
		[void]$data.Load($xmlReader)
	}
	catch {
		Log -type "ERROR" -message "Failed to get XML data! $($_)"
        if( $null -ne $xmlReader ) {
		    Log -type "INFO" -message "Closing XMLReader stream and disposing ..."
		    $xmlReader.Close()
		    $xmlReader.Dispose()
        }
		Log -type "INFO" -message "Done."
		throw
	}
	# [System.Xml.XmlTextReader] $xmlReader = [System.Xml.XmlTextReader]::new($path)

	if (-not $data) {
		Log -type "ERROR" -message "Marshalled XML object is null."
		throw "Marshalled XML object is null."
	}

	Log -type "INFO" -message "XML file successfully read!"
	return $data
}

function XASettingsExport () {
	Log -type "INFO" -message "Exporting XenApp settings to ./res ..." -fnName "XASettingsExport"
	try {
		Add-PSSnapin citrix.xenapp.commands -ErrorAction SilentlyContinue -ErrorVariable "CitrixException"
		if ($CitrixException) {
			throw
		}

		if (![System.IO.Directory]::Exists("./res")) {
			[System.IO.Directory]::CreateDirectory("./res")
		}
		Get-XAApplicationReport * | Export-Clixml "./res/applications.xml" -ErrorAction SilentlyContinue -ErrorVariable "CitrixException" | Out-Null
		Log -type "INFO" -message "Exported XAApplicationReport."

		Get-XAZone | Export-Clixml "./res/zones.xml" -ErrorAction SilentlyContinue -ErrorVariable "CitrixException" | Out-Null
		Log -type "INFO" -message "Exported XAZones."

		Get-XAServer * | Export-Clixml "./res/servers.xml" -ErrorAction SilentlyContinue -ErrorVariable "CitrixException" | Out-Null
		Log -type "INFO" -message "Exported XAServers."

		Get-XAWorkerGroup | Export-Clixml "./res/workergroups.xml" -ErrorAction SilentlyContinue -ErrorVariable "CitrixException" | Out-Null
		Log -type "INFO" -message "Exported XAWorkerGroups."
		Log -type "INFO" -message "Settings export complete!"
		return $true
	}
	catch {
		Log -type "ERROR" -message "Something when wrong during export!"
		Log -type "ERROR" -message "Error was: $($CitrixException)"
		Log -type "ERROR" -message "Cannot continue without necessary export settings :("
		return $false
	}
}

function ImportRASAdmin () {

	$PSAdminModule = "RASAdmin"
	if($(Get-Module -ListAvailable).Name.Contains("PSAdmin"))
	{
		$PSAdminModule = "PSAdmin"
	}

	try {
		Import-Module $PSAdminModule -ErrorVariable "PSAdminError" -ErrorAction SilentlyContinue
		if ($PSAdminError) {
			throw
		}
	}
	catch [Exception] {
		Log -type "ERROR" -message "No installed $PSAdminModule module was found!"
		Log -type "ERROR" -message "Error was: $($PSAdminError)"
		Log -type "ERROR" -message "Cannot continue :("
		return $false
	}
}

function CreateRASSession ([string]$server, [string] $username, [securestring]$password) {
	if (-not (ImportRASAdmin)) {
		return $false
	}

	try {
		Log -type "INFO" -message "Removing existent RAS sessions."
		# Remove-RASSession

		try {
			Log -type "INFO" -message "Connecting you to your RAS farm. ---> $($server)"
			New-RASSession -Server $server -Username $username -Password $password
		}
		catch [Exception] {
			Log -type "ERROR" -message "Something went wrong when creating a RAS session :("
			Log -type "ERROR" -message "Error was: $($_)"
			Log -type "ERROR" -message "Cannot continue :("
			return $false
		}
	}
	catch {
		Log -type "ERROR" -message "Something went wrong when removing a RAS session :("
		Log -type "ERROR" -message "Error was: $($_)"
		Log -type "ERROR" -message "Trying to continue anyways ..."
		return $true
	}
	return $true
}

function GUIDToOctetString([guid] $guid) {
	return ("\" + ([System.String]::Join('\', ($guid.ToByteArray() | ForEach-Object { $_.ToString('x2') }))));
}

function LDAPSearcherFromPath([string]$path) {
	if (-not $path) {
		return LDAPSearcher($null)
	}

	$root = [adsi]$path
	$searcher = [adsisearcher]$root
	return $searcher
}

function LDAPSearcher ([string] $domain = $null) {
	Log -type "INFO" -message "Initializing LDAP ..."
	if (-not $domain) {
		$domain = $domain = (Get-WmiObject Win32_ComputerSystem).Domain
		Log -type "INFO" -message "No domain specified using current domain instead" -fnName "LDAPSearcher"
		Log -type "INFO" -message "Domain: '$($domain)'" -fnName "LDAPSearcher"
	}
	$root = [adsi]"LDAP://$domain"
	$searcher = [adsisearcher]$root
	Log -type "INFO" -message "OK. LDAP path '$($root.Path)'"
	return $searcher
}

function GetOUByGUID ([guid] $ouGUID) {
	$hexString = GUIDToOctetString($ouGUID)
	$searcher = LDAPSearcher($null)

	$searcher.Filter = "(&(objectclass=organizationalUnit)(objectguid=$hexString))"
	Log -type "INFO" -message "Searching for organizational units (OU) by GUID."
	Log -type "INFO" -message "LDAP search query -> '$($searcher.Filter)'"
	$ret = $searcher.FindOne()
	if ($ret) {
		Log -type "INFO" -message "OU search successful!"
	}
	else {
		Log -type "WARNING" -message "OU was not found!"
	}
	return $ret
}

function GetOUFromPath([string] $path) {
	Log -type "INFO" -message "Searching for organizational units (OU) in path ..."
	Log -type "INFO" -message "LDAP search path '$($path)'"
	$searcher = LDAPSearcher($path)
	$res = $searcher.FindOne()

	if ($ret) {
		Log -type "INFO" -message "OU search successful!"
	}
	else {
		Log -type "INFO" -message "OU was not found!"
	}

	return $res
}

function GetOUServers($path = $null) {
	if (-not $path) {
		Log -type "WARNING" -message "path was null or empty."
		return
	}
	Log -type "INFO" -message "Searching for servers from an OU ..."
	$searcher = LDAPSearcherFromPath($path)
	$searcher.Filter = "(&(objectclass=computer)(operatingSystem=*Server*))"
	Log -type "INFO" -message "LDAP search query -> '$($searcher.Filter)'"
	$res = $searcher.FindAll()

	if ($res) {
		Log -type "INFO" -message "$($res.Count) servers were found!"
	}
	else {
		Log -type "INFO" -message "No servers were found. $($res.Count)"
	}

	return $res
}

function IsGUID ([string] $value) {
	$regex = "(\{{0,1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{0,1})"
	return ($value -match $regex)
}
