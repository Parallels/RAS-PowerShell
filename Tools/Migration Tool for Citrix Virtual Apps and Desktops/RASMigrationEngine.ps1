. "./Utilities.ps1"
<#
.SYNOPSIS
Citrix Virtual Apps and Desktops to RAS migration engine.

.DESCRIPTION
This is a migration script to transfer Citrix Virtual Apps and Desktops -Applications.

.EXAMPLE
# Import the script
. "./PrepareImport.ps1"

.NOTES
General notes
Utilities.ps1 is required to operate this script!
#>

## when using clixml via import-clixml we don't get a reference id so create a script wide unique one
[int]$script:refID = 0

$PSAdminModule = "RASAdmin"
if($(Get-Module -ListAvailable -Verbose:$false -Debug:$false).Name.Contains("PSAdmin"))
{
	$PSAdminModule = "PSAdmin"
}

function Get-Hash
{
    ## Modified from https://xpertkb.com/compute-hash-string-powershell/ as MD5 in original not FIPS compliant

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='Bytes to hash')]
        [object]$BytesToHash
    )

    [string]$result = $null
    if($null -ne $BytesToHash)
    {
        $hasher = $null
        $hasher = New-Object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
        if( $hasher )
        {
            $hashByteArray = $hasher.ComputeHash( $BytesToHash )
            foreach( $byte in $hashByteArray )
            {
              $result = "$result{0:X2}" -f $byte
            }
        }
    }
    $result ## return
}

function CreateTypeName ([System.Xml.XmlElement] $xml) {
	$Name = ''
	$refID = 0
	if (-not $xml.TN) {
		$refID = $xml.TNRef.RefId
		# Write-HOst "REFID: $($refID)"
		try {
			$Name = $xml.SelectSingleNode("//TN[@RefId='$($refID)']").T[0]
		}
		catch {
			# Write-Host "RefID unknown $($refID)"
			return "Unknown"
		}
	}
	else {
		$refID = $xml.TN.RefId
		$Name = $xml.TN.T[0]
	}
	return $Name
}

function CreateXAObject([System.Xml.XmlElement]$xml) {

	$ret = @{
		Type       = ""
		ObjectName = $xml.SelectSingleNode("ToString") | Select-Object -ExpandProperty InnerText -ErrorAction SilentlyContinue
		RefId      = $xml.RefId
		xml        = $xml
	}
	return $ret
}

function CreateXAFarm
{
    Param
    (
        [System.Xml.XmlElement]$xml ,
        $clixml 
    )
	$super = CreateXAObject -xml $xml
	$result = @{
		Type          = $super.Type
		ObjectName    = $super.ObjectName
		RefId         = $super.RefId
		FarmName      = $clixml | Select-Object -ExpandProperty FarmName -ErrorAction SilentlyContinue
		ServerVersion = $clixml | Select-Object -ExpandProperty ServerVersion -ErrorAction SilentlyContinue
		MachineName   = $clixml | Select-Object -ExpandProperty MachineName -ErrorAction SilentlyContinue
	}
    $result ## return
}

function CreateXAZone
{
    Param
    (
        [System.Xml.XmlElement]$xml ,
        $clixml 
    )
	$super = CreateXAObject -xml $xml
	$result = [pscustomobject] @{
		Type          = $super.Type
		ObjectName    = $super.ObjectName
		RefId         = $super.RefId
		ZoneName    = $clixml | Select-Object -ExpandProperty ZoneName -ErrorAction SilentlyContinue
		MachineName = $clixml | Select-Object -ExpandProperty MachineName -ErrorAction SilentlyContinue
	}
    $result ## return
}

function CreateXAServer 
{
    Param
    (
        [System.Xml.XmlElement]$xml ,
        $clixml 
    )
	$super = CreateXAObject($xml)
	$obj = @{
		Type        = $super.Type
		Name        = $super.ObjectName
		RefId       = $super.RefId
		MachineName = $clixml | Select-Object -ExpandProperty MachineName -ErrorAction SilentlyContinue
		ServerName  = $clixml | Select-Object -ExpandProperty ServerName -ErrorAction SilentlyContinue
		ServerId    = $clixml | Select-Object -ExpandProperty ServerId -ErrorAction SilentlyContinue
		FolderPath  = $clixml | Select-Object -ExpandProperty FolderPath -ErrorAction SilentlyContinue
		ZoneName    = $clixml | Select-Object -ExpandProperty ZoneName -ErrorAction SilentlyContinue
        }

    if( $obj -and [string]::IsNullOrEmpty( $obj.ServerName ) ) {
        $obj.ServerName = $obj.MachineName
	}
	return $obj
}

function CreateXAApplication {
    Param
    (
        [System.Xml.XmlElement] $xml ,
        $clixml
    )
	$super = CreateXAObject($xml)
	$ret = [PSCustomObject] @{
		Type                            = $super.Type
		Name                            = $clixml | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue
		Description                     = $clixml | Select-Object -ExpandProperty Description -ErrorAction SilentlyContinue
		WorkingDirectory                = $clixml | Select-Object -ExpandProperty WorkingDirectory -ErrorAction SilentlyContinue
		RefId                           = $super.RefId
        ## strip trailing backslash as later lookup fails because it does lookup without the backslash
		FolderPath                      = ($clixml | Select-Object -ExpandProperty AdminFolderName <#FolderPath#> -ErrorAction SilentlyContinue) -replace '\\+$'
		ClientFolder                    = $clixml | Select-Object -ExpandProperty ClientFolder -ErrorAction SilentlyContinue
		StartMenuFolder                 = $clixml | Select-Object -ExpandProperty StartMenuFolder -ErrorAction SilentlyContinue
		CommandLineExecutable           = $clixml | Select-Object -ExpandProperty CommandLineExecutable -ErrorAction SilentlyContinue
		ApplicationType                 = $clixml | Select-Object -ExpandProperty ApplicationType -ErrorAction SilentlyContinue
		IconData                        = $clixml | Select-Object -ExpandProperty IconData -ErrorAction SilentlyContinue
		ApplicationId                   = $clixml | Select-Object -ExpandProperty ApplicationId -ErrorAction SilentlyContinue
		ContentAddress                  = $clixml | Select-Object -ExpandProperty ContentAddress -ErrorAction SilentlyContinue
		WindowType                      = $clixml | Select-Object -ExpandProperty WindowType -ErrorAction SilentlyContinue
		InstanceLimit                   = [int]($clixml | Select-Object -ExpandProperty InstanceLimit -ErrorAction SilentlyContinue)
		UserFilter                      = $clixml | Select-Object -ExpandProperty Accounts -ErrorAction SilentlyContinue | Select-Object -ExpandProperty AccountId -ErrorAction SilentlyContinue ## (GetPropertyOrNull $clixml "Accounts.AccountId")
		UserFilterByAccountName         = $clixml | Select-Object -ExpandProperty Accounts -ErrorAction SilentlyContinue | Select-Object -ExpandProperty AccountId -ErrorAction SilentlyContinue  ## (GetPropertyOrNull $clixml "Accounts.AccountName")
		ColorDepth                      = $clixml | Select-Object -ExpandProperty ColorDepth -ErrorAction SilentlyContinue
		Enabled                         = $false
		Width                           = $null
		Height                          = $null
		MaximizedOnStartup              = $false
		WaitOnPrinterCreation           = $false
		AddToClientDesktop              = $false
		AddToClientStartMenu            = $false
		MultipleInstancesPerUserAllowed = $false
		Servers                         = @()
		WorkerGroups                    = @()
		Extensions                      = $null
	}
	
	$servers = GetPropertyOrNull $clixml "ServerNames"
	$wgNames = GetPropertyOrNull $clixml "WorkerGroupNames"
	$filetypes = GetPropertyOrNull $clixml "Extensions"

	if ($filetypes) {
		[System.Collections.ArrayList] $extensions = $filetypes
		for ($i = 0; $i -lt $extensions.Count; $i++) {
			$extensions[$i] = $extensions[$i].Remove(0, 1)
		}
		$ret.Extensions = [string]::Join(",", $extensions.ToArray())
	}

	if ($ret.UserFilter) {
		$ret.UserFilter = @($ret.UserFilter)
		for ($i = 0; $i -lt $ret.UserFilter.Count; $i++) {
			$ret.UserFilter[$i] = $ret.UserFilter[$i].Split("/")[-1] # last item
		}
	}

	$tmpBool = $false
	[bool]::TryParse((GetPropertyOrNull $clixml "AddToClientDesktop"), [ref] $tmpBool) | Out-Null
	$ret.AddToClientDesktop = $tmpBool

	[bool]::TryParse((GetPropertyorNull $clixml "AddToClientStartMenu"), [ref] $tmpBool) | Out-Null
	$ret.AddToClientStartMenu = $tmpBool

	[bool]::TryParse((GetPropertyOrNull $clixml "MultipleInstancesPerUserAllowed"), [ref] $tmpBool) | Out-Null
	$ret.MultipleInstancesPerUserAllowed = $tmpBool

	[bool]::TryParse((GetPropertyOrNull $clixml "WaitOnPrinterCreation"), [ref] $tmpBool) | Out-Null
	$ret.WaitOnPrinterCreation = $tmpBool

	[bool]::TryParse($clixml.Enabled, [ref] $tmpBool) | Out-Null
	$ret.Enabled = $tmpBool

	[bool]::TryParse((GetPropertyOrNull $clixml "MaximizedOnStartup"), [ref] $tmpBool) | Out-Null
	$ret.MaximizedOnStartup = $tmpBool

	$windowType = GetPropertyOrNull $clixml "WindowType"
	if ($windowType) {
		$widthHeight = $windowType.Split("x")
		$ret.Width = $widthHeight[0]
		$ret.Height = $widthHeight[1]
	}

	if (-not $ret.InstanceLimit -or $ret.InstanceLimit -lt 0) {
		$ret.InstanceLimit = 0
	}

	if ($servers) {
		$ret.Servers += @($servers)
	}

	if ($wgNames) {
		$ret.WorkerGroups += @($wgNames)
	}
    
	#[bool]($myObject.PSobject.Properties.name -match "myPropertyNameToTest")
	if($null -eq $ret.ApplicationId){
		$ret.ApplicationId = [guid]::NewGuid()	
	}
	# if ( -Not $ret.PSObject.Properties[ 'ApplicationId' ] -or $ret.ApplicationId -eq $null) {
	# 	## generate the same file name rather than random - names should be unique
	# 	$ret.ApplicationId = (Get-Hash -textToHash $ret.Name)
	# }
	
	if ( -Not $ret.PSObject.Properties[ 'WindowType' ] -or $ret.WindowType -eq $null) {
		$ret.WindowType = 'Normal'
	}
	

	return $ret
}

function GetGroupServers ([string] $group) {
	Log -type "INFO" -message "Searching for servers on active directory group '$($group)' ..."
	if (-not $group) {
		Log -type "ERROR" -message "Input group was null or empty!"
		throw "Group is null"
	}

	if ($group.Contains('\')) {
		$group = $group.Substring($group.IndexOf('\') + 1)
	}
	elseif ($group.Contains('/')) {
		$group = $group.Substring($group.IndexOf('/') + 1)
	}

	Log -type "INFO" -message "Searching for group '$($group)' first ..."
	$searcher = LDAPSearcher($null)
	$searcher.Filter = ("(&(objectclass=group)(name=$group))")
	Log -type "INFO" -message "LDAP search query -> '$($searcher.Filter)'"

	$groupObj = $searcher.FindOne()

	if ($groupObj) {
		Log -type "INFO" -message "Found matching group!"
	}
	else {
		Log -type "INFO" -message "No matching group found!"
		return $null
	}

	$dgname = $groupObj.Properties.distinguishedname

	Log -type "INFO" -message "Now searching for servers on group using distinguished name ..."
	$searcher.Filter = ("(&(objectclass=computer)(memberof=$dgname))")
	Log -type "INFO" -message "LDAP search query -> '$($searcher.Filter)'"

	$ret = $searcher.FindAll()
	if ($ret.Count -eq 0) {
		Log -type "INFO" -message "No servers were found on $($dgname)"
		return $null
	}
	Log -type "INFO" -message "$($ret.Count) servers found on $($group)"
	return $ret
}

function CreateXAWorkGroup {
    Param
    (
        [System.Xml.XmlElement] $xml ,
        $clixml
    )
	$super = CreateXAObject($xml)
	$ret = @{
		Type         = $super.Type
		Name         = $clixml | Select-Object -ExpandProperty WorkerGroupName -ErrorAction SilentlyContinue
		RefId        = $super.RefId
		xml          = $super.xml
		MachineName  = $clixml | Select-Object -ExpandProperty MachineName -ErrorAction SilentlyContinue
		OUs          = @()
		ServerGroups = @()
		ServerNames  = @()
	}
	$OUs = $object.OUs
	if ($OUs) {
		$ret.OUs = @($OUs)
	}

	$serverGroups = $object.ServerGroups
	if ($serverGroups) {
		$ret.ServerGroups = @($serverGroups)
	}

	$servers = $object.ServerNames
	if ($servers) {
		$ret.ServerNames = @($servers)
	}
	else {
	     $ret.ServerNames = @($ret.MachineName) # if servernames is empty use machine name instead as a server
	}
	return $ret
}

<#
.SYNOPSIS
Initializes the database schema.

.DESCRIPTION
The schema is structred based on relations between entities. i.e. Workgroup has many servers, but server can be only in one group, and so on.
Please see the documentation for the visual representation of the DB structure.

.EXAMPLE
$db = InitializeDatabase
#>
function InitializeDatabase() {
	Log -type "INFO" -message "Initializng Database schema ..."
	[System.Data.DataSet] $dataset = New-Object System.Data.DataSet

	$tbl_zone = CreateTable -TableName "tbl_zone" -Columns @(
		"Name", "MachineName", "RASId"
	)
	$tbl_zone.Columns['Name'].Unique = $true


	$tbl_server = CreateTable -TableName "tbl_server" -Columns @(
		"Name", "MachineName", "ZoneName", "RASId", "WorkgroupName"
	)
	$tbl_server.Columns['Name'].Unique = $true

	$tbl_workgroup = CreateTable -TableName "tbl_workgroup" -Columns @(
		"Name", "Description", "MachineName", "RASId"
	)
	$tbl_workgroup.Columns['Name'].Unique = $true

	$tbl_wg_server = CreateTable -TableName "tbl_wg_server" -Columns @(
		"WorkgroupName", "ServerName"
	)

	$tbl_ou = CreateTable -TableName "tbl_ou" -Columns @(
		"Name", "GUID", "OctetString", "WorkgroupName", "RASId"
	)

	$tbl_serverGroup = CreateTable -TableName "tbl_serverGroup" -Columns @(
		"Name", "WorkgroupName", "RASId"
	)
	$tbl_serverGroup.Columns['Name'].Unique = $true

	$tbl_ADMachine = CreateTable -TableName "tbl_ADMachine" -Columns @(
		"Name", "OUName", "OUGuid", "ServerGroupName", "RASId"
	)

	$tbl_folder = CreateTable -TableName "tbl_folder" -Columns @(
		"Name", "Path", "Parent", "RASId", "IsAdministrative"
	)
	$tbl_folder.Columns['Path'].Unique = $true
	$tbl_folder.Columns["IsAdministrative"].DataType = [bool]

	$tbl_application = CreateTable -TableName "tbl_application" -Columns @(
		"Id", "Name", "Description", "WorkingDirectory", "Enabled", "WindowType", "MaximizedOnStartup",
		"WaitOnPrinterCreation", "Width", "Height", "AddToClientDesktop", "InstanceLimit", "AddToClientStartMenu", "ColorDepth",
		"MultipleInstancesPerUserAllowed", "Extensions", "UserFilter", "UserFilterByAccountName", "FolderPath", "Parameters", "Target", "Type", "RASFolderID", "IconPath", "RASId"
	)
	$tbl_application.Columns["Target"].DataType = [string]
	$tbl_application.Columns["Target"].DefaultValue = [string]::Empty

	$tbl_application.Columns["ColorDepth"].DataType = [string]
	$tbl_application.Columns["ColorDepth"].DefaultValue = [string]::Empty
	$tbl_application.Columns["ColorDepth"].AllowDBNull = $true
	
	$tbl_application.Columns["WindowType"].DataType = [string]
	$tbl_application.Columns["WindowType"].DefaultValue = [string]::Empty

	$tbl_application.Columns["Parameters"].DataType = [string]
	# $tbl_application.Columns["ServerAndGroup"].DataType = [bool]
	$tbl_application.Columns["AddToClientDesktop"].DataType = [bool]
	$tbl_application.Columns["AddToClientDesktop"].DefaultValue = $false
	
	$tbl_application.Columns["Enabled"].DataType = [bool]
	$tbl_application.Columns["Enabled"].DefaultValue = $false
	
	$tbl_application.Columns["AddToClientStartMenu"].DataType = [bool]
	$tbl_application.Columns["AddToClientStartMenu"].DefaultValue = $true
	
	$tbl_application.Columns["MultipleInstancesPerUserAllowed"].DataType = [bool]
	$tbl_application.Columns["MultipleInstancesPerUserAllowed"].DefaultValue = $false
	
	$tbl_application.Columns["WaitOnPrinterCreation"].DataType = [bool]
	$tbl_application.Columns["WaitOnPrinterCreation"].DefaultValue = $false
	
	$tbl_application.Columns["RASFolderID"].DataType = [string]
	# $tbl_application.Columns["RASFolderID"].DefaultValue = 0
	
	$tbl_application.Columns["InstanceLimit"].DataType = [int]
	$tbl_application.Columns["InstanceLimit"].DefaultValue = 0
	
	$tbl_application.Columns["UserFilter"].DataType = [string]
	$tbl_application.Columns["UserFilter"].AllowDBNull = $true
	$tbl_application.Columns["UserFilter"].DefaultValue = [System.DBNull]::Value

	$tbl_application.Columns["UserFilterByAccountName"].DataType = [string]
	$tbl_application.Columns["UserFilterByAccountName"].AllowDBNull = $true
	$tbl_application.Columns["UserFilterByAccountName"].DefaultValue = [System.DBNull]::Value


	$tbl_app_server = CreateTable -TableName "tbl_app_server" -Columns @(
		"AppId", "AppName", "ServerName"
	)

	$tbl_app_workgroup = CreateTable -TableName "tbl_app_workgroup" -Columns @(
		"AppId", "AppName", "WorkgroupName"
	)

	$dataset.Tables.Add($tbl_zone)
	$dataset.Tables.Add($tbl_server)
	$dataset.Tables.Add($tbl_workgroup)
	$dataset.Tables.Add($tbl_wg_server)
	$dataset.Tables.Add($tbl_ou)
	$dataset.Tables.Add($tbl_serverGroup)
	$dataset.Tables.Add($tbl_ADMachine)
	$dataset.Tables.Add($tbl_folder)
	$dataset.Tables.Add($tbl_application)
	$dataset.Tables.Add($tbl_app_server)
	$dataset.Tables.Add($tbl_app_workgroup)
	Log -type "INFO" -message "In-memory database setup is complete."
	return $dataset
}

function ParseZones([string] $xmlPath, [System.Data.DataSet] $table) {
	Log -type "INFO" -message "Parsing XA zones ..." -fnName "ParseZones"
	$xml = LoadXML($xmlPath)
	$zones = @($xml.Objs.ChildNodes)
    $zonesFromXML = Import-Clixml -Path $xmlPath

	[System.Data.DataTable]$tbl_zone = $table.Tables["tbl_zone"]

	for ($i = 0; $i -lt $zones.Count; $i++) {
		$zone = CreateXAZone -xml $zones[$i] -clixml $zonesFromXML[$i]
		$row = $tbl_zone.NewRow()
		$row.Name = $zone.ZoneName
		$row.MachineName = $zone.MachineName
		[void]$tbl_zone.Rows.Add($row)
	}
	Log -type "INFO" -message "$($tbl_zone.Rows.Count) zones were added to 'tbl_zone'"
}

function ParseServers ([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing XA servers ... " -fnName "ParseServers"

	$xml = LoadXML($xmlPath)
	$servers = @($xml.Objs.ChildNodes)
    $serversFromXML = Import-Clixml -Path $xmlPath

	[System.Data.DataTable] $tbl_server = $db.Tables["tbl_server"]
	for ($i = 0; $i -lt $servers.Count; $i++) {
		$server = CreateXAServer -xml $servers[$i] -clixml $serversFromXML[$i]
		$row = $tbl_server.NewRow()
		$row.Name = $server.ServerName
		$row.ZoneName = $server.ZoneName
		$row.MachineName = $server.MachineName
		$row.WorkgroupName = [string]::Empty
		$tbl_server.Rows.Add($row)
	}
	Log -type "INFO" -message "$($tbl_server.Rows.Count) servers were added to 'tbl_server'"
}
function ParseWorkgroups ([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing XA worker groups ..." -fnName "ParseWorkgroups"

	$xml = LoadXML($xmlPath)
	$workgroups = @($xml.Objs.ChildNodes)
    $workgroupsFromXml = Import-Clixml -Path $xmlPath

	[System.Data.DataTable] $tbl_workgroup = $db.Tables["tbl_workgroup"]
	[System.Data.DataTable] $tbl_wg_server = $db.Tables["tbl_wg_server"]

	for ($i = 0; $i -lt $workgroups.Count; $i++) {
		$workgroup = CreateXAWorkGroup -xml $workgroups[$i] -clixml $workgroupsFromXml[$i]

		foreach ($server in $workgroup.ServerNames) {
			$wgsrv_row = $tbl_wg_server.NewRow()
			$wgsrv_row.WorkgroupName = $workgroup.Name
			$wgsrv_row.ServerName = $server
			$tbl_wg_server.Rows.Add($wgsrv_row)
		}

		$row = $tbl_workgroup.NewRow()
		$row.Name = $workgroup.Name
		$row.MachineName = $workgroup.MachineName
		$tbl_workgroup.rows.Add($row)
	}
	Log -type "INFO" -message "$($tbl_workgroup.Rows.Count) worker groups were added to 'tbl_workgroup'"
}

<#
.SYNOPSIS
This a vital algorithm which is called at the end of the ParseXMLtoDB.

.DESCRIPTION
The algorithm will go through each server in tbl_server and ensure that
the relation is one-to-many between tbl_wg_server and tbl_server. This algorithm
solves the difference in architechture between the two applications, where in 'Citrix XenApp'
an application can be publish from multiple servers as well as multiple groups. This algorithm
will extract the server entry as separate workgroup if a server is common to more than one group.

.NOTES
Servers extracted as a group a prefixed GRP_<originial server name>
#>
function WorkgroupServerCleanup ([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Force one to many relation for Workgroup ---< Server ..." -fnName "WorkgroupServerCleanup"
	[System.Data.DataTable] $tbl_wg_server = $db.Tables["tbl_wg_server"]
	[System.Data.DataTable] $tbl_server = $db.Tables["tbl_server"]

	# these tables will be updated with new group name references ...
	[System.Data.DataTable] $tbl_app_workgroup = $db.Tables["tbl_app_workgroup"]
	[System.Data.DataTable] $tbl_workgroup = $db.Tables["tbl_workgroup"]
	[System.Data.DataTable] $tbl_ou = $db.Tables["tbl_ou"]
	[System.Data.DataTable] $tbl_serverGroup = $db.Tables["tbl_serverGroup"]


	foreach ($server in $tbl_server) {
		[System.Data.DataRow[]] $groups = $tbl_wg_server.Select("[ServerName] like '$($server.Name)'")

		if ($groups.Count -eq 0) {
			continue
		}
		if ($groups.Count -eq 1) {
			# server.WorkgroupName is empty initially.
			$server.WorkgroupName = $groups[0].WorkgroupName
		}
		# if a server is common with more than one group
		else <#($groups.Count -gt 1)#> {
			$newGroupName = "GRP_$($server.Name)"
			$description = "$($newGroupName) was created because $($server.Name) was common to $($groups.Count) other groups : $([string]::Join(', ', $groups.WorkgroupName))"
			# Update tbl_server with new WorkgroupName
			$server.WorkgroupName = $newGroupName
			# Update tbl_wg_server
			# use the server name as a new group name of the form GRP_*
			$row = $tbl_wg_server.NewRow()
			$row.WorkgroupName = $newGroupName
			$row.ServerName = $server.Name
			$tbl_wg_server.Rows.Add($row)

			# Update tbl_workgroup
			$row = $tbl_workgroup.NewRow()
			$row.Name = $newGroupName
			$row.Description = $description
			$tbl_workgroup.Rows.Add($row)

			# Remove all the groups that are common to a server, and update references to other tables
			foreach ($group in $groups) {
				$app_workgroups = $tbl_app_workgroup.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
				$workgroup = $tbl_workgroup.Select("[Name] like '$($group.WorkgroupName)'")
				$ous = $tbl_ou.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
				$ad_groups = $tbl_serverGroup.Select("[WorkgroupName] like '$($group.WorkgroupName)'")

				[System.Data.DataRow[]] $wgs = $tbl_wg_server.Select("[WorkgroupName] like '$($group.WorkgroupName)'")

				# update reference for tbl_app_workgroup
				foreach	($row in $app_workgroups) {
					# If tbl_wg_server has only one entry for $group this indicates that this current $group is
					# associated with only one server. simply update the group name of the application.
					#
					# If tbl_wg_server has more than one entry for $group, this means that the $group
					# has unique servers associated with it appart from the common one. Therefore the application row needs to be coppied,
					# to preserve the reference to the unique servers of the group. The copied row will get its workgroup name
					# updated with the $newGroupName.
					if ($wgs.Count -eq 1) {
						$row.WorkgroupName = $newGroupName
					}
					else {
						$newRow = $tbl_app_workgroup.NewRow()
						$newRow.AppId = $row.AppId
						$newRow.AppName = $row.AppName
						$newRow.WorkgroupName = $newGroupName
						$tbl_app_workgroup.Rows.Add($newRow)
					}
				}

				# update reference for tbl_ou
				foreach ($row in $ous) {
					$row.WorkgroupName = $newGroupName
				}

				# update reference for tbl_serverGroup
				foreach ($row in $ad_groups) {
					$row.WorkgroupName = $newGroupName
				}

				# clean up the orphan groups
				$tbl_wg_server.Rows.Remove($group)

				# if this condition fails, it would mean that this workgroup has a server
				# which is not common to other workgroups. Therefore, we retain the workgroup in the
				# tbl_workgroup.
				if ($wgs.Count -eq 1) {
					foreach ($row in $workgroup) {
						$tbl_workgroup.Rows.Remove($row)
					}
				}
			}

		}
	}
	Log -type "INFO" -message "Created one-to-many link between 'tbl_workgroup' and 'tbl_server'"
}

function ParseOUs ([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing OUs ..." -fnName "ParseOUs"

	$xml = LoadXML($xmlPath)
	$workgroups = @($xml.Objs.ChildNodes)
    $workgroupsFromXml = Import-Clixml -Path $xmlPath

	[System.Data.DataTable] $tbl_ou = $db.Tables["tbl_ou"]

	for ($i = 0; $i -lt $workgroups.Count; $i++) {
		$workgroup = CreateXAWorkGroup -xml $workgroups[$i] -clixml $workgroupsFromXml[$i]
		foreach ($ou in $workgroup.OUs) {
			if (IsGUID($ou)) {
				$ou_obj = GetOUByGUID($ou)
			}
			else {
				try {
					$ou_obj = GetOUFromPath($ou)
				}
				catch {
					Log -type "ERROR" -message "Failed to get an OU from $ou. $($_)"
					continue
				}
				$ou = [guid]$ou_obj.Properties.objectguid[0]
			}
			$row = $tbl_ou.NewRow()
			$row.Name = [string]$ou_obj.Properties.name
			if ($row.Name -eq [string]::Empty) {
				$row.Name = $ou
			}
			$row.Name = "OU_$($row.Name)"
			$row.GUID = $ou
			$row.WorkgroupName = $workgroup.Name
			$row.OctetString = GUIDToOctetString($ou)
			$tbl_ou.Rows.Add($row)
		}
	}
	Log -type "INFO" -message "Successfully parsed $($tbl_ou.Rows.Count) OUs!"
}

function ParseServerGroups([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing Server groups (Active Directory Groups) ..."
	$xml = LoadXML($xmlPath)
	$workgroups = @($xml.Objs.ChildNodes)
    $workgroupsFromXml = Import-Clixml -Path $xmlPath

	[System.Data.DataTable] $tbl_serverGroup = $db.Tables["tbl_serverGroup"]

	for ($i = 0; $i -lt $workgroups.Count; $i++) {
		$workgroup = CreateXAWorkGroup -xml $workgroups[$i] -clixml $workgroupsFromXml[$i]
		foreach ($serverGroup in $workgroup.ServerGroups) {
			$row = $tbl_serverGroup.NewRow()
			$row.Name = "ADG_$serverGroup"
			$row.WorkgroupName = $workgroup.Name
			$tbl_serverGroup.Rows.Add($row)
		}
	}

	Log -type "INFO" -message "Successfully parsed $($tbl_serverGroup.Rows.Count) server groups!"
}

function ParseADMachines ([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing Active directory machines (from OUs and AD Groups) ..."
	$tbl_ou = $db.Tables["tbl_ou"]
	$tbl_serverGroup = $db.Tables["tbl_serverGroup"]
	$tbl_ADMachine = $db.Tables["tbl_ADMachine"]

	Log -type "INFO" -message "Searching for machines on server groups ..."
	foreach ($sg_row in $tbl_serverGroup.Rows) {
		$servers = GetGroupServers($sg_row.Name)
		$count = 0;
		foreach ($server in $servers) {
			$row = $tbl_ADMachine.NewRow()
			$row.Name = [string]$server.Properties.name
			$row.OUGuid = [string]::Empty
			$row.ServerGroupName = $sg_row.Name
			$tbl_ADMachine.Rows.Add($row)
			$count = $count + 1
		}
		Log -type "INFO" -message "$($count) machines found on $($sg_row.Name)!"
	}

	Log -type "INFO" -message "Searching for machines on OUs ..."
	foreach ($ou_row in $tbl_ou.Rows) {
		$ou = GetOUByGUID($ou_row.GUID)
		$servers = GetOUServers($ou.Path)
		$count = 0
		if (-not $servers) {
			if ([string]$ou_row.Name -eq [string]::Empty) {
				Log -type "INFO" -message "OU $($ou_row.GUID) has no servers!"
			}
			else {
				Log -type "INFO" -message "OU $($ou_row.Name) has no servers!"
			}
			continue
		}
		foreach ($server in $servers) {
			$row = $tbl_ADMachine.NewRow()
			$row.Name = [string]$server.Properties.name
			$row.OUGuid = $ou_row.GUID
			$row.OUName = $ou_row.Name
			$row.ServerGroupName = [string]::Empty
			$tbl_ADMachine.Rows.Add($row)
			$count = $count + 1
			if ([string]$ou_row.Name -eq [string]::Empty) {
				Log -type "INFO" -message "$($count) machines found in OU ---> $($ou_row.GUID)!"
			}
			else {
				Log -type "INFO" -message "$($count) machines found in OU ---> $($ou_row.Name)!"
			}
		}
	}
}

function ExtractFolders([PSCustomObject]$app, [System.Data.DataSet]$db) {
    if( $null -ne $app -and  $app.PSObject.Properties[ 'FolderPath' ] -and -Not [string]::IsNullOrEmpty( $app.FolderPath ) ) {
	    [System.Collections.ArrayList]$folders = ($app.FolderPath -Split '[\\/]+' ) ## split on / or \ as XA 6.x was / but XD 7.x is \
        ## Commented out for XD 7.x as we don't have an empry folder

		if ($app.PSObject.Properties.ColorDepth -Contains "ColorDepth" -and -not [string]::IsNullOrWhiteSpace($app.ColorDepth)) ## Color Depth is only available for version 6.X
		{
			$folders.RemoveAt(0) # remove the XenApp root folder.
		}

	    # first check if the app has a client folder.
	    if ([string]$app.ClientFolder) {
		    $tbl_folder = $db.Tables['tbl_folder']
		    $row = $tbl_folder.NewRow()
		    $row.Name = $app.ClientFolder
		    $folderPath = [string]::Empty

		    # if path to application has more than 0 folders
		    # Folder1/Folder2/application
		    if ($folders.Count -gt 0 ) {
			    $folderPath = [string]::Join("/", $folders.ToArray())
			    # full path to client folder.
			    $row.Path = [string]::Join("/", @($folderPath, $app.ClientFolder))
		    }
		    # if path to application is made up of 0 folders
		    # the path to application will be ClientFolder/
		    else {
			    $row.Path = $app.ClientFolder
		    }

		    $row.Parent = $folderPath
		    $row.IsAdministrative = $false
		    try {
			    $tbl_folder.Rows.Add($row)
		    }
		    catch {
			    Log -type "INFO" -message "Folder $($row.Name) was already added [expected]"
		    }
	    }

	    while ($folders.Count -ne 0) {
            if( -Not [string]::IsNullOrEmpty( $folders[$folders.Count - 1] )) { ## do not process empty path elements
		        $row = $tbl_folder.NewRow()
		        $row.Name = $folders[$folders.Count - 1]
		        $row.Path = [string]::Join("/", $folders.ToArray())
		        $row.IsAdministrative = $true
		        $folders.RemoveAt($folders.Count - 1)
		        $row.Parent = [string]::Join("/", $folders.ToArray())
		        try {
			        $tbl_folder.Rows.Add($row)
		        }
		        catch {
			        Log -type "INFO" -message "Folder $($row.Name) was already added [expected]"
		        }
            }
            else {
		        $folders.RemoveAt($folders.Count - 1)
            }
	    }
    }
}

function ParseAppFolders([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing Application folders ..." -fnName "ParseAppFolders"
	$xml = LoadXML($xmlPath)
	$applications = @($xml.Objs.ChildNodes)
    $applicationsFromXml = Import-Clixml -Path $xmlPath
	[System.Data.DataTable]$tbl_folder = $db.Tables["tbl_folder"]

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication -xml $applications[$i] -clixml $applicationsFromXml[$i]
		ExtractFolders -app $app -db $db
	}
}

function GetParameters ([string] $cle) {
	if (-not $cle) {
		return $null
	}
	if ($cle.StartsWith('"')) {
		$i = $cle.IndexOf('"')
		$j = $cle.IndexOf('"', $i+1)
		if ($j -eq ($cle.Length - 1)) {
			return $null
		}
		$cle = $cle.Remove($i, $j+2)
		if ($cle.StartsWith(" ")) {
			$cle = $cle.Remove(0, 1)
		}
		if($cle.Contains("'")){
			$cle = $cle.Replace("'",'')
		}
	}
	else {
		$j = $cle.IndexOf(' ')
		if ($j -lt 0) {
			return $null
		}
		$cle = $cle.Remove(0, ++$j)
	}
	return $cle
}

function GetTarget ([string] $cle) {
	if (-not $cle) {
		return $null
	}
	if ($cle.StartsWith('"')) {
		$i = $cle.IndexOf('"')
		$j = $cle.IndexOf('"', ++$i)
		if ($j -ne $cle.Length - 1) {
			$cle = $cle.Substring($i, ++$j)
		}
	}
	else {
		$j = $cle.IndexOf(' ')
		if ($j -ne -1) {
			$cle = $cle.Substring(0, ++$j)
		}
	}

	if ($cle.EndsWith(" ")) {
		$cle = $cle.Remove($cle.Length - 1)
	}
	$cle = $cle.Replace('"', '')
	return $cle
}
$settings = @{
	ScriptPath = ""
	IconPath   = ""
	FarmInfo   = $null
}

function SetSettings ([hashtable] $set) {
	$settings.ScriptPath = $set.ScriptPath
	$settings.IconPath = $set.IconPath
	return $settings
}

function ParseApplications ([string] $xmlPath, [System.Data.DataSet]$db) {
	Log -type "INFO" -message "Parsing applications ..." -fnName "ParseApplications"
	$xml = LoadXML($xmlPath)
	$applications = @($xml.Objs.ChildNodes)
    $applicationsFromXml = Import-Clixml -Path $xmlPath

	$tbl_application = $db.Tables["tbl_application"]

	$dir = $settings.IconPath
	if (! ( Test-Path -Path $dir -PathType Container)) {
		New-Item -Path $dir -ItemType Directory | Out-Null
	}
    ## in case in an FS provider path format
	$dir = (Resolve-Path $dir) -replace '^Microsoft\.PowerShell\.Core\\FileSystem::'

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication -xml $applications[$i] -clixml $applicationsFromXml[ $i ]
		
		$parameters = GetParameters -cle (GetPropertyOrNull $app "CommandLineExecutable")
		$target = GetTarget -cle (GetPropertyOrNull $app "CommandLineExecutable")

		$bytes = $null
		if($app.IconData -is [System.String]){
			$bytes = [System.Convert]::FromBase64String($app.IconData)

		}else{
			$bytes = $app.IconData
		}
        if( $bytes ) {
			#$hash = "1A495F7E01F6D5FCA260152C6EF8D3E992648571CEF6D70129C779772E84AC3A"
			$hash = Get-Hash $bytes
			$path = "$($settings.IconPath)\$($hash).ico"
			[System.IO.File]::WriteAllBytes("$dir/$($hash).ico", $bytes)
        }else{
			$path = $null
		}
		$row = $tbl_application.NewRow()
		$row.Id = GetPropertyOrNull $app "ApplicationId"
		$row.Name = GetPropertyOrNull $app "Name"
		$row.Description = GetPropertyOrNull $app "Description"
		$row.Width = GetPropertyOrNull $app "Width"
		$row.Height = GetPropertyOrNull $app "Height"
		$row.Enabled = GetPropertyOrDBNull $app "Enabled"
		$row.Parameters = $parameters
		$row.ColorDepth = GetPropertyOrNull $app "ColorDepth"
		$row.Extensions = GetPropertyOrNull $app "Extensions"
		
		if ( $app.PSObject.Properties[ 'UserFilterByAccountName' ] -and $app.UserFilterByAccountName) {
			$userFilterAccountNameCSV = [string]::Join(",", $app.UserFilterByAccountName)
			$row.UserFilterByAccountName = $userFilterAccountNameCSV
		}

		if ( $app.PSObject.Properties[ 'UserFilter' ] -and $app.UserFilter) {
			$userFilterCSV = [string]::Join(",", $app.UserFilter)
			$row.UserFilter = $userFilterCSV
		}
		
		$row.InstanceLimit = GetPropertyOrDBNull $app "InstanceLimit" 
		$row.WorkingDirectory = GetPropertyOrNull $app "WorkingDirectory"
		$row.AddToClientDesktop = GetPropertyOrDBNull $app "AddToClientDesktop"
		$row.AddToClientStartMenu = GetPropertyOrDBNull $app "AddToClientStartMenu"
		$row.MultipleInstancesPerUserAllowed = GetPropertyOrDBNull $app "MultipleInstancesPerUserAllowed"
		$row.WaitOnPrinterCreation = GetPropertyOrDBNull $app "WaitOnPrinterCreation"
		$row.MaximizedOnStartup = GetPropertyOrDBNull $app "MaximizedOnStartup"
		$row.WindowType = GetPropertyOrDBNull $app "WindowType"	
		$row.IconPath = $path
		
		#[System.Collections.ArrayList]$folders = $(if( $app.PSObject.Properties[ 'UserFilterByAccountName' ] -and -Not [string]::IsNullOrEmpty( $app.FolderPath )) { $app.FolderPath.Split('/') } )
		[System.Collections.ArrayList]$folders = $null

		$folderpath = GetPropertyOrNull $app "FolderPath"
		if(![string]::IsNullOrEmpty($folderpath)){
			[System.Collections.ArrayList]$folders = $folderpath -Split '[\\/]+'
		}


		if( $folders -and $folders.Count -gt 0 ) {
			if ($app.PSObject.Properties.ColorDepth -Contains "ColorDepth" -and -not [string]::IsNullOrWhiteSpace($app.ColorDepth)) ## Color Depth is only available for version 6.X
			{
				$folders.RemoveAt(0) 
			} 
        }

		if ([string](GetPropertyOrNull $app "ClientFolder")) {
			if (-Not $folders -or $folders.Count -eq 0) {
				$row.FolderPath = $app.ClientFolder
			}
			else {
				$row.FolderPath = [string]::Join("/", @([string]::Join("/", $folders.ToArray()), $app.ClientFolder))
			}
		}
		elseif( $folders ) {
			$row.FolderPath = [string]::Join("/", $folders.ToArray())
		}

		if ((GetPropertyOrNull $app "ApplicationType") -eq "Content") {
			$target = $app[$app.length - 1].ContentAddress
		}
		$row.Target = $target
		$row.Type = GetPropertyOrNull $app "ApplicationType"

		$tbl_application.Rows.Add($row)
	}
	Log -type "INFO" -message "Parseing applications completed."
}

function GetPropertyOrNull([object]$object, [string]$property) {
	if($object.PSobject.Properties.name -match $property) {
		return $object.$property
	} Else {
		return $null
	}
}

function GetPropertyOrDBNull([object]$object, [string]$property) {
	$ret = GetPropertyOrNull $object $property
	if($null -eq $ret){
		return [DBNull]::Value
	}else{
		return $ret
	}
}

function ParseAppServers ([string] $xmlPath, [System.Data.DataSet]$db) {
	Log -type "INFO" -message "Parsing application servers ..." -fnName "ParseAppServers"
	$xml = LoadXML($xmlPath)
	$applications = @($xml.Objs.ChildNodes)
    $applicationsFromXml = Import-Clixml -Path $xmlPath

	$tbl_app_server = $db.Tables["tbl_app_server"]

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication -xml $applications[$i] -clixml $applicationsFromXml[$i]

		foreach ($serverName in (GetPropertyOrNull $app "Servers")) {
			$row = $tbl_app_server.NewRow()
			$row.AppId = $app.ApplicationId
			$row.AppName = $app.Name
			$row.ServerName	= $serverName
			$tbl_app_server.Rows.Add($row)
		}
	}
	Log -type "INFO" -message "Created one-to-many link between 'tbl_application' and 'tbl_app_server'! $($tbl_app_server.Rows.Count) rows added."
}

function ParseAppWorkgroup ([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing application workgroups ..." -fnName "ParseAppWorkgroup"

	$xml = LoadXML($xmlPath)
	$applications = @($xml.Objs.ChildNodes)
    $applicationsFromXml = Import-Clixml -Path $xmlPath

	$tbl_app_workgroup = $db.Tables["tbl_app_workgroup"]

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication -xml $applications[$i] -clixml $applicationsFromXml[$i]

		foreach ($workgroup in (GetPropertyOrNull $app "WorkerGroups")) {
			$row = $tbl_app_workgroup.NewRow()
			$row.AppId = $app.ApplicationId
			$row.AppName = $app.Name
			$row.WorkgroupName = $workgroup
			$tbl_app_workgroup.Rows.Add($row)
		}
	}
	Log -type "INFO" -message "Created one-to-many link between 'tbl_application' and 'tbl_app_workgroup'! $($tbl_app_workgroup.Rows.Count) rows added."
}

function ParseXMLtoDB ([string] $zonesXmlPath, [string] $serversXmlPath, [string] $workgroupXmlPath, [string] $applicationXmlPath, [string] $farmXmlPath = "") {
	if ($farmXmlPath) {
		$xml = LoadXML($farmXmlPath)
        $farmFromXML = Import-Clixml -Path $farmXmlPath
		$farms = @($xml.Objs.ChildNodes)
		$settings.FarmInfo = CreateXAFarm -xml $farms[0] -clixml $farmFromXML
	}

	[System.Data.DataSet]$Database = InitializeDatabase

	if ([System.IO.File]::Exists($zonesXmlPath)) {
		Log -type "INFO" -message "PARSING ZONES"
		ParseZones  $zonesXmlPath $Database
	}

	if ([System.IO.File]::Exists($serversXmlPath)) {
		Log -type "INFO" -message "PARSING SERVERS"
		ParseServers $serversXmlPath $Database
	}

	if ([System.IO.File]::Exists($workgroupXmlPath)) {
		Log -type "INFO" -message "PARSING Workergroups"
		ParseWorkgroups $workgroupXmlPath $Database
	
		Log -type "INFO" -message "PARSING OUs"
		ParseOUs $workgroupXmlPath $Database
	
		Log -type "INFO" -message "PARSING Server groups"
		ParseServerGroups $workgroupXmlPath $Database
	}

	Log -type "INFO" -message "PARSING AD Machines"
	ParseADMachines $Database

	if ([System.IO.File]::Exists($applicationXmlPath)) {
		Log -type "INFO" -message "PARSING Application Folders"
		ParseAppFolders $applicationXmlPath $Database

		Log -type "INFO" -message "PARSING Applications"
		ParseApplications $applicationXmlPath $Database

		Log -type "INFO" -message "PARSING Application servers"
		ParseAppServers $applicationXmlPath $Database

		Log -type "INFO" -message "PARSING Applications"
		ParseAppWorkgroup $applicationXmlPath $Database
	}

	Log -type "INFO" -message "Workgroup server cleanup."
	WorkgroupServerCleanup $Database

	return $Database
}

function InitializeScript ([scriptblock]$appendToMain) {
	if ([System.IO.File]::Exists($settings.ScriptPath)) {
		[System.IO.File]::Delete($settings.ScriptPath)
	}

	$date = Get-Date
	WriteScript @"
<#
.SYNOPSIS
Citrix migration script
Script was generated on $($date)
.DESCRIPTION
This is an auto-generated script.
$(
	if ($settings.FarmInfo) {
		"Script was generated based on the exported settings for:`n"
		"Farm : $($settings.FarmInfo.FarmName)`n"
		"Server Version: $($settings.FarmInfo.ServerVersion)`n"
		"Machine Name: $($settings.FarmInfo.MachineName)`n"
	}
)
.PARAMETER server
RAS remote server ip or name

.PARAMETER username
RAS username

.PARAMETER password
RAS password

.EXAMPLE
./ImportToRas.ps1

.NOTES
Be aware that any removal of commands, will require the user to search for any missing reference of variables.
#>
"@

	WriteScript @"
`$PSADMIN_MIN_VERSION = @(16, 2)
`$verbosePreference = "Continue"
`$Global:FEATURES_16_5 = `$false
function Initialize() {

	Import-Module $PSAdminModule -ErrorVariable "PSAdminError" -ErrorAction SilentlyContinue
	if (`$PSAdminError) {
		Write-Host "Parallels RAS PowerShell Module is not installed on this system." -ForegroundColor Red
		Write-Host "Error: `$PSAdminError" -ForegroundColor Red
		return `$false
}


	`$str = Get-RASVersion
	`$spaceIndex = `$str.IndexOf(' ')
	`$str = `$str.Substring(0, `$spaceIndex)
	`$versionArray = `$str.Split('.')

	if (`$versionArray[0] -lt `$PSADMIN_MIN_VERSION[0] -or 
		( `$versionArray[0] -eq `$PSADMIN_MIN_VERSION[0] -and `$versionArray[1] -lt `$PSADMIN_MIN_VERSION[1] )) {
		Write-Host "Parallels RAS Version `$str is not supported. Please install `$([string]::Join('.', `$PSADMIN_MIN_VERSION))."
		return `$false
	}
	elseif (`$versionArray[0] -eq `$PSADMIN_MIN_VERSION[0] -and `$versionArray[1] -eq `$PSADMIN_MIN_VERSION[1]) {
		`$Global:FEATURES_16_5 = `$false
		Write-Host "Parallels RAS version `$str does not support migration of icons" -ForegroundColor Red
	}
	else {
		`$Global:FEATURES_16_5 = `$true
	}

	return `$true
}
"@

	WriteScript @"
function Main () {
	Param(
		[Parameter(Mandatory=`$true)]
		[string] `$server,
		[Parameter(Mandatory=`$true)]
		[string] `$username,
		[Parameter(Mandatory=`$true)]
		[securestring] `$password
	)
	Write-Host "$("="*100)" -ForegroundColor Cyan
	Write-Host "Welcome to Parallels Citrix-RAS Migration tool." -ForegroundColor Green
	Write-Host "Script was generated on $((Get-Date))" -ForegroundColor Green
	Write-Host "$("="*100)" -ForegroundColor Cyan
	try {
		New-RASSession -Server `$server -Username `$username -Password `$password
	}
	catch {
		Write-Host "Error: `$(`$_)" -ForegroundColor Red
		return
	}
"@
	$appendToMain.Invoke()
	WriteScript @"
	Write-Host "$("="*100)" -ForegroundColor Cyan
	Write-Host "Migration complete." -ForegroundColor Green
	Write-Host "$("="*100)" -ForegroundColor Cyan
}
"@

	WriteScript @"
if (-not (Initialize)) {
	return
}

Main
"@
}
function CreateStream ([scriptblock] $writeCallback ) {
	$writer = New-Object 'System.IO.StreamWriter' $settings.ScriptPath, $true
	$writer.AutoFlush = $true
	$writeCallback.Invoke($writer)
	$writer.Close()
}

$Global:BAKE_TO_FILE = $true
$Global:VAR_COUNTER = 0

function WriteComment ([string] $comment, [int] $tabs = 1) {
	CreateStream -writeCallback {
		Param(
			[System.IO.StreamWriter]$writer
		)
		$writer.WriteLine("")
		$writer.WriteLine("$("`t"*$tabs)$('#' * 100)")
		$writer.WriteLine("$("`t"*$tabs)# $($comment)$(' '*(100-$comment.Length-3))#")
		$writer.WriteLine("$("`t"*$tabs)$('#'*100)")
	}
}
function WriteCommentLite([string] $comment, [int] $tabs = 1) {
	CreateStream -writeCallback {
		Param(
			[System.IO.StreamWriter]$writer
		)
		$writer.WriteLine("")
		$writer.WriteLine("$("`t"*$tabs)# [$($comment)] ")
	}
}

function WriteScript ([string] $script, [string] $comment) {
	CreateStream -writeCallback {
		Param(
			[System.IO.StreamWriter]$writer
		)

		if ($comment) {
			$writer.WriteLine("# $($comment)")
		}
		$writer.WriteLine("$($script.ToString())")
	}
}

function GetVarNameFromCmdlet ([string] $cmdLet) {
	$dashIndex = $cmdLet.IndexOf('-')
	$varName = "`$$($cmdLet.Substring(++$dashIndex))"
	if (-not $vars.ContainsKey($varName)) {
		$vars.Add($varName, 0);
	}
	else {
		$vars[$varName] += 1
	}
	$varName = "$($varName)_$($vars[$varName])"
	return $varName
}

$vars = New-Object 'System.Collections.Generic.Dictionary[string, int]'
function WriteToScript ([string]$command, [switch] $useVar, [int]$tabs = 1) {
try {
     ## replace single quote with double to escape it since the result becomes a script line
     ## using positive lookbehind and lookahead of an alphabetic character since do not want to change where ' is at start or end of a string
	[scriptblock]$script = [scriptblock]::Create( ($command -replace "(?<=\w)'(?=\w)" , "''") )
} catch {
    Write-Warning -Message "Problem with command: $command"
    Write-Error $_
    return
}
	$spaceIndex = $command.IndexOf(' ')
    $varname = $null

	if ($spaceIndex -ne -1) {
		$cmdLetString = $command.Substring(0, $spaceIndex)
	}
	else {
		$cmdLetString = $command
	}

	if (-not $useVar) {
		CreateStream -writeCallback {
			Param(
				[System.IO.StreamWriter] $writer
			)
			$writer.WriteLine("$("`t"*$tabs)$($script.ToString()) | Out-Null")
		}
	}
	else {
		$varName = GetVarNameFromCmdlet $cmdLetString
		if ($cmdLetString.StartsWith("New-")) {
			$cmd = "$("`t"*$tabs)$($varName) = ($($script.ToString())).Id"
		}
		else {
			$cmd = "$("`t"*$tabs)$($varName) = $($script.ToString())"
		}

		CreateStream -writeCallback {
			Param(
				[System.IO.StreamWriter] $writer
			)
			$writer.WriteLine($cmd)
		}
	}
	return $varName
}


function MigrateFolders([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Migrating folders ..." -fnName "MigrateFolders"
	[System.Data.DataTable]$tbl_folder = $db.Tables["tbl_folder"]
	[System.Data.DataTable]$tbl_application = $db.Tables["tbl_application"]

	$root_folders = $tbl_folder.Select("[Parent] = ''")
	$child_folders = $tbl_folder.Select("[Parent] not like ''")
	Log -type "INFO" -message "First adding root folders ..."
	WriteComment -comment "FOLDER MIGRATION COMMANDS START"
	foreach ($folder in $root_folders) {
		WriteCommentLite "Commands related to root folder$($folder.Name)"
		try {
			if ([bool]$folder.IsAdministrative) {

				$res = WriteToScript "New-RASPubFolder -AdminOnly -Name '$($folder.Name)'" -useVar
			}
			else {
				$res = WriteToScript "New-RASPubFolder -Name '$($folder.Name)'" -useVar
			}
			$apps = $tbl_application.Select("[FolderPath] like '$($folder.Path)'")
			Log -type "INFO" -message "Updating 'tbl_applications' RASFolder ID ..."
			foreach ($app in $apps) {
				$app["RASFolderID"] = $res
			}
			$folder["RASId"] = $res
		}
		catch {
			Log -type "ERROR" -message "Failed to create new pubfolder" $_.Exception
		}

	}

	Log -type "INFO" -message "Now adding sub folders ..."
	foreach	($folder in $child_folders) {
		WriteCommentLite "Commands related to child folder $($folder.Name)"
		try {
			$parent = $tbl_folder.Select("Path = '$($folder.Parent)'")
			$rasFolder = WriteToScript "Get-RASPubFolder -Id $($parent.RASId)" -useVar

			if ([bool]$folder.IsAdministrative) {
				$res = WriteToScript "New-RASPubFolder -AdminOnly -Name '$($folder.Name)' -ParentFolder $rasFolder" -useVar
			}
			else {
				$res = WriteToScript "New-RASPubFolder -Name '$($folder.Name)' -ParentFolder $rasFolder" -useVar
			}

			Log -type "INFO" -message "Created '$($folder.Name)'. Path ---> $($folder.Path). RAS Folder ID ---> $($res)"
			$folder["RASId"] = $res
			$apps = $tbl_application.Select("[FolderPath] like '$($folder.Path)'")
			Log -type "INFO" -message "Updating 'tbl_applications' RASFolder ID ..."
			foreach ($app in $apps) {
				$app["RASFolderID"] = $res
			}
		}
		catch [Exception] {
			Log -type "ERROR" -message "$($_)"
			throw
		}
	}
	Log -type "INFO" -message "Folders Migrated! Applications are now aware of RAS folder IDs and can be linked to them."
}

function PublishLocalApp($app, $parentfolder){
	$httpsItems = $app.ItemArray | Where-Object { $_ -match "^https://" }
	$httpItems = $app.ItemArray | Where-Object { $_ -match "^http://" }
	
	if(-Not ($httpItems -OR $httpsItems)){
		return
	}

	if($httpsItems){
		$cmd = "New-RASPubLocalApp -Name '$($app.Name)' -URL '$httpsItems'"
	}

	if($httpItems){
		$cmd = "New-RASPubLocalApp -Name '$($app.Name)' -URL '$httpItems'"
	}

	if ($app.Description -ne '') {
		$cmd += " -Description '$($app.Description)'"
	}

	if($app.Enabled -eq $true){
		$cmd += " -EnabledMode Enabled"
	}

	if($app.Enabled -eq $false){
		$cmd += " -EnabledMode Disabled"
	}
	
	if ($parentfolder -ne $null) {
		$cmd += " -ParentFolder $parentFolder"
	}
	$res = WriteToScript $cmd -useVar
	
	$features16_5 = "Set-RASPubLocalApp -id $res -Icon '$($app.IconPath -replace '^Microsoft\.PowerShell\.Core\\FileSystem::')'"
	WriteScript @"
	if (`$FEATURES_16_5) {
		$features16_5
	}
"@
	return $res
}

function PublishRDSApp ($app, $from, $publishSource, $parentFolder) {
	$cmd = "New-RASPubRDSApp -PublishFrom '$from' -Target '$($app.Target)' -Name '$($app.Name)'"
	if ($app.Description -ne '') {
		$cmd += " -Description '$($app.Description)'"
	}
	if ($parentFolder -ne $null) {
		$cmd += " -ParentFolder $parentFolder"
	}

	if ($publishSource) {
		$publishSource = [string]::Join(',', $publishSource)
	}
	switch ($from) {
		Group {

			$cmd += " -PublishFromGroup $publishSource"
		}
		Server {

			$cmd += " -PublishFromServer $publishSource"
		}
	}

	$res = WriteToScript $cmd -useVar

	$windowType = "Normal"
	if ($app.MaximizedOnStartup) {
		$windowType = "Maximized"
	}

	WriteToScript "Set-RASPubRDSApp -Id $res -CreateShortcutOnDesktop `$$($app.AddToClientDesktop)  -InheritShortcutDefaultSettings `$false -CreateShortcutInStartFolder `$$($app.AddToClientStartMenu) -OneInstancePerUser `$$($app.MultipleInstancesPerUserAllowed) -InheritLicenseDefaultSettings `$false -ConCurrentLicenses $($app.InstanceLimit) -InheritDisplayDefaultSettings `$false -WaitForPrinters `$$($app.WaitOnPrinterCreation) -EnabledMode `Enabled -StartIn '$($app.WorkingDirectory)' -Parameters '$($app.Parameters)' -WinType $windowType"
	$features16_5 = "Set-RASPubRDSApp -id $res -Icon '$($app.IconPath -replace '^Microsoft\.PowerShell\.Core\\FileSystem::')'"
	if ($app.ColorDepth -ne [System.DBNull]::Value -and $app.ColorDepth -ne '') {
		$features16_5 += " -ColorDepth '$($app.ColorDepth)'"
	}

	if ($app.Extensions -ne [System.DBNull]::Value -and $app.Extensions -ne '') {
		$features16_5 += " -FileExtensions '$($app.Extensions)'"
	}
	WriteScript @"
	if (`$FEATURES_16_5) {
		$features16_5
	}
"@
	return $res
}

function PublishRDSDesktop ($app, $from, $publishSource, $parentFolder) {
	if ($publishSource) {
		$publishSource = [string]::Join(',', $publishSource)
	}
	$cmd = "New-RASPubRDSDesktop -Name '$($app.Name)'"

	if ($app.Description -ne '') {
		$cmd += " -Description '$($app.Description)'"
	}

	if ($parentFolder -ne $null) {
		$cmd += " -ParentFolder $parentFolder"
	}

	switch ($from) {
		Group {
			$cmd += " -PublishFrom '$from'"
			$cmd += " -PublishFromGroup $publishSource"
			$res = WriteToScript $cmd -useVar
		}
		Server {
			$cmd += " -PublishFrom '$from'"
			$cmd += " -PublishFromServer $publishSource"
			$res = WriteToScript $cmd -useVar
		}
		All {
			$res = WriteToScript $cmd -useVar
		}
	}

	WriteToScript "Set-RASPubRDSDesktop -Id $($res) -CreateShortcutOnDesktop `$$($app.AddToClientDesktop) -InheritShortcutDefaultSettings `$false -CreateShortcutInStartUpFolder `$$($app.AddToClientStartMenu) -Width $($app.Width) -Height $($app.Height) -DesktopSize Custom -EnabledMode `Enabled"
	return $res
}

function AddPubItemUserFilter ($userFilter, $rdsApp, $userFilterAccountNames = $null ) {
	if ($userFilter -ne [System.DBNull]::Value) {
		WriteToScript "Set-RASPubItemFilter -Id $rdsApp -Default Deny"
		WriteToScript "Add-RASRule -RuleName rule -ObjType PubItem -Id $rdsApp"
		$rule = WriteToScript "Get-RASRule -ObjType PubItem -Id $rdsApp" -useVar
		WriteToScript "Set-RASCriteria -ObjType PubItem -Id $rdsApp -RuleId $rule.Id -SecurityPrincipalsEnabled `$true -SecurityPrincipalsMatchingMode IsOneOfTheFollowing"
		WriteScript @"
		if (`$FEATURES_16_5) {
"@
		foreach ($sid in $userFilter.Split(',')) {
			try {
				WriteToScript "Add-RASCriteriaSecurityPrincipal -ObjType PubItem -Id $rdsApp -SID '$sid' -RuleId $rule.Id"
			}
			catch {
				Log -type "WARNING" -message "Error occured during Add-RASPubItemUserFilter" -exception $_.Exception
			}
		}
		WriteScript @"
	}
"@
	}
	if ($userFilterAccountNames -ne [System.DBNull]::Value) {
		WriteScript @"
		else {
"@
		foreach ($acc in $userFilterAccountNames.Split(',')) {
			try {
				WriteToScript "Add-RASCriteriaSecurityPrincipal -ObjType PubItem -Id $rdsApp -Account '$acc' -RuleId $rule.Id"
			}
			catch {
				Log -type "WARNING" -message "Error occured during Add-RASPubItemUserFilter" -exception $_.Exception
			}
		}
		WriteScript @"
	}
"@
	}
}

function PublishPubItem ($app, $from, $publishSource, $parentFolder) {
    $res = $null
	switch -regex ($app.Type) {
		"ServerInstalled|HostedOnDesktop" {
			if ( $null -eq $publishSource -or $publishSource.Count -eq 0) {
				$res = PublishRDSApp -app $app -from ALL -parentFolder $parentFolder
			}
			else {
				$res = PublishRDSApp -app $app -from $from  -publishSource $publishSource -parentFolder $parentFolder
			}
		}
		"ServerDesktop" {
			if ( $null -eq $publishSource -or $publishSource.Count -eq 0) {
				$res = PublishRDSDesktop -app $app -from ALL -publishSource $publishSource -parentFolder $parentFolder
			}
			else {
				$res = PublishRDSDesktop -app $app -from $from -publishSource $publishSource -parentFolder $parentFolder
			}
		}
		"^Content$" {
			$res = PublishRDSApp -app $app -from All -parentFolder $parentFolder
		}
		"^PublishedContent$" {
			$res = PublishLocalApp -app $app -parentFolder $parentFolder
		}
        Default {
            Write-Warning -Message "Unsupported app type `"$($app.Type)`" found for app `"$($app.name)`""
        }
	}
	return $res
}

function ExtractRelatedApplicationGroups ($app) {
	[System.Data.DataTable]$tbl_ou = $db.Tables["tbl_ou"]
	[System.Data.DataTable]$tbl_serverGroup = $db.Tables["tbl_serverGroup"]
	[System.Data.DataTable]$tbl_wg_server = $db.Tables["tbl_wg_server"]
	[System.Data.DataTable]$tbl_app_workgroup = $db.Tables["tbl_app_workgroup"]

	$app_workgroups = $tbl_app_workgroup.Select("[AppName] like '$($app.Name)'")

	$RDSHostPools = @()
	foreach ($group in $app_workgroups) {
		# try {
		$res = $tbl_wg_server.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
		if ($res -and $res.Count -gt 0) {
			$rdsHostPool = WriteToScript "Get-RASRDSHostPool -Name '$($group.WorkgroupName)'" -useVar
			$RDSHostPools += $rdsHostPool
		}
		# }
		# catch {
		# Log -type "WARNING" -message "Group for application '$($app.Name)' not found." -exception $_.Exception
		# }

		# foreach workgroup that we found, get the AD and OU groups
		$ADGroups = $tbl_serverGroup.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
		foreach ($adgroup in $ADGroups) {
			# try {
			$rdsHostPool = WriteToScript "Get-RASRDSHostPool -Name '$($adgroup.Name)'" -useVar
			# $RDSHostPools += GetVarNameFromCmdlet "Get-RASRDSHostGroup"
			$RDSHostPools += $rdsHostPool
			# }
			# catch {
			# Log -type "ERROR" -message "Failed to get RDS Group." -exception $_.Exception
			# }
		}

		$OUs = $tbl_ou.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
		foreach ($ou in $OUs) {
			try {
				$name = [string]$ou.Name
				if (-not [string]$ou.Name) {
					$name = $ou.GUID
				}

				$rdsHostPool = WriteToScript "Get-RASRDSHostPool -Name '$name'" -useVar
				$RDSHostPools += $rdsHostPool
			}
			catch {
				Log -type "ERROR" -message "Failed to get RDS Group." -exception $_.Exception
			}
		}
	}
	return $RDSHostPools
}

function ExtractRelatedApplicationServers ($app) {
	[System.Data.DataTable]$tbl_app_server = $db.Tables["tbl_app_server"]
	$app_servers = $tbl_app_server.Select("[AppName] like '$($app.Name)'")

	$RDSServers = @()
	foreach ($server in $app_servers) {
		try {
			$rds = WriteToScript "Get-RASRDS -Server '$($server.ServerName)'" -useVar
			$RDSServers += $rds
		}
		catch {
			Log -type "ERROR" -message "Failed to get RDS Group." -exception $_.Exception
		}
	}
	return $RDSServers
}

function MigrateApplications([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Migrating applications ..." -fnName "MigrateApplications"
	[System.Data.DataTable]$tbl_application = $db.Tables["tbl_application"]
	WriteComment -comment "APPLICATION MIGRATION COMMANDS START"

	foreach ($app in $tbl_application) {
		WriteCommentLite -comment "Commands related to $($app.Name)"
		$RDSHostPool = ExtractRelatedApplicationGroups -app $app
		$RDSServers = ExtractRelatedApplicationServers -app $app
		try {
            $res = $null
			$parentFolder = $null
			if ($app.RASFolderID -ne 0 -and $app.RASFolderID -ne [DBnull]::Value) {
				$parentFolder = WriteToScript "Get-RASPubFolder -Id $($app.RASFolderID)" -useVar
			}

			# if application is published from servers AND groups
			if ( $RDSHostPool -and ( ($RDSServers -and $RDSServers.Count -gt 0 -and $RDSHostPool.Count -gt 0) -or $RDSHostPool.Count -gt 0)) {
				Log -type "INFO" -message "Publishing application $($app.Name) from groups"
				Log -type "INFO" -message "RDSGroups.Count : $($RDSHostPool.Count)."
				$res = PublishPubItem -app $app -from Group -publishSource $RDSHostPool -parentFolder $parentFolder
				$app.RASId = $res
			}
			else {
				Log -type "INFO" -message "Publishing application $($app.Name) from servers."
				Log -type "INFO" -message "RDSServers.Count : $(if( $RDSServers ) { $RDSServers.Count } else { 0 })."

				$res = PublishPubItem -app $app -from Server -publishSource $RDSServers -parentFolder $parentFolder
				$app.RASId = $res
			}
            if( $res ) {
			AddPubItemUserFilter -userFilter $app.UserFilter -rdsApp $res -userFilterAccountNames $app.UserFilterByAccountName
            } else {
                Log -type "WARNING" -message "Failed to publish application $($app.Name) of type $($app.Type)."
            }

		}
		catch {
			Log -type "ERROR" -message "Error occured during migration of applications" -exception $_.Exception
			# Write-Host $_.Exception.StackTrace
			throw ## TODO should this be a fatal error ?
		}
	}
	Log -type "INFO" -message "Applications migrated!"
}

function MigrateSites([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Migrating sites ..." -fnName "MigrateSites"
	$tbl_zone = $db.Tables["tbl_zone"]
	WriteComment "SITES MIGRATION COMMANDS START"
	# $tbl_server = $db.Tables["tbl_server"]
	try {
		$rasSite = WriteToScript "Get-RASSite -Id 1" -useVar
		# Write-HOst $rasSite.Name, "ASDASDASD"
	}
	catch {
		Log -type "ERROR" -message "Failed to get RAS site." -exception $_.Exception
		return
	}
	if(!$tbl_zone.Rows){
		if($($tbl_zone.Rows[0].Name)){
			WriteToScript "Set-Site $rasSite -NewName '$($tbl_zone.Rows[0].Name)'"
		}
	}

	foreach ($zone in $tbl_zone) {
		# $server = $tbl_server.Select("ZoneName = '$($zone.Name)'")
		try {
			$rasSite = $rasSite[0]
			$zone["RASId"] = $rasSite
		}
		catch [Exception] {
			Log -type "ERROR" -message "$($_)"
		}
	}
	Log -type "INFO" -message "Sites migrated!"
}

function MigrateServers([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Migrating servers" -fnName "MigrateServers"
	$tbl_server = $db.Tables["tbl_server"]
	# $tbl_zone = $db.Tables["tbl_zone"]
	$tbl_ADMachine = $db.Tables["tbl_ADMachine"]

	Log -type "INFO" -message "Migrating XAServers ..."
	WriteComment "SERVERS MIGRATION COMMANDS START"

	# Migarte XA Servers
	foreach	($server in $tbl_server) {
		WriteCommentLite -comment "Commands related $($server.Name)"
		# $siteId = $tbl_zone.Select("[Name] like '$($server.ZoneName)'")[0].RASId
		try {
			Log -type "INFO" -message "Adding new RDS for XA server ---> $($server.Name)"

			$RDSServer = WriteToScript "New-RASRDSHost -Server '$($server.Name)' -NoInstall"
			$server["RASId"] = $RDSServer
		}
		catch [Exception] {
			Log -type "ERROR" -message "Migrating XA server failed $($server.Name)"
			Log -type "ERROR" -message "Error was: $($_)"
		}
	}

	Log -type "INFO" -message "Migrating active directory servers ..."
	# Migrate AD servers
	foreach ($machine in $tbl_ADMachine) {
		try {
			Log -type "INFO" -message "Adding new RDS for AD server ---> $($machine.Name)"
			$RDSServer = WriteToScript "New-RASRDSHost -Server '$($machine.Name)' -NoInstall"
			$machine["RASId"] = $RDSServer
		}
		catch [Exception] {
			Log -type "WARNING" -message "Migrating AD server failed $($machine.Name)"
			Log -type "WARNING" -message "Error was: $($_)"
		}
	}
}

function MigrateGroups ([System.Data.DataSet] $db) {
	Log -type "INFO" -message "Migrating groups ..." -fnName "MigrateGroups"
	$tbl_workgroup = $db.Tables["tbl_workgroup"]
	$tbl_server = $db.Tables["tbl_server"]
	$tbl_ADMachine = $db.Tables["tbl_ADMachine"]

	Log -type "INFO" -message "Migrating worker groups ..."
	WriteComment -comment "GROUPS MIGRATION COMMANDS START"
	foreach ($workgroup in $tbl_workgroup) {
		try {
			[System.Data.DataRow[]]$servers = $tbl_server.Select("[WorkgroupName] like '$($workgroup.Name)'")
			if ($servers -and $servers.Count -gt 0) {
				$cmd = "New-RASRDSHostPool -Name '$($workgroup.Name)'"
				if ($workgroup.Description -ne [System.DBNull]::Value -and $workgroup.Description -ne '') {
					WriteCommentLite -comment $workgroup.Description
					$script = @"
	if (`$FEATURES_16_5) {
		Set-RDSGroup -Name '$($workgroup.Name)' -Description '$($workgroup.Description)'
	}
"@
				}
				$hostpool = WriteToScript $cmd
				WriteScript $script
				$workgroup["RASId"] = $hostpool

				foreach ($server in $servers) {
					WriteToScript "Move-RASRDSHostPoolMember -GroupName '$($workgroup.Name)' -RDSServer '$($server.Name)'"
				}
			}
		}
		catch [Exception] {
			Log -type "ERROR" -message "Failed to create new RDS group $($workgroup.Name)." -exception $_.Exception
		}
	}

	Log -type "INFO" -message "Migrating active directory groups ..."
	# Migrate server groups
	$tbl_serverGroup = $db.Tables["tbl_serverGroup"]

	foreach ($serverGroup in $tbl_serverGroup) {
		try {
			Log -type "INFO" -message "Creating new RDS group ---> $($serverGroup.Name)"
			$hostpool = WriteToScript "New-RASRDSHostPool -Name '$($serverGroup.Name)'"
			$serverGroup["RASId"] = $hostpool
			$machines = $tbl_ADMachine.Select("[ServerGroupName] like '$($serverGroup.Name)'")
			foreach ($machine in $machines) {
				if ([string]$machine.Name -ne [string]::Empty) {
					WriteToScript "Move-RASRDSHostPoolMember -GroupName '$($serverGroup.Name)' -RDSServer '$($machine.Name)'"
				}
			}
		}
		catch [Exception] {
			Log -type "ERROR" -message "$($_)"
		}
	}

	Log -type "INFO" -message "Migrating OUs ..."
	# Migrate OUs
	$tbl_ou = $db.Tables["tbl_ou"]
	foreach ($ou in $tbl_ou) {
		try {
			if ([string]$ou.Name -eq [string]::Empty) {
				$hostpool = WriteToScript "New-RASRDSHostPool -Name '$($ou.GUID)'"
				$machines = $tbl_ADMachine.Select("[OUGuid] like '$($ou.GUID)'")
			}
			else {
				$hostpool = WriteToScript "New-RASRDSHostPool -Name '$($ou.Name)'"
				$machines = $tbl_ADMachine.Select("[OUName] like '$($ou.Name)'")
			}
			$ou["RASId"] = $hostpool

			foreach ($machine in $machines) {
				if ([string]$machine.Name -ne [string]::Empty) {
					WriteToScript "Move-RASRDSHostPoolMember -GroupName '$($ou.Name)' -RDSServer '$($machine.Name)'"
				}
			}
		}
		catch [Exception] {
			Log -type "ERROR" -message "$($_)"
		}
	}
	Log -type "INFO" -message "Groups migrated!"
}
