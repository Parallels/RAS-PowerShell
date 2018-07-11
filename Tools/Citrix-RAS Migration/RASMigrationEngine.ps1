. "./Utilities.ps1"
<#
.SYNOPSIS
XenApp to RAS migration script module.

.DESCRIPTION
This is a migration script to transfer XA -Applications, -Workergroups, -Zones, -Servers.
!!IMPORTANT
At the time of writing the script, BY DEFAULT, the first Zone name (in zones xml) will be used as the default RASSite name.
Due to difference in software architechture it is not possible to do a 1:1 mapping.

.EXAMPLE
# Import the script
. "./CitrixMigrate.ps1"

# Database schema is created, XMLs are parsed and inserted into the database.
$db = ParseXMLtoDB -zonesXmlPath "./XmlMock/zones.xml" -serversXmlPath "./XmlMock/servers.xml" -workgroupXmlPath "./XmlMock/workergroups.xml" -applicationXmlPath "./XmlMock/applications.xml"
# Generate script.
InitializeScript -appendToMain {
	MigrateSites($db)
	MigrateServers($db)
	MigrateGroups($db)
	MigrateFolders($db)
	MigrateApplications($db)
	WriteToScript "Remove-RASSession"
}

.NOTES
General notes
Utilities.ps1 is required to operate this script!
#>

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
		ObjectName = $xml.SelectSingleNode("ToString").InnerText
		RefId      = $xml.RefId
		xml        = $xml
	}
	return $ret
}

function CreateXAFarm([System.Xml.XmlElement]$xml) {
	$super = CreateXAObject $xml
	return @{
		Type          = $super.Type
		ObjectName    = $super.ObjectName
		RefId         = $super.RefId
		FarmName      = $xml.Props.SelectSingleNode("S[@N='FarmName']").InnerText
		ServerVersion = $xml.Props.SelectSingleNode("Version[@N='ServerVersion']").InnerText
		MachineName   = $xml.Props.SelectSingleNode("S[@N='MachineName']").InnerText
	}
}

function CreateXAZone([System.Xml.XmlElement] $xml) {
	$super = CreateXAObject($xml)
	return [pscustomobject] @{
		Type        = $super.Type
		ObjectName  = $super.ObjectName
		RefId       = $super.RefId
		ZoneName    = $xml.Props.SelectSingleNode("S[@N='ZoneName']").InnerText
		MachineName = $xml.Props.SelectSingleNode("S[@N='MachineName']").InnerText
	}
}

function CreateXAServer ([System.Xml.XmlElement] $xml) {
	$super = CreateXAObject($xml)
	$obj = @{
		Type        = $super.Type
		Name        = $super.ObjectName
		RefId       = $super.RefId
		MachineName = $xml.Props.SelectSingleNode("S[@N='MachineName']").InnerText
		ServerName  = $xml.Props.SelectSingleNode("S[@N='ServerName']").InnerText
		ServerId    = $xml.Props.SelectSingleNode("S[@N='ServerId']").InnerText
		FolderPath  = $xml.Props.SelectSingleNode("S[@N='FolderPath']").InnerText
		ZoneName    = $xml.Props.SelectSingleNode("S[@N='ZoneName']").InnerText
	}
	return $obj
}

function CreateXAApplication ([System.Xml.XmlElement] $xml) {
	$super = CreateXAObject($xml)

	$ret = [PSCustomObject] @{
		Type                            = $super.Type
		Name                            = $xml.Props.SelectSingleNode("S[@N='DisplayName']").InnerText
		Description                     = $xml.Props.SelectSingleNode("S[@N='Description']").InnerText
		WorkingDirectory                = $xml.Props.SelectSingleNode("S[@N='WorkingDirectory']").InnerText
		RefId                           = $super.RefId
		FolderPath                      = $xml.Props.SelectSingleNode("S[@N='FolderPath']").InnerText
		ClientFolder                    = $xml.Props.SelectSingleNode("S[@N='ClientFolder']").InnerText
		StartMenuFolder                 = $xml.Props.SelectSingleNode("S[@N='StartMenuFolder']").InnerText
		CommandLineExecutable           = $xml.Props.SelectSingleNode("S[@N='CommandLineExecutable']").InnerText
		ApplicationType                 = $xml.Props.SelectSingleNode("Obj[@N='ApplicationType']").SelectSingleNode("ToString").InnerText
		IconData                        = $xml.Props.SelectSingleNode("BA[@N='IconData']").InnerText
		ApplicationId                   = $xml.Props.SelectSingleNode("S[@N='ApplicationId']").InnerText
		ContentAddress                  = $xml.Props.SelectSingleNode("S[@N='ContentAddress']").InnerText
		InstanceLimit                   = [int]$xml.Props.SelectSingleNode("I32[@N='InstanceLimit']").InnerText
		UserFilter                      = $xml.Props.SelectNodes("Obj[@N='Accounts']/LST/Obj/Props/S[@N='AccountId']").InnerText
		UserFilterByAccountName         = $xml.Props.SelectNodes("Obj[@N='Accounts']/LST/Obj/Props/S[@N='AccountName']").InnerText
		ColorDepth                      = $xml.Props.SelectNodes("Obj[@N='ColorDepth']")
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
	
	if ($ret.ColorDepth -ne $null) {
		$ret.ColorDepth = $ret.ColorDepth.SelectSingleNode("ToString").InnerText
	}

	if ($ret.ColorDepth.Count -eq 0) {
		$ret.ColorDepth = $null
	}


	if ($ret.ApplicationId -eq $null) {
		$ret.ApplicationId = [guid]::NewGuid()
	}

	$servers = $xml.Props.SelectNodes("Obj[@N='ServerNames']").LST.S
	$wgNames = $xml.SelectSingleNode("Props/Obj[@N='WorkerGroupNames']").LST.S
	$filetypes = $xml.Props.SelectNodes("Obj[@N='FileTypes']/LST/Obj/Props/Obj/LST/S").InnerText

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
	[bool]::TryParse($xml.Props.SelectSingleNode("B[@N='AddToClientDesktop']").InnerText, [ref] $tmpBool)
	$ret.AddToClientDesktop = $tmpBool

	[bool]::TryParse($xml.Props.SelectSingleNode("B[@N='AddToClientStartMenu']").InnerText, [ref] $tmpBool)
	$ret.AddToClientStartMenu = $tmpBool

	[bool]::TryParse($xml.Props.SelectSingleNode("B[@N='MultipleInstancesPerUserAllowed']").InnerText, [ref] $tmpBool)
	$ret.MultipleInstancesPerUserAllowed = $tmpBool

	[bool]::TryParse($xml.Props.SelectSingleNode("B[@N='WaitOnPrinterCreation']").InnerText, [ref] $tmpBool)
	$ret.WaitOnPrinterCreation = $tmpBool

	[bool]::TryParse($xml.Props.SelectSingleNode("B[@N='Enabled']").InnerText, [ref] $tmpBool)
	$ret.Enabled = $tmpBool

	[bool]::TryParse($xml.Props.SelectSingleNode("B[@N='MaximizedOnStartup']").InnerText, [ref] $tmpBool)
	$ret.MaximizedOnStartup = $tmpBool

	$windowType = $xml.Props.SelectSingleNode("S[@N='WindowType']").InnerText
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

function CreateXAWorkGroup ([System.Xml.XmlElement] $xml) {
	$super = CreateXAObject($xml)

	$ret = @{
		Type         = $super.Type
		Name         = $super.ObjectName
		RefId        = $super.RefId
		xml          = $super.xml
		MachineName  = $xml.Props.SelectSingleNode("S[@N='MachineName']").InnerText
		OUs          = @()
		ServerGroups = @()
		ServerNames  = @()
	}
	$OUs = $xml.SelectSingleNode("Props/Obj[@N='OUs']").LST.S
	if ($OUs) {
		$ret.OUs = @($OUs)
	}

	$serverGroups = $xml.SelectSingleNode("Props/Obj[@N='ServerGroups']").LST.S
	if ($serverGroups) {
		$ret.ServerGroups = @($serverGroups)
	}

	$servers = $xml.Props.SelectSingleNode("Obj[@N='ServerNames']").LST.S
	if ($servers) {
		$ret.ServerNames = @($servers)
	}
	# else {
	#     $ret.ServerNames = @($ret.MachineName) # if servernames is empty use machine name instead as a server
	# }
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

	[System.Data.DataTable]$tbl_zone = $table.Tables["tbl_zone"]

	for ($i = 0; $i -lt $zones.Count; $i++) {
		$zone = CreateXAZone($zones[$i])
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

	[System.Data.DataTable] $tbl_server = $db.Tables["tbl_server"]
	for ($i = 0; $i -lt $servers.Count; $i++) {
		$server = CreateXAServer($servers[$i])
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

	[System.Data.DataTable] $tbl_workgroup = $db.Tables["tbl_workgroup"]
	[System.Data.DataTable] $tbl_wg_server = $db.Tables["tbl_wg_server"]

	for ($i = 0; $i -lt $workgroups.Count; $i++) {
		$workgroup = CreateXAWorkGroup ($workgroups[$i])

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

	[System.Data.DataTable] $tbl_ou = $db.Tables["tbl_ou"]

	for ($i = 0; $i -lt $workgroups.Count; $i++) {
		$workgroup = CreateXAWorkGroup ($workgroups[$i])
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

	[System.Data.DataTable] $tbl_serverGroup = $db.Tables["tbl_serverGroup"]

	for ($i = 0; $i -lt $workgroups.Count; $i++) {
		$workgroup = CreateXAWorkGroup ($workgroups[$i])
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
	[System.Collections.ArrayList]$folders = $app.FolderPath.Split('/')
	$folders.RemoveAt(0) # remove the XenApp root folder.

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
}

function ParseAppFolders([string] $xmlPath, [System.Data.DataSet] $db) {
	Log -type "INFO" -message "Parsing Application folders ..." -fnName "ParseAppFolders"
	$xml = LoadXML($xmlPath)
	$applications = @($xml.Objs.ChildNodes)

	[System.Data.DataTable]$tbl_folder = $db.Tables["tbl_folder"]

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication($applications[$i])
		ExtractFolders -app $app -db $db
	}
}

function GetParameters ([string] $cle) {
	if (-not $cle) {
		return
	}
	if ($cle.StartsWith('"')) {
		$i = $cle.IndexOf('"')
		$j = $cle.IndexOf('"', ++$i)
		if ($j -eq $cle.Length - 1) {
			return
		}
		$cle = $cle.Remove($i, ++$j)
		if ($cle.StartsWith(" ")) {
			$cle = $cle.Remove(0, 1)
		}
	}
	else {
		$j = $cle.IndexOf(' ')
		if ($j -lt 0) {
			return
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

	$tbl_application = $db.Tables["tbl_application"]

	$dir = $settings.IconPath
	if (![System.IO.Directory]::Exists($dir)) {
		[System.IO.Directory]::CreateDirectory($dir) | Out-Null
	}
	$dir = Resolve-Path $dir
	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication($applications[$i])

		$parameters = GetParameters -cle $app.CommandLineExecutable
		$target = GetTarget -cle $app.CommandLineExecutable

		$bytes = [System.Convert]::FromBase64String($app.IconData)
		$path = "$($settings.IconPath)/$($app.ApplicationId).ico"
		[System.IO.File]::WriteAllBytes("$dir/$($app.ApplicationId).ico", $bytes)
		$row = $tbl_application.NewRow()
		$row.Id = $app.ApplicationId
		$row.Name = $app.Name
		$row.Description = $app.Description
		$row.Width = $app.Width
		$row.Height = $app.Height
		$row.Enabled = $app.Enabled
		$row.Parameters = $parameters
		$row.ColorDepth = $app.ColorDepth
		$row.Extensions = $app.Extensions
		
		if ($app.UserFilterByAccountName) {
			$userFilterAccountNameCSV = [string]::Join(",", $app.UserFilterByAccountName)
			$row.UserFilterByAccountName = $userFilterAccountNameCSV
		}

		if ($app.UserFilter) {
			$userFilterCSV = [string]::Join(",", $app.UserFilter)
			$row.UserFilter = $userFilterCSV
		}
		
		$row.InstanceLimit = $app.InstanceLimit
		$row.WorkingDirectory = $app.WorkingDirectory
		$row.AddToClientDesktop = $app.AddToClientDesktop
		$row.AddToClientStartMenu = $app.AddToClientStartMenu
		$row.MultipleInstancesPerUserAllowed = $app.MultipleInstancesPerUserAllowed
		$row.WaitOnPrinterCreation = $app.WaitOnPrinterCreation
		$row.MaximizedOnStartup = $app.MaximizedOnStartup
		$row.WindowType = $app.WindowType

		$row.IconPath = $path
		[System.Collections.ArrayList]$folders = $app.FolderPath.Split('/')
		$folders.RemoveAt(0)

		if ([string]$app.ClientFolder) {
			if ($folders.Count -eq 0) {
				$row.FolderPath = $app.ClientFolder
			}
			else {
				$row.FolderPath = [string]::Join("/", @([string]::Join("/", $folders.ToArray()), $app.ClientFolder))
			}
		}
		else {
			$row.FolderPath = [string]::Join("/", $folders.ToArray())
		}

		if ($app.ApplicationType -eq "Content") {
			$target = $app.ContentAddress
		}
		$row.Target = $target
		$row.Type = $app.ApplicationType

		$tbl_application.Rows.Add($row)
	}
	Log -type "INFO" -message "Parseing applications completed."
}

function ParseAppServers ([string] $xmlPath, [System.Data.DataSet]$db) {
	Log -type "INFO" -message "Parsing application servers ..." -fnName "ParseAppServers"
	$xml = LoadXML($xmlPath)
	$applications = @($xml.Objs.ChildNodes)

	$tbl_app_server = $db.Tables["tbl_app_server"]

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication($applications[$i])

		foreach ($serverName in $app.Servers) {
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

	$tbl_app_workgroup = $db.Tables["tbl_app_workgroup"]

	for ($i = 0; $i -lt $applications.Count; $i++) {
		$app = CreateXAApplication($applications[$i])

		foreach ($workgroup in $app.WorkerGroups) {
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

	if (![System.IO.File]::Exists($zonesXmlPath)) {
		throw "File '$($zonesXmlPath)' not found"
	}

	if (![System.IO.File]::Exists($serversXmlPath)) {
		throw "File '$($serversXmlPath)' not found"
	}

	if (![System.IO.File]::Exists($workgroupXmlPath)) {
		throw "File '$($workgroupXmlPath)' not found"
	}

	if (![System.IO.File]::Exists($applicationXmlPath)) {
		throw "File '$($applicationXmlPath)' not found"
	}

	if ($farmXmlPath) {
		$xml = LoadXML($farmXmlPath)
		$farms = @($xml.Objs.ChildNodes)
		$settings.FarmInfo = CreateXAFarm $farms[0]
	}

	[System.Data.DataSet]$Database = InitializeDatabase

	Log -type "INFO" -message "PARSING ZONES"
	ParseZones  $zonesXmlPath $Database

	Log -type "INFO" -message "PARSING SERVERS"
	ParseServers $serversXmlPath $Database

	Log -type "INFO" -message "PARSING Workergroups"
	ParseWorkgroups $workgroupXmlPath $Database

	Log -type "INFO" -message "PARSING OUs"
	ParseOUs $workgroupXmlPath $Database

	Log -type "INFO" -message "PARSING Server groups"
	ParseServerGroups $workgroupXmlPath $Database

	Log -type "INFO" -message "PARSING AD Machines"
	ParseADMachines $Database

	Log -type "INFO" -message "PARSING Application Folders"
	ParseAppFolders $applicationXmlPath $Database

	Log -type "INFO" -message "PARSING Applications"
	ParseApplications $applicationXmlPath $Database

	Log -type "INFO" -message "PARSING Application servers"
	ParseAppServers $applicationXmlPath $Database

	Log -type "INFO" -message "PARSING Applications"
	ParseAppWorkgroup $applicationXmlPath $Database

	Log -type "INFO" -message "Workgroup server cleanup."
	WorkgroupServerCleanup $Database

	return $Database
}

function InitializeScript ([scriptblock]$appendToMain) {
	if ([System.IO.File]::Exists("./MigrationScript.ps1")) {
		[System.IO.File]::Delete("./MigrationScript.ps1")
	}

	$scriptPath = "./MigrationScript.ps1"
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
./MigrationScript.ps1

.NOTES
Be aware that any removal of commands, will require the user to search for any missing reference of variables.
#>
"@

	WriteScript @"
`$PSADMIN_MIN_VERSION = @(16, 2)
`$verbosePreference = "Continue"
`$Global:FEATURES_16_5 = `$false
function Initialize() {

	Import-Module PSAdmin -ErrorVariable "PSAdminError" -ErrorAction SilentlyContinue
	if (`$PSAdminError) {
		Write-Host "Parallels RAS PowerShell Module is not installed on this system." -ForegroundColor Red
		Write-Host "Error: `$PSAdminError" -ForegroundColor Red
		return `$false
}


	`$str = Get-RASVersion
	`$spaceIndex = `$str.IndexOf(' ')
	`$str = `$str.Substring(0, `$spaceIndex)
	`$versionArray = `$str.Split('.')

	if (`$versionArray[0] -lt `$PSADMIN_MIN_VERSION[0] -or `$versionArray[1] -lt `$PSADMIN_MIN_VERSION[1]) {
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

	[scriptblock]$script = [scriptblock]::Create($command)
	$spaceIndex = $command.IndexOf(' ')

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

				$res = WriteToScript "New-PubFolder -AdminOnly -Name '$($folder.Name)'" -useVar
			}
			else {
				$res = WriteToScript "New-PubFolder -Name '$($folder.Name)'" -useVar
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
			$rasFolder = WriteToScript "Get-PubFolder -Id $($parent.RASId)" -useVar

			if ([bool]$folder.IsAdministrative) {
				$res = WriteToScript "New-PubFolder -AdminOnly -Name '$($folder.Name)' -ParentFolder $rasFolder" -useVar
			}
			else {
				$res = WriteToScript "New-PubFolder -Name '$($folder.Name)' -ParentFolder $rasFolder" -useVar
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

function PublishRDSApp ($app, $from, $publishSource, $parentFolder) {
	$cmd = "New-PubRDSApp -PublishFrom '$from' -Target '$($app.Target)' -Name '$($app.Name)'"
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

	WriteToScript "Set-PubRDSApp -Id $res -CreateShortcutOnDesktop `$$($app.AddToClientDesktop)  -InheritShorcutDefaultSettings `$false -CreateShortcutInStartFolder `$$($app.AddToClientStartMenu) -OneInstancePerUser `$$($app.MultipleInstancesPerUserAllowed) -InheritLicenseDefaultSettings `$false -ConCurrentLicenses $($app.InstanceLimit) -InheritDisplayDefaultSettings `$false -WaitForPrinters `$$($app.WaitOnPrinterCreation) -Enable `$$($app.Enabled) -StartIn '$($app.WorkingDirectory)' -Parameters '$($app.Parameters)' -WinType $windowType"
	$features16_5 = "Set-PubRDSApp -id $res -Icon '$($app.IconPath)'"
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
	$cmd = "New-PubRDSDesktop -Name '$($app.Name)'"

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

	WriteToScript "Set-PubRDSDesktop -Id $($res) -CreateShortcutOnDesktop `$$($app.AddToClientDesktop) -InheritShorcutDefaultSettings `$false -CreateShortcutInStartUpFolder `$$($app.AddToClientStartMenu) -Width $($app.Width) -Height $($app.Height) -DesktopSize Custom -Enable `$$($app.Enabled)"
	return $res
}

function AddPubItemUserFilter ($userFilter, $rdsApp, $userFilterAccountNames = $null ) {
	if ($userFilter -ne [System.DBNull]::Value) {
		WriteScript @"
		if (`$FEATURES_16_5) {
"@
		foreach ($sid in $userFilter.Split(',')) {
			try {
				WriteToScript "Add-PubItemUserFilter -Id $rdsApp -SID '$sid'"
			}
			catch {
				Log -type "WARNING" -message "Error occured during Add-PubItemUserFilter" -exception $_.Exception
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
				WriteToScript "Add-PubItemUserFilter -Id $rdsApp -Account '$acc'"
			}
			catch {
				Log -type "WARNING" -message "Error occured during Add-PubItemUserFilter" -exception $_.Exception
			}
		}
		WriteScript @"
	}
"@
	}
}

function PublishPubItem ($app, $from, $publishSource, $parentFolder) {
	switch ($app.Type) {
		"ServerInstalled" {
			if ($publishSource.Count -eq 0) {
				$res = PublishRDSApp -app $app -from ALL -parentFolder $parentFolder
			}
			else {
				$res = PublishRDSApp -app $app -from $from  -publishSource $publishSource -parentFolder $parentFolder
			}
		}
		"ServerDesktop" {
			if ($publishSource.Count -eq 0) {
				$res = PublishRDSDesktop -app $app -from ALL -publishSource $publishSource -parentFolder $parentFolder
			}
			else {
				$res = PublishRDSDesktop -app $app -from $from -publishSource $publishSource -parentFolder $parentFolder
			}
		}
		"Content" {
			$res = PublishRDSApp -app $app -from All -parentFolder $parentFolder
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

	$RDSGRoups = @()
	foreach ($group in $app_workgroups) {
		# try {
		$res = $tbl_wg_server.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
		if ($res -and $res.Count -gt 0) {
			$rdsGroup = WriteToScript "Get-RDSGroup -Name '$($group.WorkgroupName)'" -useVar
			$RDSGRoups += $rdsGroup
		}
		# }
		# catch {
		# Log -type "WARNING" -message "Group for application '$($app.Name)' not found." -exception $_.Exception
		# }

		# foreach workgroup that we found, get the AD and OU groups
		$ADGroups = $tbl_serverGroup.Select("[WorkgroupName] like '$($group.WorkgroupName)'")
		foreach ($adgroup in $ADGroups) {
			# try {
			$rdsGroup = WriteToScript "Get-RDSGroup -Name '$($adgroup.Name)'" -useVar
			# $RDSGRoups += GetVarNameFromCmdlet "Get-RDSGroup"
			$RDSGRoups += $rdsGroup
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

				$rdsGroup = WriteToScript "Get-RDSGroup -Name '$name'" -useVar
				$RDSGRoups += $rdsGroup
			}
			catch {
				Log -type "ERROR" -message "Failed to get RDS Group." -exception $_.Exception
			}
		}
	}
	return $RDSGRoups
}

function ExtractRelatedApplicationServers ($app) {
	[System.Data.DataTable]$tbl_app_server = $db.Tables["tbl_app_server"]
	$app_servers = $tbl_app_server.Select("[AppName] like '$($app.Name)'")

	$RDSServers = @()
	foreach ($server in $app_servers) {
		try {
			$rds = WriteToScript "Get-RDS -Server '$($server.ServerName)'" -useVar
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
		$RDSGroups = ExtractRelatedApplicationGroups -app $app
		$RDSServers = ExtractRelatedApplicationServers -app $app
		try {
			$parentFolder = $null
			if ($app.RASFolderID -ne 0 -and $app.RASFolderID -ne [DBnull]::Value) {
				$parentFolder = WriteToScript "Get-PubFolder -Id $($app.RASFolderID)" -useVar
			}

			# if application is published from servers AND groups
			if (($RDSServers.Count -gt 0 -and $RDSGroups.Count -gt 0) -or $RDSGroups.Count -gt 0) {
				Log -type "INFO" -message "Publishing application $($app.Name) from groups"
				Log -type "INFO" -message "RDSGroups.Count : $($RDSGroups.Count)."
				$res = PublishPubItem -app $app -from Group -publishSource $RDSGroups -parentFolder $parentFolder
				$app.RASId = $res
			}
			else {
				Log -type "INFO" -message "Publishing application $($app.Name) from servers."
				Log -type "INFO" -message "RDSServers.Count : $($RDSServers.Count)."

				$res = PublishPubItem -app $app -from Server -publishSource $RDSServers -parentFolder $parentFolder
				$app.RASId = $res
			}
			AddPubItemUserFilter -userFilter $app.UserFilter -rdsApp $res -userFilterAccountNames $app.UserFilterByAccountName
		}
		catch {
			Log -type "ERROR" -message "Error occured during migration of applications" -exception $_.Exception
			# Write-Host $_.Exception.StackTrace
			throw
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
		$rasSite = WriteToScript "Get-Site -Id 1" -useVar
		# Write-HOst $rasSite.Name, "ASDASDASD"
	}
	catch {
		Log -type "ERROR" -message "Failed to get RAS site." -exception $_.Exception
		return
	}

	WriteToScript "Set-Site $rasSite -NewName '$($tbl_zone.Rows[0].Name)'"

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

			$RDSServer = WriteToScript "New-RDS -Server '$($server.Name)' -NoInstall"
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
			$RDSServer = WriteToScript "New-RDS -Server '$($machine.Name)' -NoInstall"
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
				$cmd = "New-RDSGroup -Name '$($workgroup.Name)'"
				if ($workgroup.Description -ne [System.DBNull]::Value -and $workgroup.Description -ne '') {
					WriteCommentLite -comment $workgroup.Description
					$script = @"
	if (`$FEATURES_16_5) {
		Set-RDSGroup -Name '$($workgroup.Name)' -Description '$($workgroup.Description)'
	}
"@
				}
				$RDSGroup = WriteToScript $cmd
				WriteScript $script
				$workgroup["RASId"] = $RDSGroup

				foreach ($server in $servers) {
					WriteToScript "Add-RDSGroupMember -GroupName '$($workgroup.Name)' -RDSServer '$($server.Name)'"
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
			$RDSGroup = WriteToScript "New-RDSGroup -Name '$($serverGroup.Name)'"
			$serverGroup["RASId"] = $RDSGroup
			$machines = $tbl_ADMachine.Select("[ServerGroupName] like '$($serverGroup.Name)'")
			foreach ($machine in $machines) {
				if ([string]$machine.Name -ne [string]::Empty) {
					WriteToScript "Add-RDSGroupMember -GroupName '$($serverGroup.Name)' -RDSServer '$($machine.Name)'"
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
				$RDSGroup = WriteToScript "New-RDSGroup -Name '$($ou.GUID)'"
				$machines = $tbl_ADMachine.Select("[OUGuid] like '$($ou.GUID)'")
			}
			else {
				$RDSGroup = WriteToScript "New-RDSGroup -Name '$($ou.Name)'"
				$machines = $tbl_ADMachine.Select("[OUName] like '$($ou.Name)'")
			}
			$ou["RASId"] = $RDSGroup

			foreach ($machine in $machines) {
				if ([string]$machine.Name -ne [string]::Empty) {
					WriteToScript "Add-RDSGroupMember -GroupName '$($ou.Name)' -RDSServer '$($machine.Name)'"
				}
			}
		}
		catch [Exception] {
			Log -type "ERROR" -message "$($_)"
		}
	}
	Log -type "INFO" -message "Groups migrated!"
}