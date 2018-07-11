<#
.SYNOPSIS
CitrixMigrate.ps1 Unit tests.

.DESCRIPTION
This script contains various unit tests which check the integrity of Parsing and migration algorithms.

.EXAMPLE
./CitrixMigrationTests.ps1

.NOTES
This script includes general unit test functions created in UnitTestFramework.ps1.
To see how to use these functions, refer to the documentation in UnitTestFramework.ps1
#>
[System.IO.Directory]::SetCurrentDirectory($PSScriptRoot)
Import-Module PSAdmin
. "./UnitTestFramework.ps1"
. "./Utilities.ps1"
. "./RASMigrationEngine.ps1"
. "../../Scripts/Defaults.ps1"


# create table strucutre with no content.
[System.Data.DataSet]$db = InitializeDatabase
SetSettings @{
	ScriptPath = "./MigrationScript.ps1"
	IconPath   = "./icons"
}
$TestSuiteResults = @()
$TestSuiteResults += TestSuite -name "Citrix Migration Tool Parsing the XML." -unitTests @(
 {
		UnitTest "Load Xml" {
			[xml]$xml = LoadXML("./XmlMock/servers.xml")
			Assert -expression ($xml -ne $null) -message "Load XML returned null"
		}
	}
 {
		UnitTest "Create XAZone object from xml" {
			[xml] $xml = LoadXML('./XmlMock/zones.xml')

			$res = $xml.SelectSingleNode("Objs")
			Assert ($res -ne $null) -message "Xml must contain 'Objs' node."

			[System.Xml.XmlNodeList]$zones = $res.ChildNodes
			Assert ($zones.Count -gt 0) -message "'Objs' has no child nodes."

			foreach ($zone in $zones) {
				$zone = CreateXAZone($zone)
				Assert ($zone -ne $null) -message "CreateXAZone returned null."
				Assert ($zone.ObjectName -ne [string]::Empty) -message "'zone.ObjectName' was empty."
				Assert ($zone.RefId -ne [string]::Empty) -message "'zone.RefId' was empty."
				Assert ($zone.ZoneName -ne [string]::Empty) -message "'zone.ZoneName' was empty."
				Assert ($zone.MachineName -ne [string]::Empty) -message "'zone.MachineName' was empty."
			}
		}
	}
 {
		UnitTest "Create XA Server" {
			[xml] $xml = LoadXML('./XmlMock/servers.xml')

			$res = $xml.SelectSingleNode("Objs")
			Assert ($res -ne $null) -message "Xml must contain 'Objs' node"

			[System.Xml.XmlNodeList] $servers = $res.ChildNodes
			Assert ($servers.Count -gt 0) -message "'Objs' has no child nodes."

			foreach ($server in $servers) {
				$server = CreateXAServer($server)
				Assert ($server -ne $null) -message "'CreateXAServer returned null."
				Assert ($server.Name -ne [string]::Empty) -message "'server.Type' was empty."
				Assert ($server.RefId -ne [string]::Empty) -message "'server.Type' was empty."
				Assert ($server.MachineName -ne [string]::Empty) -message "'server.Type' was empty."
				Assert ($server.ServerName -ne [string]::Empty) -message "'server.Type' was empty."
				Assert ($server.ServerId -ne [string]::Empty) -message "'server.ServerId' was empty."
				Assert ($server.FolderPath -ne [string]::Empty) -message "'server.FolderPath' was empty."
				Assert ($server.ZoneName -ne [string]::Empty) -message "'server.ZoneName' was empty."
				Assert ($server.Version -ne [string]::Empty) -message "'server.Type' was empty."
			}
		}
	}
 {
		UnitTest "Create XA application" {
			[xml] $xml = LoadXML("./XmlMock/applications.xml")
			$objs = $xml.SelectSingleNode("Objs")
			Assert ($objs -ne $null) -message "Xml must contain 'Objs' node"

			[System.Xml.XmlNodeList]$applications = $objs.ChildNodes
			Assert ($applications.Count -gt 0) -message "'Objs' has no child nodes."

			foreach ($app in $applications) {
				$app = CreateXAApplication($app)

				Assert ($app -ne $null) -message "CreateXAApplication returned null."
				Assert ($app.Name -ne [string]::Empty) -message "app.Name was empty"
				Assert ($app.RefId -ne [string]::Empty) -message "app.RefId was empty"
				Assert ($app.FolderPath -ne [string]::Empty) -message "app.FolderPath was empty"
				Assert ($app.CommandLineExecutable -ne [string]::Empty) -message "app.CommandLineExecutable was empty"
				Assert ($app.ApplicationType -ne [string]::Empty) -message "app.ApplicationType was empty"
				Assert ($app.IconData -ne [string]::Empty) -message "app.IconData was empty"
				Assert ($app.ApplicationId -ne [string]::Empty) -message "app.ApplicationId was empty"
				Assert ($app.ContentAddress -ne [string]::Empty) -message "app.ContentAddress was empty"
			}
		}
	}
 {
		UnitTest "Create LDAP Searcher from path" {
			$domain = $domain = (Get-WmiObject Win32_ComputerSystem).Domain
			$path = "LDAP://$domain"
			[adsisearcher]$searcher = LDAPSearcherFromPath $path
			Assert ($searcher -ne $null) -message "LDAPSearcherFromPath returned null."
			$res = $searcher.FindAll()
			Assert ($res.Count -ne 0) -message "searcher.FindAll is emtpy"
		}
	}
 {
		UnitTest "Create LDAPSearcher" {
			$searcher = LDAPSearcher
			Assert ($searcher -ne $null) -message "LDAPSearcher returned null."
			$res = $searcher.FindAll()
			Assert ($res.Count -ne 0) -message "searcher.FindAll is emtpy"
		}
	}
 {
		UnitTest "Get OU by GUID" {
			$searcher = LDAPSearcher
			$searcher.Filter = "(&(objectclass=organizationalUnit))"
			$ou_a = $searcher.FindOne();
			Assert ($ou_a -ne $null) -message "OU was null."

			$objectguid_a = $ou_a.Properties["objectguid"]

			$ou_b = GetOUByGUID $objectguid_a[0]
			Assert ($ou_b -ne $null) -message "GetOUByGUID returned null."

			$objectguid_b = $ou_b.Properties.objectguid

			$objectguid_a = [string]::Join(' ', $objectguid_a[0])
			$objectguid_b = [string]::Join(' ', $objectguid_b[0])
			Assert ($objectguid_a[0] -eq $objectguid_b[0]) -message "ObjectGUIDs dont match"
		}
	}
 {
		UnitTest "Get OU Servers" {
			$domain = $domain = (Get-WmiObject Win32_ComputerSystem).Domain
			$path = "LDAP://$domain"
			$servers = GetOUServers $path
			Assert ($servers.Count -ne 0) -message "GetOUServers is returned empty."
		}
	}
 {
		UnitTest "Get Group Servers" {
			$searcher = LDAPSearcher
			$searcher.Filter = ("(&(objectclass=computer))")
			$servers = $searcher.FindOne()

			$searcher.Filter = ("(&(objectclass=group)(distinguishedName=$($servers.Properties.memberof[0])))")
			$group = $searcher.FindOne()

			$servers = GetGroupServers $group.Properties.name

			foreach ($server in $servers) {
				$dgname = $group.Properties.distinguishedname
				[System.DirectoryServices.ResultPropertyValueCollection] $memberof = $server.Properties.memberof
				Assert ($server.Properties.memberof.Contains($dgname[0])) -message "GetGroupServers returned a server which is not a member of '$($group.Properties.name)'."
			}
		}
	}
 {
		UnitTest "Create XA Workgroup" {
			[xml] $xml = LoadXML "./XmlMock/workergroups.xml"
			$objs = $xml.SelectSingleNode("Objs")
			Assert ($objs -ne $null) -message "Xml must contains 'Objs' node."

			$workgroups = $objs.ChildNodes
			Assert ($workgroups.Count -ne 0) -message "'Objs' has no child nodes."

			foreach ($wg in $workgroups) {
				$wg = CreateXAWorkGroup $wg
				Assert ($wg -ne $null) -message "CreateXAWorkGroup returned null."

				Assert ($wg.Name -ne [string]::Empty) -message "wg.Name was empty"
				Assert ($wg.RefId -ne [string]::Empty) -message "wg.RefId was empty"
				Assert ($wg.MachineName -ne [string]::Empty) -message "wg.MachineName was empty"
				switch ($wg.Name) {
					"WG All" {
						Assert ($wg.OUs.Count -ne 0) -message "wg.OUs was empty"
						Assert ($wg.ServerGroups.Count -ne 0) -message "wg.ServerGroups was empty"
						Assert ($wg.ServerNames.Count -ne 0) -message "wg.ServerNames was empty"
					}
					"wg ADC" {
						Assert ($wg.OUs.Count -ne 0) -message "wg.OUs was empty"
						Assert ($wg.ServerGroups.Count -eq 0) -message "wg.ServerGroups was empty"
						Assert ($wg.ServerNames.Count -eq 0) -message "wg.ServerNames was empty"
					}
					"wg ADSG" {
						Assert ($wg.OUs.Count -eq 0) -message "wg.OUs was empty"
						Assert ($wg.ServerGroups.Count -ne 0) -message "wg.ServerGroups was empty"
						Assert ($wg.ServerNames.Count -eq 0) -message "wg.ServerNames was empty"
					}
					"WG FS" {
						Assert ($wg.OUs.Count -eq 0) -message "wg.OUs was empty"
						Assert ($wg.ServerGroups.Count -eq 0) -message "wg.ServerGroups was empty"
						Assert ($wg.ServerNames.Count -ne 0) -message "wg.ServerNames was empty"
					}
					"WG_1" {

						Assert ($wg.OUs.Count -eq 0) -message "wg.OUs was empty"
						Assert ($wg.ServerGroups.Count -eq 0) -message "wg.ServerGroups was empty"
						Assert ($wg.ServerNames.Count -ne 0) -message "wg.ServerNames was empty"
					}
					"WG_2" {

						Assert ($wg.OUs.Count -eq 0) -message "wg.OUs was empty"
						Assert ($wg.ServerGroups.Count -eq 0) -message "wg.ServerGroups was empty"
						Assert ($wg.ServerNames.Count -ne 0) -message "wg.ServerNames was empty"
					}
					Default {}
				}
			}
		}
	}
 {
		UnitTest "Create Table" {
			[System.Data.DataTable]$table = CreateTable "table" @("col1", "col2")
			Assert (!-not $table) -message "CreateTable returned null."
			Assert ($table.TableName.Equals("table")) -message "Table name was not set correctly."
			Assert $($table.Columns[0].ColumnName -eq "col1") -message "1st column was not correctly named."
			Assert $($table.Columns[1].ColumnName -eq "col2") -message "2nd column was not correctly named."
		}
	}
 {
		UnitTest "Parse zones" {
			ParseZones "./XmlMock/zones.xml" $db

			[System.Data.DataTable]$table = $db.Tables['tbl_zone']
			Assert ($table.Rows.Count -eq 2) -message "Zones count is not equal to 2. Actual count : $($table.Rows.Count)"
			foreach ($zone in $table.Rows ) {
				switch ($zone.Name) {
					"MOSCOW" {
						Assert ($zone.MachineName -eq "XA65-TARGET") -message "Machine name in zone '$($zone.Name)' is not correct."
					}
					"MALTA" {
						Assert ($zone.MachineName -eq "XA65-TARGET") -message "Machine name in zone '$($zone.Name)' is not correct."
					}
					Default {
						Assert ($false) -message "Un expected zone name! $($zone.Name)"
					}
				}
			}
		}
	}
 {
		UnitTest "Parse servers" {
			ParseServers "./XmlMock/servers.xml" $db
			[System.Data.DataTable]$table = $db.Tables["tbl_server"]
			Assert ($table.Rows.Count -eq 2) -message "Servers count is not equalt to 2. Actual count : $($table.Rows.Count)"

			foreach ($server in $table.Rows) {
				Assert ($server.Name -ne [string]::Empty) -message "server name is empty"

				switch ($server.name) {
					"APPD-LIC01" {
						Assert ($server.MachineName -eq "XA65-TARGET") -message "Machine name for server '$($server.Name)' is incorrect. Actual : $($server.MachineName)"
						Assert ($server.ZoneName -eq "Moscow") -message "Zone name for server '$($server.Name)' is incorrect. Actual : $($server.ZoneName)"
						Assert ($server.WorkgroupName -eq "") -message "Workgroup name for server '$($server.Name)' is not empty. Actual : $($server.WorkgroupName)"
					}
					"XA65-TARGET" {
						Assert ($server.MachineName -eq "XA65-TARGET") -message "Machine name for server '$($server.Name)' is incorrect. Actual : $($server.MachineName)"
						Assert ($server.ZoneName -eq "Malta") -message "Zone name for server '$($server.Name)' is incorrect. Actual : $($server.ZoneName)"
						Assert ($server.WorkgroupName -eq "") -message "Workgroup name for server '$($server.Name)' is not empty. Actual : $($server.WorkgroupName)"
					}
					Default {
						Assert ($false) -message "Unexpected server name! $($server.Name)"
					}
				}
			}
		}
	}
 {
		UnitTest "Parse Workgroups" {
			ParseWorkgroups "./XmlMock/workergroups.xml" $db

			[System.Data.DataTable]$table = $db.Tables["tbl_workgroup"]
			Assert ($table.Rows.Count -eq 6) -message "Number of groups is not equal to 6. Actual : $($table.Rows.Count)."

			foreach ($wg in $table) {
				Assert ($wg.Name -ne [string]::Empty) -message "Workgroup name was empty."
				Assert ($wg.MachineName -eq "XA65-TARGET") -message "Workgroup $($wg.Name) has incorrect machine name. Actual : $($wg.MachineName)"
			}

			[System.Data.DataTable]$table = $db.Tables["tbl_wg_server"]
			Assert ($table.Rows.Count -eq 4) -message "Workgroup server count is incorrect. Actual : $($table.Rows.Count)"
			foreach ($row in $table) {
				Assert ($row.WorkgroupName -ne [string]::Empty) -message "Workgroup name was empty."
				Assert ($row.ServerName -ne [string]::Empty) -message "Server name was empty."
				switch ($row.WorkgroupName) {
					"WG All" {
						Assert ($row.ServerName -eq "APPD-LIC01") -message "Incorrect server name. Actual: $($row.ServerName)"
					}
					"WG FS" {
						Assert ($row.ServerName -eq "APPD-LIC01") -message "Incorrect server name. Actual: $($row.ServerName)"
					}
					"WG_1" {
						Assert ($row.ServerName -eq "XA65-TARGET") -message "Incorrect server name. Actual: $($row.ServerName)"
					}
					"WG_2" {
						Assert ($row.ServerName -eq "XA65-TARGET") -message "Incorrect server name. Actual: $($row.ServerName)"
					}
					Default {
						Assert ($false) -message "Unexpected workgroup name!"
					}
				}
			}
		}
	}
 {
		UnitTest "ParseOU" {
			ParseOUs "./XmlMock/workergroups.xml" $db
			[System.Data.DataTable] $table = $db.Tables['tbl_ou']
			Assert ($table.Rows.Count -eq 3) -message "Incorrect number of OUs. Actual $($table.Rows.Count)"

			foreach ($ou in $table) {
				Assert ($ou.Name -ne [string]::Empty) -message "OU name was emtpy!"
				Assert ($ou.GUID -ne [string]::Empty) -message "OU GUID was emtpy!"
				Assert ($ou.OctetString -ne [string]::Empty) -message "OU octet string was emtpy!"
				Assert ($ou.WorkgroupName -ne [string]::Empty) -message "OU workgroup name was emtpy!"
			}
		}
	}
 {
		UnitTest "ParseServerGroups" {
			ParseServerGroups "./XmlMock/workergroups.xml" $db
			[System.Data.DataTable] $table = $db.Tables['tbl_serverGroup']
			Assert ($table.Rows.Count -eq 2) -message "Incorrect number of AD groups. Actual : $($table.Rows.Count)."

			foreach ($sg in $table) {
				Assert ($sg.Name -ne [string]::Empty) -message "AD group name was empty."
				Assert ($sg.WorkgroupName -ne [string]::Empty) -message "AD group workgroup name was empty."
				switch ($sg.Name) {
					"RASLAB\rdsh hosts upd" {
						Assert ($sg.WorkgroupName -eq "WG All") -message "AD group workgroup name was incorrect. Actual : $($sg.WorkgroupName)"
					}
					"RASLAB\Domain Computers" {
						Assert ($sg.WorkgroupName -eq "WG ADSG") -message "AD group workgroup name was incorrect. Actual : $($sg.WorkgroupName)"

					}
					Default {
						Assert ($false) -message "Unexpected AD group name!"
					}
				}
			}
		}
	}
 {
		UnitTest "ParseADMachines" {
			ParseADMachines $db

			$table = $db.Tables['tbl_ADMachine']
			# Assert ($table.Rows.Count -ne 0) -message "No AD machines were resolved."

			foreach ($machine in $table) {
				if ($machine.OUGuid -eq [string]::Empty) {
					Assert ($machine.ServerGroupsName -ne [string]::Empty) -message "Server group name was empty!"
				}
				elseif ($machine.ServerGroupsName -eq [string]::Empty) {
					Assert ($machine.OUGuid -ne [string]::Empty) -message "OUGuid was empty!"
				}
			}
		}
	}
 {
		UnitTest "ExtractFolders" {
			[xml] $xml = LoadXML "./XmlMock/applications.xml"
			$objs = $xml.SelectSingleNode("Objs")
			Assert ($objs -ne $null) -message "Xml has no 'Objs' node"

			[System.Xml.XmlNodeList]$applications = $objs.ChildNodes
			Assert ($applications.Count -ne 0) -message "No application nodes were found!"

			$app = CreateXAApplication $applications[3]
			Assert ($app -ne $null) -message "CreateXAApplication returned null."
			ExtractFolders $app $db

			$table = $db.Tables['tbl_folder']
			Assert ($table.Rows.Count -ne 0) -message "tbl_folder table is empty!"
			$folder = $table.Select("[Name] like 'Documents'")
			Assert (!-not $folder) -message "Folder Documents was not found!"

			$parent = $table.Select("[Name] like '$($folder.Parent)'")
			Assert (!-not $parent) -message "Parent folder for $($folder.Name) was not found!"
			Assert ($parent.Name -eq "DepartmentY") -message "Parent folder name is not correct!" -actual $($parent.Name) -expected "DepartmentY"
			Assert ($parent.Parent -eq [string]::Empty) -message "Parent folder should have been empty!"
		}
	}
 {
		UnitTest "Parse Application folders" {
			ParseAppFolders "./XmlMock/applications.xml" $db

			$tbl_folder = $db.Tables['tbl_folder']
			$Editors = $tbl_folder.Select("[Name] like 'Editors' and [Parent] like ''")

			Assert (!-not $Editors) -message "No folder named Editors was found!"
			Assert ($Editors.Parent -eq [string]::Empty) -message "This folder should have no parent folder."

			$Editors = $tbl_folder.Select("[Name] like 'Editors' and [Parent] like 'DepartmentX'")

			Assert (!-not $Editors) -message "No folder named Editors was found!"
			Assert ($Editors.Parent -eq "DepartmentX") -message "This folder should have parent folder 'DepartmentX'."

			$Editors = $tbl_folder.Select("[Name] like 'Editors' and [Parent] like 'DepartmentY/Documents'")
			Assert ($Editors.Parent -eq "DepartmentY/Documents") -message "This folder should have parent folder 'DepartmentY/Documents'."

			$Desktops = $tbl_folder.Select("[Name] like 'Desktops'")
			Assert (!-not $Desktops) -message "No folder named 'Desktops' was found!"
			Assert ($Desktops.Parent -eq [string]::Empty) -message "This folder should have no parent folder."

			$Documents = $tbl_folder.Select("[Name] like 'Documents'")
			Assert (!-not $Documents) -message "No folder named 'Documents' was found!"
			Assert ($Documents.Parent -eq "DepartmentY") -message "This folder should have parent folder called 'Department Y'. Actual : $($Documents.Parent)"

			$DepartmentY = $tbl_folder.Select("[Name] like 'DepartmentY'")
			Assert (!-not $DepartmentY) -message "No folder named 'DepartmentY' was found!"
			Assert ($DepartmentY.Parent -eq [string]::Empty) -message "This folder should have no parent folder."

			$DepartmentX = $tbl_folder.Select("[Name] like 'DepartmentX'")
			Assert (!-not $DepartmentX) -message "No folder named 'DepratmentX' was found!"
			Assert ($DepartmentX.Parent -eq [string]::Empty) -message "This folder should have no parent folder. Actual : $($DepartmentX.Parent)"
		}
	}
 {
		UnitTest "Parse Applications" {
			ParseApplications "./XmlMock/applications.xml" $db
			# "Id", "Name", "FolderPath", "Parameters", "Target", "Type", "RASFolderID", "ServerAndGroup", "IconPath", "RASId"

			$tbl_application = $db.Tables['tbl_application']

			foreach ($app in $tbl_application) {
				Assert ($app.Id -ne [string]::Empty) -message "Application id should not be empty"
				Assert ($app.Name -ne [string]::Empty) -message "Application name should not be empty"
				Assert ($app.IconPath -ne [string]::Empty) -message "Application name should not be empty"

				switch ($app.Id) {
					"2b85-0006-000001d6" {
						Assert ($app.Name -eq "Notepad") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "Editors") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						# Assert ($app.Parameters.Equals('"%*"')) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq 'C:\Windows\notepad.exe') -message "Application $($app.Id) has incorrect target." -actual $($app.Target) -expected 'C:\Windows\notepad.exe "%*"'
						Assert ($app.Parameters -eq '"%*"')
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter -ne [System.DBNull]::Value) -message "Application user filter is null for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577"
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260") -message "Application UserFilter mismatch for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577" -actual $($app.UserFilter)
					}
					"2b85-0006-000001d9" {
						Assert ($app.Name -eq "Desktop") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "Desktops") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq [System.DBNull]::Value) -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerDesktop") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.Description -eq "Desktop Desc") -message "Invalid description." -expected "Desktop Desc" -actual $app.Description
						Assert ($app.Enabled -eq $true) -message "Enabled should be true for app id $($app.Id)" -expected "$true" -actual "$($app.Enabled)"
						Assert ($app.AddToClientStartMenu -eq $true) -message "AddToClientStartMenu is invalid!" -expected "$true" -actual "$($app.AddToClientStartMenu)"
						Assert ($app.AddToClientDesktop -eq $true) -message "AddToClientDesktop is invalid!" -expected "$true" -actual "$($app.AddToClientDesktop)"
						Assert ($app.MaximizedOnStartup -eq $true) -message "MaximizedOnStartup is invalid!" -expected $true -actual $app.MaximizedOnStartup
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-1000") -message "Application UserFilter mismatch $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-1000" -actual $app.UserFilter
					}
					"2b85-0006-000001df" {
						Assert ($app.Name -eq "Calculator App") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Windows\System32\calc.exe") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5186") -message "Application UserFilter mismatch $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5186" -actual $app.UserFilter
					}
					"2b85-0006-000001ea" {
						Assert ($app.Name -eq "Notepad from WG") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "DepartmentY/Documents/Editors") -message "Application '$($app.Id)' must be in 'Editors' folder." -expected "DepartmentY/Documents/Editors" -actual $app.FolderPAth
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Windows\System32\notepad.exe") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5218") -message "Application UserFilter mismatch for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5218" -actual $app.UserFilter
					}
					"2b85-0006-000001f5" {
						Assert ($app.Name -eq "Link to Yahoo") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "http://yahoo.com") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "Content") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260") -message "Application UserFilter mismatch for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5260" -actual $app.UserFilter
					}
					"110d-0006-00000207" {
						Assert ($app.Name -eq "Notepad") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "DepartmentX/Editors") -message "Application '$($app.Id)' must be in  folder." -expected 'DepartmentX/Editors' Actual : $($app.FolderPAth)
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Windows\System32\notepad.exe") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter -eq [System.DBNull]::Value) -message "Application UserFilter must be null for $($app.Id)"
					}
					"110d-0006-0000020b" {
						Assert ($app.Name -eq "OfficeDocument") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "c:\test\test.doc") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "Content") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter -eq [System.DBNull]::Value) -message "Application UserFilter must be null for $($app.Id)"
					}
					"110d-0006-00000211" {
						Assert ($app.Name -eq "calcultator") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "DepartmentY") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Windows\System32\calc.exe") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter -eq [System.DBNull]::Value) -message "Application UserFilter must be null for $($app.Id)"
					}
					"2b85-0006-00000214" {
						Assert ($app.Name -eq "Word2013") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Program Files\Microsoft Office\Office15\WINWORD.EXE") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260,S-1-5-21-1457957886-1659725039-1882262577-5186") -message "Application UserFilter mismatch for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5260,S-1-5-21-1457957886-1659725039-1882262577-5186" -actual $app.UserFilter
					}
					"2b85-0006-00000216" {
						Assert ($app.Name -eq "Excel2013") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Program Files\Microsoft Office\Office15\EXCEL.EXE") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260") -message "Application UserFilter mismatch for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5260" -actual $app.UserFilter
					}
					"2b85-0006-00000219" {
						Assert ($app.Name -eq "PowerPoint2013") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)."
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260") -message "Application UserFilter mismatch for $($app.Id)." -expected "S-1-5-21-1457957886-1659725039-1882262577-5260" -actual $app.UserFilter
					}
					"110d-0006-0000021c" {
						Assert ($app.Name -eq "word2016") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null"
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260") -message "Application UserFilter mismatch for $($app.Id)." -expected "S-1-5-21-1457957886-1659725039-1882262577-5260" -actual $app.UserFilter
					}
					"110d-0006-0000021e" {
						Assert ($app.Name -eq "excel2016") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application userfilt was null for $($app.Id)"
						Assert ($app.UserFilter -eq "S-1-5-21-1457957886-1659725039-1882262577-5260") -message "Application UserFilter mismatch for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5260" -actual $app.UserFilter
					}
					"110d-0006-00000220" {
						Assert ($app.Name -eq "PowerPoint2016") -message "Application '$($app.Id)' has incorrect name. Actual : $($app.Name)"
						Assert ($app.FolderPath -eq "") -message "Application '$($app.Id)' must be in 'Editors' folder. Actual : $($app.FolderPAth)"
						Assert ($app.Parameters -eq [System.DBNull]::Value) -message "Application '$($app.Id)' has invalid parameter. Actual : $($app.Parameters)"
						Assert ($app.Target -eq "C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE") -message "Application $($app.Id) has incorrect target. Actual : $($app.Target)"
						Assert ($app.Type -eq "ServerInstalled") -message "Application $($app.Id) has incorrect type. Actual : $($app.Type)"
						Assert ($app.UserFilter) -message "Application UserFilter was null for $($app.Id)" -expected "S-1-5-21-1457957886-1659725039-1882262577-5260" -actual $app.UserFilter
					}
					Default {
						Assert $false -message "Unexpected application ID!"
					}
				}
			}
		}
	}
 {
		UnitTest "Parse Application Servers" {
			ParseAppServers "./XmlMock/applications.xml" $db
			# "AppId", "AppName", "ServerName"

			$tbl_app_server = $db.Tables['tbl_app_server']

			Assert ($tbl_app_server.Rows.Count -eq 10) -message "There must be 3 entires in tbl_app_server." -actual $tbl_app_server.rows.Count -expected 10

			foreach ($appServ in $tbl_app_server) {
				Assert ($appServ.AppId -ne [string]::Empty) -message "AppId must not be empty!"
				Assert ($appServ.AppName -ne [string]::Empty) -message "AppName must not be empty!"
				Assert ($appServ.ServerName -ne [string]::Empty) -message "ServerName must not be empty!"
			}

			$app = $tbl_app_server.Select("[AppId] like '2b85-0006-000001d6'")
			Assert (-not $app) -message "There should be no entry for app id '2b85-0006-000001d6'."

			$app = $tbl_app_server.Select("[AppId] like '2b85-0006-000001df'")
			Assert (-not $app) -message "There should be no entry for app id '2b85-0006-000001df'."

			$app = $tbl_app_server.Select("[AppId] like '2b85-0006-000001ea'")
			Assert (-not $app) -message "There should be no entry for app id '2b85-0006-000001ea'."

			$app = $tbl_app_server.Select("[AppId] like '2b85-0006-000001f5'")
			Assert (-not $app) -message "There should be no entry for app id '2b85-0006-000001f5'."

			$app = $tbl_app_server.Select("[AppId] like '110d-0006-0000020b'")
			Assert ($app) -message "No rows with AppId '110d-0006-0000020b' where found!."

			$app = $tbl_app_server.Select("[AppId] like '2b85-0006-000001d9'")
			Assert (!-not $app) -message "No rows with AppId '2b85-0006-000001d9' where found!"
			Assert ($app.Count -eq 2) -message "There must be 2 rows!"

			$app = $tbl_app_server.Select("[AppId] like '110d-0006-00000207'")
			Assert (!-not $app) -message "No rows with AppId '110d-0006-00000207' where found!"
			Assert ($app.Count -eq 1) -message "There must be 1 row!"
		}
	}
 {
		UnitTest "Parse Application Workgroups" {
			ParseAppWorkgroup "./XmlMock/applications.xml" $db

			$tbl_app_workgroup = $db.Tables['tbl_app_workgroup']

			foreach ($app in $tbl_app_workgroup) {
				Assert ($app.AppId -ne [string]::Empty) -message "AppId should not be empty!"
				Assert ($app.AppName -ne [string]::Empty) -message "AppName should not be empty!"
				Assert ($aoo.ServerName -ne [string]::Empty) -message "ServerName should not be empty!"
			}

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001d6'")
			Assert (!-not $app) -message "No rows for AppId 2b85-0006-000001d6 were found!"
			Assert ($app.Count -eq 1) -message "There must be 1 row."

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001d9'")
			Assert ($app.Count -eq 0) -message "There should be no entries for app id '2b85-0006-000001d9'." -expected 0 -actual $app.Count

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001df'")
			Assert (!-not $app) -message "No rows for AppId 2b85-0006-000001df were found!"
			Assert ($app.Count -eq 1) -message "There must be 1 row."

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001ea'")
			Assert (!-not $app) -message "No rows for AppId 2b85-0006-000001ea were found!"
			Assert ($app.Count -eq 1) -message "There must be 1 row."

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001f5'")
			Assert (-not $app) -message "There should be no entry for app id 2b85-0006-000001f5."

			$app = $tbl_app_workgroup.Select("[AppId] like '110d-0006-00000207'")
			Assert (-not $app) -message "There should be no entry for app id 110d-0006-00000207."

			$app = $tbl_app_workgroup.Select("[AppId] like '110d-0006-0000020b'")
			Assert (-not $app) -message "There should be no entry for app id 110d-0006-0000020b."
		}
	}
 {
		UnitTest "Workgroup server cleanup" {

			WorkgroupServerCleanup $db

			######################################################
			# Asserting contents of tbl_wg_server after cleanup #
			######################################################
			$tbl_wg_server = $db.Tables['tbl_wg_server']
			Assert ($tbl_wg_server.Rows.Count -eq 2) -message "tbl_wg_server should have 2 entries after cleanup! Actual : $($tbl_wg_server.Rows.Count)"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'GRP_APPD-LIC01'")
			Assert (!-not $grp) -message "No entries found for GRP_APPD-LIC01"
			Assert ($grp.Count -eq 1) -message "There must be 1 entry for GRP_APPD-LIC01."

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'GRP_XA65-TARGET'")
			Assert (!-not $grp) -message "No entries found for GRP_XA65-TARGET"
			Assert ($grp.Count -eq 1) -message "There must be 1 entry for GRP_XA65-TARGET."

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'WG All'")
			Assert (-not $grp) -message "There should be no entries for WG All"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'WG FS'")
			Assert (-not $grp) -message "There should be no entries for WG FS"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'wg ADC'")
			Assert (-not $grp) -message "There should be no entries for wg ADC"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'wg ADC'")
			Assert (-not $grp) -message "There should be no entries for wg ADC"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'wg ADSG'")
			Assert (-not $grp) -message "There should be no entries for wg ADSG"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'WG_1'")
			Assert (-not $grp) -message "There should be no entries for WG_1"

			$grp = $tbl_wg_server.Select("[WorkgroupName] like 'WG_2'")
			Assert (-not $grp) -message "There should be no entries for WG_2"

			######################################################
			# Asserting contents of tbl_server after cleanup    #
			######################################################
			$tbl_server = $db.Tables['tbl_server']
			Assert ($tbl_server.Rows.Count -eq 2) -message "tbl_server must have 2 entires."

			$server = $tbl_server.Select("[Name] like 'APPD-LIC01'")
			Assert (!-not $server) -message "Entry for server APPD-LIC01 must be present!"
			Assert ($server.WorkgroupName -eq "GRP_APPD-LIC01") -message "Server APPD-LIC01 has incorrect workgroup name. Actual $($server.WorkgroupName)"

			$server = $tbl_server.Select("[Name] like 'XA65-TARGET'")
			Assert (!-not $server) -message "Entry for server XA65-TARGET must be present!"
			Assert ($server.WorkgroupName -eq "GRP_XA65-TARGET") -message "Server XA65-TARGET has incorrect workgroup name. Actual $($server.WorkgroupName)"

			##########################################################
			# Asserting contents of tbl_app_workgroup after cleanup #
			##########################################################
			$tbl_app_workgroup = $db.Tables['tbl_app_workgroup']
			Assert ($tbl_app_workgroup.Rows.Count -eq 3) -message "tbl_app_workgroup must have 3 rows. Actual : $($tbl_app_workgroup.Rows.Count)"

			$app = $tbl_app_workgroup.Select("[WorkgroupName] like 'WG All'")
			Assert (-not $app) -message "No entries with group name 'WG All' should be present!"

			$app = $tbl_app_workgroup.Select("[WorkgroupName] like 'WG ADSG'")
			Assert (-not $app) -message "No entries with group name 'WG ADSG' should be present!"

			$app = $tbl_app_workgroup.Select("[WorkgroupName] like 'WG FS'")
			Assert (-not $app) -message "No entries with group name 'WG FS' should be present!"

			$app = $tbl_app_workgroup.Select("[WorkgroupName] like 'WG_1'")
			Assert (-not $app) -message "No entries with group name 'WG_1' should be present!"

			$app = $tbl_app_workgroup.Select("[WorkgroupName] like 'WG_2'")
			Assert (-not $app) -message "No entries with group name 'WG_2' should be present!"

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001df' and [WorkgroupName] like 'GRP_XA65-TARGET'")
			Assert (!-not $app) -message "Application with id 2b85-0006-000001df and group name GRP_XA65-TARGET must exist."

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001d6' and [WorkgroupName] like 'wg ADC'")
			Assert (!-not $app) -message "Application with id 2b85-0006-000001d6 and group name wg ADC must exist!"

			$app = $tbl_app_workgroup.Select("[AppId] like '2b85-0006-000001ea' and [WorkgroupName] like 'GRP_APPD-LIC01'")
			Assert (!-not $app) -message "Application with id 2b85-0006-000001ea and group name GRP_APPD-LIC01 must exist!"

			######################################################
			# Asserting contents of tbl_workgroup after cleanup #
			######################################################
			$tbl_workgroup = $db.Tables['tbl_workgroup']
			Assert ($tbl_workgroup.Rows.Count -eq 4) -message "tbl_workgroup must have 4 entries! Actual : $($tbl_workgroup.Rows.Count)"

			$group = $tbl_workgroup.Select("[Name] like 'WG All'")
			Assert (-not $group) -message "Workgroup WG All must not exist!"

			$group = $tbl_workgroup.Select("[Name] like 'WG FS'")
			Assert (-not $group) -message "Workgroup WG FS must not exist!"

			$group = $tbl_workgroup.Select("[Name] like 'WG_1'")
			Assert (-not $group) -message "Workgroup WG FS must not exist!"

			$group = $tbl_workgroup.Select("[Name] like 'WG_2'")
			Assert (-not $group) -message "Workgroup WG FS must not exist!"

			$group = $tbl_workgroup.Select("[Name] like 'WG ADC'")
			Assert (!-not $group) -message "Workgroup WG ADC must exist!"

			$group = $tbl_workgroup.Select("[Name] like 'WG ADSG'")
			Assert (!-not $group) -message "Workgroup WG ADSG must exist!"

			$group = $tbl_workgroup.Select("[Name] like 'GRP_APPD-LIC01'")
			Assert (!-not $group) -message "Workgroup 'GRP_APPD-LIC01' must exist!"

			$group = $tbl_workgroup.Select("[Name] like 'GRP_XA65-TARGET'")
			Assert (!-not $group) -message "Workgroup 'GRP_XA65-TARGET' must exist!"

			######################################################
			# Asserting contents of tbl_ou after cleanup 		 #
			######################################################
			$tbl_ou = $db.Tables['tbl_ou']
			Assert ($tbl_ou.Rows.Count -eq 3) -message "tbl_ou must have 3 entries! Actual : $($tbl_ou.Rows.Count)"

			$ou = $tbl_ou.Select("[GUID] like 'e3e6ed90-a3da-41d4-a341-bf39b0036148' and [WorkgroupName] like 'GRP_APPD-LIC01'")
			Assert (!-not $ou) -message "OU with GUID 'e3e6ed90-a3da-41d4-a341-bf39b0036148' and group name 'GRP_APPD-LIC01' must be present!"

			$ou = $tbl_ou.Select("[GUID] like 'adf263e9-3351-4f62-91e5-b338bdc219cc' and [WorkgroupName] like 'wg ADC'")
			Assert (!-not $ou) -message "OU with GUID 'adf263e9-3351-4f62-91e5-b338bdc219cc' and group name 'wg ADC' must be present!"

			$ou = $tbl_ou.Select("[GUID] like '2e44ed95-195c-488e-b5c7-87b3e2c74c27' and [WorkgroupName] like 'wg ADC'")
			Assert (!-not $ou) -message "OU with GUID '2e44ed95-195c-488e-b5c7-87b3e2c74c27' and group name 'wg ADC' must be present!"

			#######################################################
			# Asserting contents of tbl_serverGroup after cleanup #
			#######################################################
			$tbl_serverGroup = $db.Tables['tbl_serverGroup']
			Assert ($tbl_serverGroup.Rows.Count -eq 2) -message "tbl_serverGroup must have 2 entries!"

			$group = $tbl_serverGroup.Select("[Name] like 'RASLAB\rdsh hosts upd'")
			Assert (!-not $group) -message "AD group with name 'RASLAB\rdsh hosts upd' must be present!"

			$group = $tbl_serverGroup.Select("[Name] like 'RASLAB\Domain Computers'")
			Assert (!-not $group) -message "AD group with name 'RASLAB\Domain Computers' must be present!"
		}
	}
)

$db = ParseXMLtoDB -zonesXmlPath "./XmlMock/zones.xml" -serversXmlPath "./XmlMock/servers.xml" -workgroupXmlPath "./XmlMock/workergroups.xml" -applicationXmlPath "./XmlMock/applications.xml"


foreach ($TestSuiteResult in $TestSuiteResults) {
	Log -type "INFO" -message $TestSuiteResult.Header
	foreach ($test in $TestSuiteResult.UnitTests) {
		Log -type "INFO" -message "[UnitTest: $($test.TestName)]" -testInfo $test.Info -testStatus $test.Status
	}
	Log -type "INFO" -message $TestSuiteResult.Footer
}

Remove-Module PSAdmin