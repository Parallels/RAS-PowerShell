# Parallels RAS Citrix Migration Tool

## Prerequesites

#### Citrix XenApp version & components

* Citrix XenApp 6.*
* XenApp SDK  6.*
* PowerShell v2

#### Parallels RAS version & components
* Parallels RAS PowerShell 16.5
* .Net 4.5
* PowerShell v3

#### Preparing target RAS farm

* Create VM with Windows 2008 R2 or later
* Download [RASInstaller](https://www.parallels.com/products/ras/download/links/)
* Deploy RAS components on a single host following installer steps
* Configure RAS farm and activate license (trial can be used)

# Migration Process

## Export Citrix XenApp settings
Migration tool requires 4 XML settings files to operate. These are obtained using Citrix XenApp PowerShell SDK. Run the following commands to extract the settings.

```powershell
Add-PSSnapin citrix.xenapp.commands # load Citrix PowerShell Cmdlets

# OPTIONAL export all farm settings. Used to add more information
# to the header of the generated script.
Get-XAFarm | Export-Clixml "./farm.xml"

# Export all application settings
Get-XAApplicationReport * | Export-Clixml "./applications.xml"

# Export all zone settings
Get-XAZone | Export-Clixml "./zones.xml"

# Export all server settings
Get-XAServer * | Export-Clixml "./servers.xml"

# Export all workergroup settings
Get-XAWorkerGroup | Export-Clixml "./workergroups.xml"

```


## Running the Citrix-RAS Migration tool

1. Download Parallels RAS Citrix Migration Tool, and move the exported settings into its directory.
2. Launch **PowerShell** and change current path to the **Citrix-RAS Migration Tool** directory.

3. In PowerShell console window, execute `Run.ps1` script as below.
	```powershell
	.\Run.ps1 -XmlPathWorkgroups ./workergroups.xml -XmlPathZones ./zones.xml -XmlPathServers ./servers.xml -XmlPathApplications ./applications.xml -XmlPathFarm ./farm.xml
	```
	Running this script will generate a `MigrationScript.ps1` script along with the exported icons in the `icons` folder. `MigrationScript.ps1` can be then modified to your needs if required.
	If the script is going be executed on a different machine, the icons folder must be available too.

## Migrating Citrix XenApp settings to Parallels RAS
1. Execute `MigrationScript.ps1` script and provide your RAS server hostname/ip and credentials when prompted.

2. If the Parallels RAS version is supported, the migration will begin. **Minimal supported version is 16.2**. Note icon support is available in 16.5 and above. Parallels RAS 16.2 does not support file extension filtering, and setting of color depth through PowerShell.

## Migrating RDS host servers to RAS
### Uninstalling Citrix components.
Refer to [this document](https://www.parallels.com/fileadmin/docs/ras/resources/WP_MigrationfromCitrix_EN_A4.pdf)
for more details on how to remove Citrix XenApp 6.5

###  Installing RAS Agents.
* Open RAS Console
* Click `Upgrade all agents` from the `Tasks` menu.
* Select the hosts you want to push the "RAS RD Session Host Agent" to.
* Click `OK`, to start installation process.

As soon as RAS RD Session Hosts Agent are installed on the target, hosts and their status is verified, the servers are ready to host RDP connections.

# Migration details
## Migrated components.
| Citrix XenApp	| Parallels RAS 		|
| -------------	|-------------			|
|Zones			| Sites (Partial)		|
|Server			| RDS hosts				|
|Worker Groups	| RDS groups			|
|Applications	| Published applications|

## Components that are not migrated.

| Citrix XenApp 				| Parallels RAS 				|
| ------------- 				|-------------					|
| Administrator     			| Different permissions schema 	|
| Load Balancing Policies      	| Not Available					|
| Load Evaluators      			| Not available					|
|Policies						| Not available					|

## Migration process of Sites (Zones)
This version only migrates the first site that appears in the `zones.xml`, with the site name updated. This is because it is not possible to resolve under which zone certain components reside. i.e. A workgroup can have multiple farm servers that belong to different zones, similiarly, applications being published from worker groups have no way of knowing to which zone they belong.

## Migration process of Servers
The servers are migrated to RAS as RD Session Hosts set with the primary site.

## Migration process of Application Folders
The folder strucutre for Administrative folder and Client Application folders are merged. To make a distinction, Administrative folders are marked as `Use for administrative purpose`

## Migration process of Worker Groups
Citrix XenApp worker groups are adjusted to the Parallels RAS RD Session Host Groups structure due to their differences.

#### Key differences
* Citirx XenApp allows sharing servers between 2 or more workgroups (Not allowed in Parallels RAS 16.5)
	Solving this problem required to extract servers that are common to 2 or more groups into their own server group.
	Such servers are extracted into a group prefixed `GRP_<servername>`. This adaptation allows to have a similar configuration on RAS for RDS Groups. Also note, that OUs, and AD groups are prefixed with `OU_<oiGUID>` and `ADG_<groupName>` respectively.

* Ctirix XenApp worker groups are not bound to zones. This makes it difficult to have a 1:1 mapping to Parallels RAS.
	In this version of migration tool the issue is solved by using the first zone name that appears in the `zones.xml`. **Everything is migrated under one zone**.

## Migration process of Applications
Parallels RAS Citrix Migration Tool was tested with Parallels RAS 16.2, Parallels RAS 16.5, Citrix XenApp 6.0, Citrix XenApp 6.5

All the settings that are available to Parallels RAS in xml settings were migrated.

|Property				|Status			|
|-----------------------|---------------|
|Name					|Migrated		|
|Description			|Migrated		|
|Application Type		|Migrated		|
|Command Line			|Migrated		|
|Working directory		|Migrated		|
|Servers				|Migrated		|
|Groups					|Migrated		|
|User Filtering			|Migrated		|
|Shortcut presentation	|Migrated		|
|File types				|Migrated (16.5)|
|Licence limits			|Migrated		|
|Printer settings		|Migrated		|
|Color 					|Migrated (16.5)|
|Resolution				|Migrated		|
|Start up settings		|Migrated		|
|App state				|Migrated		|


For more information please see [this document]()

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
