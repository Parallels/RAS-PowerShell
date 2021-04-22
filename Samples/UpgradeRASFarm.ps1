## ==================================================================
##
## Copyright (c) 2005-2020 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Upgrades Paralles RAS environmment
.DESCRIPTION  
    Upgrades Paralles RAS environmment
.NOTES  
    File Name  : UpgradeRASFarm.ps1
.EXAMPLE
	.\UpgradeRASFarm.ps1 -RASAdminUsernam admin -RASAdminDomain myDomain -RASInstaller RASInstaller-XX.X.XXXX.msi
#>

[CmdletBinding()]
    Param(
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RASAdminUsernam,
        [parameter(Mandatory=$true)][AllowNull()][AllowEmptyString()][string]$RASAdminDomain,
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [securestring]$RASAdminPassword,
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RASInstaller,
		[parameter(Mandatory=$false)][string]$RASInstallationPath = "c:\Program Files (x86)\Parallels\ApplicationServer"
    )

if ([string]::IsNullOrEmpty($RASAdminDomain)){
	$RASAdminUser = $RASAdminUsernam
}
else{
	$RASAdminUser = $RASAdminUsernam + '@' + $RASAdminDomain
}

#Configure logging
function log
{
   param([string]$message)
   "`n$(get-date -f o)  $message" 
}

log "`nUpgrading RAS Farm"

$pathIsValid = Test-Path $RASInstaller
if( -Not $pathIsValid ){
	Write-Error -Message "`nERROR: Installer path not valid." -ErrorAction Stop
}

$logFile = $PSScriptRoot + "\UpgradeRASFarm.log"

#Update RAS Agents
$InstalledServerList=@( hostname )
function UpdateRASAgent
{
   param([string]$server,[int]$siteId)

   if ($InstalledServerList -NotContains $server) {
	   Update-RASAgent -Server $server -SiteId $siteId 
	   $InstalledServerList  += $server
   }
}

#upgrade Primary licensing server 
log "Upgrading RAS Licensing Server"
$processPath = "msiexec"

$process = (Start-Process -FilePath $processPath -ArgumentList @("/i", "$RASInstaller", "/qn", "/norestart", "/l*vx $logFile") -PassThru -Wait -NoNewWindow)

if($process.ExitCode -eq 3010){
	log "Successful Upgrading. Reboot required."
}
elseif($process.ExitCode -ne 0){
	Write-Error -Message "Thrown error from msiexec." -ErrorId $process.ExitCode -ErrorAction Stop
}

#install PowerShell
log "Installing PowerShell Module"
$process = (Start-Process -FilePath $processPath -ArgumentList @("/i", "$RASInstaller", "ADDLOCAL=F_PowerShell", "/qn", "/norestart", "/l*vx $logFile") -PassThru -Wait -NoNewWindow)

if($process.ExitCode -eq 3010){
	log "Successful Upgrading. Reboot required."
}
elseif($process.ExitCode -ne 0){
	Write-Error -Message "Thrown error from msiexec." -ErrorId $process.ExitCode -ErrorAction Stop
}

# Import RAS PowerShell module
log "Loading PowerShell Module"

# import PSAdmin (before v18)
$PSModulePath = $RASInstallationPath + "\Modules\PSAdmin\PSAdmin.psd1"
$PSModulePathIsValid = Test-Path $PSModulePath
if($PSModulePathIsValid){
	Import-Module $PSModulePath
}

# RASAdmin (from v18)
$RASModulePath = $RASInstallationPath + "\Modules\RASAdmin\RASAdmin.psd1"
$RASModulePathIsValid = Test-Path $RASModulePath
if($RASModulePathIsValid){
	Import-Module $RASModulePath
}

if( (-Not $PSModulePathIsValid ) -and ( -Not $RASModulePathIsValid ) ){
	Write-Error -Message "`nERROR: RAS PowerShell module path not valid." -ErrorAction Stop
}


#Logon to RAS
log "Connecting with RAS Licensing server"
New-RASSession -Username $RASAdminUser -Password $RASAdminPassword

#Go through all sites
$Sites = Get-RASSite
foreach ($Site in $Sites) {
	
	#Upgrade RAS Publishing agents in order of Priority
	log "Upgrading RAS Publishing agents in order of Priority"
	$PAs = Get-RASPA -Site $Site.Id | Sort-Object -Property Priority
	foreach($PA in $PAs) {
		UpdateRASAgent  $PA.Server $Site.Id 
	}
	
	#Upgrade RAS SecureClientGateways
	log "Upgrading RAS SecureClientGateways"
	$GWs = Get-RASGW -Site $Site.Id
	foreach($GW in $GWs) {
		UpdateRASAgent $GW.Server $Site.Id 
	}
	
	#Upgrade RAS RemoteDesktopServers Agents
	log "Upgrading RAS RemoteDesktopServers Agents"
	$RDServers = Get-RASRDS -Site $Site.Id
	foreach($RDServer in $RDServers) {
		UpdateRASAgent $RDServer.Server $Site.Id 
	}  
}

log "`nRAS Farm Upgrade finished.`n"




