## ==================================================================
##
## Copyright (c) 2005-2020 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Powers on a percentage of RDS Machines in an RDS Group on Paralles RAS environmment
.DESCRIPTION  
    Powers on RDS Machines in an RDS Group based on:
    - RDS Group
    - Percentage of Group
.NOTES  
    File Name  : InitialPowerOn.ps1
.EXAMPLE
	.\InitialPowerOn.ps1 -RASAdminUsername admin -RDSGroupName RDSGroup -BootServerPercent 10
#>

[CmdletBinding()]
    Param(
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RASAdminUsername,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [securestring]$RASAdminPassword,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RDSGroupName,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$BootServerPercent

    )
#Configure logging
function log
{
   param([string]$message)
   "$(get-date -f o)  $message" 
}

function PowerOnMachine
{
   param(
        [string]$Server,
        [string]$Description)

    log "Powering On $Server $Description"

    #Add code to power on Machine using Cloud Provider API
}


log "Loading RAS PowerShell Module"
Import-Module RASAdmin

log "Connecting with RAS Licensing server"
New-RASSession -Username $RASAdminUsername -Password $RASAdminPassword

log "Retrieving list of Servers in '$RDSGroupName'" 
$RDSGroupMembers = Get-RASRDSHostPoolMember -HostPoolName $RDSGroupName

log "Sort Server by ID"
$SortedRDSGroupMembers = [System.Collections.ArrayList] ($RDSGroupMembers | Sort-Object -Property Id)

$BootServerCount =  [int]($SortedRDSGroupMembers.Count / 100 * $BootServerPercent)
Log "Booting $BootServerCount RDS Servers:"

log "Retrieving RDS Status"
$RDSStatusList = Get-RASRDSHostStatus

For ($i=0; $i -lt $BootServerCount; $i++) {
    $RDS = $SortedRDSGroupMembers[0]
    $SortedRDSGroupMembers.RemoveAt(0)

    $RDSStatus = $RDSStatusList | Where-Object {$_.Id -eq $RDS.Id}

    if ($RDSStatus -eq $null) {
        Log "$($RDS.Server) is not applied yet."
        continue
    }

    if ($RDSStatus.AgentState -eq "NotVerified") {        
        PowerOnMachine($RDS.Server, $RDS.Description)
    }
}
