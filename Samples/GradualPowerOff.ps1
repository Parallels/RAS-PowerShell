## ==================================================================
##
## Copyright (c) 2005-2020 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Gradually powers off RDS Machines on Paralles RAS environmment
.DESCRIPTION  
    Gradually powers off RDS Machines on Paralles RAS environmment based on:
    - Monitors Machines in an RDS Group
    - Number of sessions running on the RDS
    - Ensuring X% RDS Machines are still kept on
.NOTES  
    File Name  : GradualPowerOff.ps1
.EXAMPLE
	.\GradualPowerOff.ps1 -RASAdminUsername admin -RDSGroupName RDSGroup -IgnoreServerPercent 10
#>

[CmdletBinding()]
    Param(
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RASAdminUsername,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [securestring]$RASAdminPassword,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RDSGroupName,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$IgnoreServerPercent

    )
#Configure logging
function log
{
   param([string]$message)
   "$(get-date -f o)  $message" 
}

function PowerOffMachine
{
   param(
        [string]$Server,
        [string]$Description)

    log "Powering Off $Server $Description"

    #Add code to power off Machine using Cloud Provider API
}


log "Loading RAS PowerShell Module"
Import-Module RASAdmin

log "Connecting with RAS Licensing server"
New-RASSession -Username $RASAdminUsername -Password $RASAdminPassword

log "Retrieving list of Servers in '$RDSGroupName'" 
$RDSGroupMembers = Get-RASRDSGroupMember -GroupName $RDSGroupName

log "Sort Server by ID"
$SortedRDSGroupMembers = [System.Collections.ArrayList] ($RDSGroupMembers | Sort-Object -Property Id)

$IgnoreServerCount =  [int]($SortedRDSGroupMembers.Count / 100 * $IgnoreServerPercent)
Log "Ignoring $IgnoreServerCount RDS Server:"
For ($i=0; $i -lt $IgnoreServerCount; $i++) {
    Log "- $($SortedRDSGroupMembers[0].Server)"
    $SortedRDSGroupMembers.RemoveAt(0)
}

log "Retrieving RDS Status"
$RDSStatusList = Get-RASRDSStatus

log "Checking Sessions on RDS Machines"
ForEach ($RDS in $SortedRDSGroupMembers) {
    
    $RDSStatus = $RDSStatusList | Where-Object {$_.Id -eq $RDS.Id}
    if ($RDSStatus -eq $null) {
        Log "$($RDS.Server) is not applied yet."
        continue
    }

    if ($RDSStatus.AgentState -ne "OK") {
        Log "$($RDS.Server) does not have Agent State as OK"
        continue
    }

    if ($RDSStatus.ActiveSessions -eq 0 ) {
        PowerOffMachine($RDS.Server, $RDS.Description)
    }
    else {
        Log "$($RDS.Server) has still $($RDSStatus.ActiveSessions) Active Sessions"
    }
}
