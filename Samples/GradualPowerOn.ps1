## ==================================================================
##
## Copyright (c) 2005-2020 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Gradually powers on RDS Machines on Paralles RAS environmment
.DESCRIPTION  
    Gradually powers on RDS Machines on Paralles RAS environmment based on:
    - Monitors Machines in an RDS Group
    - Ensuring X% RDS Machines are always on
    - Power on RDS Maches when average usage is higher then Y%

.NOTES  
    File Name  : GradualPowerOn.ps1
.EXAMPLE
	.\GradualPowerOn.ps1 -RASAdminUsername admin -RDSGroupName RDSGroup -MaxThresholdPercent 80 -MinThresholdPercent 60
#>

[CmdletBinding()]
    Param(
		[parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RASAdminUsername,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [securestring]$RASAdminPassword,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$RDSGroupName,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$MaxSessionPerMachine,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$MaxThresholdPercent,
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$MinThresholdPercent
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

log "Retrieving RDS Status"
$RDSStatusList = Get-RASRDSHostStatus

$TotalMachinesOn = 0
$TotalActiveSessions = 0

$PoweredOffRDSGroupMembers = @()

ForEach ($RDS in $SortedRDSGroupMembers) {
    $RDSStatus = $RDSStatusList | Where-Object {$_.Id -eq $RDS.Id}

    if ($RDSStatus -eq $null) {
        Log "$($RDS.Server) is not applied yet."
        continue
    }

    if ($RDSStatus.AgentState -eq "NotVerified") {
            $PoweredOffRDSGroupMembers += $RDS
    }
    elseif ($RDSStatus.AgentState -eq "OK") {
        Log "$($RDS.Server) is switched on"
        $TotalMachinesOn = $TotalMachinesOn + 1
        $TotalActiveSessions =  $TotalActiveSessions + $RDSStatus.ActiveSessions
    }
    else {
        Log "$($RDS.Server) is being ignored has state $($RDSStatus.AgentState)"
    }
}

Log "$TotalMachinesOn are powered on hosting $TotalActiveSessions sessions"

if ($TotalMachinesOn -eq 0 ){
    $CurrentLoad = 100
}
else{
    $CurrentLoad = [math]::Ceiling((100 *$TotalActiveSessions) / ($MaxSessionPerMachine * $TotalMachinesOn))
}
Log "Current Load - $CurrentLoad"

if ($CurrentLoad -ge $MaxThresholdPercent){

    $TotalRequiredMachinesOn = [math]::Ceiling( (100  * $TotalActiveSessions) / ($MaxSessionPerMachine * $MinThresholdPercent))

    $NewMachinesToPowerOnCount = $TotalRequiredMachinesOn - $TotalMachinesOn.Count
    if ($NewMachinesToPowerOnCount -ge 0){
        Log "Powering on $NewMachinesToPowerOnCount RDS Machines"
        $NewMachinesToPowerOn = $PoweredOffRDSGroupMembers | Get-Random -Count $NewMachinesToPowerOnCount
        ForEach ($RDS in $NewMachinesToPowerOn) {
            Log "$($RDS.Server) triggered to Power on"
            PowerOnMachine($RDS.Server, $rds.Description)
        }
    }
    else {
        Log "No Machines available to power on"
    }
}
