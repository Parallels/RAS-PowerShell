## ==================================================================
##
## Copyright (c) 2005-2019 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    RAS PowerShell Publishing Examples
.DESCRIPTION  
    Examples to demonstrates how to manage published resources and use filtering options.
.NOTES  
    File Name  : PublishingSample.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\PublishingSample.ps1
#>

CLS


#Pre-set Params
$RDSServer1 = "rds1.company.dom" 	#(replace 'rds1.company.dom' with a valid FQDN, computer name, or IP address).
$RDSServer2 = "rds2.company.dom" 	#(replace 'rds2.company.dom' with a valid FQDN, computer name, or IP address).
$RDSGroupName = "My RDS Group"		#(replace with a more specific name).
$RDSDefSettMaxSessions = 100		#(replace default value with preferred max sessions).
$RDSDefSettAppMonitor = $true		#(replace default value with preferred App Monitoring value (Enabeld/Disabled)).
$VMNameFormat       = "Win10-%ID:3%"

$AccFldName = "AccDept"
$AccFldDesc = "Accounting"
$SalesFldName = "SalesDept"
$SalesFldDesc = "Sales"

$AccDeskName = "AccPubDesktop"
$SalesDeskName = "SalesPubDesktop"

$AccAppName = "AccPubApp"
$AccAppTarget = "C:\Windows\System32\calc.exe"
$AccAppIPFilter = "10.0.0.1-10.0.0.12"
$SalesAppName = "SalesPubApp"
$SalesAppTarget = "C:\Windows\System32\notepad.exe"




#Configure logging
function log
{
   param([string]$message)
   "`n`n$(get-date -f o)  $message" 
}


Import-Module RASAdmin

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

###### FARM CONFIGURATION ######

#Add two RD Session Host servers.
log "Adding two RD Session host servers"
$RDS1 = New-RASRDSHost -Server $RDSServer1
$RDS2 = New-RASRDSHost -Server $RDSServer2

#Get the list of RD Session Host servers. The $RDSList variable receives an array of objects of type RDS.
log "Retrieving the list of RD Session servers"
$RDSList = Get-RASRDSHost

log "Print the list of RD Session servers retrieved"
Write-Host ($RDSList | Format-Table | Out-String)

#Create an RD Session Host Group and add both RD Session Host objects to it.
log "Add an RD Session host group (with list of RD Sessions)"
New-RASRDSHostPool -Name $RDSGroupName -Description "RDSTemplates Pool" -WorkLoadThreshold 50 -ServersToAddPerRequest 2 `
-WorkLoadToDrain 20 -HostsToCreate 1 -HostName $VMNameFormat -MinServersFromTemplate 2 -MaxServersFromTemplate 2 -Autoscale $true -RDSObject $RDSList

#Update default settings used to configure RD Session Host agents.
log "Updating RDS default settings"
Set-RASRDSDefaultSettings -MaxSessions $RDSDefSettMaxSessions -EnableAppMonitoring $RDSDefSettAppMonitor

###### PUBLISHING CONFIGURATION ######

#Add published folders to be used by different departments.
log "Adding published folders for different departments"
$Fld_Acc = New-RASPubFolder -Name $AccFldName -Description $AccFldDesc
$Fld_Sales = New-RASPubFolder -Name $SalesFldName -Description $SalesFldDesc

#Add published desktops within their respective folders.
log "Adding published desktops for their respective department folders"
$Desk_Acc = New-RASPubRDSDesktop -Name $AccDeskName -ParentFolder $Fld_Acc -DesktopSize FullScreen -PublishFrom Group -PublishFromGroup $RDSGrp
$Desk_Sales = New-RASPubRDSDesktop -Name $SalesDeskName -ParentFolder $Fld_Sales -DesktopSize Custom -Width 600 -Height 400 -PublishFrom All

#Add published applications within their respective folders.
log "Adding published applications for their respective department folders"
$App_Acc = New-RASPubRDSApp -Name $AccAppName -Target $AccAppTarget -ParentFolder $Fld_Acc -PublishFrom All -WinType Maximized -StartOnLogon
$App_Sales = New-RASPubRDSApp -Name $SalesAppName -Target $SalesAppTarget -ParentFolder $Fld_Sales -PublishFrom Server -PublishFromServer $RDS1

#Update default settings used to configure published resources.
log "Updating Publishing default settings"
Set-RASPubDefaultSettings -CreateShortcutOnDesktop $true

#Override shortcut default settings for a specific published application.
log "Overriding shortcut default settings for the Sales published application."
Set-RASPubRDSApp -InputObject $App_Sales -InheritShortcutDefaultSettings $false -CreateShortcutOnDesktop $false

###### PUB FILTERING CONFIGURATION ######

#Set AD account filters by ID.
log "Set Active Directory filters for Accounts published desktop"
Set-RASCriteria -ObjType PubItem -InputObject $Desk_Sales -Enable $true -Replicate $true
Add-RASCriteriaSecurityPrincipal -Account "Accounts" -SID "10" -Id $Desk_Acc.Id -ObjType PubItem -RuleId 1

#Set AD account filters by object.
log "Set Active Directory filters for Sales published desktop"
Set-RASCriteria -InputObject $Desk_Sales -Enable $true -Replicate $true
Add-RASCriteriaSecurityPrincipal -InputObject $App_Acc -IP "10.0.0.1-10.0.0.12"

#Set an IP filter (with range) on application.
log "Set IP filters for Accounts published application"
Set-RASCriteria -ObjType PubItem -InputObject $App_Acc -Enable $true
Add-RASCriteriaSecurityPrincipal -InputObject $App_Acc -IP "10.0.0.1-10.0.0.12"

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-RASApply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"
