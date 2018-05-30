## ==================================================================
##
## Copyright (c) 2005-2018 Parallels Software International, Inc.
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


Import-Module PSAdmin

#Establish a connection with Parallels RAS (NB. User will be prompted for Username and Password)
log "Creating RAS session"
New-RASSession

###### FARM CONFIGURATION ######

#Add two RD Session Host servers.
log "Adding two RD Session host servers"
$RDS1 = New-RDS -Server $RDSServer1
$RDS2 = New-RDS -Server $RDSServer2

#Get the list of RD Session Host servers. The $RDSList variable receives an array of objects of type RDS.
log "Retrieving the list of RD Session servers"
$RDSList = Get-RDS

log "Print the list of RD Session servers retrieved"
Write-Host ($RDSList | Format-Table | Out-String)

#Create an RD Session Host Group and add both RD Session Host objects to it.
log "Add an RD Session host group (with list of RD Sessions)"
New-RDSGroup -Name $RDSGroupName -RDSObject $RDSList

#Update default settings used to configure RD Session Host agents.
log "Updating RDS default settings"
Set-RDSDefaultSettings -MaxSessions $RDSDefSettMaxSessions -EnableAppMonitoring $RDSDefSettAppMonitor

###### PUBLISHING CONFIGURATION ######

#Add published folders to be used by different departments.
log "Adding published folders for different departments"
$Fld_Acc = New-PubFolder -Name $AccFldName -Description $AccFldDesc
$Fld_Sales = New-PubFolder -Name $SalesFldName -Description $SalesFldDesc

#Add published desktops within their respective folders.
log "Adding published desktops for their respective department folders"
$Desk_Acc = New-PubRDSDesktop -Name $AccDeskName -ParentFolder $Fld_Acc -DesktopSize FullScreen -PublishFrom Group -PublishFromGroup $RDSGrp
$Desk_Sales = New-PubRDSDesktop -Name $SalesDeskName -ParentFolder $Fld_Sales -DesktopSize Custom -Width 600 -Height 400 -PublishFrom All

#Add published applications within their respective folders.
log "Adding published applications for their respective department folders"
$App_Acc = New-PubRDSApp -Name $AccAppName -Target $AccAppTarget -ParentFolder $Fld_Acc -PublishFrom All -WinType Maximized -StartOnLogon
$App_Sales = New-PubRDSApp -Name $SalesAppName -Target $SalesAppTarget -ParentFolder $Fld_Sales -PublishFrom Server -PublishFromServer $RDS1

#Update default settings used to configure published resources.
log "Updating Publishing default settings"
Set-PubDefaultSettings -CreateShortcutOnDesktop $true

#Override shortcut default settings for a specific published application.
log "Overriding shortcut default settings for the Sales published application."
Set-PubRDSApp -InputObject $App_Sales -InheritShorcutDefaultSettings $false -CreateShortcutOnDesktop $false

###### PUB FILTERING CONFIGURATION ######

#Set AD account filters by ID.
log "Set Active Directory filters for Accounts published desktop"
Set-PubItemUserFilter -Id $Desk_Acc.Id -Enable $true
Add-PubItemUserFilter -Id $Desk_Acc.Id -Account "Accounts"

#Set AD account filters by object.
log "Set Active Directory filters for Sales published desktop"
Set-PubItemUserFilter -InputObject $Desk_Sales -Enable $true -Replicate $true
Add-PubItemUserFilter -InputObject $Desk_Sales -Account "Sales"

#Set an IP filter (with range) on application.
log "Set IP filters for Accounts published application"
Set-PubItemIPFilter -InputObject $App_Acc -Enable $true
Add-PubItemIPFilter -InputObject $App_Acc -IP $AccAppIPFilter

#Apply all settings. This cmdlet performs the same action as the Apply button in the RAS console.
log "Appling settings"
Invoke-Apply

#End the current RAS session.
log "Ending RAS session"
Remove-RASSession

log "All Done"