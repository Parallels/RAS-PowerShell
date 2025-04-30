## ==================================================================
##
## Copyright (c) 2005-2024 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Parallels RAS auto-deploy script for Azure MarketPlace Deployments
.DESCRIPTION  
    This script is part of the Parallels RAS auto-deploy script for Azure MarketPlace Deployments and performs the installation and configuration of the primary Parallels RAS Connection Broker.
.NOTES  
    File Name  : RAS_Azure_MP_Primary_CB.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\RAS_Azure_MP_Primary_CB.ps1
#>

#Collect Parameters
param(
    [Parameter(Mandatory = $true)]
    [string]$domainJoinUserName,

    [Parameter(Mandatory = $true)]
    [string]$domainJoinPassword,

    [Parameter(Mandatory = $true)]
    [string]$domainName,

    [Parameter(Mandatory = $true)]
    [string]$numberofCBs,

    [Parameter(Mandatory = $true)]
    [string]$numberofSGs,

    [Parameter(Mandatory = $true)]
    [string]$prefixCBName,

    [Parameter(Mandatory = $true)]
    [string]$prefixSGName,

    [Parameter(Mandatory = $true)]
    [string]$RasAdminsGroupAD,

    [Parameter(Mandatory = $true)]
    [string]$downloadURLRAS,

    [Parameter(Mandatory = $false)]
    [string]$license,

    [Parameter(Mandatory = $false)]
    [string]$maU,

    [Parameter(Mandatory = $false)]
    [string]$maP
   
)

#Set variables
$Temploc = 'C:\install\RASInstaller.msi'
$installPath = "C:\install"
$secdomainJoinPassword = ConvertTo-SecureString $domainJoinPassword -AsPlainText -Force
$primaryConnectionBroker = $prefixCBName + "-1" + "." + $domainName

#Set Windows Update to "NoAutoUpdate" to prevent automatic installation of updates
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Value 1

function New-ImpersonateUser {

    [cmdletbinding()]
    Param( 
        [Parameter(ParameterSetName = "ClearText", Mandatory = $true)][string]$Username, 
        [Parameter(ParameterSetName = "ClearText", Mandatory = $true)][string]$Domain, 
        [Parameter(ParameterSetName = "ClearText", Mandatory = $true)][string]$Password, 
        [Parameter(ParameterSetName = "Credential", Mandatory = $true, Position = 0)][PSCredential]$Credential, 
        [Parameter()][Switch]$Quiet 
    ) 
 
    #Import the LogonUser Function from advapi32.dll and the CloseHandle Function from kernel32.dll
    if (-not ([System.Management.Automation.PSTypeName]'Import.Win32').Type) {
        Add-Type -Namespace Import -Name Win32 -MemberDefinition @'
            [DllImport("advapi32.dll", SetLastError = true)]
            public static extern bool LogonUser(string user, string domain, string password, int logonType, int logonProvider, out IntPtr token);
  
            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool CloseHandle(IntPtr handle);
'@ -ErrorAction SilentlyContinue
    }
    #Set Global variable to hold the Impersonation after it is created so it may be ended after script run
    $Global:ImpersonatedUser = @{} 
    #Initialize handle variable so that it exists to be referenced in the LogonUser method
    $tokenHandle = 0 
 
    #Pass the PSCredentials to the variables to be sent to the LogonUser method
    if ($Credential) { 
        Get-Variable Username, Domain, Password | ForEach-Object { 
            Set-Variable $_.Name -Value $Credential.GetNetworkCredential().$($_.Name) } 
    } 
 
    #Call LogonUser and store its success. [ref]$tokenHandle is used to store the token "out IntPtr token" from LogonUser.
    $returnValue = [Import.Win32]::LogonUser($Username, $Domain, $Password, 2, 0, [ref]$tokenHandle) 
 
    #If it fails, throw the verbose with the error code
    if (!$returnValue) { 
        $errCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error(); 
        Write-Host "Impersonate-User failed a call to LogonUser with error code: $errCode" 
        throw [System.ComponentModel.Win32Exception]$errCode 
    } 
    #Successful token stored in $tokenHandle
    else { 
        #Call the Impersonate method with the returned token. An ImpersonationContext is returned and stored in the
        #Global variable so that it may be used after script run.
        $Global:ImpersonatedUser.ImpersonationContext = [System.Security.Principal.WindowsIdentity]::Impersonate($tokenHandle) 
     
        #Close the handle to the token. Voided to mask the Boolean return value.
        [void][Import.Win32]::CloseHandle($tokenHandle) 
 
        #Write the current user to ensure Impersonation worked and to remind user to revert back when finished.
        if (!$Quiet) { 
            Write-Host "You are now impersonating user $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" 
            Write-Host "It is very important that you call Remove-ImpersonateUser when finished to revert back to your user."
        } 
    } 

    Function Global:Remove-ImpersonateUser { 
        <#
        .SYNOPSIS
        Used to revert back to the orginal user after New-ImpersonateUser is called. You can only call this function once; it is deleted after it runs.
  
        .INPUTS
        None. You cannot pipe objects to Remove-ImpersonateUser
  
        .OUTPUTS
        None. Remove-ImpersonateUser does not generate any output.
        #> 
 
        #Calling the Undo method reverts back to the original user.
        $ImpersonatedUser.ImpersonationContext.Undo() 
 
        #Clean up the Global variable and the function itself.
        Remove-Variable ImpersonatedUser -Scope Global 
        Remove-Item Function:\Remove-ImpersonateUser 
    } 
}

# Check if the install path already exists
if (-not (Test-Path -Path $installPath)) { New-Item -Path $installPath -ItemType Directory }

#Configute logging
$Logfile = "C:\install\RAS_Azure_MP_CB_prereq.log"
function WriteLog {
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
}

#Set Windows Update to "Download Only" to prevent automatic installation of updates
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AUOptions" -Value 2

#Create Firewall Rules
WriteLog "Configuring Firewall Rules"
New-NetFirewallRule -DisplayName "Parallels RAS Administration (TCP)" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 68, 80, 81, 1234, 135, 443, 445, 20001, 20002, 20003, 20009, 20020, 20030, 20443, 30004, 30006
New-NetFirewallRule -DisplayName "Parallels RAS Administration (TCP)" -Direction Inbound -Action Allow -Protocol UDP -LocalPort 80, 443, 20000, 20009, 30004, 30006

#Download the latest RAS installer
WriteLog "Dowloading most recent Parallels RAS Installer"
$RASMedia = New-Object net.webclient
$RASMedia.Downloadfile($downloadURLRAS, $Temploc)

#Impersonate user with local admin permissins to install RAS
WriteLog "Impersonating user"
Add-LocalGroupMember -Group "Administrators" -Member $domainJoinUserName
New-ImpersonateUser -Username $domainJoinUserName -Domain $domainName  -Password $domainJoinPassword

#Install RAS Connection Broker role
WriteLog "Install Connection Broker role"
Start-Process msiexec.exe -ArgumentList "/i C:\install\RASInstaller.msi ADDFWRULES=1 ADDLOCAL=F_Controller,F_PowerShell /norestart /qn /log C:\install\RAS_Install.log" -Wait 
cmd /c "`"C:\Program Files (x86)\Parallels\ApplicationServer\x64\2XRedundancy.exe`" -c -AddRootAccount $domainJoinUserName"
start-sleep -Seconds 30

# Replace instances of '../4.0' with './4.0'
$filePath = "C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1"
$content = Get-Content -Path $filePath
$updatedContent = $content -replace "../4.0", "./4.0"
Set-Content -Path $filePath -Value $updatedContent

# Enable RAS PowerShell module
WriteLog "Import RAS PowerShell Module"
Import-Module 'C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1'

#Create new RAS PowerShell Session
start-sleep -Seconds 10
WriteLog "Create new RAS PowerShell Session"
New-RASSession -Username $domainJoinUserName -Password $secdomainJoinPassword

#Add AD group as RAS Admins
WriteLog "Add AD group as RAS Admins"
New-RASAdminAccount $RasAdminsGroupAD

#Add Secure Gateways
WriteLog "Add secure gateways"
for ($i = 1; $i -le $numberofSGs; $i++) {
    $paddedNumber = $i.ToString("D2")
    $secureGateway = "$prefixSGName-$paddedNumber.$domainName"
    New-RASGateway -Server $secureGateway
    Start-Sleep -Seconds 10
}
Invoke-RASApply

#Add secondary Connection Brokers
WriteLog "Add secondary Connection Brokers"
for ($i = 2; $i -le $numberofCBs; $i++) {
    $paddedNumber = $i.ToString("D2")
    $connectionBroker = "$prefixCBName-$paddedNumber.$domainName"
    New-RASBroker -Server $connectionBroker -Takeover
    Start-Sleep -Seconds 10
    if ($i -eq 4) { Set-RASBroker -Server $connectionBroker -enabled $false }
}
Invoke-RASApply

WriteLog "Remove RAS PowerShell Session"
Remove-RASSession

WriteLog "Remove impersonation"
Remove-ImpersonateUser

try {
    if ($license -eq 'trial') {
        #Create new RAS PowerShell Session
        start-sleep -Seconds 10
        WriteLog "Create new RAS PowerShell Session"
        New-RASSession -Username $domainJoinUserName -Password $secdomainJoinPassword
        #Activate 30 day trial using Azure MP Parallels Business account
        WriteLog "Activating RAS License"
        $maPSecure = ConvertTo-SecureString $maP -AsPlainText -Force
        Invoke-RASLicenseActivate -Email $maU -Password $maPSecure
        invoke-RASApply
        start-sleep -Seconds 10
    }
}
Catch {
    WriteLog $_.Exception.Message
    exit
}

WriteLog "Restart to finish installation of RAS Connection Broker role"
shutdown -r -f -t 0
