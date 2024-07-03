<#  
.SYNOPSIS  
    PArallels RAS auto-deploy script for Azure MarketPlace Deployments
.NOTES  
    File Name  : RAS_Azure_MP_AIO_Install.ps1
    Author     : Freek Berson
    Version    : v0.0.1
    Date       : Jul 3 2024
.EXAMPLE
    .\RAS_Azure_MP_AIO_Install.ps1
#>

#Collect Parameters
param(
    [Parameter(Mandatory = $true)]
    [string]$addsSelection,

    [Parameter(Mandatory = $false)]
    [string]$domainJoinUserName,

    [Parameter(Mandatory = $false)]
    [string]$domainJoinPassword,

    [Parameter(Mandatory = $false)]
    [string]$RasAdminsGroupAD,

    [Parameter(Mandatory = $false)]
    [string]$domainName,

    [Parameter(Mandatory = $false)]
    [string]$localAdminUser,

    [Parameter(Mandatory = $false)]
    [string]$localAdminPassword,

    [Parameter(Mandatory = $true)]
    [string]$maU,

    [Parameter(Mandatory = $true)]
    [string]$maP
)

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
function Set-RunOnceScriptForAllUsers {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    # Ensure the script file exists
    if (-not (Test-Path $ScriptPath)) {
        Write-Error "Script file does not exist at the specified path: $ScriptPath"
        return
    }

    # Registry path for RunOnce in HKLM
    $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"

    # Create a command to run PowerShell with the specified script
    $command = "PowerShell -File `"$ScriptPath`""

    # Add the command to the RunOnce registry key
    try {
        Set-ItemProperty -Path $registryPath -Name "RunMyScriptOnceForAllUsers" -Value $command
        Write-Host "The script at '$ScriptPath' will be executed at the next logon of any user."
    }
    catch {
        Write-Error "Failed to set registry value. Error: $_"
    }
}

#Set variables
$Temploc = 'C:\install\RASInstaller.msi'
$installPath = "C:\install"
$downloadURLRAS = 'https://download.parallels.com/ras/latest/RASInstaller.msi'
$hostname = hostname

# Check if the install path already exists
if (-not (Test-Path -Path $installPath)) { New-Item -Path $installPath -ItemType Directory }

#Configute logging
$Logfile = "C:\install\RAS_Azure_MP_Install.log"
function WriteLog {
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
}

#Disable Server Manager from starting at logon
WriteLog "Disabling Server Manager from starting at logon"
schtasks /Change /TN "Microsoft\Windows\Server Manager\ServerManager"  /Disable

#Download the latest RAS installer 
WriteLog "Dowloading most recent Parallels RAS Installer"
$RASMedia = New-Object net.webclient
$RASMedia.Downloadfile($downloadURLRAS, $Temploc)
WriteLog "Dowloading most recent Parallels RAS Installer done"

if ($addsSelection -eq "adds") {
    #Impersonate user with admin permissions to install RAS and administrators to manage RAS
    WriteLog "Impersonating user"
    Add-LocalGroupMember -Group "Administrators" -Member $domainJoinUserName
    New-ImpersonateUser -Username $domainJoinUserName -Domain $domainName -Password $domainJoinPassword
}

#Install RAS Console & PowerShell role
WriteLog "Install Parallels RAS Console and Powershell role"
Start-Process msiexec.exe -ArgumentList "/i C:\install\RASInstaller.msi /quiet /passive /norestart ADDFWRULES=1 /log C:\install\RAS_Install.log" -Wait

# Enable RAS PowerShell module
Import-Module 'C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1'

# Add permissions to the RAS Admins group
if ($addsSelection -eq "adds") {
    WriteLog "New RAS Session for ADDS user"
    New-RASSession -Username $localAdminUser -Password $localAdminPassword
    New-RASAdminAccount $RasAdminsGroupAD
    invoke-RASApply
}
#add permissions to the local admin group
if ($addsSelection -eq "workgroup") {
    WriteLog "New RAS Session for workgroup user"
    New-RASSession -Username $localAdminUser -Password $localAdminPassword
    Set-RASAuthSettings -AllTrustedDomains $false -Domain Workgroup/$hostname
    invoke-RASApply
}
#Activate 30 day trial using Azure MP Parallels Business account
WriteLog "Activating RAS License"
Invoke-RASLicenseActivate -Email $maU -Password $maPSecure
invoke-RASApply

#Add VM Appliance RDS Server
writelog "Adding VM Appliance RDS Server"
New-RASRDS "localhost" -NoInstall -ErrorAction Ignore
invoke-RASApply

# Publish Applications & RDSH Desktop
WriteLog "Publishing Applications & RDSH Desktop"
New-RASPubRDSDesktop -Name "Published Desktop"
New-RASPubRDSApp -Name "Calculator" -Target "C:\Windows\System32\calc.exe" -PublishFrom All -WinType Maximized
New-RASPubRDSApp -Name "Paint" -Target "C:\Windows\System32\mspaint.exe" -PublishFrom All -WinType Maximized
New-RASPubRDSApp -Name "WordPad" -Target "C:\Program Files\Windows NT\Accessories\wordpad.exe"  -PublishFrom All -WinType Maximized 
invoke-RASApply

if ($addsSelection -eq "adds") {
    WriteLog "Add AD group as RAS Admins"
    New-RASAdminAccount $RasAdminsGroupAD
}

invoke-RASApply

#Deploy Run Once script to launch post deployment actions at next admin logon
WriteLog "Deploying Run Once script to launch post deployment actions at next admin logon"
$basePath = 'C:\Packages\Plugins\Microsoft.Compute.CustomScriptExtension'
$latestVersionFolder = Get-ChildItem -Path $basePath -Directory | Sort-Object Name -Descending | Select-Object -First 1

if ($null -ne $latestVersionFolder) {
    # Construct the full script path
    $scriptPath = Join-Path -Path $latestVersionFolder.FullName -ChildPath 'Downloads\0\RAS_Azure_MP_AIO_Login.ps1'

    # Run the command with the constructed script path
    Set-RunOnceScriptForAllUsers -ScriptPath $scriptPath
}
else {
    WriteLog "No version subfolders found in '$basePath'."
}

if ($addsSelection -eq "adds") {
    #Remove impersonation
    WriteLog "Removing impersonation"
    Remove-ImpersonateUser
}

WriteLog "Finished installing RAS..."
