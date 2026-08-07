﻿## ==================================================================
##
## Copyright (c) 2005-2025 Parallels Software International, Inc.
## Released under the terms of MIT license (see LICENSE for details)
##
## ==================================================================

<#  
.SYNOPSIS  
    Parallels RAS auto-deploy script for AWS MarketPlace Deployments
.DESCRIPTION  
    This script is part of the Parallels RAS auto-deploy script for AWS MarketPlace Deployments and performs an all-in-one installation of Parallels RAS on a Windows Server VM.
.NOTES  
    File Name  : RAS_AWS_MP_AIO_Install.ps1
    Author     : www.parallels.com
.EXAMPLE
    .\RAS_AWS_MP_AIO_Install.ps1
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$localAdminUser,

    [Parameter(Mandatory = $false)]
    [string]$localAdminPassword,

    [Parameter(Mandatory = $true)]
    [string]$downloadURLRAS
)

function Set-RunOnceScriptForAllUsers {

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    if (-not (Test-Path $ScriptPath)) {
        WriteLog "Script file not found: $ScriptPath"
        return
    }

    $registryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"

    $command = "PowerShell.exe -ExecutionPolicy Bypass -File `"$ScriptPath`""

    Set-ItemProperty `
        -Path $registryPath `
        -Name "RunParallelsRASLogin" `
        -Value $command
}

# Set variables
$Temploc = 'C:\install\RASInstaller.msi'
$installPath = "C:\install"
$hostname = hostname

$localAdminPasswordSecure =
    ConvertTo-SecureString `
        $localAdminPassword `
        -AsPlainText `
        -Force

# Create install folder
if (-not (Test-Path -Path $installPath)) {
    New-Item `
        -Path $installPath `
        -ItemType Directory `
        -Force
}

# Logging
$Logfile = "C:\install\RAS_AWS_MP_Install.log"

function WriteLog {
    Param ([string]$LogString)

    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")

    Add-Content `
        $LogFile `
        "$Stamp $LogString"
}

WriteLog "Starting AWS Marketplace deployment"

# Disable Server Manager
schtasks `
    /Change `
    /TN "Microsoft\Windows\Server Manager\ServerManager" `
    /Disable

# Download latest RAS
WriteLog "Downloading latest Parallels RAS"

$RASMedia = New-Object Net.WebClient

$RASMedia.DownloadFile(
    $downloadURLRAS,
    $Temploc
)

WriteLog "Download complete"

# Install RAS
WriteLog "Installing Parallels RAS"

Start-Process `
    msiexec.exe `
    -ArgumentList "/i C:\install\RASInstaller.msi /quiet /passive /norestart ADDFWRULES=1 /log C:\install\RAS_Install.log" `
    -Wait

Start-Sleep -Seconds 30

# Fix module path issue
$filePath =
"C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1"

if (Test-Path $filePath) {

    $pattern = [regex]::

    $replacement = "./4.0"

    $content =
        Get-Content `
            $filePath `
            -Raw

    $updatedContent =
        $content -replace `
            $pattern,
            $replacement

    Set-Content `
        -Path $filePath `
        -Value $updatedContent
}

# Import RAS PowerShell module
Import-Module `
"C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\RASAdmin.psd1"

# Add all local admins as RAS root admins
WriteLog "Configuring Root Admins"

$allLocalAdmins =
    Get-LocalGroupMember `
        -Group "Administrators"

foreach ($localAdmin in $allLocalAdmins) {

    cmd /c `
    "`"C:\Program Files (x86)\Parallels\ApplicationServer\x64\2XRedundancy.exe`" -c -AddRootAccount $localAdmin"
}

# Create RAS Session
WriteLog "Creating RAS Session"

New-RASSession `
    -Username $localAdminUser `
    -Password $localAdminPasswordSecure

Set-RASAuthSettings `
    -AllTrustedDomains $false `
    -Domain Workgroup/$hostname

Invoke-RASApply

# Trial activation intentionally removed
WriteLog "Skipping activation (AWS Marketplace v0)"

# Add localhost as RDS Host
WriteLog "Adding localhost RDS host"

New-RASRDSHost `
    "localhost" `
    -NoInstall `
    -ErrorAction Ignore

Invoke-RASApply

# Publish Desktop
New-RASPubRDSDesktop `
    -Name "Published Desktop"

# Publish apps
New-RASPubRDSApp `
    -Name "Calculator" `
    -Target "C:\Windows\System32\calc.exe" `
    -PublishFrom All `
    -WinType Maximized

New-RASPubRDSApp `
    -Name "Paint" `
    -Target "C:\Windows\System32\mspaint.exe" `
    -PublishFrom All `
    -WinType Maximized

New-RASPubRDSApp `
    -Name "Notepad" `
    -Target "C:\Windows\System32\notepad.exe" `
    -PublishFrom All `
    -WinType Maximized

New-RASPubRDSApp `
    -Name "Snipping Tool" `
    -Target "C:\Windows\System32\SnippingTool.exe" `
    -PublishFrom All `
    -WinType Maximized

Invoke-RASApply

# Configure login helper
WriteLog "Configuring RunOnce login helper"

$scriptPath =
"C:\install\RAS_AWS_MP_AIO_Login.ps1"

Set-RunOnceScriptForAllUsers `
    -ScriptPath $scriptPath

# Disable OOBE noise
Set-ItemProperty `
    -Path "HKLM:\SOFTWARE\Microsoft\ServerManager" `
    -Name "DoNotOpenServerManagerAtLogon" `
    -Value 1 `
    -Type DWord `
    -Force

# Install RDSH Role
# No -Restart for CloudFormation deployments
WriteLog "Installing RDS-RD-Server Role"

Add-WindowsFeature `
    -Name "RDS-RD-Server"

WriteLog "Finished AWS Marketplace deployment"
