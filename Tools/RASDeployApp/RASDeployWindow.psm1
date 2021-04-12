using assembly "presentationframework"
using assembly windowsbase
using module "..\PSUI\PSWindow\PSWindow.psm1"
using module BitsTransfer
using namespace Microsoft.Win32
using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Media.Imaging

#requires -version 5
<#
.SYNOPSIS
Contains definition of RASDeployWindow
.DESCRIPTION
TODO:
.INPUTS
None
.OUTPUTS
PSWindow UI
.NOTES
Version:        1.0
Purpose:        Showcase powershell capabilities with RASAdmin

.EXAMPLE
using module .\PSUI\PSUI.psd1
#>

class RASDeployWindow : PSWindow {
	[string] $scriptPath
	[string] $pathToMsi
	[string] $msiName

    RASDeployWindow() :base("resources\MainWindow.xaml") {
        $this.scriptPath = (Resolve-Path .\).Path
        $bmpLogo = [BitmapImage]::new("$($this.scriptPath)\logo.png");
        $this.Form.Icon = $bmpLogo
        $this.txtMasterPAGW.Text = [System.Net.Dns]::GetHostName()
        $this.txtLicenseKey.Text = "Trial"
        $this.OnLicenseKeyUpdated()
    }

    hidden [bool] IsEmpty([string] $value) {
        if ([string]::IsNullOrEmpty($value)) {
            return $true
        }
        else {
            return $false
        }
    }

    hidden DeployButtonState([bool] $state) {
        $this.btnDeploy.IsEnabled= $state
    }
    
    hidden [ListBoxItem] CreateListItem([string] $content) {
        $item = [ListBoxItem]::new()
        $item.Content = $content
        return $item
    }
	hidden [void] Log([string] $type, [string] $message) {
		$date = Get-Date
		$dateChunk = "[$($date)]"
		$typeChunk = "[$($type)]"
		$arrows = " ---> "

		Write-Host $dateChunk -ForegroundColor Yellow -NoNewline
		Write-Host $typeChunk -ForegroundColor Green -NoNewline
		Write-Host $arrows -NoNewline
		Write-Host $message
		$this.txtInfo.Text = $message
	}
	
	hidden [void] Log([string] $type, [string] $message, [Exception] $exception) {
		$date = Get-Date
		$dateChunk = "[$($date)]"
		$typeChunk = "[$($type)]"
		$arrows = " ---> "
		Write-Host $dateChunk -ForegroundColor Gray -NoNewline
		Write-Host $typeChunk -ForegroundColor Red -NoNewline
		Write-Host $arrows -NoNewline
		Write-Host "$message. Error was: $($exception.Message)"
		$this.txtInfo.Text = "$message. Error: $($exception.Message)"
	}
	
    [void] DownloadLatest() {
        try {            

            $this.msiName = ""
			$msiPath = ""
			 
			$uriVersion = "http://download.parallels.com/ras/RAS_16.5.xml"
			$versionFileName = "product_version.xml"
			$this.Log("INFO", "Getting $versionFileName from $uriVersion.")
			Invoke-RestMethod -Method Get -Uri $uriVersion -ContentType "text/xml; charset=utf-8" -OutFile "$($env:TEMP)\$($versionFileName)"
			$this.Log("INFO", "Retrieved $versionFileName. File saved in $($env:TEMP)\$($versionFileName)")

			$msiPath = $env:TEMP
			[string] $res = (Get-Content "$($msiPath)\$($versionFileName)")
			[xml] $productXML = [xml]::new()
			$productXML.InnerXml = $res

			$version = $productXML.Product.Version
			# $displayName =$productXML.Product.DisplayName
			# $message = $productXML.Product.Message
			$result = [MessageBox]::Show("Are you you want to download Parallels RAS version $version ?", "Parallels RAS.", [System.Windows.MessageBoxButton]::YesNo)
			if ($result -eq [MessageBoxResult]::No) {
				return
			}
			$this.Log("INFO", "RAS installer version $version")
			$this.msiName= "RASInstaller-$($version).msi"
			
			if ([System.IO.File]::Exists("$($msiPath)\$($this.msiName)") -eq $false) {
				$this.Log("INFO", "No installer found in $msiPath.")
				$this.Log("INFO", "Downloading latest Parallels RAS installer from $($productXML.Product.MSIPackageURL).")
				Start-BitsTransfer -Source $productXML.Product.MSIPackageURL -Destination "$msiPath\$($this.msiName)"
				$this.Log("INFO", "Download complete. File saved in $msiPath\$($this.msiName)")
			}
			$this.pathToMsi = "$msiPath\$($this.msiName)"
        }
        catch [Exception] {
			$this.Log("ERROR", "DownloadLatest() failed.", $_.Exception)
            throw $_.Exception
        }
    }
	
	hidden [bool] InstallMsi([string]$path) {
		$this.Log("INFO", "Started installtion using $path")
		$process = Start-Process msiexec.exe -Wait -ArgumentList "/I $path /qb ADDLOCAL=""F_PowerShell"""

		if ($process.HasExited) {
			$this.Log("INFO", "Installation process has exited with exit code $($process.ExitCode)")
			if ($process.ExitCode -ne 0) {
				$this.Log("INFO", "Installation was not completed.")
				return $false
			}
		}

		$rasAdmin = "C:\Program Files (x86)\Parallels\ApplicationServer\Modules\RASAdmin\2.0\RASAdmin.dll"
		$this.Log("INFO", "Trying to import RASAdmin module from path $($rasAdmin)")
		$importError = $null
		Import-Module "$($rasAdmin)" -ErrorAction Continue -ErrorVariable "importError"
		if ($importError) {
			$this.Log("ERROR", "Failed to import RASAdmin.", $importError)
			return $false
		}
		$this.Log("INFO", "RASAdmin module imported.")
		return $true
	}

    hidden [void] NewGW ([string] $server) {
		if ($this.IsEmpty($server)) {
			throw "Server is null or empty."
		}

		try {
			$this.Log("INFO", "Creating New gateway for $server ...")
			New-RASGW $server
			$this.Log("INFO", "Gateway created.")
		}
		catch {
			$this.Log("ERROR", "Failed to create new gateway.", $_.Exception)
			throw "Failed to add new GW. $($_.Exception)"
		}
    }
    
    hidden [void] NewPA([string] $server) {
        if ($this.IsEmpty($server)) {
            throw "Server was null or empty"
        }

        try {
			$this.Log("INFO", "Creating new publishing agent for $server ...")
			New-RASPA $server
			$this.Log("INFO", "Publishing agent created.")
        }
        catch {
			$this.Log("ERROR", "Failed to create publishing agent.", $_.Exception)
            throw "NewPA Failed. Server [$($server)]. $($_.Exception)"
        }
    }
    
    hidden [void] NewPAGW([string] $server) {
        try {
            if ($this.IsEmpty($server)) {
                throw "server is empty or null."
            }

            $this.NewPA($server)
            $this.NewGW($server)
        }
        catch {
            throw "Failed to create new PA & GW for [$($server)].$($_.Exception)"
        }        
    }

    hidden [void] NewRASFarm ([string]$username, [securestring]$pswd) {
        if ($this.IsEmpty($username)) {
            throw "Username is empty"
        }
        
        if ($this.IsEmpty($pswd)) {
            throw "Password is empty"
        }

        $masterPA = $this.txtMasterPAGW.Text
        try {
            if ($this.IsEmpty($masterPA)) {
                throw "server is empty or null."
			}
			$this.Log("INFO", "Creating new ras farm on $masterPA for $username")
			New-RASFarm -Server $masterPA -Username $username -Password  $pswd
			$this.Log("INFO", "Farm created.")
        }
        catch {
			$this.Log("ERROR", "Failed to create new RAS farm.", $_.Exception)
            throw "Failed to create new Farm on Primary PA [$($masterPA)].$($_.Exception)"
        }    
    }

    hidden [void] NewSecondaryPAGW() {
        $secondaryPA_GW = $this.txtSecondaryPAGW.Text
        if ($this.IsEmpty($secondaryPA_GW) -eq $false) {
            $this.NewPAGW($secondaryPA_GW)
        }
    }
    
    hidden [void] NewRASSession([string]$username, [securestring]$pswd) {
        if ($this.IsEmpty($username)) {
            throw "Username is empty"
        }
        
        if ($this.IsEmpty($pswd)) {
            throw "Password is empty"
        }

        try {
			$masterPA = $this.txtMasterPAGW.Text
			$this.Log("INFO", "Creating new RAS session to $masterPA for $username ...")
			New-RASSession -server $masterPA -Username $username -Password $pswd -Retries 5
			$this.Log("INFO", "Session created.")			
        }
        catch {
			$this.Log("ERROR", "Failed to create RAS session.", $_.Exception)
            throw "NewRASSession Failed. $($_.Exception)"
        }
    }
    
    hidden [void] ActivateLicense([string] $key) {
        if ($this.IsEmpty($this.txtPrlsEmail.Text)) {
            throw "Parallels email is empty."
        }

        if ($this.IsEmpty($this.txtPrlsPassword.Password)) {
            throw "Parallels password is empty"
        }

        $prlsEmail = $this.txtPrlsEmail.Text
		$prlsPass = ConvertTo-SecureString -AsPlainText $this.txtPrlsPassword.Password -Force
		try {
			$this.Log("INFO", "Activating license as trial for $prlsEmail")
			if ($key -eq 'Trial' -or [string]::IsNullOrEmpty($key)) {
				Invoke-RASLicenseActivate -Email $prlsEmail -Password $prlsPass
			} else {
				Invoke-RASLicenseActivate -Email $prlsEmail -Password $prlsPass -Key $key
			}
			$this.Log("INFO", "Activation successful.")
		}
		catch {
			$this.Log("ERROR", "Failed to activate license.", $_.Exception)
			throw
		}
    }
    
    hidden [void] AddRDSHList() {
		foreach ($rdsh in $this.lstRDSH.Items) {
			if ($this.IsEmpty($rdsh.Content) -eq $false) {
				try {
					$this.Log("INFO", "Adding new RDS $($rdsh.Content)")
					New-RASRDS $rdsh.Content
					$this.Log("INFO", "RDS added.")
				}
				catch {
					$this.Log("ERROR", "Failed to add RDS.", $_.Exception)
					throw "Failed to add [$($rdsh.Content)]. $($_.Exception)"
				}
			}
		}            
    }
    
    hidden [void] PubApp([string] $name, [string] $target) {
        if ($this.IsEmpty($name)) {
            throw "Application name was null or empty."
        }

        if ($this.IsEmpty($target)) {
            throw "Target was null or empty."
        }
        
        try {
			$this.Log("INFO", "Creating new published application $name using $target ...")
			New-RASPubRDSApp -Name $name -Target $target
			$this.Log("INFO", "Published application created.")
        }
        catch {
			$this.Log("ERROR", "Failed to created published application.")
            throw "Failed to add [$($name) | $($target)]. $($_.Exception)"
        }
    }
    
    hidden [void] PubDesktop([string] $name) {
        if ($this.IsEmpty($name)) {
            throw "Desktop name was null or empty."
        }
        
		try {
			$this.Log("INFO", "Creating new published desktop ...")
			New-RASPubRDSDesktop -Name "Desktop"
			$this.Log("INFO", "Published desktop created.")
		}
		catch {
			$this.Log("ERROR", "Failed to create publsihed desktop.", $_.Exception)
			throw "Failed to add desktop [Desktop]. $($_.Exception)"
		}
    }

    [void] OnLicenseKeyUpdated() {
        if ($this.IsEmpty($this.txtLicenseKey.Text) -or 
            $this.txtLicenseKey.Text.Equals("Trial") -or $this.txtLicenseKey.Text.Equals("trial"))
        {
            $this.lblLicenseStatus.Content = "No key specified. Trial version will be installed."
            $this.lblLicenseStatus.Visibility = "Visible"
        }
        else {
            $this.lblLicenseStatus.Content = [string]::Empty
            $this.lblLicenseStatus.Visibility = "Hidden"
        }
    }

    [void] OnAddRDSH() {
        try {
            $rdsh = $this.txtRdSessionHost.Text
            if ($this.IsEmpty($rdsh) -eq $false) {
                $item = $this.CreateListItem($rdsh)
                $this.lstRDSH.Items.Add($item)
                $this.txtRdSessionHost.Text = [string]::Empty
            }
        }
        catch [Exception] {
            throw "OnAddRDSH failed. $($_.Exception)"
        }
    }

    [void] OnRemoveRDSH() {
        try {
            while ($this.lstRDSH.SelectedItems.Count -ne 0) {
                $this.lstRDSH.Items.Remove($this.lstRDSH.SelectedItems[0])
            }
        }
        catch [Exception] {
            throw $_.Exception
        }
	}
	
	[void] OnBrowseMsi() {
		$fileDialogue = [OpenFileDialog]::new()
		$fileDialogue.DefaultExt = ".msi"
		$result = $fileDialogue.ShowDialog();
		if ($result) {
			  $this.pathToMsi = $fileDialogue.FileName
			  $this.txtPath.Text = $this.pathToMsi
		}
	}

	[void] OnDownloadMsi() {
		$this.DownloadLatest()
		$this.txtPath.Text = $this.pathToMsi
	}

    [void] OnDeploy() {
        try {
			Clear-Host
			$this.Log("INFO", "Deployment started.")
            if ($this.chkEnableVerbosity.IsChecked -eq $true) {
                $VerbosePreference = 'Continue'
            }

            $username = $this.txtUsername.Text
            $password = $this.txtPassword.Password
            $licenseKey = $this.txtLicenseKey.Text

            $pswd = ConvertTo-SecureString $password -AsPlainText -Force
            
			# $this.DownloadLatest()
			$this.Log("INFO", "Started InstallMsi() ...")
			$this.InstallMsi($this.pathToMsi)

			$this.Log("INFO", "Started NewRASFarm() ...")
            $this.NewRASFarm( $username, $pswd)

			$this.Log("INFO", "Started NewRASSession() ...")
            $this.NewRASSession($username, $pswd)

			$this.Log("INFO", "Started ActivateLicense() ...")
            $this.ActivateLicense($licenseKey)

			$this.Log("INFO", "Started NewSecondaryPAGW() ...")
			$this.NewSecondaryPAGW()

			$this.Log("INFO", "Started AddRDSHList() ...")
            $this.AddRDSHList()

            if ($this.chkPubDesktop.IsChecked -eq $true) {
				$this.Log("INFO", "Started PubDesktop() ...")
                $this.PubDesktop("Desktop")
            }
            if ($this.chkPubCalc.IsChecked -eq $true) {
				$this.Log("INFO", "Started PubApp() for calc.exe ...")
                $this.PubApp("calc", "C:\Windows\System32\calc.exe")
            }
            if ($this.chkPubNotepad.IsChecked -eq $true) {
				$this.Log("INFO", "Started PubApp() for notepad.exe ...")
                $this.PubApp("notepad", "C:\Windows\System32\notepad.exe")
			}
			$this.Log("INFO", "Applying settings ...")
			Invoke-RASApply
			$this.Log("INFO", "Settings applied")
			$this.Log("INFO", "Deployment completed!")
        }
        catch [Exception] {
			$this.DeployButtonState($true)
			$this.Log("ERROR", "Deployment failed.", $_.Exception)
            throw  "OnDeploy Failed. $($_.Exception)"
        } 
        finally {
            $VerbosePreference = 'SilentlyContinue'
        }
    }
}
