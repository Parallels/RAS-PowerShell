using assembly "presentationframework"
using namespace System.Windows
using namespace System.Windows.Controls
using namespace Microsoft.Win32
using module ".\RASDeployWindow.psm1"
# $VerbosePreference = 'Continue'
# $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
# $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
# if($isAdmin -eq $false) {
#     Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
#     return
# }
try {
	$mainWindow = [RASDeployWindow]::new();
	
    $mainWindow.btnDeploy.Add_Click({ 
        try {
            $mainWindow.OnDeploy()
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    });

    $mainWindow.btnAddRDSH.Add_Click({
        try {
            $mainWindow.OnAddRDSH()
        }
        catch {
            Write-Warning $_.Exception
        }
    });

    $mainWindow.btnRemoveRDSH.Add_Click({ 
        try {
            $mainWindow.OnRemoveRDSH()
        }
        catch {
            Write-Warning $_.Exception
        }
    });

    $mainWindow.txtLicenseKey.Add_TextChanged({
        try {
            $mainWindow.OnLicenseKeyUpdated()
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    })

	$mainWindow.btnBrowseMsi.Add_Click({
		try {
			$mainWindow.OnBrowseMsi()
		}
		catch {
			Write-Warning $_.Exception.Message
		}
	});

	$mainWindow.btnDownloadMsi.Add_Click({
		try {
			$mainWindow.OnDownloadMsi()
		}
		catch {
			Write-Warning $_.Exception.Message
		}
	});

    $mainWindow.Show();
}
catch [Exception]{
    Write-Warning $_.Exception
}

