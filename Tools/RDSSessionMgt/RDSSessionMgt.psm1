#requires -version 5
<#
.SYNOPSIS
  This is the "client" module to go with "client" script. This module contains 
  class definitions specific to RAS. 
.DESCRIPTION
  This module is imported by client.ps1, where all the events are initialized.
.INPUTS
    None
.OUTPUTS
    class definitions
    RASMessageUI
    RASSessionAppUI
    RDSSessionMgt
    !IMPORTANT - Classes can only be imported with the "using module" syntax. Import-Module does not import classes.
.NOTES
  Version:        1.0
  Purpose:        Showcase powershell capabilities with PSAdmin
  
.EXAMPLE
  using module ./client.psm1
#>
using assembly presentationFramework
using assembly windowsbase
using module '../PSUI/PSWindow/PSWindow.psm1'
using namespace 'PSAdmin'
using namespace 'RASAdminEngine.Core.OutputModels'
using namespace System.Security
using namespace System.Collections.ObjectModel
using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Collections.Generic

#################################################################
# Send message UI
#################################################################
class RASMessageUI : PSWindow { 
    RASMessageUI() :base("./resources/MessageWindow.xaml") {}
}

#################################################################
# Session applications UI
#################################################################
class RASSessionAppUI : PSWindow {

    [ObservableCollection[ServerAppInfo]] $Applications
    RASSessionAppUI([string] $path) :base($path) {}

    [void] SetAppsDataGrid([List[ServerAppInfo]] $apps) {
        [ObservableCollection[ServerAppInfo]]$oc = [ObservableCollection[ServerAppInfo]]::new()
        foreach ($app in $apps) {
            $oc.Add($app)
        }
        $this.Applications = $oc
        $this.dtApplications.ItemsSource = $this.Applications


    }
}

#################################################################
# Main application window
#################################################################
class RDSSessionMgt : PSWindow {
    ##################################################
    # public proerties
    ##################################################
    [ObservableCollection[RDPSession]] $Sessions
    ##################################################

    ##################################################
    # ctor
    ##################################################
    RDSSessionMgt([string] $path) :base($path) {
        if ([string]::IsNullOrEmpty($this.txtServer.Text)) {
            $s = [System.Net.Dns]::GetHostName()
            $this.txtServer.Text = $s
        }
    }
    ##################################################

    # Get the scripblock text and assigns it to the the txtCmd.
    # Values contained in the scriptblock will not be evaluated,
    # i.e. some variable $a will be shown as '$a' and not the 
    # value it holds.
    [void] UpdateCommandText([scriptblock]$cmd) {
            $this.txtCmd.Text = $cmd.ToString()
    }

    # Refreshes the dtSessions Items with RDSSession
    [void] Refresh() {
        $tmp = $this.GetRDSSessions()
        if ($tmp -eq $null) {
            Write-Host "Failed to refresh" -ForegroundColor red
            return
        }
        $this.SetSessionsDataGrid($tmp)
        $this.dtSessions.ItemsSource = $tmp
        Write-Host "Refreshed" -ForegroundColor Green
    }

    # Sends a message to the selected session.
    [void] OnBtnSendMessageClick() {
        try {
            if (([DataGrid]$this.dtSessions).SelectedItems.Count -eq 0) {
                Write-Warning "Please select a session"
                return
            }
            $msgWindow = [RASMessageUI]::new()
            $msg = ""
            $rootwnd = $this
            $msgWindow.btnSend.Add_Click({
                $msg = $msgWindow.txtMessage.Text
                if ([String]::IsNullOrEmpty($msg)) {
                    Write-Warning "Please input a message text"
                    return
                }
                [RDPSession]$item = ([DataGrid]$rootwnd.dtSessions).SelectedItem
                $res = {Invoke-RDSSessionCmd -InputObject $item -Command SendMsg -MsgTitle "Message from $(whoami)" -Message $msg}
                $res.Invoke()
                $rootwnd.UpdateCommandText($res)
                $msgWindow.Close()
            })
            $msgWindow.Show()
        } catch [Exception] {
            Write-Host "Failed to send message" -ForegroundColor red
            Write-Host $_.Exception.Message -ForegroundColor red
        }
    }

    # Creates a New-RASSession. Default server value is 127.0.0.1
    [void] OnBtnLoginClick() {
        try {
            $u = $this.txtUsername.Text
            $p = ConvertTo-SecureString -AsPlainText $this.txtPassword.Password -Force
            $s = $this.txtServer.Text
            if ([string]::IsNullOrEmpty($s)) {
                $s = [System.Net.Dns]::GetHostName()
                $this.txtServer.Text = $s
            }
            $cmd = {New-RASSession -Username $u -Password $p -Server $s}
            $cmd.Invoke()
            $this.UpdateCommandText($cmd)
            [MessageBox]::Show("Session created for $($u).")
            Write-Host "Created session for $u" -ForegroundColor Green
            $this.Refresh()
        }
        catch [Exception] {
            Write-Host "Could not login." -ForegroundColor red
            Write-Host $_.Exception.Message -ForegroundColor red  
            [MessageBox]::Show("Could not login. $($_.Exception.Message)")
            
        }
    }

    [void] OnBtnRefreshClick () {
        $this.Refresh()
    }

    # Disconnects from the selected session.
    [void] OnBtnDisconnectSessionClick() {
        try {
            [RDPSession]$item = ([DataGrid]$this.dtSessions).SelectedItem
            if(!$item) {
                Write-Warning "Please select a session"
                return
            }
            $cmd = {Invoke-RDSSessionCmd -InputObject $item -Command Disconnect}
            $cmd.Invoke()
            $this.UpdateCommandText($cmd)            
        }
        catch [Exception] {
            Write-Host "Failed to disconnect session" -ForegroundColor red
            Write-Host $_.Exception.Message -ForegroundColor red
        }
    }
 
    # Log-out a selected session
    [void] OnBtnLogoffSessionClick() {
        try {
            [RDPSession]$item = ([DataGrid]$this.dtSessions).SelectedItem
            if(!$item) {
                Write-Warning "Please select a session"
                return
            }
            $cmd = {Invoke-RDSSessionCmd -InputObject $item -Command Logoff}
            $cmd.Invoke()
            $this.UpdateCommandText($cmd)
            Write-Host "Removed $($item.User)" -ForegroundColor Green
            $this.Sessions.RemoveAt(([DataGrid]$this.dtSessions).SelectedIndex)

        } catch [Exception] {
            Write-Host "Failed to Logoff session" -ForegroundColor red
            Write-Error $_.Exception.Message
        }
    }

    # Retrievies a list of application running on the selected session.
    [void] OnBtnGetAppsClick() {
        try {
            [RDPSession]$item = ([DataGrid]$this.dtSessions).SelectedItem
            if(!$item) {
                Write-Warning "Please select a session"
                return
            }
            $appsWindow = [RASSessionAppUI]::new("./resources/SessionApps.xaml")
            $cmd = {Get-RDSStatus -StatusLevel Level3}
            $this.UpdateCommandText($cmd)
            $tmpStat = $cmd.Invoke()
            foreach($s in $tmpStat.Sessions) {
                if ($s.SessionID -eq $item.SessionID) {
                    $item = $s
                }
            }
            $rootwnd = $this # inside the script block '$this' would refer to the scriptblock instance.
            $appsWindow.SetAppsDataGrid($item.Applications)
            $appsWindow.Form.Title = "Applications [$($item.SessionID) | $($item.User)]"

            $appsWindow.btnKillProcess.Add_Click({
                [ServerAppInfo]$app = ([DataGrid]$appsWindow.dtApplications).SelectedItem
                if (!$app) {
                    Write-Warning "Please select an application"
                    return
                }
                $cmd = {Invoke-RDSKillProcessCmd -InputObject $app}
                $cmd.Invoke()
                $rootwnd.UpdateCommandText($cmd)
                $appsWindow.Applications.RemoveAt(([DataGrid]$appsWindow.dtApplications).SelectedIndex)
                $rootwnd.Refresh()
            })

            $appsWindow.btnRefresh.Add_Click({
                $cmd = {Get-RDSStatus -StatusLevel Level3}
                Write-Host "refresh apps"
                $tmpStat = $cmd.Invoke()
                foreach($s in $tmpStat.Sessions) {
                    if ($s.SessionID -eq $item.SessionID) {
                        $item = $s
                        Write-Host "item: "
                        Write-Host $item

                    }
                }
                $appsWindow.SetAppsDataGrid($item.Applications)
            })

            $appsWindow.Show()
        } catch [Exception] {
            Write-Host "Failed to get a list of applications" -ForegroundColor red
            Write-Host $_.Exception.Message -ForegroundColor red
        }
    }

    # Gets a list of RDSSessions.
    [ObservableCollection[RDPSession]] GetRDSSessions() {
        try {
            $cmd = {Get-RDSStatus -StatusLevel Level2}
            $tmpSessions = $cmd.Invoke()
            $this.UpdateCommandText($cmd)
            [ObservableCollection[RDPSession]]$oc = [ObservableCollection[RDPSession]]::new()
            foreach ($st in $tmpSessions) {
                foreach ($session in $st.Sessions) {
                    $oc.Add($session)
                }
            }
            return $oc
        } catch [Exception] {
            Write-Host "Failed to get a list of RDSSessions" -ForegroundColor red
            Write-Host $_.Exception.Message -ForegroundColor red
            return $Null
        }
    }

    [void] SetSessionsDataGrid([ObservableCollection[RDPSession]] $sessions) {
        if ($sessions -eq $null) {
            return
        }
        if ($sessions.Count -eq 0) {
            return
        }
        $this.Sessions = $sessions
        $this.dtSessions.ItemsSource = $this.Sessions
    }
}

