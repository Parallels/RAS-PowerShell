#requires -version 5
<#
.SYNOPSIS
    This is the initialization script for RDSSessionMgt class
.DESCRIPTION
  In stance of RDSSessionMgt is created here and button click events are
  initialized.
.INPUTS
    None
.OUTPUTS
    PSWindow UI
.NOTES
  Version:        1.0
  Purpose:        Showcase powershell capabilities with PSAdmin
  
.EXAMPLE
  ./client.ps1
#>
using assembly presentationFramework
using module 'RASAdmin'
using module './RDSSessionMgt.psm1'


[string] $xaml = "./resources/MainWindow.xaml"

try {
    $wnd = [RDSSessionMgt]::new($xaml)
}
catch [Exception]{
        Write-Host "Could not create RDSSessionMgt" -ForegroundColor red
        return
}

$wnd.btnLogin.Add_Click({
    try {
        $wnd.OnBtnLoginClick()
    }
    catch {
        Write-Warning $_.Exception.Message
    }
})

$wnd.btnGetApps.Add_Click({
    try {
        $wnd.OnBtnGetAppsClick()
    }
    catch {
        Write-Warning $_.Exception.Message
    }
})

$wnd.btnLogoff.Add_Click({
    try {
        $wnd.OnBtnLogoffSessionClick();
    }
    catch {
        Write-Warning $_.Exception.Message
    }
})

$wnd.btnDisconnect.Add_Click({
    try {
        $wnd.OnBtnDisconnectSessionClick();
    }
    catch {
        Write-Warning $_.Exception.Message
    }
})

$wnd.btnRefresh.Add_Click({
    try {
        $wnd.OnBtnRefreshClick();
    }
    catch {
        Write-Warning $_.Exception.Message
    }
})

$wnd.btnSendMessage.Add_Click({
    try {
        $wnd.OnBtnSendMessageClick();
    }
    catch {
        Write-Warning $_.Exception.Message        
    }
})

$wnd.Form.ShowDialog()
