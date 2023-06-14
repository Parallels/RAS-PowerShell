'#######################################################################
'# Author - Paul Fisher - Parallels                                    #
'# Date   - 28/02/2023                                                 #
'# Description: Logon script for Imprivata with Parallels RAS.         #
'# Description: Replace ApplicationID with desired Parallels RAS appID #
'#######################################################################
 
On Error Resume Next

  'Set connection properties.
  Server = "servername.here.com"
  Port = "443"
 
  'Get SSO credentials - from OneSign agent variables.
  'Set authentication properties.
  Domain = "domain.local"
  Username = "Username"
  Password = "Password"
  ApplicationID = "#2" 
  RASClientPath = "C:\Program Files\Parallels\Client\TSClient.exe"
 
  
'Launch tsclient and wait for session to end and logoff.
Set RASClientShell = CreateObject("Wscript.Shell")

'Build Argument String
Args = " s!='" & Server & "' " & "a!='" & ApplicationID & "' " & "t!='" & Port & "' " & "d!='" & Domain & "' " & "u!='" & Username & "' " & "q!='" & Password & "'  m!='2'"

'Connect to Published Desktop
 
RASClientShell.Run "cmd /c " & chr(34) & RASClientPath & chr(34) & Args, 0, False
wscript.quit