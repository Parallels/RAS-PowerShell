'###########################################
'# Author - Paul Fisher - Parallels        #
'# Date   - 30/05/2023                     #
'# Description: Disconnect sessions & Lock #
'###########################################
 
On Error Resume Next
 
'AppServerClient Path
RASClientPath = "C:\Program Files\Parallels\Client\APPServerClient.exe"
 
Set RASClientShell = CreateObject("Wscript.Shell")
 
'Build Argument String
Args = " -disconnectallsessions"

'Call AppServerClient to disconnect sessions and lock workstation
RASClientShell.Run "cmd /c " & chr(34) & RASClientPath & chr(34) & Args, 0, False
wscript.quit
 