@{
	# Points CustomProvider.psm1 (and every Test-*.ps1 script in this folder) at the
	# provider script to test, plus the CustomSettings passed to it on provider/connect.
	# Edit CommandPath/CommandArgs/CustomSettings for your own environment before running
	# any test script here -- see this project's own README.md for a worked example.
	#
	# host/username/password below are placeholders. Point them at a real (ideally
	# non-production) HPE VME appliance before running anything beyond Test-Connect.ps1
	# against provider/initialize alone. access_token can be used instead of
	# username/password; ssl_check = $false skips TLS validation for a lab appliance.
	CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
	CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\hpe-vme\Parallels-RAS-CPF-HPE-VME-Advanced.ps1"'
	CustomSettings = @{ host = 'vme.example.com'; username = 'ras-service'; password = 'XXX' }
}
