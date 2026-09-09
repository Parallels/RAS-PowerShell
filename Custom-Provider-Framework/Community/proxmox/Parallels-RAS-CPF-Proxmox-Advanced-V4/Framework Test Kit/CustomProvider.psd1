@{
	# Points CustomProvider.psm1 (and every Test-*.ps1 script in this folder) at the
	# provider script to test, plus the CustomSettings passed to it on provider/connect.
	# Edit CommandPath/CommandArgs/CustomSettings for your own environment before running
	# any test script here -- see this project's own README.md for a worked example.
	#
	# host/username/token_name/token_secret below are placeholders. Point them at a real
	# (ideally non-production) Proxmox VE cluster and API token before running anything
	# beyond Test-Connect.ps1 against provider/initialize alone.
	CommandPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
	CommandArgs = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\CPF_Scripts\Parallels-RAS-CPF-Proxmox-Advanced.ps1"'
	CustomSettings = @{ host = 'proxmox.example.com'; username = 'root@pam'; token_name = 'automation'; token_secret = 'XXX' }
}
