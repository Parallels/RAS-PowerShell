@{
	CommandPath = 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe'
	CommandArgs = '-File "C:\Scripts\OLVM-CustomProvider.ps1"'
	CustomSettings = @{
		engine_url     = 'https://olvm.example.com/ovirt-engine/api'
		username       = 'admin@internal'
		password       = 'secret'
		insecure       = $true
		cluster        = 'Default'
		storage_domain = 'data'
	}
}
