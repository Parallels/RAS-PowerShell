PowerShell code that reads clixml files produced by a policy module found on Citrix Delivery Controllers and produces a script that will create them in Parallels RAS

Import-Module -Name "$env:ProgramFiles\Citrix\Telemetry Service\TelemetryModule\Citrix.GroupPolicy.Commands.psm1"
New-PSDrive -PSProvider CitrixGroupPolicy -Name LocalFarmGpo -Root \ -Controller $env:COMPUTERNAME
Export-CtxGroupPolicy -FolderPath "H:\Citrix Policies Export"

To process these exported policies (it should create 3 xml files in the folder specified), run the script on a machine which has the RASAdmin PowerShell module but it does not need to be run with a user that has RAS permissions

& 'C:\Parallels\Citrix Policy conversion to RAS.ps1' -mappingFile "C:\Parallels\RAS Policy Mapping.csv" -citrixPolicyFolder "H:\Citrix Policies Export"

Which will create the file CreateRASPoliciesFromCitrix.ps1 which can be run by a user with RAS permission to create the corresponding RAS policies where they exist.

The script will error if CreateRASPoliciesFromCitrix.ps1 already exists and -overwrite is not specified. The output file name can be changed with -outputFile
