<#  
.SYNOPSIS  
    PArallels RAS register script for Azure MarketPlace Deployments
.NOTES  
    File Name  : RAS_Azure_MP_Register.ps1
    Author     : Freek Berson
    Version    : v0.0.19
    Date       : Jun 27 2024
.EXAMPLE
    .\RAS_Azure_MP_Register.ps1
#>

function IsSupportedOS {
    if ([System.Environment]::Is64BitOperatingSystem) {
        try {
            $processorArchitecture = (Get-CimInstance -ClassName Win32_Processor).Architecture
            if ($processorArchitecture -eq 5) {
                Write-Host "ARM Based operating systems are supported by this script." -ForegroundColor red
                return $false
            }
        }
        catch {
            Write-Host "Failed to retrieve processor architecture: $_" -ForegroundColor red
            return $true
        }
    }
    return $true
}

function Test-InternetConnection {
    param (
        [int]$TimeoutMilliseconds = 5000
    )

    $request = [System.Net.WebRequest]::Create("http://www.google.com")
    $request.Timeout = $TimeoutMilliseconds

    try {
        $response = $request.GetResponse()
        $response.Close()
        return $true
    }
    catch {
        Write-Host "Internet connectivity is not available, check connectivity and try again." -ForegroundColor Red
        return $false
    }
}

function ConfigureNuGet {
    param()

    $requiredVersion = '2.8.5.201'
    $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue -Force

    if (!$nugetProvider) {
        Install-PackageProvider -Name NuGet -Force
    }
    else {
        $installedVersion = $nugetProvider.Version

        if ($installedVersion -lt $requiredVersion) {
            Write-Host "The installed NuGet provider version is $($installedVersion). Required version is $($requiredVersion) or higher."
            Install-PackageProvider -Name NuGet -Force
        }
    }
}

function import-AzureModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,

        [Parameter(Mandatory = $false)]
        [string]$moduleVersion
    )

    # Check if the module is already imported
    $module = Get-Module -Name $ModuleName -ListAvailable
    if (-not $module) {
        Write-Host "Required module '$ModuleName' is not imported. Installing and importing..."
        # Install the module if not already installed
        if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
            if ($PSBoundParameters.ContainsKey('moduleVersion')) {
                Install-Module -Name $ModuleName -RequiredVersion $moduleVersion -Scope CurrentUser -Force
            }
            else {
                Install-Module -Name $ModuleName -Scope CurrentUser -Force
            }
        }
        # Import the module
        Import-Module -Name $ModuleName -Force
    }
}

function get-AzureDetailsFromJSON {
    try {
        # Define the path to the JSON file
        $jsonFilePath = "C:\install\output.json"

        # Read the JSON content from the file
        $jsonContent = Get-Content -Path $jsonFilePath | Out-String

        # Convert JSON content to a PowerShell object
        $data = $jsonContent | ConvertFrom-Json

        return $data
    }
    catch {
        Write-Host "Error reading JSON file with Azure details." -ForegroundColor Red
        return $false
    }
}

function get-resourceUsageId {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId,
        [Parameter(Mandatory = $true)]
        [string]$appPublisherName,
        [Parameter(Mandatory = $true)]
        [string]$appProductName

    )
    Set-AzContext -SubscriptionId $SubscriptionId    
    $filteredResources = Get-AzResource -ResourceType 'Microsoft.Solutions/applications' | 
    Where-Object { 
            ($_.Plan.Publisher -match $appPublisherName) -and 
            ($_.Plan.Product -match $appProductName) -and 
            ($_.Kind -match 'MarketPlace') 
    }

    if ($filteredResources.Count -eq 0) {
        Write-Host "No matching resources found."
    }
    elseif ($filteredResources.Count -eq 1) {
        $managedAppName = $filteredResources[0].Name
        Write-Host "Only one matching resource found: $managedAppName"
    }
    else {
        Write-Host "Choose a resource by entering its corresponding number:"

        # Display numbered list of resource names
        for ($i = 0; $i -lt $filteredResources.Count; $i++) {
            Write-Host "$($i + 1). $($filteredResources[$i].Name)"
        }

        $validChoice = $false
        while (-not $validChoice) {
            $choice = Read-Host "Enter the number of the resource you want to select"
            $choice = [int]$choice

            if ($choice -gt 0 -and $choice -le $filteredResources.Count) {
                $managedAppName = $filteredResources[$choice - 1].Name
                Write-Host "You have selected: $managedAppName"
                $validChoice = $true
            }
            else {
                Write-Host "Invalid choice. Please enter a valid number."
            }
        }
    }
    $managedAppResourceGroupName = (get-azresource -ResourceType 'Microsoft.Solutions/applications' -Name $managedAppName).ResourceGroupName
    $resource = (Get-AzResource -ResourceType "Microsoft.Solutions/applications" -ResourceGroupName $managedAppResourceGroupName -Name $managedAppName)
    $resourceUsageId = $resource.Properties.billingDetails.resourceUsageId
    return $resourceUsageId
}

function get-keyVaultSecret {
    param (
        [Parameter(Mandatory = $true)]
        [string]$keyVaultName,
        [Parameter(Mandatory = $true)]
        [string]$secretName

    )
    return Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName
}


function CreateVMReaderRole {
    param(        
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId
    )
    #Create custom role definition
    $existingRoleDefinition = Get-AzRoleDefinition -Name "VM Reader Parallels RAS" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if ($null -eq $existingRoleDefinition) {
        $role = Get-AzRoleDefinition "Virtual Machine Contributor"
        $role.Id = $null
        $role.Name = "VM Reader Parallels RAS"
        $role.Description = "Provides read access to Microsoft.Compute"
        $role.Actions.Clear()
        $role.Actions.Add("Microsoft.Compute/*/read")
        $role.AssignableScopes.clear()
        $role.AssignableScopes.Add("/subscriptions/$SubscriptionId")
        New-AzRoleDefinition -Role $role | Out-Null
    }
}


function New-AzureAppRegistration {
    param(        
        [Parameter(Mandatory = $true)]
        [string]$appName
    )
 
    # Check if the AzADServicePrincipal already exists
    $ADServicePrincipal = Get-AzADServicePrincipal -DisplayName $appName
    if ($null -ne $ADServicePrincipal) {
        Write-Host "AD Service Principal with name '$appName' already exists. Please choose a different name."
        return
    }

    if (!($myApp = Get-AzADServicePrincipal -DisplayName $appName -ErrorAction SilentlyContinue)) {
        $myApp = New-AzADServicePrincipal -DisplayName $appName
    }
    return (Get-AzADServicePrincipal -DisplayName $appName)
}

function Set-AzureVNetResourcePermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$vnetId,

        [Parameter(Mandatory = $true)]
        [string]$ObjectId
    )
    
    #Add contributor permissions to the vnet for the app registration 
    $roleAssignment = New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName "Contributor" -Scope $vnetId | Out-Null
        
    # Return the selected vnet
    return $roleAssignment

}

function Add-AzureAppRegistrationPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$appName
    )
    # Get the app registration
    $applicationID = (Get-AzADApplication -DisplayName $appName).AppId

    #Add Group.Read.All permission
    Add-AzADAppPermission -ApplicationId $applicationID -ApiId "00000003-0000-0000-c000-000000000000" -PermissionId 5b567255-7703-4780-807c-7be8301ae99b -Type Role
    Add-AzADAppPermission -ApplicationId $applicationID -ApiId "00000003-0000-0000-c000-000000000000" -PermissionId df021288-bdef-4463-88db-98f22de89214 -Type Role
}
function New-AzureADAppClientSecret {
    param(     
        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$applicationID
    )
    #Remove-AzADAppCredential -ObjectId ((Get-AzADAppCredential -ApplicationId $applicationID).ObjectId| Where-Object CustomKeyIdentifier -EQ $null).KeyId | Out-Null
    $secretStartDate = Get-Date
    $secretEndDate = $secretStartDate.AddYears(1)
    $webApiSecret = New-AzADAppCredential -StartDate $secretStartDate -EndDate $secretEndDate -ApplicationId $applicationID -CustomKeyIdentifier ([System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("Parallels RAS secret")))
    return $webApiSecret    
}


function Set-azureResourceGroupPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$objectId,

        [Parameter(Mandatory = $true)]
        [string]$resourceGroupID
    )

    # Assign the contributor role to the service principal on the resource group
    New-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName "Contributor" -Scope $resourceGroupID | Out-Null
}

function Add-UserAccessAdministrationRole {
    param(
        [Parameter(Mandatory = $true)]
        [string]$appId,
        
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId
    )

    # Assign User Access Administrator role to the app registration at the subscription level
    New-AzRoleAssignment -ObjectId $appId -RoleDefinitionName "User Access Administrator" -Scope "/subscriptions/$SubscriptionId" | Out-Null
}

function Add-VMReaderRole {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Objectid,
        
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId
    )
    # Assign VM Reader role to the app registration at the subscription level
    New-AzRoleAssignment -Objectid $Objectid -RoleDefinitionName "VM Reader Parallels RAS" -Scope "/subscriptions/$SubscriptionId" | Out-Null
}

function Set-azureKeyVaultSecret {
    [CmdletBinding()]
    param (
               
        [Parameter(Mandatory = $true)]
        [string]$keyVaultName,
    
        [Parameter(Mandatory = $true)]
        [string]$SecretValue,
        
        [Parameter(Mandatory = $true)]
        [string]$SecretName
        
    )

    # Add the secret to the Key Vault
    $secret = ConvertTo-SecureString -String $SecretValue -AsPlainText -Force
    Set-AzKeyVaultSecret -VaultName $keyVaultName  -Name $SecretName -SecretValue $secret | Out-Null

    return $KeyVaultName
}

function set-AdminConsent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ApplicationId,

        [Parameter(Mandatory)]
        [string]$TenantId
    )

    $Context = Get-AzContext

    $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
        $context.Account, $context.Environment, $TenantId, $null, "Never", $null, "74658136-14ec-4630-ad9b-26e160ff0fc6")

    $headers = @{
        'Authorization'          = 'Bearer ' + $token.AccessToken
        'X-Requested-With'       = 'XMLHttpRequest'
        'x-ms-client-request-id' = [guid]::NewGuid()
        'x-ms-correlation-id'    = [guid]::NewGuid()
    }

    $url = "https://main.iam.ad.ext.azure.com/api/RegisteredApplications/$ApplicationId/Consent?onBehalfOfAll=true"
    Invoke-RestMethod -Uri $url -Headers $headers -Method POST -ErrorAction Stop
}

# BEGIN SCRIPT

# Disable IE ESC for Administrators and users
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}' -Name 'IsInstalled' -Value 0
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}' -Name 'IsInstalled' -Value 0

# Disable Edge first run experience
New-item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Force | Out-Null
New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -Name 'HideFirstRunExperience' -Value 1 -Force | Out-Null


if (-not (IsSupportedOS)) {
    Read-Host "Press any key to continue..."
    exit
}

if (-not (Test-InternetConnection)) {
    Read-Host "Press any key to continue..."
    exit
}

#Check if NuGet 2.8.5.201 or higher is installed, if not install it
try {
    Write-Host 'Installing required Azure Powershell modules.' `n
    ConfigureNuGet
}
Catch {
    Write-Host "ERROR: trying to install latest NuGet version"
    Write-Host $_Write-Host $_.Exception.Message
    exit
}

# Check and import the required Azure PowerShell module
try {
    import-AzureModule -ModuleName "Az.accounts" -moduleVersion "2.13.2"
    import-AzureModule -ModuleName "Az.Resources" -moduleVersion "6.12.0"
    import-AzureModule -ModuleName "Az.KeyVault" -moduleVersion "5.0.0"
}
Catch {
    Write-Host "ERROR: trying to import required modules import Az.Accounts, AzureAD, Az.Resources, Az.network, and Az.keyVault"
    Write-Host $_.Exception.Message
    exit
}

# Get Azure details from JSON file
try {
    $retreivedData = get-AzureDetailsFromJSON 
}
Catch {
    Write-Host "ERROR: retreiving Azure details from JSON file"
    Write-Host $_.Exception.Message
    exit
}

# Connect to Azure and Azure AD
try {
    Write-Host 'Please authenticate towards Azure to complete the setup.' `n
    $currentUser = Connect-AzAccount -Tenant $retreivedData.tenantID -AuthScope MicrosoftGraphEndpointResourceId
}
Catch {
    Write-Host "ERROR: trying to run Connect-AzAccount and Connect-AzureAD"
    Write-Host $_.Exception.Message
}

#Get the resourceUsageId
try {
    Write-Host 'Performing post-installation steps...' `n
    $appPublisherName = $retreivedData.appPublisherName
    $appProductName = $retreivedData.appProductName
    $resourceUsageId = get-resourceUsageId -SubscriptionId $retreivedData.SubscriptionId -appPublisherName $appPublisherName -appProductName $appProductName
}
Catch {
    Write-Host "ERROR: trying to read resource usage id from managed app"
    Write-Host $_.Exception.Message
    exit
}

#Get the keyvault secret
try {
    $localAdminPasswordSecure = (get-keyVaultSecret -keyVaultName $retreivedData.keyVaultName -secretName $retreivedData.secretName).secretValue
}
Catch {
    Write-Host "ERROR: trying to read resource usage id from managed app"
    Write-Host $_.Exception.Message
    exit
}


#Create Azure app registration if specified
if ($retreivedData.providerSelection -ne "noProvider") {

    Write-Host 'Performing post deployment configuration in Parallels RAS, please wait...' `n

    # Create a custom role to allow reading all compute resource
    try {
        CreateVMReaderRole -SubscriptionId $retreivedData.SubscriptionId
    }
    Catch {
        Write-Host "ERROR: creating custom role to allow reading VM resources"
        Write-Host $_.Exception.Message
        exit
    }

    # Create the app registration
    try {
        $app = New-AzureAppRegistration -appName $retreivedData.providerAppRegistrationName
    }
    Catch {
        Write-Host "ERROR: trying to create the App Registration"
        Write-Host $_.Exception.Message
        exit
    }

    # Set permissions on the virtual network
    try {
        Set-azureVNetResourcePermissions -vnetId $retreivedData.vnetId -objectId $app.Id
    }
    Catch {
        Write-Host "ERROR: trying to configure contributor permissons on vnet"
        Write-Host $_.Exception.Message
        exit
    }

    # Set the required Graph API permissions on the created app registration
    try {
        Add-AzureAppRegistrationPermissions -appName $app.DisplayName
    }
    Catch {
        Write-Host "ERROR: trying to set app registration Graph API permissions"
        Write-Host $_.Exception.Message
        exit
    }

    # Create a client secret on the app registration and capture the secret key
    try {
        $secret = New-AzureADAppClientSecret -TenantId $retreivedData.tenantID -applicationID $app.AppId
    }
    Catch {
        Write-Host "ERROR: trying to create the App Registration client secret"
        Write-Host $_.Exception.Message
        exit
    }

    # Add app registration contributor permissions on resource group
    try {
        $rg = set-azureResourceGroupPermissions -resourceGroupID $retreivedData.mgrID -objectId $app.Id
    }
    Catch {
        Write-Host "ERROR: trying to create the resource group and set contributor permissions"
        Write-Host $_.Exception.Message
        exit
    }

    # Add User Access Administratrion permission on subscription to the app registration
    try {
        Add-UserAccessAdministrationRole -appId $app.Id -SubscriptionId $retreivedData.SubscriptionId
    }
    Catch {
        Write-Host "ERROR: trying to set User Access Administration role"
        Write-Host $_.Exception.Message
        exit
    }

    # Add VM Reader permission on subscription to the app registration
    try {
        Add-VMReaderRole -Objectid $app.Id -SubscriptionId $retreivedData.SubscriptionId
    }
    Catch {
        Write-Host "ERROR: trying to set VM Reader role"
        Write-Host $_.Exception.Message
        exit
    }

    # Store client secret in Azure KeyVault
    try {
        $selectedKeyVaultName = Set-azureKeyVaultSecret -keyVaultName $retreivedData.keyVaultName -SecretValue $secret.SecretText -SecretName $retreivedData.providerAppRegistrationName
    }
    Catch {
        Write-Host "ERROR: trying to create a new Azure KeyVault and adding the client secret"
        Write-Host $_.Exception.Message
        exit
    }

    #Wait to make sure the API permisissons are configgured before doing a consent
    sleep 60

    # Grant admin consent to an the app registration
    try {
        set-AdminConsent -ApplicationId $app.AppId -TenantId $retreivedData.tenantID
    }
    Catch {
        Write-Host "ERROR: trying to grant admin consent to an the app registration"
        Write-Host $_.Exception.Message
        exit
    }

}

# Check if the license type is 0 (AZMP) then register Parallels RAS with the license AZMP key
# Check if the license type is 1 (BYOL) then allow to register license key manually or use trial
if ($retreivedData.licenseType -eq 0) {
    # Register Parallels RAS with the license key
    New-RASSession -Username $retreivedData.domainJoinUserName -Password $localAdminPasswordSecure -Server $retreivedData.primaryConnectionBroker

    #Set Azure Marketplace related settings in RAS db
    Set-RASAzureMarketplaceDeploymentSettings -SubscriptionID $retreivedData.SubscriptionId -TenantID $retreivedData.tenantID -CustomerUsageAttributionID $retreivedData.customerUsageAttributionID -ManagedAppResourceUsageID $resourceUsageId[1]

    # Invoke-apply
    invoke-RASApply
}

#Create Azure or AVD in RAS if specified
if ($retreivedData.providerSelection -eq "AVDProvider") {
    $appPassword = ConvertTo-SecureString -String $secret.SecretText -AsPlainText -Force
    Set-RASAVDSettings -Enabled $true
    invoke-RASApply
    New-RASProvider $retreivedData.providerName -AVDVersion AVD -AVD -TenantID $retreivedData.tenantID -SubscriptionID $retreivedData.SubscriptionId -ProviderUsername $app.AppId -ProviderPassword $appPassword -NoInstall | Out-Null
    invoke-RASApply
}
if ($retreivedData.providerSelection -eq "AzureProvider") {
    $appPassword = ConvertTo-SecureString -String $secret.SecretText -AsPlainText -Force
    New-RASProvider $retreivedData.providerName -AzureVersion Azure -Azure -TenantID $retreivedData.tenantID -SubscriptionID $retreivedData.SubscriptionId -ProviderUsername $app.AppId -ProviderPassword $appPassword -NoInstall | Out-Null
    invoke-RASApply
}


# Invoke-apply and remove session
invoke-RASApply
Remove-RASSession

#restart secundary RAS servers to complete installation
for ($i = 2; $i -le $retreivedData.numberofCBs; $i++) {
    $connectionBroker = $retreivedData.prefixCBName + "-" + $i + "." + $retreivedData.domainName
    restart-computer -computername $connectionBroker -WsmanAuthentication Kerberos -force
}
for ($i = 1; $i -le $retreivedData.numberofSGs; $i++) {
    $secureGateway = $retreivedData.prefixSGName + "-" + $i + "." + $retreivedData.domainName
    restart-computer -computername $secureGateway -WsmanAuthentication Kerberos -force
}

#Clean up JSON file
remove-item "C:\install\output.json" -Force

Write-Host 'Registration of Parallels RAS is completed.' `n
Read-Host -Prompt "Press any key to open the Parallels RAS console..." | Out-Null

Start-Process -FilePath "C:\Program Files (x86)\Parallels\ApplicationServer\2XConsole.exe"
