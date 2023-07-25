<# 
.SYNOPSIS 
    Parallels RAS script to configure prerequisites for the AVD Provider.
.DESCRIPTION 
    The script will ask for Tenant, Subscription, App name, location, and resource group name, virtual network, 
    and Keyvalt and configures all prerequisites required for the AVD Provider in Parallels RAS. It outputs the
    information needed to complete the AVD Provider wizard in Parallels RAS.
.PARAMETER   
    None, all parameters are collected at runtime.
.OUTPUTS 
    - Azure AD Tenant ID
    - Azure Subscription ID
    - App registration ID
    - App registration Secret
.NOTES 
    Version: 1.0
    Author: Freek Berson  
    Last update: 24/07/23
    Changelog:  1.0 -   Initial published version
.LICENSE
    Released under the terms of MIT license (see LICENSE for details)
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
        [string]$ModuleName
    )

    # Check if the module is already imported
    $module = Get-Module -Name $ModuleName -ListAvailable
    if (-not $module) {
        Write-Host "Required module '$ModuleName' is not imported. Installing and importing..."
        # Install the module if not already installed
        if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force
        }
        # Import the module
        Import-Module -Name $ModuleName -Force
    }
}

function set-AzureTenant {
    # Retrieve Azure tenants
    $tenants = Get-AzTenant

    # Display the list of tenants and prompt the user to select one
    $i = 1
    $selectedTenant = $null

    Write-Host "Azure Tenants:" -ForegroundColor Yellow
    foreach ($tenant in $tenants) {
        Write-Host "$i. $($tenant.Name) - $($tenant.TenantId)"
        $i++
    }

    $validSelection = $false
    while (-not $validSelection) {
        $selection = Read-Host ('>> Select a tenant by entering the corresponding number')
        
        if ($selection -match '^\d+$') {
            $selection = [int]$selection
            if ($selection -ge 1 -and $selection -le $tenants.Count) {
                $validSelection = $true
            }
        }
        
        if (-not $validSelection) {
            Write-Host "Invalid input. Please enter a valid number between 1 and $($tenants.Count)" -ForegroundColor Red
        }
    }

    $selectedTenant = $tenants[$selection - 1]

    # Store the selected tenant ID in tenantId variable
    $tenantId = $selectedTenant.TenantId

    Write-Host "Selected Tenant ID: $tenantId`n" -ForegroundColor Green

    # Return the selected tenant ID
    return $tenantId
}
function Set-AzureSubscription {
    # Check if the user is authenticated
    if (-not (Get-AzContext)) {
        Write-Host "Failed to authenticate against Azure. Please check your credentials and try again."
        return
    }

    # Get the list of Azure subscriptions
    $subscriptions = Get-AzSubscription

    # Check if the user has access to any subscriptions
    if ($subscriptions) {
        # Display the list of subscriptions and prompt the user to select one
        $i = 1
        $selectedSubscription = $null

        Write-Host "Azure Subscriptions:" -ForegroundColor Yellow
        foreach ($subscription in $subscriptions) {
            Write-Host "$i. $($subscription.Name) - $($subscription.Id)"
            $i++
        }

        $validSelection = $false
        while (-not $validSelection) {
            $selection = Read-Host ('>> Select a subscription by entering the corresponding number')

            if ($selection -match '^\d+$') {
                $selection = [int]$selection
                if ($selection -ge 1 -and $selection -le $subscriptions.Count) {
                    $validSelection = $true
                }
            }

            if (-not $validSelection) {
                Write-Host "Invalid input. Please enter a valid number between 1 and $($subscriptions.Count)" -ForegroundColor Red
            }
        }

        $selectedSubscription = $subscriptions[$selection - 1]

        # Store the selected subscription ID in subscriptionId variable
        $subscriptionId = $selectedSubscription.Id

        Write-Host "Selected Subscription ID: $subscriptionId`n" -ForegroundColor Green

        Set-AzContext -SubscriptionId $subscriptionId

        # Return the selected subscription object
        return $selectedSubscription
    }
    else {
        Write-Host "You do not have access to any Azure subscriptions."
    }
}

function set-AzureLocation {
    # Retrieve Azure locations
    $locations = Get-AzLocation | Where-Object { $_.Providers -contains "Microsoft.DesktopVirtualization" } | Select-Object -ExpandProperty Location | Sort-Object

    # Display the list of locations and prompt the user to select one
    $selectedLocation = $null

    # Determine the number of columns for display
    $columnCount = 3
    $rowCount = [Math]::Ceiling($locations.Count / $columnCount)

    # Display the list of locations in multiple columns
    Write-Host "Azure Locations:" -ForegroundColor Yellow

    for ($row = 0; $row -lt $rowCount; $row++) {
        for ($col = 0; $col -lt $columnCount; $col++) {
            $index = $row + ($col * $rowCount)

            if ($index -lt $locations.Count) {
                $location = $locations[$index]
                $label = ($index + 1).ToString().PadRight(3)

                Write-Host "$label. $location" -NoNewline

                $padding = 20 - $location.Length
                Write-Host (" " * $padding) -NoNewline
            }
        }

        # Stop if the total count is reached
        if ($index -eq $locations.Count - 1) {
            break
        }

        Write-Host
    }
    Write-Host `n

    $validSelection = $false
    while (-not $validSelection) {
        $selection = Read-Host ('>> Select the location of where you want to deploy the AVD resources')
        if ($selection -match '^\d+$') {
            $selection = [int]$selection
            if ($selection -ge 1 -and $selection -le $locations.Count) {
                $validSelection = $true
                $selectedLocation = $locations[$selection - 1]
                Write-Host "Selected Location: $selectedLocation" -ForegroundColor Green
                return $selectedLocation
            }
        }
        if (-not $validSelection) {
            Write-Host "Invalid input. Please enter location number between 1 and $($locations.Count)" -ForegroundColor Red
        }
    }
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
    $validAppName = $false
    $invalidChars = @('<', '>', ';', '&', '%')

    Write-Host `n"App registrations:" -ForegroundColor Yellow
    
    while (-not $validAppName) {
        $appName = Read-Host '>> Provide the App Registration name'
    
        if (-not [string]::IsNullOrWhiteSpace($appName) -and $appName.Length -gt 0) {
            if ($appName.Length -le 120) {
                $containsInvalidChars = $false
                foreach ($invalidChar in $invalidChars) {
                    if ($appName.Contains($invalidChar)) {
                        $containsInvalidChars = $true
                        break
                    }
                }
    
                if (-not $containsInvalidChars) {
                    # Check if the app name already exists
                    $existingAppName = Get-AzureADApplication | Where-Object { $_.DisplayName -eq $appName }
                    if ($existingAppName) {
                        Write-Host "The provided App Registration name already exists. Please provide a different name." -ForegroundColor Red
                    }
                    else {
                        $validAppName = $true
                    }
                }
                else {
                    Write-Host "The provided App Registration name contains invalid characters. Please avoid using <, >, ;, &, or %." -ForegroundColor Red
                }
            }
            else {
                Write-Host "The provided App Registration name exceeds the maximum allowed length of 120 characters." -ForegroundColor Red
            }
        }
        else {
            Write-Host "The App Registration name cannot be empty or have a length of 0." -ForegroundColor Red
        }
    }    
    
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
        [string]$vnetLocation,

        [Parameter(Mandatory = $true)]
        [string]$ObjectId
    )
    
    $i = 1

    # Get all virtual networks in the specified location
    $vnets = Get-AzVirtualNetwork | Where-Object { $_.Location -eq $vnetLocation }
    
    # Check if any vnets are found in the specified location
    if ($vnets.Count -eq 0) {
        Write-Host `n"Virtual Networks:" -ForegroundColor Yellow
        Write-Host "No virtual networks found in the specified location: $vnetLocation"`n
        return $null
    }
    
    # Display the list of vnets with their index numbers starting from 1
    Write-Host `n"Virtual Networks in the Azure Subscription located in $vnetLocation :" -ForegroundColor Yellow
    foreach ($vnet in $vnets) {
        Write-Host "$i. $($vnet.Name)"
        $i++
    }
    
    $validSelection = $false
    while (-not $validSelection) {
        $selection = Read-Host ('>> Enter the index number of the vnet you want to use (Hit Enter to Skip)')
        if ($selection -match '^\d+$') {
            $selection = [int]$selection
            if (($selection -ge 1 -and $selection -le $vnets.Count) -or $selection -eq '') {
                $validSelection = $true
            }            
        }

        if ($selection -eq '') {
            Write-Host "Vnet selection skipped"`n -ForegroundColor Green
            return $null
        }

        if (-not $validSelection) {
            Write-Host "Invalid input. Please enter a valid number between 1 and $($subscriptions.Count)" -ForegroundColor Red
        }
    }
    #Add contributor permissions to the vnet for the app registration 
    $selectedVNet = $vnets[$selection - 1]
    $roleAssignment = New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName "Contributor" -Scope $selectedVNet.id | Out-Null
        
    Write-Host "Selected Vnet: "$selectedVNet.Name`n -ForegroundColor Green

    # Return the selected vnet
    return $roleAssignment

}

function Add-AzureAppRegistrationPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$appName
    )
    # Get the app registration
    $applicationID = (Get-AzureADApplication -Filter "displayName eq '$appName'").AppId

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
    Remove-AzureADApplicationPasswordCredential -ObjectId (Get-AzureADApplication -Filter "AppId eq '$applicationID'").ObjectId -KeyId (Get-AzureADApplicationPasswordCredential -ObjectId (Get-AzureADApplication -Filter "AppId eq '$applicationID'").ObjectId | Where-Object CustomKeyIdentifier -EQ $null).KeyId
    $secretStartDate = Get-Date
    $secretEndDate = $secretStartDate.AddYears(1)
    $webApiSecret = New-AzADAppCredential -StartDate $secretStartDate -EndDate $secretEndDate -ApplicationId $applicationID -CustomKeyIdentifier ([System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("Parallels RAS secret")))
    return $webApiSecret    
}

function New-AzureResourceGroupWithPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$objectId,

        [Parameter(Mandatory = $true)]
        [string]$resourceGroupLocation
    )

    $validResourceGroupName = $false

    while (-not $validResourceGroupName) {
        $resourceGroupName = Read-Host '>> Provide the name of the resource group you want to create'
    
        if (-not [string]::IsNullOrWhiteSpace($resourceGroupName)) {
            if ($resourceGroupName -match '^[A-Za-z0-9_-]+$' -and $resourceGroupName.Length -le 90) {
                # Check if the resource group already exists
                $existingResourceGroup = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
    
                if ($existingResourceGroup) {
                    $validInput = $false
                    while (!$validInput) {
                        $confirm = Read-Host "The resource group '$resourceGroupName' already exists. Do you want to use it? (Y/N)"
                        if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                            $validInput = $true
                            $validResourceGroupName = $true
                        }
                        elseif ($confirm -eq 'N' -or $confirm -eq 'n') {
                            $validInput = $true
                            $validResourceGroupName = $false
                        }
                        else {
                            Write-Host "Invalid input. Please enter 'Y' or 'N'." -ForegroundColor red
                        }
                    }
                }
                else {
                    $validResourceGroupName = $true
                }
            }
            else {
                Write-Host "The provided resource group name contains invalid characters, or is too long. Please use only alphanumeric characters, hyphens, underscores, and up to 90 characters." -ForegroundColor Red
            }
        }
        else {
            Write-Host "The resource group name cannot be empty." -ForegroundColor Red
        }
    }
    
    # Create the resource group
    $resourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $resourceGroupLocation -Force


    # Assign the contributor role to the service principal on the resource group
    New-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName "Contributor" -Scope $resourceGroup.ResourceId | Out-Null

    return $resourceGroup
}

function Add-UserAccessAdministrationRole {
    param(
        [Parameter(Mandatory = $true)]
        [string]$objectId,
        
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId
    )

    # Assign User Access Administrator role to the app registration at the subscription level
    New-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName "User Access Administrator" -Scope "/subscriptions/$SubscriptionId" | Out-Null
}

function Add-VMReaderRole {
    param(
        [Parameter(Mandatory = $true)]
        [string]$objectId,
        
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId
    )
    # Assign VM Reader role to the app registration at the subscription level
    New-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName "VM Reader Parallels RAS" -Scope "/subscriptions/$SubscriptionId" | Out-Null
}

function New-AzureKeyVaultWithSecret {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ResourceGroupName,
        
        [Parameter(Mandatory = $true)]
        [string]$Location,
        
        [Parameter(Mandatory = $true)]
        [string]$SecretValue,
        
        [Parameter(Mandatory = $true)]
        [string]$SecretName,
        
        [Parameter(Mandatory = $true)]
        [object]$CurrentUser
    )

    Write-Host `n"Azure Keyvault:" -ForegroundColor Yellow

    # Prompt the user to enter the Key Vault name and validate it
    $validSelection = $false
    while (-not $validSelection) {
        $KeyVaultName = Read-Host ">> Enter the name for the new Azure Key Vault to store secrets (Hit Enter to Skip)"
        if ($KeyVaultName -match '^[A-Za-z][\w-]{1,22}[A-Za-z0-9]$' -or $KeyVaultName -eq '') {
            $validSelection = $true
        }

        if ($KeyVaultName -eq '') {
            Write-Host "Keyvault configuration skipped"`n -ForegroundColor Green
            return $null
        }

        if (-not $validSelection) {
            Write-Host "Invalid Key Vault name. Key Vault names must be between 3 and 24 characters in length. They must begin with a letter, end with a letter or digit, and contain only alphanumeric characters and dashes. Consecutive dashes are not allowed." -ForegroundColor Red
        }
    }


    # Check if the Key Vault already exists
    $existingKeyVault = Get-AzKeyVault -ResourceGroupName $ResourceGroupName -VaultName $KeyVaultName -ErrorAction SilentlyContinue

    if ($existingKeyVault) {
        # Key Vault already exists
        $useExisting = Read-Host "A Key Vault with the name '$KeyVaultName' already exists. Do you want to use the existing Key Vault? (Y/N)"
        if ($useExisting -eq 'Y') {
            Write-Output "Using the existing Key Vault '$KeyVaultName'."
            $keyVault = $existingKeyVault
        }
        else {
            Write-Output "Aborting operation."
            return
        }
    }
    else {
        # Create a new Key Vault
        $keyVault = New-AzKeyVault -ResourceGroupName $ResourceGroupName -VaultName $KeyVaultName -Location $Location
    }

    # Add the user as a Key Vault administrator
    $objectId = (Get-AzADUser -UserPrincipalName $CurrentUser.Context.Account.Id).Id
    Set-AzKeyVaultAccessPolicy -VaultName $KeyVault.VaultName -ObjectId $objectId -PermissionsToSecrets @('Get', 'Set', 'List', 'Delete')

    # Add the secret to the Key Vault
    $secret = ConvertTo-SecureString -String $SecretValue -AsPlainText -Force
    Set-AzKeyVaultSecret -VaultName $KeyVault.VaultName  -Name $SecretName -SecretValue $secret | Out-Null
    Write-Host "Added a new secret with the name $($SecretName) to the Key Vault $($KeyVaultName.VaultName)." -ForegroundColor Green

    return $KeyVaultName
}

Clear-Host

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
    ConfigureNuGet
}
Catch {
    Write-Host "ERROR: trying to install latest NuGet version"
    Write-Host $_Write-Host $_.Exception.Message
    exit
}

# Check and import the required Azure PowerShell module
try {
    import-AzureModule "Az.Accounts"
    import-AzureModule "AzureAD"
    import-AzureModule "Az.Resources"
    import-AzureModule "Az.network"
    import-AzureModule "Az.keyVault"
}
Catch {
    Write-Host "ERROR: trying to import required modules import Az.Accounts, AzureAD, Az.Resources, Az.network, and Az.keyVault"
    Write-Host $_.Exception.Message
    exit
}

# Connect to Azure and Azure AD
try {
    $currentUser = Connect-AzAccount
    Connect-AzureAD | Out-Null
}
Catch {
    Write-Host "ERROR: trying to run Connect-AzAccount and Connect-AzureAD"
    Write-Host $_.Exception.Message
    exit
}

# Connect to Azure and Azure AD
try {
    $selectedTenantId = set-AzureTenant
}
Catch {
    Write-Host "ERROR: trying to get Azure Tenants"
    Write-Host $_.Exception.Message
    exit
}

# Provide list of available Azure subscriptions and allow setting active subscription
try {
    $selectedsubscriptionID = (set-AzureSubscription).Id
}
Catch {
    Write-Host "ERROR: trying to set Azure subscription"
    Write-Host $_.Exception.Message
    exit
}

# Provide list of avalable Azure locations and allow setting active location
try {
    $selectedAzureLocation = set-AzureLocation
}
Catch {
    Write-Host "ERROR: trying to get Azure Location"
    Write-Host $_.Exception.Message
    exit
}

# Create a custom role to allow reading all compute resource
try {
    CreateVMReaderRole -SubscriptionId $selectedsubscriptionID
}
Catch {
    Write-Host "ERROR: creating custom role to allow reding VM resources"
    Write-Host $_.Exception.Message
    exit
}

# Prompt for the app name and create the app registration
try {
    $app = New-AzureAppRegistration
    Write-Host "App registration name: "$app.DisplayName -ForegroundColor Green
}
Catch {
    Write-Host "ERROR: trying to create the App Registration"
    Write-Host $_.Exception.Message
    exit
}

# Prompt for the vnet to be used to set permissions for the app registration
try {
    Set-AzureVNetResourcePermissions -vnetLocation $selectedAzureLocation -objectId $app.Id
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
    $secret = New-AzureADAppClientSecret -TenantId $selectedTenantId -applicationID $app.AppId
}
Catch {
    Write-Host "ERROR: trying to create the App Registration client secret"
    Write-Host $_.Exception.Message
    exit
}

# Prompt for the resource group name, create the Resource Group and add the app registration contributor permissions
try {
    Write-Host "Azure Resource Group:" -ForegroundColor Yellow
    $rg = new-AzureResourceGroupWithPermissions -resourceGroupLocation $selectedAzureLocation -objectId $app.Id
    Write-Host "Resource Group name: "$rg.ResourceGroupName -ForegroundColor Green
    Write-Host "* NOTE: If AVD Session Hosts and/or templates VMs are going to be stored in different Resource Group(s) later, add Contributor permissions to those Resource Group for the following app registration:" $app.DisplayName "No additional permissions are required if RAS template or AVD Session Hosts are part of the newly created Resource Group" $rg.ResourceGroupName -ForegroundColor Cyan
}
Catch {
    Write-Host "ERROR: trying to create the resource group and set contributor permissions"
    Write-Host $_.Exception.Message
    exit
}

# Add User Access Administratrion permission on subscription to the app registration
try {
    Add-UserAccessAdministrationRole -objectId $app.Id -SubscriptionId $selectedsubscriptionID
}
Catch {
    Write-Host "ERROR: trying to set User Access Administration role"
    WWrite-Host $_.Exception.Message
    exit
}

# Add VM Reader permission on subscription to the app registration
try {
    Add-VMReaderRole -objectId $app.Id -SubscriptionId $selectedsubscriptionID
}
Catch {
    Write-Host "ERROR: trying to set VM Reader role"
    WWrite-Host $_.Exception.Message
    exit
}

# Add an Azure Keyvault and store the Client Secret in it
try {
    $selectedKeyVaultName = New-AzureKeyVaultWithSecret -ResourceGroupName $rg.ResourceGroupName -Location $selectedAzureLocation -SecretValue $secret.SecretText -SecretName "Parallels-RAS-AVD-Provider" -CurrentUser $currentUser
}
Catch {
    Write-Host "ERROR: trying to create a new Azure KeyVault and adding the client secret"
    Write-Host $_.Exception.Message
    exit
}

#Create summary information
Write-Host "`n* App registration created, permissions configured and secret created." -ForegroundColor Cyan
Write-host "* Below is the information to create Parallels RAS AVD Provider, COPY THE INFORMATION BELOW and store in a safe place before continuing!" -ForegroundColor Cyan
if ($selectedKeyVaultName) {
    Write-Host "`* The App registration ID is also securely stored in the keyvault"$selectedKeyVaultName -ForegroundColor Cyan
}
Write-Host "- 1. Azure AD Tenant ID:`t "$selectedTenantId
Write-Host "- 2. Azure Subscription ID:`t "$selectedsubscriptionID
Write-Host "- 3. App registration ID:`t "$app.AppId
Write-Host "- 4. App registration Secret:`t "$secret.SecretText

#Open Web browser towards App Regisration for perform final consent
Write-Host "A browser will now be opened to Azure to provide permission consent, make sure you log on with a Global Admin Account." -ForegroundColor Cyan
Read-Host "Press any key to continue..."
Start-Process "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)/isMSAApp~/false"
