# Required Microsoft Windows Server modules.

Write-Host "Install RSAT-AD-PowerShell and RSAT-RDS-Tools if necessary."
Install-WindowsFeature RSAT-AD-PowerShell
Install-WindowsFeature RSAT-RDS-Tools

# This should be run on a Remote Application Server with a single farm.
Import-Module RASAdmin

$PSW = "Parallels Secure Workspace"
$RAS = "Parallels Remote Application Server"

# Start a new session.
try {

    $Upn = Read-Host -Prompt "Enter the Microsoft Windows username for the admin account (format: UPN, e.g. jsmith@example.org )"
    $Password = Read-Host -Prompt "Enter the password for the admin account" -MaskInput
    $Server = Read-Host -Prompt "Enter the FQDN or IP address of the $RAS server"

    New-RASSession -Username $Upn -Password ($Password | ConvertTo-SecureString -AsPlainText -Force) -Server $Server -ErrorAction Stop

}
catch {

    Write-Host -ForegroundColor Red "Unable to connect to the $RAS server. Check credentials and IP address of the $RAS server."
    exit

}

function Get-RDSCollectionBySessionHost {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SessionHost,

        [Parameter(Mandatory = $false)]
        [string]$ConnectionBroker = $env:COMPUTERNAME
    )


        try {
            # Get all session collections for this broker
            $collections = Get-RDSessionCollection -ConnectionBroker $Broker

            foreach ($collection in $collections) {

                # Get session hosts for this collection
                $sessionHosts = Get-RDSessionHost `
                    -CollectionName $collection.CollectionName `
                    -ConnectionBroker $Broker

                foreach ($sessionHostObj in $sessionHosts) {

                    # Match FQDN or short name (case-insensitive)
                    if (
                        $sessionHostObj.SessionHost -ieq $SessionHost -or
                        $sessionHostObj.SessionHost.Split('.')[0] -ieq $SessionHost
                    ) {
                        [PSCustomObject]@{
                            ConnectionBroker = $Broker
                            CollectionName = $collection.CollectionName
                            SessionHost = $sessionHostObj.SessionHost
                            CollectionType = $collection.CollectionType
                        }
                    }
                }
            }
        } 
        catch {
            Write-Warning "Failed to query broker $Broker : $_"
        }
    
}




# Select and confirm the Workspace domains to import.
$Domains = (Get-Content -Path "data/domains.json" | ConvertFrom-Json).results
$TwoFactorProviders = (Get-Content -Path "data/twofactor-providers.json" | ConvertFrom-Json).results
$AppServers = (Get-Content -Path "data/app-servers.json" | ConvertFrom-Json).results
$Apps = (Get-Content -Path "data/apps.json" | ConvertFrom-Json).results
$SSLCerts = (Get-Content -Path "data/ssl-offloader-certificates.json" | ConvertFrom-Json).results

# Currently, only domains pointing to the same Microsoft Windows back end infrastructure are supported in this script.
$Selection = $Domains | Select-Object name, fqdn, netbios | Out-GridView -Title "Select $PSW domains" -PassThru

# To keep the migration script simple: Create a theme for each Workspace domain. 
# This can be used to map the login permissions.
$Domains = $Domains | Where-Object { $Selection.name -contains $_.name }

# As of now, assume there is only 1 site.
$SiteId = 1

$Domains | ForEach-Object {

    $Domain = $_

    # Derive and set identifier.
    $Domain | Add-Member -NotePropertyName "pk" -NotePropertyValue $($_.app_icons -split "domain=")[1]

    Write-Host -ForegroundColor Green "Processing Workspace domain $($Domain.name)"

    # Skip if login permissions were never configured.
    if($Domain.sign_in_user_labels -eq "") {
        Write-Host -ForegroundColor Yellow "Warning: Login permissions were never set? Skipping this domain."
        return
    }
    
    # A theme will need to be created.
    # Let's do so even if only one domain was selected.
    try {
            
        $Theme = Get-RASTheme `
            -Name $Domain.name `
            -SiteId $SiteId `
            -ErrorAction Stop

    }
    catch {

        $Theme = New-RASTheme `
            -SiteId $SiteId `
            -Name $Domain.name
    
    }
    
    
    # Check if TOTP was selected.
    # Not supported: HOTP, direct DUO or RADIUS integration.
    # Even if it was just selected and no "mfa:required" label was added above, this MFA provider will be pre-created.
    $TwoFactorProvider = $TwoFactorProviders | Where-Object { $_.uri -eq $Domain.twofactor }

    if($TwoFactorProvider.key -eq "TOTP" -and $Domain.sign_in_context_labels -match "mfa:required") {

        $ExistingMFAProvider = Get-RASMFA -Name "Standard TOTP"

        if($null -eq $ExistingMFAProvider) {

            # To decide: Create a different TOTP for every single domain?
            # Pro: Trusted IPs could be added.
            # Con: Complexity, duplicate configs, ...
            New-RASMFA ``
                -SiteId $SiteId `
                -Name "Standard TOTP" `
                -DisplayName "One-time password" `
                -Description "A standard configuration, supported by many different vendors." `
                -Enabled $true `
                -TOTP `
                -TOTPType TOTP

            $ExistingMFAProvider = Get-RASMFA -Name "Standard TOTP"

        }

        Set-RASTheme `
            -SiteId $SiteId `
            -Name $Theme.Name `
            -MfaId $ExistingMFAProvider.Id


    }

    

    # While instead of "Domain Users" a group restriction could just NOT be configured, 
    # administrators should be encouraged to go for maximum security and revise this.
    Set-RASTheme `
        -SiteId $SiteId `
        -Name $Theme.Name `
        -GroupEnabled $true

    $Domain.sign_in_user_labels -split " " | ForEach-Object {

        $Split = $_ -split ":"
        $Key = $Split[0]
        $Value = $Split[1]

        Write-Host "Processing login permission $($Key):$($Value)"

        if($Key -match "user") {
            Write-Host -ForegroundColor Yellow "Warning: Only group permissions are allowed in $RAS. Skipping specific login permission for $Value ."
        }
        
        if($key -match "group") {

            $SAMAccountName = "$($Domain.netbios)\$Value"

            try {

                $Permission = Get-RASThemeGroupFilter `
                    -SiteId $SiteId `
                    -Name $Domain.name `
                    -Account $SAMAccountName

                $Permission | Format-List

                if($null -eq $Permission) {
                        
                    Add-RASThemeGroupFilter `
                        -SiteId $SiteId `
                        -Name $Domain.name `
                        -GroupName "$($Domain.netbios)\$Value" `
                        -ErrorAction Stop

                }

            }
            catch {
                Write-Host -ForegroundColor Yellow "Warning: Unable to add login permission for group $SAMAccountName ."
            }

        }

        if($key -match "all") {

            try {

                Add-RASThemeGroupFilter `
                    -SiteId $SiteId `
                    -Name $Domain.name `
                    -GroupName "$($_.netbios)\Domain Users" `
                    -ErrorAction Stop

            }
            catch {
                Write-Host -ForeGroundColor Yellow "Warning: Unable to add login permission for group Domain Users ."
            }

        }

    }

    # PSW has the appication servers defined on the Workspace domain level.


    # The applications published in PSW.
    # Here, filtering should be added.
    # It's probably easiest if individual "Allow if theme is ..." rules are added; with a Deny All at the end.

    Write-Host "Processing app servers"



    

    $WinRemoteApps = New-Object System.Collections.ArrayList
    
    $AppServers | Where-Object {
        $_.domain -match "/$($Domain.pk)/$"
    } | ForEach-Object {

        $AppServer = $_

        Write-Host "Processing app server: $($AppServer.host)"
        
        $RASHosts = New-Object System.Collections.ArrayList

        try {
 
            $IsRDSH = $false
            $IsRDCB = $false

            $Feature = Invoke-Command -ComputerName $AppServer.host -ScriptBlock {
                Get-WindowsFeature RDS-RD-Server
            } -ErrorAction Stop

            $IsRDSH = $Feature.Installed

            
            $Feature = Invoke-Command -ComputerName $AppServer.host -ScriptBlock {
                Get-WindowsFeature RDS-Connection-Broker
            } -ErrorAction Stop

            $IsRDCB = $Feature.Installed


            if($IsRDSH) {
                Write-Host "RD Session Host role installed."
            }
            if($IsRDCB) {
                Write-Host "RD Connection Broker role installed."
            }

        }
        catch {

            # WinRM failure, offline, DNS failure, firewall, etc.
            # Unfortunately, Linux machines (xRDP) will also end up in this category.
            Write-Host -ForegroundColor Red "Unable to detect server type. Server does not seem to be reachable, or is not a Microsoft Windows machine."
            return

        }
        

        if($IsRDCB) {

            # In case of a Microsoft Remote Desktop Connection Broker:
            # Both the RDCB and RDSH roles can be installed on the same machine.
            # But, as a best practice, RDCB should have been installed on a separate machine.
            # If the PSW environment points to an RDCB, the session hosts belonging to the collections should also be created in RAS.

            $Broker = $AppServer.host

            $Collections = Get-RDSessionCollection `
                -ConnectionBroker $Broker

            foreach($Collection in $Collections) {

                $SessionHosts = Get-RDSessionHost `
                    -CollectionName $Collection.CollectionName `
                    -ConnectionBroker $Broker

                foreach($SessionHost in $SessionHosts) {

                    [PSCustomObject]@{
                        ConnectionBroker = $AppServer.host
                        CollectionName = $Collection.CollectionName
                        SessionHost = $sessionHost.SessionHost
                    }

                    # Try to create this session host in RAS.
                    # Check if it exists in RAS.
                    try {

                        # This will throw an error if the host doesn't exist yet.
                        $RASHost = Get-RASRDSHost `
                            -SiteId $SiteId `
                            -Server $SessionHost.SessionHost `
                            -ErrorAction Stop

                        $RASHosts.Add($RASHost)

                        Write-Host "For broker $($Broker): RDS Host $($SessionHost.SessionHost) already exists."

                    }
                    catch {
                        
                        try {

                            Write-Host "For broker $($Broker): Create RDS Host $($SessionHost.SessionHost)"

                            $Params = @{
                                SiteId = $SiteId
                                Server = $SessionHost.SessionHost
                                Username = $Upn
                                Password = $($Password | ConvertTo-SecureString -Force -AsPlainText)
                                ErrorAction = "Stop"
                            }

                            $RASHost = New-RASRDSHost @Params

                            $RASHosts.Add($RASHost)
                            
                            Write-Host "Host successfully added."

                        }
                        catch {
                            Write-Host -ForegroundColor Red $_
                            $Params | Format-Table
                            Write-Host "Unable to create host (managed by broker)."
                        }


                        if($null -ne $RASHost) {
                                
                            try {
                                
                                Write-Host "Configure RDS Host."
                            
                                # The "name" (PSW) can not be directly mapped in RAS; it could only be added into the description.
                                
                                $Params = @{
                                    Id = $RASHost.Id
                                    Enabled = $AppServer.enabled
                                    Description = ($null -eq $AppServer.Description ? "" : $AppServer.Description)
                                    Port = $AppServer.port
                                    MaxSessions = ($AppServer.max_connections -eq 0 ? 1000 : $AppServer.max_connections)
                                    ErrorAction = "Stop"
                                }
                                
                                Set-RASRDSHost @Params

                                
                                Write-Host "Host successfully configured."

                            }
                            catch {
                                Write-Host -ForegroundColor Red $_
                                $Params | Format-Table
                                Write-Host "Unable to configure host."
                            }

                        }

                    }

                }
            }

            # Todo: Consider a host pool?
            
            $SpecificBrokerRemoteApps = Get-RDRemoteApp `
                -ConnectionBroker $Broker

            $WinRemoteApps += $SpecificBrokerRemoteApps



        }
        elseif($IsRDSH) {


            # Todo: Consider a host pool?
            
            # If it is *exclusively* a Remote Desktop Session Host (and NOT a Connection Broker):
            # Check if it exists in RAS.
            try {

                # This will throw an error if the host doesn't exist yet.
                $RASHost = Get-RASRDSHost `
                    -SiteId $SiteId `
                    -Server $_.host `
                    -ErrorAction Stop

                $RASHosts.Add($RASHost)

                Write-Host "RDS Host already exists."

            }
            catch {
                
                try {

                    Write-Host "Create Host $($AppServer.host) (RDSH)."

                    $Params = @{
                        SiteId = $SiteId
                        Server = $AppServer.host
                        Username = $Upn
                        Password = $($Password | ConvertTo-SecureString -Force -AsPlainText)
                        ErrorAction = "Stop"
                    }

                    $Params | Format-Table

                    $RASHost = New-RASRDSHost @Params

                    $RASHosts.Add($RASHost)

                    Write-Host "Configure Host."

                    $Params | Format-Table

                    $Params = @{
                        SiteId = $SiteId
                        Id = $RASHost.Id
                        Enabled = $AppServer.enabled
                        Description = $AppServer.description
                        Port = $AppServer.port
                        MaxSessions = $AppServer.max_connections
                        ErrorAction = "Stop"
                    }
                
                    # The "name" (PSW) can not be directly mapped in RAS; it could only be added into the description.
                    Set-RASRDSHost @Params
                    
                    Write-Host "Host successfully added."

                }
                catch {
                    Write-Host -ForegroundColor Red $_
                    $Params | Format-Table
                    Write-Host "Unable to create and configure host."
                }

            }

        }
        else {

            
            # RemotePC creation through PowerShell requires an explicit provider.
            $ProviderName = "Provider for remote PC $($AppServer.host)"

            try {
                
                $Params = @{
                    SiteId = $SiteId
                    Name = $ProviderName
                    ErrorAction = "Stop"
                }

                $Provider = Get-RASProvider @Params

                Write-Host "Found existing RAS Provider: $($ProviderName)"

            }
            catch {

                Write-Host "Create new RAS Provider: $($ProviderName)"

                # "name" is not allowed here.
                $Params = @{
                    RemotePCStaticVersion = "RemotePCStatic"
                    SiteId = $SiteId
                    Server = $Server
                    ProviderUsername = $Upn
                    ProviderPassword = $($Password | ConvertTo-SecureString -AsPlainText -Force)
                    ErrorAction = "Stop"
                }

                $Params | Format-Table

                $Provider = New-RASProvider -RemotePCStatic @Params

                Invoke-RASApply

                Write-Host "Created new RAS Provider: $($ProviderName)"
                
                $Params = @{
                    Id = $Provider.Id
                    NewName = $ProviderName
                    Description = "A provider created during migration from $PSW to $RAS"
                }

                Set-RASProvider @Params

                Invoke-RASApply
                
                # Re-query now that the name has changed, and the property is read-only.

                $Params = @{
                    SiteId = $SiteId
                    Name = $ProviderName
                    ErrorAction = "Stop"
                }

                $Provider = Get-RASProvider @Params
                

            }

            Write-Host "Provider details:"
            $Provider | Format-List

            $Params = @{
                SiteId = $SiteId
                Name = $Provider.Name
            }
            
            $Params | Format-Table

            if((Get-RASProviderRemotePCStatic @Params).Name -contains $AppServer.host) {

                Write-Host "Server already exists."

            }
            else {


                # Create a provider.
                try {
                    
                    Write-Host "Create Host $($AppServer.host) (RemotePC)."

                    $Params = @{
                        RemotePCStaticName = $AppServer.host
                        Id = $Provider.Id
                        ErrorAction = "Stop"
                    }

                    Add-RASProviderRemotePCStatic @Params

                    Write-Host "Host successfully added."

                }
                    
                catch {
                    Write-Host -ForegroundColor Red $_
                    Write-Host "Unable to create and configure host."
                }


            }

            
            # Create host pool with one provider (that contains one remote PC).
            $HostPoolName = ($AppServer.host -replace '[^a-zA-Z0-9-]', '-')

            try {

                $RemotePCPool = Get-RASVDIHostPool `
                    -SiteId $SiteId `
                    -Name $HostPoolName `
                    -ErrorAction Stop

                Write-Host "Found existing RAS VDI Host Pool: $($HostPoolName)"

            }
            catch {

                $Params = @{
                    SiteId = $SiteId
                    Name = $HostPoolName
                    Enabled = $AppServer.enabled
                    ProvisioningType = "Standalone"
                    ErrorAction = "Stop"
                }

                $Params | Format-Table

                $RemotePCPool = New-RASVDIHostPool @Params

                Write-Host "Created new RAS VDI Host Pool: $($HostPoolName)"

            }

            $RemotePCPool | Format-List

            $Params = @{
                VdiHostPoolId = $RemotePCPool.Id
            }

            if((Get-RASVDIHostPoolMember @Params).length -eq 0) {

                $Params = @{
                    Name = $Provider.Name
                    VDIHostPoolId = $RemotePCPool.Id
                    ProviderId = $Provider.Id
                    Type = "AllHostsInProvider" # Note: The provider only contains one remote PC.
                }

                $Params | Format-Table
                
                Add-RASVDIHostPoolMember @Params

            }

            $AppServer | Add-Member -MemberType NoteProperty -Name "_ras_host_pool" -Value $RemotePCPool


        }


        $AppServer | Add-Member -MemberType NoteProperty -Name "_ras_hosts" -Value $RASHosts

        Invoke-RASApply

    }


    $WinRemoteApps = $WinRemoteApps | Select-Object -Unique


    # Loop through the apps.

    # If a new app must be created, and it's a RemoteApp, the configuration must be derived from the Microsoft Connection Broker.

    # If brokers are stand-alone session hosts; it's pretty straight-forward.
    # If the appserver: label is used while this is only a session host and *not* the connection broker, this will most likely fail.


    $Apps | Where-Object {
        $_.domain -match "/$($Domain.pk)/$" -And (@("DESKTOP", "REMOTE-APP", "RDP") -contains $_.type)
    } | ForEach-Object {

        $App = $_ 

        Write-Host "$("*" * 25) Processing app $($App.name) ( $($App.type) )"

        $LinkedHosts = New-Object System.Collections.ArrayList
        $LinkedHostPools = New-Object System.Collections.ArrayList

        # Get the linked session host(s).
        
        Write-Host "Server labels: $($App.server_labels)"

        $App.server_labels -split " " | ForEach-Object {
            
            $ServerLabel = $_
            $Split = $ServerLabel -split ":"
            $Key = $Split[0]

            # Find the application server(s) with the same label.
            # In a correct configuration; appserver: should point to Microsoft Windows Servers with the RD Session Host role.
            # The appserver: label should be more or less 1:1.
            if($Key -eq "appserver" -or $Key -eq "rdscollection") {

                # Find the application server(s) in the same domain with the same label.
                $AppServers | Where-Object { $_.domain -match "/$($Domain.pk)/" -and $_.labels -match "(^| )$ServerLabel( |$)" } | ForEach-Object {

                    $AppServer = $_
                    Write-Host "Found linked app server: $($AppServer.name)"

                    # The RAS host(s) mapping to this application server.
                    $AppServer._ras_hosts | ForEach-Object {
                        $LinkedHosts.Add($_) | Out-Null
                    }

                    $LinkedHostPools.Add($AppServer._ras_host_pool) | Out-Null

                }

            }
            else {
                Write-Host -ForegroundColor Red "Unexpected key: $($Key)"
            }
        
        }

        Write-Host "Only keeping unique hosts"

        # It is possible that for instance the Remote Desktop Connection brokers were added as application servers.
        # They would share an rdscollection: label; the script would add the (same) linked RAS hosts, ...
        $LinkedHosts = $LinkedHosts | Select-Object -Unique
        $LinkedHostPools = $LinkedHostPools | Select-Object -Unique

        # To consider: categories, auto start (in PSW, it relies on user labels!), max number of sessions (?), ...


        switch($App.color_depth) {
            "16" {
                $ColorDepth = "Colors16Bit"
            }
            "24" {
                $ColorDepth = "Colors24Bit"
            }
            "32" {
                $ColorDepth = "Colors32Bit"
            }
        }


        if($App.type -eq "DESKTOP") {

            $RASApp = Get-RASPubRDSDesktop -SiteId $SiteId -Name $App.name

            if($RASApp.length -eq 0) {

                Write-Host "Publish RAS RDS Desktop"
                
                try {
                        
                    $Params = @{
                        EnabledMode = "Enabled"
                        Name = $App.Name
                        Description = $App.Description
                        SiteId = $SiteId
                        ErrorAction = "Stop"
                    }

                    if($LinkedHosts.Count -gt 0) {
                        $Params.PublishFrom = "Host"
                        $Params.PublishFromServerIds = $LinkedHosts.Id
                    }
                    elseif ($LinkedHostPools.Count -gt 0) {
                        $Params.PublishFrom = "HostPool"
                        $Params.PublishFromHostPoolIds = $LinkedHostPools.Id
                    }
                    else {
                        Write-Host -ForegroundColor Red "Unable to create RAS RDS Desktop. No linked hosts nor host pools available."
                    }

                    if ($Params.PublishFrom) {
                        $RASApp = New-RASPubRDSDesktop @Params
                        Write-Host "RAS RDS Desktop created."
                    }

                }
                catch {

                    Write-Host -ForegroundColor Red $_
                    $Params | Format-Table
                    Write-Host "Unable to create RAS RDS Desktop."

                }
                
            }
            else {

                Write-Host -ForegroundColor Yellow "RAS RDS Desktop already exists."

            }

        }
        elseif($App.type -eq "RDP") {
                
            # Todo: Decide: What to do with apps that were published as RDP?
            # Just rely on the default approach in RAS?

            # As of now, just return.
            Write-Host -ForeGroundColor Yellow "RDP (non RemoteApp) is currently not supported in this migration script."
            return


        }
        elseif($App.type -eq "REMOTE-APP") {

            # Try to get the app.
            # It would only throw an error if the name is e.g. empty; but not if the app simply doesn't exist.
            $RASApp = Get-RASPubRDSApp -SiteId $SiteId -Name $App.name

            if($null -eq $RASApp) {

                # See if there is matching Microsoft Windows RemoteApp configuration.
                $WinRemoteApp = $WinRemoteApps | Where-Object { $_.Alias -eq $App.command }
                
                if($null -eq $WinRemoteApp) {
                    
                    Write-Host -ForegroundColor Red "Unable to find RemoteApp with alias $($App.command)"
                    return

                }
                
                switch($WinRemoteApp.CommandLineSetting) {
                    "Allow" {
                        # If nothing is set, it doesn't matter.
                        $Parameters = ""
                    }
                    "DoNotAllow" {
                        $Parameters = ""
                    }
                    "Require" {
                        $Parameters = $WinRemoteApp.RequiredCommandLine
                    }
                }


                # $WinRemoteApp | fl

                if($LinkedHosts.length -eq 0) {
                    Write-Host -ForegroundColor Red "The app will NOT be created. No linked hosts."
                    return
                }

                # Debug
                Write-Host "Linked hosts:"
                $LinkedHosts | Format-List
                $LinkedHosts.Id

                
                $LinkedHosts | ForEach-Object {
                    Write-Host "Type: $($_.GetType().FullName)"
                    Write-Host "Value: [$_]"
                }


                Write-Host "Found $($LinkedHosts.length) or $($LinkedHosts.count) linked hosts"

                $RASApp = New-RASPubRDSApp `
                    -SiteId $SiteId `
                    -PublishFrom Host `
                    -EnabledMode "Enabled" `
                    -Name $App.name `
                    -Description $App.description `
                    -PublishFromServerIds $LinkedHosts.Id `
                    -Target $WinRemoteApp.FileVirtualPath `
                    -StartIn $(Split-Path $WinRemoteApp.FileVirtualPath -Parent) `
                    -Parameters $Parameters `
                    -ErrorAction Stop

                Write-Host "App created."
                
                # Mapping for color depth
                Set-RASPubRDSApp `
                    -SiteId $SiteId `
                    -Id $RASApp.Id `
                    -ColorDepth $ColorDepth `
                    -OneInstancePerUser $App.concurrent_usage

                Write-Host "RAS RDS App created."

            }
            else {

                Write-Host -ForegroundColor Yellow "RAS RDS App $($App.name) already exists."

            }

        }

        if($null -eq $RASApp) {

            Write-Host -ForegroundColor Red "The app was NOT created. Perhaps a host was unavailable?"
            return
        }

        $RASApp | Format-List

        # For maximum security, the default rule will deny access.
        Set-RASPubItemFilter -SiteId $SiteId -Id $RASApp.Id -Default Deny
        

        # Link the application to a theme.
        # Criteria are added to a rule.

        $RASRuleName = "Access"
        $RASRule = Get-RASRule `
            -SiteId $SiteId `
            -ObjType PubItem `
            -Id $RASApp.Id | `
            Where-Object { $_.Name -eq $RASRuleName }

        if($null -eq $RASRule) {

            Write-Host "Create rule $RASRuleName for app $($RASApp.Name)"

            $Params = @{
                SiteId = $SiteId
                ObjType = "PubItem"
                Id = $RASApp.Id
                Enabled = $true
                RuleName = $RASRuleName
                ErrorAction = "Stop"
            }

            $Params | Format-Table

            Add-RASRule @Params

            Invoke-RASApply

            $RASRule = Get-RASRule `
                -SiteId $SiteId `
                -ObjType PubItem `
                -Id $RASApp.Id | `
                Where-Object { $_.Name -eq $RASRuleName }



        }

        Write-Host "Rule:"
        $RASRule | Format-List

        if($null -eq $RASRule) {
            Write-Host -ForegroundColor Red "Failed to create or retrieve RAS rule."
            return
        }

        # Note: Do not use -SiteId here, it breaks things.
        $RASRuleCriteriaTheme = Get-RASCriteriaTheme `
            -ObjType PubItem `
            -Id $RASApp.Id `
            -RuleId $RASRule.Id

        if($RASRuleCriteriaTheme.length -eq 0) {

            $Params = @{
                ObjType = "PubItem"
                Id = $RASApp.Id
                RuleId = $RASRule.Id
                ThemeId = $Theme.Id
                ErrorAction = "Stop"
            }

            $RASRuleCriteriaTheme = Add-RASCriteriaTheme @Params

            # Do not use -SiteId here.
            $Params = @{
                ObjType = "PubItem"
                Id = $RASApp.Id
                RuleId = $RASRule.Id
                ThemesEnabled = $true
                ThemesMatchingMode = "IsOneOfTheFollowing"
                ErrorAction = "Stop"
            }

            Set-RASCriteria @Params

            Invoke-RASApply

        }

        # Prepare for user filtering too.
        # Do not use -SiteId here.

            $Params = @{
                ObjType = "PubItem"
                Id = $RASApp.Id
                RuleId = $RASRule.Id
                SecurityPrincipalsEnabled = $($App.user_labels -ne "")
                SecurityPrincipalsMatchingMode = "IsOneOfTheFollowing"
                ErrorAction = "Stop"
            }
            Set-RASCriteria @Params

        # User labels (Parallels Secure Workspace) will be converted into criteria for Parallels Remote Application Server.
        # The admin: label must be replaced with the actual usernames/admins.
        $App.user_labels -replace "admin:", $Domain.admin_user_labels -split " " | ForEach-Object {

            $Split = $_ -split ":"
            $Key = $Split[0]
            $Value = $Split[1]

            Write-Host "Processing user permission $($Key):$($Value)"

            # User labels should only have "user", "group", "all" keys.
            if($Key -match "user" -or $key -match "group") {

                $SAMAccountName = "$($Domain.netbios)\$Value"
                $SID = "SID://$($Domain.netbios)/$Value"

            }
            elseif($Key -match "all") {
                
                $SAMAccountName = "$($Domain.netbios)\Domain Users"
                $SID = "SID://$($Domain.netbios)/Domain Users"

            }
            else {

                Write-Host -BackgroundColor Red "Unsupported key: $($Key)"
                return

            }
            
            try {

                Invoke-RASApply
                
                # Check if it already exists.
                Write-Host "Check permission for security principal $SAMAccountName - Published item $($RasApp.Id) - Rule $($RASRule.Id)"

                # Note: Do not use -SiteId here, it breaks things.
                $Params = @{
                    ObjType = "PubItem"
                    Id = $RASApp.Id
                    RuleId = $RASRule.Id
                    ErrorAction = "Stop"
                }

                $Params | Format-Table

                $RASCriteriaPrincipal = Get-RASCriteriaSecurityPrincipal @Params | Where-Object { $_.Account -eq $SID }

                if($null -eq $RASCriteriaPrincipal) {
                    
                    Write-Host "Add permission for security principal $SAMAccountName"

                    $Params = @{
                        ObjType = "PubItem"
                        Id = $RASApp.Id
                        RuleId = $RASRule.Id
                        Account = $SAMAccountName
                        ErrorAction = "Stop"
                    }

                    $Params | Format-Table

                    # Note: Do not use -SiteId here, it breaks things.
                    Add-RASCriteriaSecurityPrincipal @Params
                }
                else {

                    Write-Host "Permission for security principal $SAMAccountName already exists."
                    $RASCriteriaPrincipal | Format-List

                }

                
                Invoke-RASApply


            }
            catch {
                Write-Host -ForegroundColor Yellow "Warning: Unable to add permission for $SAMAccountName ."
                $_
            }



        }


        # categories?
        



    }

}


$LetsEncryptCerts = $SSLCerts | Where-Object { $_.is_automatic -eq $true }

if($LetsEncryptCerts.length -gt 0) {
        
    $Option = Read-Host -Prompt `
        "Let's Encrypt was used to automatically request certificates. " `
        "For $RAS, you must explicitly accept the Let's Encrypt EULA. " `
        "You can find this online.`n`n" `
        "Confirm your choice:`n" `
        "1. Accept.`n" `
        "2. Reject (no migration will happen)`n`n" `
        "Enter number"

    if($Option -eq "1") {

        Write-Host -ForegroundColor Yellow `
            "Do not forget to allow outbound AND inbound connectivity (both directionos) " `
            "between the $RAS servers and Let's Encrypt, TCP port 80 and 443."

        $Email = Read-Host -Prompt `
            "$RAS requires you to specify one (preferably non personal, but some team) e-mail address to send Let's Encrypt notifications to when needed. Please enter one."

        Set-RASLetsEncryptSettings `
            -SiteId $SiteId `
            -TermsAccepted $true `
            -ExpirationEmails $Email
        
    }

    $LetsEncryptCerts | ForEach-Object {

        # In PSW, one Let's Encrypt certificate only has one name.
        New-RASCertificate `
            -SiteId $SiteId `
            -Name $_.subject_names `
            -CommonName $_.subject_names `
            -LetsEncrypt

    }



}

# Safety. Apply any unsaved changes.
Invoke-RASApply

Write-Host -ForegroundColor Green "Finished."

Remove-RASSession
