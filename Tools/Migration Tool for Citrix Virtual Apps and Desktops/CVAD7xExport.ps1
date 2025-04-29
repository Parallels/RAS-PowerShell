#requires -version 3

<#
.SYNOPSIS
    Backup XenDesktop apps to CLIXML to be imported by Parallels RAS script

.PARAMETERS ddc
    Delivery controller to connect to. If not specified will try and connect locally

.PARAMETERS folder
    The empty folder to place the exported files in. If not existent will be created.
    
.PARAMETERS cloud
    Work with Citrix DaaS rather than on-prem

.PARAMETERS customerId
    The customer id for Citrix DaaS

.PARAMETERS profileName
    The local Citrix DaaS profile name to use for authentication. Previously saved with Set-XDCredentials

.PARAMETERS overwrite
    When specified will not exit if the destination folder already contains files

.PARAMETERS maxRecordCount
    The maximum number of records to retrieve from Citrix
    
.PARAMETERS limitAppGroups
    Only export as many as this number of app groups. Designed to allow quick testing.
    
.PARAMETERS limitApps
    Only export as many as this number of apps. Designed to allow quick testing.

.EXAMPLE
    & '.\CVAD7xExport.ps1' -ddc xaddc01 -folder c:\temp\CitrixExport

    Connect to the delivery controller and export the configuration to the given folder
    
.EXAMPLE
    & '.\CVAD7xExport.ps1' -cloud -folder c:\temp\CitrixExport

    Connect to Citrix DaaS with interactive (browser) authentication and export the configuration to the given folder

.NOTES
    Requires Citrix CVAD PowerShell cmdlets to be installed or the the Remote SDK when using DaaS
#>

<#
Copyright © 2024 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='Cloud',Mandatory=$false)]
    [switch]$cloud ,
    [Parameter(ParameterSetName='OnPrem',Mandatory=$true)]
    [string]$ddc ,
    [Parameter(Mandatory=$true,HelpMessage='Base folder to write exported CLIXML files')]
    [string]$folder ,
    [Parameter(ParameterSetName='Cloud',Mandatory=$false)]
    [string]$customerId ,
    [Parameter(ParameterSetName='Cloud',Mandatory=$false)]
    [string]$profileName ,
    [int]$maxRecordCount = 50000 ,
    [int]$limitAppGroups ,
    [int]$limitApps ,
    [switch]$overwrite
    ## TODO implement parameters for filtering on specific types or names matching pattern/regex
)

## map Citrix properties to RAS ones - empty string means remove the property, not in the table means pass through
## + at start of value means replace the :: delimited property names with the value for those properties. If property does not exist or is empty, the replacement will be the empty string
## @ at start means the data in the source will be morphed as per a separate mapping entry
[hashtable]$CitrixApplicationsToRAS = @{
    'CommandLineExecutable' = '+"::CommandLineExecutable::" ::CommandLineArguments::'
    'AssociatedUserNames' = '@Accounts : ^(?<domain>.+)\\(?<user>.+)$' ## Each source item will be parsed to match the regex and each instance of the named match etc will be replaced by the corresponding $match
    'AssociatedUserSIDs' = '' ## no guarantee that order of this array is same as user names so discard SIDs and look them up afresh when mapping AssociatedUserNames
    'EncodedIconData' = 'IconData'
    'PublishedName' = 'DisplayName'
    'AllAssociatedDesktopGroupUids' = ''
    'AllAssociatedDesktopGroupUUids' = ''
    'AssociatedApplicationGroupUids' = ''
    'AssociatedApplicationGroupUUIDs' = ''
    'AssociatedDeliveryGroups' = ''
    'AssociatedDesktopGroupPriorities' = ''
    'AssociatedDesktopGroupUids' = ''
    'AssociatedDesktopGroupUUIDs' = ''
    'WaitForPrinterCreation' = ''
    'Tags' = ''
    'Uid' = ''
    'IconUid' = ''
    'AdminFolderUid' = ''
}

## mappings for @items
## strings will be "executed" after placeholders filled in so functions, etc can be called on them and return value is what goes in output. %var% replaced by $match of that named group
[hashtable]$structureMappings = @{
    'Accounts' = @{
        AccountDisplayName = '%domain%\%user%'
        AccountName = '%user%'
        AccountType = '$(Get-AccountInfo `"%user%`" ''AccountType'' )' ## don't use domain which may cause problems in multi-domain/trusts - TODO cater for domain (how do we get FQDN from just domain name?)
        AccountId = '0X2/NT/%domain%/$((New-Object System.Security.Principal.NTAccount(`"%domain%\%user%`")).Translate([System.Security.Principal.SecurityIdentifier]).value)'
        AccountAuthority = 'NTDomain/%domain%'
        OtherName = '%user%'
        SearchPath = '$(Get-AccountInfo `"%user%`" ''OU'' )'
        MachineName = $env:COMPUTERNAME
    }
}

[hashtable]$CitrixDeliveryGroupsToRAS = @{
    'DesktopGroupName' = 'WorkerGroupName'
}

[hashtable]$CitrixZonesToRAS = @{
    'Name' = 'ZoneName'
    'ControllerNames' = 'DataCollector'
}

[hashtable]$CitrixMachinesToRAS = @{
    'HostedMachineName' = 'ServerName'
    'DNSName' = 'ServerFqdn'
}

[hashtable]$CitriSitesToRAS = @{
    'Name' = 'FarmName'
    'LicensedSessionsActive' = 'SessionCount'
}

[System.Collections.Generic.List[string]]$citrixModules = @( 'Citrix.Broker.Admin.V2' , 'Citrix.Configuration.Admin.V2' )

Function Get-AccountInfo
{
    Param
    (
        [string]$accountName ,
        [ValidateSet('AccountType','OU')]
        [string]$attribute = 'AccountType'
    )
    [string]$result = ''
    Write-Verbose -Message "Get-AccountInfo( $accountName , $attribute )"
    $searcher = [ADSISearcher]"(&(sAMAccountName=$accountName))"
    $account = $null
    $account = $searcher.FindOne()
    if( $null -ne $account )
    {
        if( $attribute -eq 'AccountType' )
        {
            if( $account.Properties.objectclass -icontains 'group' )
            {
                $result = 'Group'
            }
            elseif( $account.Properties.objectclass -icontains 'person' -or $account.Properties.objectclass -icontains 'user' )
            {
                $result = 'User'
            }
            else
            {
                Write-Warning -Message "Account $accountName of unexpected type $($account.Properties.objectclass -join '#' )"
            }
        }
        elseif( $attribute -eq 'OU' )
        {
            ## remove first part of string which is name of object itself CN=Citrix_PublishedApp_PowerShell,OU=Groups,OU=Users,OU=Wakefield,OU=Sites,DC=guyrleech,DC=local
            $result = $account.Properties.distinguishedname -replace '^[^,]+,'
        }
        $result ## output
    }
    else
    {
        Write-Warning -Message "Failed to find account $accountName"
    }
}

Function Convert-CitrixToRAS
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$resourceType ,
        [Parameter(Mandatory=$true)]
        $outputObject ,
        [Parameter(Mandatory=$true)]
        [string]$folder ,
        [string]$baseFileName ,
        [Parameter(Mandatory=$true)]
        [hashtable]$mappingTable ,
        [hashtable]$extraProperties = @{}
    )
    
    if( $outputObject.PSObject.Properties[ $resourceType ] )
    {
        if( [string]::IsNullOrEmpty( $baseFileName ) )
        {
            $baseFileName = $resourceType
        }
        [string]$outputFile = Join-Path -Path $folder -ChildPath "$baseFileName.xml"
    
        Write-Verbose -Message "Writing $($outputObject.$resourceType.Count) $resourceType to `"$outputFile`""

        & {
            ForEach( $item in $outputObject.$resourceType )
            {
                [hashtable]$result =$extraProperties.Clone()

                ForEach( $property in $item.psobject.properties )
                {
                    $propertyMapping = $mappingTable[ $property.Name ]
                    if( $null -eq $propertyMapping )
                    {
                        $result.Add( $property.Name , $property.Value )
                    }
                    elseif( $propertyMapping.Length -gt 0 )
                    {
                        if( $propertyMapping[0] -eq '+' ) ## we have some morphing to do
                        {
                            [string]$expandedValue = $propertyMapping.SubString( 1 )
                            ## replace each : delimited property with its property value, if present
                            While( $true )
                            {
                                if( $expandedValue -match '::([a-z0-9]+)::' )
                                {
                                    [string]$wholeMatch = $Matches[0]
                                    [string]$propertyName = $Matches[1]
                                    $propertyValue = $null
                                    if( $item.PSObject.Properties[ $propertyName ] )
                                    {
                                        $propertyValue = $item.$propertyName
                                    }
                                    ## else property doesn't exist so it will be replaced with null
                                    [string]$replaced = $expandedValue -replace $wholeMatch , $propertyValue
                                    if( $replaced -eq $expandedValue )
                                    {
                                        Write-Warning -Message "Problems expanding $propertyName in $expandedValue"
                                        break
                                    }
                                    else
                                    {
                                        $expandedValue = $replaced
                                    }
                                }
                                else
                                {
                                    break
                                }
                            }
                            $result.Add( $property.Name , $expandedValue.Trim() )
                        }
                        elseif( $propertyMapping[0] -eq '@' ) ## we have to create a new property and populate via placeholders
                        {
                            ## @Accounts : ^(?<domain>.+)\\(?<user>.+)$
                            if( $propertyMapping -notmatch '^@\s*(\w+)\s*:\s*(.+)$' )
                            {
                                Write-Warning -Message "Unparsable format in `"$propertyMapping`""
                            }
                            else
                            {
                                [string]$propertyName = $matches[ 1 ]
                                [string]$regex = $matches[ 2 ]
                            
                                $structureMapping = $structureMappings[ $propertyName ]

                                if( $null -eq $structureMapping )
                                {
                                    Write-Warning -Message "No mapping structure found for property name `"$propertyName`" in string $propertyName"
                                }
                                else
                                {
                                    [array]$mapped = @(ForEach( $value in $property.Value )
                                    {
                                        [hashtable]$newValues = @{}
                                        if( $value -match $regex )
                                        {
                                            ForEach( $mappedItem in $structureMapping.GetEnumerator() )
                                            {
                                                [string]$resultString = $mappedItem.Value
                                                ForEach( $match in $matches.GetEnumerator() )
                                                {
                                                    $resultString = $resultString -replace "%$($match.Name)%" , $match.Value
                                                }
                                                Write-Verbose -Message "Morphed property `"$($mappedItem.Name)`", value `"$($mappedItem.Value)`" to `"$resultString`""
                                                ## TODO now "execute" string to expand any functions,cmdlets, etc (<-- attack vector!!)
                                                $scriptBlock = [System.Management.Automation.ScriptBlock]::Create( "`"$resultString`"" )
                                                $resultString = $scriptBlock.Invoke()
                                                Write-Verbose -Message "String now $resultString"
                                                $newValues.Add( $mappedItem.Name , $resultString )
                                            }
                                        }
                                        else
                                        {
                                            Write-Warning -Message "Property value `"$value`" in property `"$($property.name)`" does not match regex $regex"
                                        }
                                        [pscustomobject]$newValues ## output
                                    })
                                    $result.Add( $propertyName , $mapped )
                                }
                            }
                        }
                        else ## simple string replacement
                        {
                            $result.Add( $propertyMapping , $property.Value )
                        }
                    }
                    # else mapping is empty string which means drop the property
                }
                [pscustomobject]$result ## output
            }
        } | Export-Clixml -Path $outputFile
    }
    else
    {
        Write-Warning -Message "No instances of resource type `"$resourceType`" found"
    }
}

if( Test-Path -Path $folder -PathType Container )
{
    if( -Not $overwrite -and $null -ne ($existingFiles = @( Get-ChildItem -Path $folder -File -Filter '*.xml' )) -and $existingFiles.Count -gt 0 )
    {
        Throw "Already got $($existingFiles.Count) files in folder `"$folder`""
    }
}
elseif( -Not (New-Item -Path $folder -ItemType Directory -Force))
{
    Throw "Failed to created folder `"$folder`""
}
if( -Not [string]::IsNullOrEmpty( $profileName ) -or -Not [string]::IsNullOrEmpty( $customerId )  )
{
    $cloud = $true
    $citrixModules.Add( 'Citrix.Sdk.Proxy.V*' )
}

## TODO Try ipmo
ForEach( $snapin in $citrixModules )
{
    if( -Not ( Get-PSSnapin -Name $snapin -ErrorAction SilentlyContinue ) )
    {
        Add-PSSnapin -Name $snapin
    }
}

## get AD so we can find accounts later if we need to
$rootDSE = [ADSI]"LDAP://RootDSE"
$domainDN = $rootDSE.Get("defaultNamingContext")

$outputObject = $null
$inputObject = $null
[int]$errors = 0

[hashtable]$citrixCommonParameters = @{ }

if( -Not [string]::IsNullOrEmpty( $ddc ) -and $ddc -ine 'localhost' -and $ddc -ne '.' )
{
    $citrixCommonParameters.Add( 'AdminAddress' , $ddc )
}
elseif( $cloud -or -not [string]::IsNullOrEmpty( $profileName ) -or -Not [string]::IsNullOrEmpty( $customerId ) )
{
    if( -Not ( Get-Command -Name Get-XDAuthentication -ErrorAction SilentlyContinue ))
    {
        Throw "Unable to attempt to authenticate to Citrix Cloud - is the Remote PowerShell SDK installed (https://download.apps.cloud.com/CitrixPoshSdk.exe) ?"
    }

    [hashtable]$authParameters = @{ ErrorVariable = 'cloudAuthError' }
    if( -Not [string]::IsNullOrEmpty( $profileName ) )
    {
        $authParameters.Add( 'ProfileName' , $profileName ) 
    }
    if( -Not [string]::IsNullOrEmpty( $customerId ) )
    {
        $authParameters.Add( 'CustomerId' , $customerId ) 
    }
    $cloudAuthError = $null
    $auth = Get-XDAuthentication @authParameters
    $cloud = $?
    if( -Not $cloud )
    {
        Throw "Authentication to Citrix Cloud failed: $cloudAuthError"
    }
}

[hashtable]$citrixDDCParameter = $citrixCommonParameters.Clone() ## used when Get-Broker* cmdlet doesn't take -MaxRecordCount

$citrixCommonParameters.Add( 'MaxRecordCount' , $maxRecordCount )

$outputObject = [pscustomobject]@{
    'ExportedBy' = $env:USERNAME
    'ExportedAt' = Get-Date -Format s
    'ExportedFrom' = $env:COMPUTERNAME
    'DeliveryController' = $ddc
    'Site' = @( Get-BrokerSite @citrixDDCParameter ) ## not used for restore, just for info
    'Zones' = @( Get-ConfigZone @citrixCommonParameters )
    ## Get-BrokerController fails for cloud so put customer id in instead
    'Controllers' = $(if( $cloud ) { @( [pscustomobject]@{ 'Customer id' = $customerId } ) } else { @( Get-BrokerController @citrixCommonParameters ) } ) ## not used for restore, just for info
    #'ProvisioningSchemes' = @( Get-ProvScheme @citrixCommonParameters )
    #'HypervisorConnections' = @( Get-BrokerHypervisorConnection @citrixCommonParameters ) ## we don't create these as need credentials but need to ensure they exist
    #'IdentityPools' = @( Get-AcctIdentityPool @citrixCommonParameters )
    #'AccessRules' = @( Get-BrokerAccessPolicyRule @citrixCommonParameters )
    #'Catalogs' = @( Get-BrokerCatalog @citrixCommonParameters )
    'DeliveryGroups' = @( Get-BrokerDesktopGroup @citrixCommonParameters )
    #'Desktops' = @( Get-BrokerEntitlementPolicyRule @citrixCommonParameters )
    'Machines' = @( Get-BrokerMachine @citrixCommonParameters )
    #'Tags' = @( Get-BrokerTag @citrixCommonParameters )
}

Add-Member -InputObject $outputObject -MemberType NoteProperty -Name Policies -Value (Export-BrokerDesktopPolicy)
    
[array]$appGroups = @( Get-BrokerApplicationGroup @citrixCommonParameters )
[int]$counter = 0

ForEach( $appGroup in $appGroups )
{
    $counter++
    Write-Verbose -Message "$counter / $($appGroups.Count): appgroup $($appGroup.Name)"
    if( $limitAppGroups -gt 0 -and $counter -gt $limitAppGroups )
    {
        Write-Warning -Message "Hit limit of $limitAppGroups on $($appGroups.Count) app groups items total"
        break
    }
    ## Get Desktop group so we can export
    ## Store delivery group names rather than GUIDs as they may change when we create new
    [string]$deliveryGroups = $null

    [array]$associatedDeliveryGroups = @( $appGroup.AssociatedDesktopGroupUids | ForEach-Object `
    {
        $AssociatedDesktopGroupUid = $_
        if( $deliveryGroup = $outputObject.DeliveryGroups.Where( { $_.Uid -eq $AssociatedDesktopGroupUid } , 1 ) )
        {
            $deliveryGroup.Name
        }
        else
        {
            Write-Warning -Message "Unable to find delivery group for uid $AssociatedDesktopGroupUid"
        }
    })

    Add-Member -InputObject $appGroup -MemberType NoteProperty -Name AssociatedDeliveryGroups -Value $associatedDeliveryGroups
}
    
Add-Member -InputObject $outputObject -MemberType NoteProperty -Name ApplicationGroups -Value $appGroups        
    
[array]$apps = @( Get-BrokerApplication @citrixCommonParameters )

Write-Verbose "Backing up $($apps.Count) apps to `"$folder`""
$counter = 0

ForEach( $app in $apps )
{
    $counter++
    Write-Verbose -Message "$counter / $($apps.Count): app $($app.Name)"
    if( $limitApps -gt 0 -and $counter -gt $limitApps )
    {
        Write-Warning -Message "Hit limit of $limitApps on $($apps.Count) apps items total"
        break
    }
    ## Get Desktop group so we can export
    ## If more than one, we'll make a delimited list so can store in this object AssociatedDesktopGroupUids 
    #[string]$deliveryGroups = $null
    ## trailing \ can break import
    $app.AdminFolderName = $app.AdminFolderName -replace '\\+$'
    $app.ClientFolder = $app.ClientFolder -replace '\\+$'
    $app.StartMenuFolder = $app.StartMenuFolder -replace '\\+$'

    [array]$associatedDeliveryGroups = @( $app.AssociatedDesktopGroupUids | ForEach-Object `
    {
        $AssociatedDesktopGroupUid = $_
        if( $deliveryGroup = $outputObject.DeliveryGroups.Where( { $_.Uid -eq $AssociatedDesktopGroupUid } , 1 ) )
        {
            $deliveryGroup.Name
        }
        else
        {
            Write-Warning -Message "Unable to find delivery group for uid $AssociatedDesktopGroupUid"
        }
    })

    Add-Member -InputObject $app -MemberType NoteProperty -Name AssociatedDeliveryGroups -Value $associatedDeliveryGroups

    [string]$appGroupsList = $null
    [array]$AssociatedApplicationGroups = @( ForEach( $AssociatedApplicationGroupUid in $app.AssociatedApplicationGroupUids )
    {
        if( $applicationGroup = $outputObject.ApplicationGroups.Where( { $_.Uid -eq $AssociatedApplicationGroupUid } , 1 ) )
        {
            $applicationGroup.Name
        }
        else
        {
            Write-Warning -Message "Unable to find application group for uid $AssociatedApplicationGroupUid for app $($app.Name)"
        }
    })

    ## we'll add the apps to the app group export later too
    Add-Member -InputObject $app -MemberType NoteProperty -Name AssociatedAppGroups -Value $AssociatedApplicationGroups

    ## dump icon to file
    if( -Not $app.IconFromClient )
    {
        Add-Member -InputObject $app -MemberType NoteProperty -Name EncodedIconData -Value (Get-BrokerIcon -Uid $app.IconUid).EncodedIconData
    }
}
        
Add-Member -InputObject $outputObject -MemberType NoteProperty -Name Applications -Value $apps

if( $outputObject )
{
    Convert-CitrixToRAS -resourceType 'applications'   -outputObject $outputObject -folder $folder -mappingTable $CitrixApplicationsToRAS
## TODO get ddc machine names and versions
    Convert-CitrixToRAS -resourceType 'zones'          -outputObject $outputObject -folder $folder -mappingTable $CitrixZonesToRAS -extraProperties @{ 'MachineName' = 'Dummy' }
    Convert-CitrixToRAS -resourceType 'deliverygroups' -outputObject $outputObject -folder $folder -mappingTable $CitrixDeliveryGroupsToRAS -baseFileName 'workergroups'
    Convert-CitrixToRAS -resourceType 'machines'       -outputObject $outputObject -folder $folder -mappingTable $CitrixMachinesToRAS       -baseFileName 'servers' 
    Convert-CitrixToRAS -resourceType 'site'           -outputObject $outputObject -folder $folder -mappingTable $CitriSitesToRAS           -baseFileName 'farm' -extraProperties @{ 'MachineName' = 'Dummy' ; 'ServerVersion' = '1.2.3.4' }
}
else
{
    Throw "No config to save"
}