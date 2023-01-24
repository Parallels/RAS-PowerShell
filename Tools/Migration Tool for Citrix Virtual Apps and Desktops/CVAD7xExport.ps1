#requires -version 3

<#
.SYNOPSIS
    Backup XenDesktop apps to CLIXML to be imported by Parallels RAS script

.PARAMETERS ddc
    Delivery controller to connect to. If not specified will try and connect locally

.PARAMETERS folder
    The empty folder to place the exported files in. If not existent will be created.

.PARAMETERS overwrite
    When specified will not exit if the destination folder already contains files

.PARAMETERS maxRecordCount
    The maximum number of records to retrieve from Citrix

.EXAMPLE
    & '.\CVAD7xExport.ps1' -ddc xaddc01 -folder c:\temp\CitrixExport

    Connect to the delivery controller and export the configuration to the given folder

.NOTES
    Requires Citrix CVAD PowerShell cmdlets to be installed (not the Remote SDK)

Modification History:

    2022/12/21  @guyrleech  First official release
    2023/01/18  @guyrleech  Make CommandLineExecutable contain executable and parameters
#>

<#
Copyright � 2022 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the �Software�), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED �AS IS�, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [string]$ddc ,
    [Parameter(Mandatory=$true,HelpMessage='Base folder to write exported CLIXML files')]
    [string]$folder ,
    [int]$maxRecordCount = 50000 ,
    [switch]$overwrite
    ## TODO implement parameters for filtering on specific types or names matching pattern/regex
)

## map Citrix properties to RAS ones - empty string means remove the property, not in the table means pass through
## + at start of value means replace the :: delimited property names with the value for those properties. If property does not exist or is empty, the replacement will be the empty string
[hashtable]$CitrixApplicationsToRAS = @{
    'CommandLineExecutable' = '+"::CommandLineExecutable::" ::CommandLineArguments::'
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
                        if( $propertyMapping[0] -eq '+' ) ## we have some moprhing to do
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

## TODO Try ipmo
if( ! ( Get-PSSnapin Citrix.Broker.Admin.V2 -EA SilentlyContinue ) )
{
    Add-PSSnapin Citrix.Broker.Admin.V2
}

$outputObject = $null
$inputObject = $null
[int]$errors = 0

[hashtable]$citrixCommonParameters = @{ }

if( -Not [string]::IsNullOrEmpty( $ddc ) -and $ddc -ine 'localhost' -and $ddc -ne '.' )
{
    $citrixCommonParameters.Add( 'AdminAddress' , $ddc )
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
    'Controllers' = @( Get-BrokerController @citrixCommonParameters ) ## not used for restore, just for info
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

ForEach( $appGroup in $appGroups )
{
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

## Write-Verbose "Backing up $($apps.Count) apps to `"$folder`""
ForEach( $app in $apps )
{
    ## Get Desktop group so we can export
    ## If more than one, we'll make a delimited list so can store in this object AssociatedDesktopGroupUids 
    #[string]$deliveryGroups = $null
            
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

    ## now dump icon to file
    if( ! $app.IconFromClient )
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
