<#
.SYNOPSIS
    Read Citrix policy export and convert to RAS policy

.DESCRIPTION
    Policies exported by Export-CtxGroupPolicy from Citrix.GroupPolicy.Commands.psm1 from Citrix Supportability Pack (https://support.citrix.com/s/article/CTX203082-citrix-supportability-pack?language=en_US)
    or from the module "$env:ProgramFiles\Citrix\Telemetry Service\TelemetryModule\Citrix.GroupPolicy.Commands.psm1" on delivery controllers (see .NOTES section)

.PARAMETER outputFile
    The name of the output file to create. Default is 'CreateRASPoliciesFromCitrix.ps1'.

.PARAMETER mappingFile
    The path to the CSV file containing the mappings of Citrix policies to RAS.

.PARAMETER citrixPolicyFolder
    The folder containing the policies already exported from Citrix.

.PARAMETER description
    The description to use for the migrated policies. Default is 'Migrated from Citrix by script'.

.PARAMETER preambleFile
    The path to a file containing preamble code to include in the output script. Default is 'Policy Preamble.ps1'.

.PARAMETER overWrite
    Switch to overwrite the output file if it already exists.

.PARAMETER prefix
    A prefix to add to the names of the migrated policies. Use this if policies with the same names already exist in RAS as they will not be overwritten by this script. Default is empty.

.PARAMETER blankLines
    The number of blank lines to insert between sections in the output script. Default is 1.

.PARAMETER enabled
    Switch to enable the migrated policies. By default, policies are disabled.

.PARAMETER baseFileName
    The name of the base file containing the list of policies. Default is 'GroupPolicy.xml'.

.PARAMETER configurationFileName
    The name of the file containing the policy configurations. Default is 'GroupPolicyConfiguration.xml'.

.PARAMETER filterFileName
    The name of the file containing the policy filters. Default is 'GroupPolicyFilter.xml'.

.PARAMETER citrixEquivalentColumnName
    The name of the column in the mapping file that contains the Citrix policy names. Default is 'Citrix Equivalent'.

.PARAMETER RASEquivalentColumnName
    The name of the column in the mapping file that contains the RAS policy names. Default is 'Set-RASClientPolicy Parameters'.

.PARAMETER RASPSModuleName
    The name of the PowerShell module for RAS. Default is 'RASAdmin'.

.PARAMETER ClientPolicyVariableName
    The name of the variable to use for the client policy in the output script. Default is 'clientPolicy'.

.PARAMETER baseRuleName
    The base name to use for the rules in the output script. Default is 'MigratedFromCitrix'.

.PARAMETER both
    The value to use for bidirectional mappings. Default is '**BOTH**'.

.PARAMETER ruleNumber
    The starting number to use for the rules in the output script. Default is 100.

.PARAMETER all
    Switch to process all pol

.EXAMPLE
    & '.\Citrix Policy conversion to RAS.ps1' -mappingFile 'c:\Parallels\RAS Policy Mapping.csv' -citrixPolicyFolder 'H:\Citrix Policies Export\'

    Process the files in the Citrix policy exported to 'H:\Citrix Policies Export\' and produce a file called "CreateRASPoliciesFromCitrix.ps1" in the current working directory
    which can be run on a Parallels RAS connection broker to create the correpsonding policies in RAS.
    The file 'c:\Parallels\RAS Policy Mapping.csv' is available on GitHub and contains the mappings of Citrix policies to RAS, where they exist

.NOTES
    Run the following on a Citrix Delivery Controller to export policies

    Import-Module -Name "$env:ProgramFiles\Citrix\Telemetry Service\TelemetryModule\Citrix.GroupPolicy.Commands.psm1"
    New-PSDrive -PSProvider CitrixGroupPolicy -Name LocalFarmGpo -Root \ -Controller $env:COMPUTERNAME
    Export-CtxGroupPolicy -FolderPath "H:\Citrix Policies Export"

    Modification History:
#>

[CmdletBinding()]

Param
(
    [string]$outputFile = 'CreateRASPoliciesFromCitrix.ps1' , ## output
    [Parameter(Mandatory=$true)]
    [string]$mappingFile ,
    [string]$citrixPolicyFolder ,
    [string]$description = 'Migrated from Citrix by script' ,
    [AllowNull()]
    [AllowEmptyString()]
    [string]$preambleFile = 'Policy Preamble.ps1', ### set a default value here if required to pull in code like disclaimers, copyright, etc ###
    [switch]$overWrite ,
    [string]$prefix ,
    [int]$blankLines = 1,
    [switch]$enabled , ## disabled by default
    [string]$baseFileName = 'GroupPolicy.xml' , ## hard coded in Export-CtxGroupPolicy
    [string]$configurationFileName = 'GroupPolicyConfiguration.xml' ,
    [string]$filterFileName = 'GroupPolicyFilter.xml' ,
    [string]$citrixEquivalentColumnName = 'Citrix Equivalent' ,
    [string]$RASEquivalentColumnName = 'Set-RASClientPolicy Parameters' ,
    [string]$RASPSModuleName = 'RASAdmin' ,
    [string]$ClientPolicyVariableName = 'clientPolicy' ,
    [string]$baseRuleName = 'MigratedFromCitrix' ,
    [string]$both = '**BOTH**' ,
    [int]$ruleNumber = 100 ,
    [switch]$all
)

$scriptOutput = New-Object -TypeName System.Collections.Generic.List[string]

## make the first value in the dictionary the default value to use if not explicitly set
[hashtable]$citrixEnumToRAS = @{
    'RASAdminEngine.Core.OutputModels.ClientRulesPolicy.Session.Display.SettingsColorDepth' = @{
        BitsPerPixel8  = 'Colors256'
        BitsPerPixel15 = 'HighColor15Bit'
        BitsPerPixel16 = 'HighColor16Bit'
        BitsPerPixel24 = 'TrueColor24Bit'
        BitsPerPixel32 = 'HighestQuality32Bit'
    }
    ## Citrix AutoCreateClientPrinters but ClientPrinterRedirection must be set too
    'RASAdminEngine.Core.OutputModels.ClientRulesPolicy.Session.UniversalPrinting' = @{
        AllPrinters               = 'All'
        DefaultPrinterOnly        = 'DefaultOnly'
        _NoDirectMappingAvailable = 'SpecificOnly'
    }
    ## Citrix has AllowWIARedirection & TwainRedirection whereas RAS just has -ScanTech
    <#
    'RASAdminEngine.Core.OutputModels.ClientRulesPolicy.Session.ScanningTechnologyUse' = @{
        'None'
        'WIA'
        'TWAIN'
        'WIAandTWAIN'
    }
    #>
    ## AudioQuality = Low,Medium,High
    'RASAdminEngine.Core.OutputModels.ClientRulesPolicy.Session.AudioQuality' = @{
        _NoDirectMappingAvailable = 'AdjustDynamically'
        Low    = 'UseMediumQuality'
        Medium = 'UseMediumQuality'
        High   = 'UseUncompressedQuality'
    }
    ## separate policies RestrictClientClipboardWrite and RestrictSessionClipboardWrite
    
    'RASAdminEngine.Core.ClipboardDirections' = @{
        _NoDirectMappingAvailable = 'None'
        RestrictClientClipboardWrite = 'ClientToServer' ## inverse
        RestrictSessionClipboardWrite = 'ServerToClient'
        $both = 'Bidirectional'
    }
    ## ClientClipboardWriteAllowedFormats / SessionClipboardWriteAllowedFormats set to CF_TEXT | CF_UNICODETEXT only
   
    'RASAdminEngine.Core.LimitClipboardToTextOnlyMode' = @{  
        _NoDirectMappingAvailable = 'NoLimit'
        ClientClipboardWriteAllowedFormats = 'ClientToServer'
        SessionClipboardWriteAllowedFormats = 'ServerToClient'
        $both = 'BothDirections'
    }
    #>
}

if( $outputFileProperties = Get-ItemProperty -Path $outputFile -ErrorAction SilentlyContinue )
{
    if( -Not $overWrite )
    {
        Throw "Output file `"$outputFile`" already exists ($([math]::Round( $outputFileProperties.Length / 1KB , 1))KB, last modified $($outputFileProperties.LastWriteTime.ToString())) - use -overwrite to overwrite"
    }
    if( $outputFileProperties.Attributes -match 'Directory' )
    {
        Throw "Output file `"$outputFile`" is a folder"
    }
}

if( -not [string]::IsNullOrEmpty( $preambleFile ) )
{
    if( -Not (Test-path -Path $preambleFile -PathType Leaf ) )
    {
        Throw "Preamble file `"$preambleFile`" not found"
    }
}
else ## no preamble file so put in the minimum required to function
{
    $scriptOutput.Add( '$errorActionPreference = "Stop"' )
    $scriptOutput.Add( 'Import-Module -Name RASAdmin' )
    $scriptOutput.Add( 'New-RASSession -username $env:Username' )
}

[array]$mappings = @()
$mappings = @( Import-Csv -Path $mappingFile -WarningAction SilentlyContinue )
if( $null -eq $mappings -or $mappings.Count -eq 0 )
{
    Throw "No mappings read from `"$mappingFile`""
}

[array]$canMap = @( $mappings | Where-Object 'Citrix Equivalent' )

Write-Verbose -Message "Got $($canMap.Count) / $($mappings.Count) Citrix settings that can be mapped"
$canMap | Select-Object -ExpandProperty 'Citrix Equivalent' | Sort-Object | Write-Verbose

## get the base file as that tells us the policies and whether user or computer
## other files will be base + Configuration and base + PolicyFilter

[string]$policiesListFile = $baseFileName
if( -Not ( Test-Path $baseFileName -PathType Leaf ) )
{
    $policiesListFile = Join-Path -Path $citrixPolicyFolder -ChildPath $baseFileName
    if( -Not ( Test-Path -Path $policiesListFile -PathType Leaf ) )
    {
        Throw "Unable to open file `"$policiesListFile`""
    }
}

[array]$policiesList = @()
$policiesList = @( Import-Clixml -Path $policiesListFile )

if( $policiesList.Count -eq 0 )
{
    Throw "No policies found in `"$policiesListFile`""
}

Write-Verbose -Message "Got $($policiesList.Count) policies from `"$policiesListFile`""
$policiesList|Format-Table -AutoSize|Out-String | Write-Verbose

[string]$configurationFile = $configurationFileName
if( -Not ( Test-Path $configurationFileName -PathType Leaf ) )
{
    $configurationFile = Join-Path -Path (Split-Path -Path $policiesListFile -Parent) -ChildPath $configurationFileName
    if( -Not ( Test-Path -Path $configurationFile -PathType Leaf ) )
    {
        Throw "Unable to open policies configuration file `"$configurationFile`""
    }
}

$policiesConfiguration = $null
$policiesConfiguration = @( Import-Clixml -Path $configurationFile )

if( $null -eq $policiesConfiguration -or $policiesConfiguration.Count -eq 0 )
{
    Throw "No policy configuration read from file `"$configurationFile`""
}

Write-Verbose -Message "Got $($policiesConfiguration.Count) configurations"

if( $policiesConfiguration.Count -ne $policiesList.Count )
{
    Write-Warning -Message "Expecting $($policiesList.Count) policy configurations but got $($policiesConfiguration.Count)"
}

$policiesFilters = $null
[string]$filtersFile = $filterFileName
if( -Not ( Test-Path $filterFileName -PathType Leaf ) )
{
    $filtersFile = Join-Path -Path (Split-Path -Path $policiesListFile -Parent) -ChildPath $filterFileName
    if( -Not ( Test-Path -Path $filtersFile -PathType Leaf ) )
    {
        Write-Warning "Unable to open policies filter file `"$filtersFile`""
    }
    else
    {
        $policiesFilters = @( Import-Clixml -Path $filtersFile )

        Write-Verbose -Message "Got $($policiesFilters.Count) filters"

        if( $policiesFilters.Count -eq 0 )
        {
            Write-Warning -Message "Got no filters from file `"$filtersFile`""
        }
    }
}

[int]$counter = 0
[int]$disabled = 0
[int]$processed = 0
[int]$errors = 0
[int]$notMapped = 0

$newSettings = New-Object -TypeName System.Collections.Generic.List[object]
$filtersToApply = New-Object -TypeName System.Collections.Generic.List[object]
[string]$enabledState = '$false'
if( $enabled )
{
    $enabledState = '$true'
}

## get the command so we can look up parameter types (eg string, bool, enum)
Import-Module -Name $RASPSModuleName -Verbose:$false -Debug:$false
$SetRASClientPolicyCommand = Get-Command -Name Set-RASClientPolicy -ErrorAction Stop

## Priority 1 is the highest and RAS policy priorities are top down so if we add the highest priority policies first, they will be correctly ordered
## https://docs.citrix.com/en-us/xenapp-and-xendesktop/7-15-ltsr/policies/policies-compare-model.html#prioritize-policies

ForEach( $policy in ($policiesList | Sort-Object -Property Priority ) )
{
    $counter++
    [int]$numberExamined = 0
    Write-Verbose -Message "$counter : $($policy.Type) : $($policy.PolicyName)"
   
    if( -Not $policy.Enabled -and -Not $all )
    {
        Write-Warning -Message "Not processing `"$($policy.PolicName)`" ($($policy.Type)) as disabled"
        $disabled++
        continue
    }
    $configuration = $policiesConfiguration | Where-Object { $_.PolicyName -ieq $policy.PolicyName -and $_.Type -ieq $policy.Type }
    if( $null -eq $configuration )
    {
        Write-Warning -Message "No configuration found for `"$($policy.PolicName)`" ($($policy.Type))"
        $errors++
        continue
    }
    [System.Collections.Generic.List[pscustomobject]]$policyItems = @()

    ForEach( $item in ( $configuration.PSObject.Properties | Where-Object { $_.MemberType -ieq 'NoteProperty' -and $_.value.psobject.properties[ 'state' ] -and $_.value.State -ine 'NotConfigured' } ))
    {
        if( $item.value.GetType().Name -ieq 'PSCustomObject' ) ## exclude non settings like PolicyName and Type
        {
            ## not just enabled settings - there will be allowed/prohibited too which need mapping
            if( $item.value.State -ine 'Disabled' ) ######
            {
                $numberExamined++
                $existingSetting = $null
                Write-Verbose -Message "+ Processing $($item.Name) : $($item.value.Path), state $($item.value.state)"
                $mapping = $canMap | Where-Object $citrixEquivalentColumnName -ieq $item.Name
                if( $null -ne $mapping )
                {
                    if( [string]::IsNullOrEmpty( $mapping.$RASEquivalentColumnName ) )
                    {
                        Write-Warning -Message "No RAS equivalent found for Citrix setting `"$($mapping.$citrixEquivalentColumnName)`" in policy `"$($policy.PolicyName)`""
                    }
                    else
                    {
                        [string]$parameterTypeName = $null
                        $parameterType = $SetRASClientPolicyCommand.Parameters[ ($mapping.$RASEquivalentColumnName).Trim('" -') ] ## trim because of leading - in one of the mappings
                        if( $null -ne $parameterType )
                        {
                            ## System.Nullable`1[RASAdminEngine.Core.OutputModels.ClientRulesPolicy.Session.Display.SettingsColorDepth]
                            $parameterTypeName = $parameterType.ParameterType.ToString() -replace '^System\.Nullable`1' -replace '^\[(.+)\]$' , '$1'
                        }
                        else
                        {
                            Write-Warning -Message "Failed to get type of argument to $($SetRASClientPolicyCommand.Name) for `"$($mapping.$RASEquivalentColumnName)`""
                        }
                        [string]$argumentValue = switch -Regex ($parameterTypeName)
                        {
                            '\bSystem\.String\[\]' { 
                                "`"$($item.Value.Value)`"" -join '","' } ## assumes RAS also takes array
                            '\bSystem\.String\\b' { 
                                "`"$($item.Value.Value)`""
                            }
                            '\bSystem\.Boolean\b' { 
                                "`$$($item.Value.State -ieq 'Allowed')"  ## so will be $false if prohibited
                            }
                            '\bSystem\.U?int\d+' { 
                                $item.Value.Value
                            }
                            '\bRASAdminEngine\.Core\.' {
                                $enumMapping = $citrixEnumToRAS[ $parameterTypeName ]
                                $existingSetting = $policyItems | Where-Object ParameterName -ieq $mapping.$RASEquivalentColumnName
                                if( $null -ne $enumMapping )
                                {
                                    if( $parameterTypeName -match 'LimitClipboardToTextOnly|ClipboardDirection' )
                                    {
                                        if( -Not $item.value.PSObject.Properties[ 'values' ] -or ($item.Value.Values -match 'TEXT$').Count -gt 0 )
                                        {
                                            if( $existingSetting )
                                            {
                                                $value = $enumMapping[ $both ] ## same policy can't have sanme setting in so it must be other direction meaning both are allowed
                                            }
                                            else
                                            {
                                                $value = $enumMapping[ $item.Name ]
                                            }
                                        }
                                        else ##  TEXT not in allowed format so don't allow it by getting default (=one starting with _) dictionary value
                                        {
                                            $value = $enumMapping.GetEnumerator() | Where-Object Key -match '^_' | Select-Object -ExpandProperty Value
                                        }
                                        $value ## output
                                    }
                                    else
                                    {
                                        if( $item.Value.PSObject.Properties[ 'value' ] )
                                        {
                                            $mappedValue = $enumMapping[ $item.Value.Value ]
                                        }
                                        else
                                        {
                                            $mappedValue = $enumMapping[ $item.Name ] ## item has no value as is a simple boolean enabled/disabled and we would not be here if it was disabled
                                        }
                                        if( $null -eq $mappedValue )
                                        {
                                            Write-Warning -Message "No mapping of `"$($item.Name)`" value `"$($item.value)`" of type `"$parameterTypeName`" in policy `"$($policy.PolicyName)`""
                                        }
                                        else
                                        {
                                            $mappedValue ## output
                                        }
                                    }
                                }
                                else
                                {
                                    Write-Warning -Message "No enum mapping for setting $($item.Name), type $parameterTypeName in policy `"$($policy.PolicyName)`""
                                }
                            }
                            default { Write-Warning -Message "Don't know how to deal with type $($parameterTypeName)" }
                        }
                        if( $null -ne $existingSetting ) ## we may have changed an existing value
                        {
                            if( [string]::IsNullOrEmpty( $argumentValue ) )
                            {
                                Write-Warning -Message "No new value for $parameterTypeName in $($item.Name) in policy `"$($policy.PolicyName)`""
                            }
                            else
                            {
                                $existingSetting.value = $argumentValue
                            }
                        }
                        else ## new value
                        {
                            ## TODO remove properties that we don't need
                            $toBeMapped = [pscustomobject]@{
                                ItemName = $item.Name
                                ##Mapping = $mapping
                                ParameterTypeName = $parameterTypeName
                                ##ParameterValueSource = $item.value.value
                                Value = $argumentValue
                                Quotes = $(if( $parameterTypeName -notmatch '\bSystem\.(Boolean\b|U?int\d+)' ) { '"' })
                                ParameterName = $mapping.$RASEquivalentColumnName.Trim( '" -' )
                            }
                            Write-Verbose "`t-> $($toBeMapped.ParameterName) $($toBeMapped.Value)"
                            $policyItems.Add( $toBeMapped )
                        }
                    }
                }
                else
                {
                    Write-Verbose -Message "   - not mapped"
                    $notMapped++
                }
            }
            ## else Note enabled so ignore TODO should we create it disabled?
        }
        else
        {
            ##Write-Verbose -Message "  Ignoring $($item.Name)"
        }
    }
    
    if( $null -ne $policyItems -and $policyItems.Count -gt 0 )
    {
        ## we don't need the client policy after we have done the filters so don't need a uscript unique variable name
        $scriptOutput.Add( "`$$ClientPolicyVariableName = New-RASClientPolicy -Name `"$prefix$($policy.PolicyName) ($($policy.Type))`" -Enabled $enabledState -Description `"$description`"" )
        ## quote everything since if a number it will be cast to one if required
        [string]$arguments = ( $policyItems | Select-Object -Property @{name='Argument';expression = { "-$($_.ParameterName ) $($_.Quotes)$($_.Value)$($_.Quotes)" }} | Select-Object -ExpandProperty Argument ) -join ' ' 
        
        $scriptOutput.Add( "Set-RASClientPolicy -Id `$$($ClientPolicyVariableName).Id $arguments" )

        $filters = @( $policiesFilters | Where-Object { $_.Type -ieq $policy.Type -and $_.PolicyName -ieq $policy.PolicyName } )
        if( $null -ne $filters -and $filters.Count -gt 0 )
        {
            Write-Verbose -Message " * got $($filters.Count) filters"
            [bool]$gotRASRule = $false
            $allowRule = $null
            $denyRule = $null

            ## May be more than one account in allow/deny rule so we keep a rule for each that we can add to rather than creating a rule for each account
            ForEach( $filter in $filters )
            {
                
                Write-Verbose "`tfilter - type $($filter.filtertype) mode $($filter.mode) enabled $($filter.enabled) value $($filter.FilterValue)"

                if( $filter.FilterType -imatch '^(Client|User)' )
                {
                    [string]$matchingMode = 'IsOneOfTheFollowing'
                    $ruleName = $allowRule
                    if( $filter.Mode -ieq 'deny' )
                    {
                        $matchingMode = 'IsNotOneOfTheFollowing'
                        $ruleName = $denyRule
                    }
                    if( [string]::IsNullOrEmpty( $ruleName ) ) ## we don't yet have a rule of the type we need so make it
                    {
                        [string]$ruleName = "$baseRuleName$rulenumber"
                        $ruleNumber++
                        $scriptOutput.Add( "Add-RASRule -ObjType ClientPolicy -Id `$$($ClientPolicyVariableName).Id -RuleName $ruleName -Description `"$description`"" )
                        if( $filter.Mode -ieq 'deny' )
                        {
                            $denyRule = $ruleName
                        }
                        else
                        {
                            $allowRule = $ruleName
                        }
                    }
                    ## else we already have the rule so we will get it again

                    [string]$criteria = $null
                    [string]$filterDetails = $null
                    [string]$filterCmdlet = $null
                    if( $filter.FilterType -ieq ' User' -or $filter.filterType -ieq 'ClientName' )
                    {
                        if( $filter.filterValue -match '\*' )
                        {
                            Write-Warning -Message "Wildcards are not supported in filters - `"$($filter.filterValue)`" (policy `"$($policy.PolicyName)`")"
                        }
                        else
                        {
                            $criteria = "-SecurityPrincipalsEnabled `$true -SecurityPrincipalsMatchingMode $matchingMode"
                            $filterCmdlet = "Add-RASCriteriaSecurityPrincipal"
                            $filterDetails = "-Account `"$($filter.FilterValue)`""
                        }
                    }
                    elseif( $filter.FilterType -ieq 'ClientIP' )
                    {
                        $criteria = "-IPsEnabled `$true -IPsMatchingMode $matchingMode"
                        $filterCmdlet = "Add-RASCriteriaIP"
                        [string]$IPtype = 'Version4'
                        if( $filter.filterValue -match ':' )
                        {
                            $IPtype = 'Version6'
                        }
                        ## TODO if value contains * then need to change to a range
                        [string]$filterValue = $filter.FilterValue
                        if( $filter.filterValue -match '^([^\*]+)(\.\*)+(.*)$' )
                        {
                            if( -Not [string]::IsNullOrEmpty( $Matches[3] )) ## can only support multiple wildcards
                            {
                                Write-Warning -Message "Unable to deal with IP pattern $filterValue due to characters after * (policy `"$($policy.PolicyName)`")"
                            }
                            else
                            {
                                [int]$numberOfDots = (Select-String -InputObject $filter.FilterValue -Pattern '\.' -AllMatches).Matches.Count
                                [string]$start = $filter.FilterValue -replace '\.\*' , '.0'
                                [string]$end = $filter.FilterValue -replace '\.\*' , '.255'
                                For( [int]$extra = 3 - $numberOfDots ; $extra -gt 0 ; $extra-- )
                                {
                                    $start += '.0'
                                    $end += '.255'
                                }
                                $filterValue = "$start-$end"
                            }
                        }
                        $filterDetails = "-IPType $IPtype -IP `"$filterValue`""
                    }
                    if( -Not [string]::IsNullOrEmpty( $criteria ) )
                    {
                        if( -Not $gotRASRule )
                        {
                            ## Add-RASRule doesn't return the id of the rule created or have -passthru so we need to fetch it
                            $scriptOutput.Add( "`$Rule = Get-RASrule -Id `$$($ClientPolicyVariableName).Id -ObjType ClientPolicy | Where-Object Name -eq '$ruleName'" )
                            $gotRASRule = $true
                        }

                        $scriptOutput.Add( "Set-RASCriteria -ObjType ClientPolicy -Id `$$($ClientPolicyVariableName).Id -RuleId `$Rule.Id $criteria" )
                        $scriptOutput.Add( "$filterCmdlet -ObjType ClientPolicy -Id `$$($ClientPolicyVariableName).Id -RuleId `$Rule.Id $filterDetails" )
                    }
                   
                }
                else ## TODO add IP addresses, maybe Gateways/SD-WAN
                {
                    Write-Warning -Message "Unsupported policy filter type $($filter.FilterType) for item `"$($filter.FilterValue)`" in policy `"$($policy.policyName)`""
                }
            }
        }
    }
    else
    {
        Write-Warning -Message "No mappable settings in the $numberExamined enabled in Citrix policy `"$($policy.PolicyName)`""
    }
}

Write-Verbose -Message "Processed $processed / $($policiesList.Count) with $notMapped not mapped, $disabled disabled skipped and $errors errors & got $($newSettings.Count) new settings"

$scriptOutput.Add( 'Remove-RASSession' )
## TODO could add a success message

if( $null -ne $outputFileProperties ) ## don't remove earlier in case script aborts so we only do it if there is new content to write
{
    Remove-Item -Path $outputFile
    if( -not $? )
    {
        Throw "Failed to remove existing output file `"$outputFile`""
    }
}

if( -Not [string]::IsNullOrEmpty( $preambleFile ) )
{
    Copy-Item -Path $preambleFile -Destination $outputFile
    if( -Not $? )
    {
        Throw "Failed to copy preamble file `"$preambleFile`""
    }
}

Write-Verbose -Message "Writing $($scriptOutput.Count) lines to $outputFile"

$( ForEach( $line in $scriptOutput )
{
    For( [int]$blankLine = 0 ; $blankLine -lt $blankLines ; $blankLine++ )
    {
        ''
    }
    $line
} ) | Add-Content -Path $outputFile

## TODO Thoughts 
## 
##  Can we merge policies with Computer and User sections into one so we can keep the ordering simple by sorting the input plocies in descending order so we can create in that order nd not have to change order??