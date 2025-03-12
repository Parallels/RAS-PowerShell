<#
.SYNOPSIS
    Create policies exported from Citrix CVAD/DaaS into Parallels RAS

.PARAMETERS

.EXAMPLE

.NOTES

    Modification History:
#>

[CmdletBinding()]

Param
(
    [string]$connectionBroker ,
    [string]$username = $env:username
)

## if -ErrorAction not specifically passed then set it to stop so that will not continue on error
if( -Not $PSBoundParameters[ 'ErrorAction' ] )
{
    $ErrorActionPreference = 'Stop'
}

Import-Module -Name RASAdmin

[hashtable]$connectionParameters = @{
    Username = $username
}

if( -not [string]::IsNullOrEmpty( $connectionBroker ) -and $connectionBroker -ne '.' -and $connectionBroker -ine 'localhost' -and $connectionBroker -ine $env:COMPUTERNAME )
{
    $connectionParameters.Add( 'Server' , $connectionBroker )
}

New-RASSession @connectionParameters
