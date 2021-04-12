#requires -version 5
<#
.SYNOPSIS
    This module defines a base PSWindow, with all basic required functions defined.
    The PSWindow is created using the .xaml resource files.
.DESCRIPTION
  import this module and inherit PSWindow class.
.INPUTS
.OUTPUTS
    PSWindow
    !IMPORTANT - classes can only be exported with the "using module" syntax. Import-Module does not
    import classes.
.NOTES
  Version:        1.0
  Purpose:        Showcase powershell capabilities with RASAdmin
  
.EXAMPLE
  using module ./PSWindow.psm1
#>
using namespace System.Xml
using namespace System.Windows
using namespace System.Windows.Markup
using namespace System.Windows.Controls
using namespace System.Collections.ObjectModel
using namespace System.IO

<#
    Any powershell UIs to be created inherit from this base class.
    This class takes care of auto-generating class properties from a given
    XAML document.
#>
class PSWindow {
    [Window] $Form

    PSWindow() {}

    PSWindow([string] $pathToXAML) {
        if ($this.Initialize($pathToXAML) -eq $false) {
            throw "Failed to initialize"
        }
    }

    <#
        reads XAML document and auto-generates class properties from
        the element's 'Name' attribute. This means that any child classes
        would allow for $this.txtElementName.Text syntax.
    #>
    hidden [bool] Initialize([string]$pathToXAML) {
        try {
            [string]$tmp = $this.Load($pathToXAML)
            if ([string]::IsNullOrEmpty($tmp)) {
                return $false
            }
            [xml]$xaml = $tmp
            [XmlNodeReader]$reader = [XmlNodeReader]::new($xaml)
            $this.Form = [XamlReader]::Load($reader)
    
            $xaml.SelectNodes("//*[@Name]") | % { 
                    Add-Member -InputObject $this -MemberType NoteProperty -Name ($_.Name) -Value $this.Form.FindName($_.Name)
            }
        }
        catch [Exception]{
            Write-Host "Failed to initialize" -ForegroundColor red
            Write-Host $_.Exception.Message -ForegroundColor red
            return $false
        }
        return $true
    }

    hidden [string] Load($pathToXAML) {
        if ($this.checkPath($pathToXAML) -ne $true) {
            Write-Host "File $($pathToXAML) does not exist." -ForegroundColor red
            return $null
        }

        [string]$content = Get-Content -Path $pathToXAML
        return $content  -replace 'x:Class.".*"$', '' -replace 'mc:Ignorable="d"','' -replace "x:N",'N'
    }

    [void] Close() {
        $this.Form.Close()
    }
    
    [void] Show() {
        $this.Form.ShowDialog()        
    }

    hidden [bool] checkPath([string] $path) {
        return (Test-Path $path)
    }
}
Write-Host 'RASAdmin module Imported' -ForegroundColor Green
