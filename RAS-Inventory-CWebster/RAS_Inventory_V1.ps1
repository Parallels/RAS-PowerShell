#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

<#
.SYNOPSIS
	Creates a complete inventory of Parallels Remote Application Server.
.DESCRIPTION
	Creates a complete inventory of Parallels Remote Application Server (RAS) using 
	Microsoft PowerShell, Word, plain text, or HTML.
	
	The script requires at least PowerShell version 3 but runs best in version 5.

	Word is NOT needed to run the script. This script will output in Text and HTML.
	The default output format is HTML.
	
	Creates an output file named RASInventory.<fileextension>.
	
	You do NOT have to run this script on a server running RAS. This script was developed 
	and run from a Windows 10 VM.

	Word and PDF document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

.PARAMETER HTML
	Creates an HTML file with an .html extension.
	
	HTML is now the default report format.
	
	This parameter is set True if no other output format is selected.
.PARAMETER MSWord
	SaveAs DOCX file
	
	Microsoft Word is no longer the default report format.
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	
	This parameter requires Microsoft Word to be installed.
	This parameter uses Word's SaveAs PDF capability.

	This parameter is disabled by default.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date timestamp to the end of the file name.
	
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2019 at 6PM is 2021-06-01_1800.
	
	Output filename will be ReportName_2021-06-01_1800.<ext>.
	
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field. 
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field. 
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name to use for the Cover Page. 
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field. 
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)

	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)

	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER From
	Specifies the username for the From email address.
	
	If SmtpServer or To are used, this is a required parameter.
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER ServerName
	Specifies which RAS server to use to run the script against.
	
	ServerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	
	If entered as localhost, the actual computer name is determined and used.
	
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	Default value is LocalHost
.PARAMETER SmtpPort
	Specifies the SMTP port for the SmtpServer. 
	The default is 25.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report(s). 
	
	If From or To are used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	
	If SmtpServer or From are used, this is a required parameter.
.PARAMETER User
	Username to use for the connection to the RAS server.

	Default value is contained in $env:username
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.EXAMPLE
	PS C:\PSScript > .\RAS_Inventory_V1.ps1
	
	Outputs, by default to HTML.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript .\RAS_Inventory_V1.ps1 -MSWord -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster" -ComputerName RAS01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		RAS server named RAS01 for the ComputerName.

	Outputs to Microsoft Word.
	Prompts for credentials for the RAS Server RAS01.
.EXAMPLE
	PS C:\PSScript .\RAS_Inventory_V1.ps1 -PDF -CN "Carl Webster Consulting" -CP "Mod" 
	-UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).

	Outputs to PDF.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript .\RAS_Inventory_V1.ps1 -CompanyName "Sherlock Holmes 
	Consulting"
	-CoverPage Exposure 
	-UserName "Dr. Watson"
	-CompanyAddress "221B Baker Street, London, England"
	-CompanyFax "+44 1753 276600"
	-CompanyPhone "+44 1753 276200"
	-MSWord
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.

	Outputs to Microsoft Word.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript .\RAS_Inventory_V1.ps1 -CompanyName "Sherlock Holmes 
	Consulting"
	-CoverPage Facet 
	-UserName "Dr. Watson"
	-CompanyEmail SuperSleuth@SherlockHolmes.com
	-PDF

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.

	Outputs to PDF.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript > .\RAS_Inventory_V1.ps1 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	
	-Text

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.

	Outputs to a text file.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript > .\RAS_Inventory_V1.ps1 
	-SmtpServer mailrelay.domain.tld
	-From Anonymous@domain.tld 
	-To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and will not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script will generate an anonymous secure password for the anonymous@domain.tld 
	account.

	Outputs, by default, to HTML.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript > .\RAS_Inventory_V1.ps1 
	-SmtpServer labaddomain-com.mail.protection.outlook.com
	-UseSSL
	-From SomeEmailAddress@labaddomain.com 
	-To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.

	Outputs, by default, to HTML.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript > .\RAS_Inventory_V1.ps1 
	-SmtpServer smtp.office365.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.

	Outputs, by default, to HTML.
	Prompts for credentials for the LocalHost RAS Server.
.EXAMPLE
	PS C:\PSScript > .\RAS_Inventory_V1.ps1 
	-SmtpServer smtp.gmail.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.

	Outputs, by default, to HTML.
	Prompts for credentials for the LocalHost RAS Server.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script. This script creates a Word, PDF, HTML, or plain 
	text document.
.NOTES
	NAME: RAS_Inventory_V1.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster
	LASTEDIT: August 4, 2020
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$ServerName="LocalHost",
	
	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[string]$User=$env:username,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(Mandatory=$False)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False
	
	)

	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on February 9, 2018

#Version 1.0 released to the community on 5-August-2020
#

Set-StrictMode -Version Latest

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$Script:emailCredentials = $Null

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "
			`n`n
			`t`t
			Folder $Folder is a file, not a folder.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			AbortScript
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
		`t`t
		Folder $Folder does not exist.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\RASDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdpath\RASInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
#endregion

#region initialize variables for Word, HTML, and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	#[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	#[int]$wdColorRed = 255
	#[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	#[int]$wdAlignParagraphLeft = 0
	#[int]$wdAlignParagraphCenter = 1
	#[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	#[int]$wdCellAlignVerticalTop = 0
	#[int]$wdCellAlignVerticalCenter = 1
	#[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	#[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	#[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	#[int]$wdAdjustFirstColumn = 2
	#[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	#[int]$Indent1TabStops = 1 * $PointsPerTabStop
	#[int]$Indent2TabStops = 2 * $PointsPerTabStop
	#[int]$Indent3TabStops = 3 * $PointsPerTabStop
	#[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155
	#[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	#[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
    $Script:htmlredmask       = "#FF0000" 4>$Null
    $Script:htmlcyanmask      = "#00FFFF" 4>$Null
    $Script:htmlbluemask      = "#0000FF" 4>$Null
    $Script:htmldarkbluemask  = "#0000A0" 4>$Null
    $Script:htmllightbluemask = "#ADD8E6" 4>$Null
    $Script:htmlpurplemask    = "#800080" 4>$Null
    $Script:htmlyellowmask    = "#FFFF00" 4>$Null
    $Script:htmllimemask      = "#00FF00" 4>$Null
    $Script:htmlmagentamask   = "#FF00FF" 4>$Null
    $Script:htmlwhitemask     = "#FFFFFF" 4>$Null
    $Script:htmlsilvermask    = "#C0C0C0" 4>$Null
    $Script:htmlgraymask      = "#808080" 4>$Null
    $Script:htmlblackmask     = "#000000" 4>$Null
    $Script:htmlorangemask    = "#FFA500" 4>$Null
    $Script:htmlmaroonmask    = "#800000" 4>$Null
    $Script:htmlgreenmask     = "#008000" 4>$Null
    $Script:htmlolivemask     = "#808000" 4>$Null

    $Script:htmlbold        = 1 4>$Null
    $Script:htmlitalics     = 2 4>$Null
    $Script:htmlred         = 4 4>$Null
    $Script:htmlcyan        = 8 4>$Null
    $Script:htmlblue        = 16 4>$Null
    $Script:htmldarkblue    = 32 4>$Null
    $Script:htmllightblue   = 64 4>$Null
    $Script:htmlpurple      = 128 4>$Null
    $Script:htmlyellow      = 256 4>$Null
    $Script:htmllime        = 512 4>$Null
    $Script:htmlmagenta     = 1024 4>$Null
    $Script:htmlwhite       = 2048 4>$Null
    $Script:htmlsilver      = 4096 4>$Null
    $Script:htmlgray        = 8192 4>$Null
    $Script:htmlolive       = 16384 4>$Null
    $Script:htmlorange      = 32768 4>$Null
    $Script:htmlmaroon      = 65536 4>$Null
    $Script:htmlgreen       = 131072 4>$Null
	$Script:htmlblack       = 262144 4>$Null

	$Script:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$Script:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break }
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_}	{$CultureCode = "ca-"}
		{$ChineseArray -contains $_}	{$CultureCode = "zh-"}
		{$DanishArray -contains $_}		{$CultureCode = "da-"}
		{$DutchArray -contains $_}		{$CultureCode = "nl-"}
		{$EnglishArray -contains $_}	{$CultureCode = "en-"}
		{$FinnishArray -contains $_}	{$CultureCode = "fi-"}
		{$FrenchArray -contains $_}		{$CultureCode = "fr-"}
		{$GermanArray -contains $_}		{$CultureCode = "de-"}
		{$NorwegianArray -contains $_}	{$CultureCode = "nb-"}
		{$PortugueseArray -contains $_}	{$CultureCode = "pt-"}
		{$SpanishArray -contains $_}	{$CultureCode = "es-"}
		{$SwedishArray -contains $_}	{$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		AbortScript
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	#fixed by MBS
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		AbortScript
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module | ForEach-Object { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	
	[string]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If($ModuleFound -ne $ModuleName) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0 4>$Null
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		Write-Error "
		`n`n
		`t`t
		The Word object could not be created. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		Write-Error "
		`n`n
		`t`t
		Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		Write-Error "
		`n`n
		`t`t
		Microsoft Word 2007 is no longer supported.
		`n`n
		`t`t
		Script will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
		`t`t
		The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
	Else
	{
		Write-Error "
		`n`n
		`t`t
		You are running an untested or unsupported version of Microsoft Word.
		`n`n
		`t`t
		Script will end.
		`n`n
		`t`t
		Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "
			`n`n
			`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "
			`n
			`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "
			`n
			`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			`n`n"
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
		`t`t
		For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		Write-Error "
		`n`n
		`t`t
		An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		Write-Error "
		`n`n
		`t`t
		An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object{$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue2
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey = $Reg.OpenSubKey($path2)
		If ($RegKey)
		{
			$Results = $RegKey.GetValue($name)

			If($Null -ne $Results)
			{
				Return $Results
			}
			Else
			{
				Return $Null
			}
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		$null = $Script:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		$null = $Script:Output.Append( $name + $value )
	}
	Else
	{
		$null = $Script:Output.AppendLine( $name + $value )
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML. They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used. Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

# to suppress $crlf in HTML documents, replace this with '' (empty string)
# but this was added to make the HTML readable
$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
# re-implemented/re-based by Michael B. Smith
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	#FIXME - long story short, this function was wrong and had been wrong for a long time. 
	## The function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites! If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		## $HTMLBody = '<p></p>'
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold
		if( $ital ) { $null = $sb.Append( '<i>' ) }
		if( $bold ) { $null = $sb.Append( '<b>' ) } 

		switch( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		$null = $sb.Append( $HTMLOpen )

		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		if( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		else                    { $null = $sb.Append( $HTMLClose ) }

		if( $ital ) { $null = $sb.Append( '</i>' ) }
		if( $bold ) { $null = $sb.Append( '</b>' ) } 

		if( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	if( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
		$rowCount = $rowInfo.Length
	}

	for( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		if( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
		}

		$subRowLength = $subRow.Length
		for( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = if( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } else { 0 }

			$text   = if( $item ) { $item.ToString() } else { '' }
			$format = if( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $Script:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics

			if( $null -eq $fixedInfo -or $fixedInfo.Length -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}
			else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}

			if( $bold ) { $null = $sb.Append( '<b>' ) }
			if( $ital ) { $null = $sb.Append( '<i>' ) }

			if( $text -eq ' ' -or $text.length -eq 0)
			{
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			else
			{
				for ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					if( $text[ $inx ] -eq ' ' )
					{
						$null = $sb.Append( '&nbsp;' )
					}
					else
					{
						break
					}
				}
				$null = $sb.Append( $text )
			}

			if( $bold ) { $null = $sb.Append( '</b>' ) }
			if( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
		}

		$null = $sb.AppendLine( '</tr>' )
	}

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith
#***********************************************************************************************************

<#
.Synopsis
	Format table for a HTML output document.
.DESCRIPTION
	This function formats a table for HTML from multiple arrays of strings.
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). Otherwise the table will be generated
	with a border (border = '1').
.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a separate array containing column headers
	(columnArray is not specified). Set this parameter equal to the number of columns in the table.
.PARAMETER rowArray
	This parameter contains the row data array for the table.
.PARAMETER columnArray
	This parameter contains column header data for the table.
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")
.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file. All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column. You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column. Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data. Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array. If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics. For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below. As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('FileName',$htmlsb,$Script:FileName,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -ne $rowArray )
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	if( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		for( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $Script:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			if( $null -eq $fixedWidth -or $fixedWidth.Length -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			if( $bold ) { $HTMLBody += '<b>' }
			if( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			if( $array )
			{
				if( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				else
				{
					for( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						if( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			if( $bold ) { $HTMLBody += '</b>' }
			if( $ital ) { $HTMLBody += '</i>' }
		}

		$HTMLBody += '</font></td>'
		$HTMLBody += $crlf
	}

	$HTMLBody += '</tr>' + $crlf

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	If( $rowArray )
	{

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HtmlFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HtmlFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:HtmlFileName -Force -InputObject $HTMLHead 4>$Null
}#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -eq $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$True, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $True; }
				If($Italic) { $Cell.Range.Font.Italic = $True; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$True, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$True, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	Remove-Variable -Name word -Scope Script 4>$Null
	Remove-Variable -Name Doc  -Scope Script 4>$Null
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = (Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}
	If($null -ne $wordprocess -and $wordprocess.Id -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess.Id)"
		Stop-Process $wordprocess.Id -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date): Setting up Text"
	[System.Text.StringBuilder] $Script:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date): Saving Text file"
	Write-Output $Script:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date): Saving HTML file"
	Out-File -FilePath $Script:HtmlFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function ProcessDocumentOutput
{
	Param([string] $Condition)
	
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	If($Condition -eq "Regular")
	{
		$GotFile = $False

		If($MSWord)
		{
			If(Test-Path "$($Script:WordFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:WordFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:WordFileName)"
				Write-Error "Unable to save the output file, $($Script:WordFileName)"
			}
		}
		If($PDF)
		{
			If(Test-Path "$($Script:PDFFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:PDFFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:PDFFileName)"
				Write-Error "Unable to save the output file, $($Script:PDFFileName)"
			}
		}
		If($Text)
		{
			If(Test-Path "$($Script:TextFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:TextFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:TextFileName)"
				Write-Error "Unable to save the output file, $($Script:TextFileName)"
			}
		}
		If($HTML)
		{
			If(Test-Path "$($Script:HTMLFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:HTMLFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:HTMLFileName)"
				Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
			}
		}
		
		#email output file if requested
		If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			If($MSWord)
			{
				$emailAttachment = $Script:WordFileName
				SendEmail $emailAttachment
			}
			If($PDF)
			{
				$emailAttachment = $Script:PDFFileName
				SendEmail $emailAttachment
			}
			If($Text)
			{
				$emailAttachment = $Script:TextFileName
				SendEmail $emailAttachment
			}
			If($HTML)
			{
				$emailAttachment = $Script:HTMLFileName
				SendEmail $emailAttachment
			}
		}
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Add DateTime         : $($AddDateTime)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name         : $($Script:CoName)"
		Write-Verbose "$(Get-Date): Company Address      : $($CompanyAddress)"
		Write-Verbose "$(Get-Date): Company Email        : $($CompanyEmail)"
		Write-Verbose "$(Get-Date): Company Fax          : $($CompanyFax)"
		Write-Verbose "$(Get-Date): Company Phone        : $($CompanyPhone)"
		Write-Verbose "$(Get-Date): Cover Page           : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): Dev                  : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile         : $($Script:DevErrorFile)"
	}
	If($MSWord)
	{
		Write-Verbose "$(Get-Date): Word FileName        : $($Script:WordFileName)"
	}
	If($HTML)
	{
		Write-Verbose "$(Get-Date): HTML FileName        : $($Script:HtmlFileName)"
	} 
	If($PDF)
	{
		Write-Verbose "$(Get-Date): PDF FileName         : $($Script:PDFFileName)"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date): Text FileName        : $($Script:TextFileName)"
	}
	Write-Verbose "$(Get-Date): Folder               : $($Folder)"
	Write-Verbose "$(Get-Date): From                 : $($From)"
	Write-Verbose "$(Get-Date): Log                  : $($Log)"
	Write-Verbose "$(Get-Date): RAS Version          : $($Script:RASVersion)"
	Write-Verbose "$(Get-Date): Save As HTML         : $($HTML)"
	Write-Verbose "$(Get-Date): Save As PDF          : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT         : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD         : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo           : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port            : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server          : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title                : $($Script:Title)"
	Write-Verbose "$(Get-Date): To                   : $($To)"
	Write-Verbose "$(Get-Date): Use SSL              : $($UseSSL)"
	Write-Verbose "$(Get-Date): User                 : $($Script:User)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name            : $($UserName)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected          : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version         : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture            : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture          : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language        : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version         : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start         : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		If(Test-Path variable:script:word)
		{
			$Script:Word.quit()
			Write-Verbose "$(Get-Date): System Cleanup"
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
			Remove-Variable -Name word -Scope Script 4>$Null
		}
		#is the winword process still running? kill it

		#find out our session (usually "1" except on TS/RDC or Citrix)
		$SessionID = (Get-Process -PID $PID).SessionId

		#Find out if winword is running in our session
		$wordprocess = $Null
		$wordprocess = (Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}
		If($null -ne $wordprocess -and $wordprocess.Id -gt 0)
		{
			Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess.Id)"
			Stop-Process $wordprocess.Id -EA 0
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region script start function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	#make sure PSAdmin module is loaded
	If(!(Check-LoadedModule "PSAdmin"))
	{
		Write-Error "
		`n`n
		`t`t
		The PSAdmin module could not be loaded.
		`n`n
		`t`t
		Please see the Prerequisites section in the ReadMe file (RAS_Inventory_V1_ReadMe.rtf).
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Write-Verbose "$(Get-Date): "
		AbortScript
	}

	#if computer name is localhost, get actual server name
	If($Script:ServerName -eq "localhost")
	{
		$Script:ServerName = $env:ComputerName
		Write-Verbose "$(Get-Date): Server name has been changed from localhost to $($Script:ServerName)"
	}
	
	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $Script:ServerName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$Script:ServerName = $Result.HostName
			Write-Verbose "$(Get-Date): Server name has been changed from $ip to $Script:ServerName"
		}
		Else
		{
			Write-Warning "Unable to resolve $Script:ServerName to a hostname"
		}
	}
	Else
	{
		#server is online but for some reason $Script:ServerName cannot be converted to a System.Net.IpAddress
	}

	If(![String]::IsNullOrEmpty($Script:ServerName)) 
	{
		#get server name
		#first test to make sure the server is reachable
		Write-Verbose "$(Get-Date): Testing to see if $Script:ServerName is online and reachable"
		If(Test-Connection -ComputerName $Script:ServerName -quiet -EA 0)
		{
			Write-Verbose "$(Get-Date): Server $Script:ServerName is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Server $Script:ServerName is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $Script:ServerName is either offline or is not a valid computer name.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			AbortScript
		}
	}

	#attempt to connect to the RAS server
	
	$creds = Get-Credential -UserName $Script:User -message "Enter credentials to connect to $Script:ServerName"

	Write-Verbose "$(Get-Date): Attempting connection to $Script:ServerName as $($creds.UserName)"
	$PSDefaultParameterValues = @{"*:Verbose"=$False}
	
	$Results = New-RASSession -username $creds.UserName -Password $creds.Password -Server $Script:ServerName -EA 0 *>$Null
	
	If(!($?))
	{
		$PSDefaultParameterValues = @{"*:Verbose"=$True}
		Write-Error "
		`n`n
		`t`t
		Unable to connect to Parallels RAS server $Script:ServerName.`
		`n`n
		`t`t
		Please rerun the script with the correct RAS server name, user name, or password.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
	Else
	{
		$PSDefaultParameterValues = @{"*:Verbose"=$True}
		$Script:User = $creds.UserName
		Write-Verbose "$(Get-Date): Successfully connected to $Script:ServerName as $Script:User"
	}

	Write-Verbose "$(Get-Date): Get RAS Version"
	$Results = Get-RASVersion -EA 0 4>$Null
	
	If($? -and $Null -ne $Results)
	{
		[double]$Script:RASVersion = $Results.SubString(0,4)
	}
	Else
	{
		$Script:RASVersion = 0
	}
	Write-Verbose "$(Get-Date): Running RAS Version $($Script:RASVersion)"
	$Script:Title = "Parallels RAS Inventory"
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\RASInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime         : $($AddDateTime)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name         : $($Script:CoName)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address      : $($CompanyAddress)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email        : $($CompanyEmail)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax          : $($CompanyFax)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone        : $($CompanyPhone)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page           : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev                  : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile         : $($Script:DevErrorFile)" 4>$Null
		}
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word FileName        : $($Script:WordFileName)" 4>$Null
		}
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTML FileName        : $($Script:HtmlFileName)" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDF Filename         : $($Script:PDFFileName)" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Text FileName        : $($Script:TextFileName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder               : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From                 : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                  : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "RAS Version          : $($Script:RASVersion)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML         : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF          : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT         : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD         : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info          : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Server               : $($Script:ServerName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port            : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server          : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title                : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                   : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL              : $($UseSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "User                 : $($Script:User)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name            : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected          : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version         : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture            : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture          : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language        : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version         : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start         : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time         : $($Str)" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}
	
	#cleanup obj variables
	$Script:Output = $Null

	If($Script:RASVersion -gt 17.1)
	{
		Write-Host "$(Get-Date):                                                         " 
		Write-Host "$(Get-Date):       This script was created on RAS Version 17.1       " 
		Write-Host "$(Get-Date):            You are running RAS Version $($Script:RASVersion)             " 
		Write-Host "$(Get-Date):Any new product features will not be in the output report" 
		Write-Host "$(Get-Date):                                                         " 
	}

	Write-Verbose "$(Get-Date):                                                                                    " 
	Write-Verbose "$(Get-Date):               This FREE script was brought to you by Conversant Group              " 
	Write-Verbose "$(Get-Date):We design, build, and manage infrastructure for a secure, dependable user experience" 
	Write-Verbose "$(Get-Date):                       Visit our website conversantgroup.com                        " 
	Write-Verbose "$(Get-Date):                                                                                    " 
}
#endregion

#region process farm
Function ProcessFarm
{
	Write-Verbose "$(Get-Date): Processing Farm"
	
	$Results = Get-RASFarmSettings -EA 0

	If(!$?)
	{
		Write-Error "
		`n`n
		`t`t
		Unable to retrieve RAS Farm Settings for Parallels RAS server $Script:ServerName.`
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
	ElseIf($? -and $Null -eq $Results)
	{
		Write-Error "
		`n`n
		`t`t
		No RAS Farm Settings retrieved for Parallels RAS server $Script:ServerName.`
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
	Else
	{
		$Script:RASFarmName = $Results.Name
		
		OutputFarm
	}
	
	ProcessFarmSites
	ProcessSites
}

Function OutputFarm
{
	Write-Verbose "$(Get-Date): `tOutput Farm"
	
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Farm - " $Script:RASFarmName
	}
	If($Text)
	{
		Line 0 "Farm - " $Script:RASFarmName
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Farm - " $Script:RASFarmName
	}
}

Function ProcessFarmSites
{
	Write-Verbose "$(Get-Date): `tProcessing Farm Sites"
	$Sites = Get-Site -EA 0 4> $Null

	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Sites for Farm $Script:RASFarmName`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Sites for Farm $Script:RASFarmName"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Sites for Farm $Script:RASFarmName"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Sites for Farm $Script:RASFarmName"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Sites were retrieved for Farm $Script:RASFarmName.`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Sites were retrieved for Farm $Script:RASFarmName"
		}
		If($Text)
		{
			Line 0 "No Sites were retrieved for Farm $Script:RASFarmName"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Sites were retrieved for Farm $Script:RASFarmName"
		}
	}
	Else
	{
		ForEach($Site in $Sites)
		{
			OutputFarmSite $Site
		}
	}
}

Function OutputFarmSite
{
	Param([object]$Site)
	
	Write-Verbose "$(Get-Date): `t`tOutput Farm Site $($Site.Name)"
	$SiteSettings = Get-SiteStatus -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Site Status for Site $Site.Name`
		`n`n
		"
		$MasterPublishingAgent = "N/A"
		$Type = "N/A"
		$State = "N/A"
		$Description = "Can't find"
		$ID = $Site.Id
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Site Status for Site $Site.Name"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Site Status for Site $Site.Name"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Site Status for Site $Site.Name"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Site Status retrieved for Site $Site.Name.`
		`n`n
		"
		$MasterPublishingAgent = "N/A"
		If($Site.LicensingSite)
		{
			$Type = "Licensing Site/Local Site/Current Site"
		}
		Else
		{
			$Type = "Secondary Site"
		}
		$State = "N/A"
		$Description = "Can't find"
		$ID = $Site.Id
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Site Status retrieved for Site $Site.Name"
		}
		If($Text)
		{
			Line 0 "No Site Status retrieved for Site $Site.Name"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Site Status retrieved for Site $Site.Name"
		}
	}
	Else
	{
		$Results = Get-PA -SiteId $Site.Id -EA 0 4> $Null
		
		If($? -and $Null -ne $Results)
		{
			$Description = $Results.Description
		}
		Else
		{
			$Description = ""
		}
		
		$MasterPublishingAgent = $SiteSettings.Server
		If($Site.LicensingSite)
		{
			$Type = "Licensing Site/Local Site/Current Site"
		}
		Else
		{
			$Type = "Secondary Site"
		}
		$State = $SiteSettings.AgentState
		$ID = $Site.Id
	}
	
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Site"; Value = $Site.Name; }) > $Null
		$ScriptInformation.Add(@{Data = "Master Publishing Agent"; Value = $MasterPublishingAgent; }) > $Null
		$ScriptInformation.Add(@{Data = "Type"; Value = $Type; }) > $Null
		$ScriptInformation.Add(@{Data = "State"; Value = $State; }) > $Null
		$ScriptInformation.Add(@{Data = "Description"; Value = $Description; }) > $Null
		$ScriptInformation.Add(@{Data = "ID"; Value = $ID; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "Site`t`t`t: " $Site.Name
		Line 1 "Master Publishing Agent`t: " $MasterPublishingAgent
		Line 1 "Type`t`t`t: " $Type
		Line 1 "State`t`t`t: " $State
		Line 1 "Description`t`t: " $Description
		Line 1 "ID`t`t`t: " $ID
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Site",($Script:htmlsb),$Site.Name,$htmlwhite)
		$rowdata += @(,( "Master Publishing Agent",($Script:htmlsb),$MasterPublishingAgent,$htmlwhite))
		$rowdata += @(,( "Type",($Script:htmlsb),$Type,$htmlwhite))
		$rowdata += @(,( "State",($Script:htmlsb),$State.ToString(),$htmlwhite))
		$rowdata += @(,( "Description",($Script:htmlsb),$Description,$htmlwhite))
		$rowdata += @(,( "ID",($Script:htmlsb),$ID,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","275")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function ProcessSites
{
	Write-Verbose "$(Get-Date): `tProcessing Sites"
	$Sites = Get-Site -EA 0 4> $Null

	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Sites`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Sites"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Sites"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Sites"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Sites were retrieved.`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Sites were retrieved"
		}
		If($Text)
		{
			Line 0 "No Sites were retrieved"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Sites were retrieved"
		}
	}
	Else
	{
		ForEach($Site in $Sites)
		{
			OutputSite $Site
		}
	}
}

Function GetVDIType
{
	Param([string] $VDIHostType)
	
	Switch ($VDIHostType)
	{
		"HyperVWin2008Std"					{$VDIType = "HyperV on Windows Server 2008 Standard Edition"; Break}
		"HyperVWin2008Ent"					{$VDIType = "HyperV on Windows Server 2008 Enterprise Edition"; Break}
		"HyperVWin2008Dtc"					{$VDIType = "HyperV on Windows Server 2008 Datacenter Edition"; Break}
		"HyperV"							{$VDIType = "HyperV"; Break}
		"HyperVWin2012Std"					{$VDIType = "HyperV on Windows Server 2012 Datacenter Edition"; Break}
		"HyperVWin2012Dtc"					{$VDIType = "HyperV on Windows Server 2012 Datacenter Edition"; Break}
		"HyperVWin2012Srv"					{$VDIType = "HyperV on Windows Server 2012"; Break}
		"HyperVWin2008R2Std"				{$VDIType = "HyperV on Windows Server 2008 R2 Standard Edition"; Break}
		"HyperVWin2008R2Ent"				{$VDIType = "HyperV on Windows Server 2008 R2 Enterprise Edition"; Break}
		"HyperVWin2008R2Dtc"				{$VDIType = "HyperV on Windows Server 2008 R2 Datacenter Edition"; Break}
		"CitrixXenUnknown"					{$VDIType = "Citrix XenServer"; Break}
		"CitrixXen5_0"						{$VDIType = "Citrix XenServer 5.0"; Break}
		"CitrixXen5_5"						{$VDIType = "Citrix XenServer 5.5"; Break}
		"CitrixXen5_6"						{$VDIType = "Citrix XenServer 5.6"; Break}
		"CitrixXen5_6_1"					{$VDIType = "Citrix XenServer 5.6.1"; Break}
		"CitrixXen6_0"						{$VDIType = "Citrix XenServer 6.0"; Break}
		"CitrixXen6_1"						{$VDIType = "Citrix XenServer 6.1"; Break}
		"CitrixXen6_2"						{$VDIType = "Citrix XenServer 6.2"; Break}
		"CitrixXen6_5"						{$VDIType = "Citrix XenServer 6.5"; Break}
		"CitrixXen7_0"						{$VDIType = "Citrix XenServer 7.0"; Break}
		"CitrixXen7_1"						{$VDIType = "Citrix XenServer 7.1"; Break}
		"CitrixXen7_2"						{$VDIType = "Citrix XenServer 7.2"; Break}
		"QemuKvmUnknown"					{$VDIType = "QEMU KVM unknown"; Break}
		"QemuKvm1_2_14"						{$VDIType = "QEMU KVM 1.2.14"; Break}
		"HyperVUnknown"						{$VDIType = "HyperV on Unknown Server"; Break}
		"HyperVWin2012R2Std"				{$VDIType = "HyperV on Windows Server 2012 Standard Edition"; Break}
		"HyperVWin2012R2Dtc"				{$VDIType = "HyperV on Windows Server 2012 R2 Datacenter Edition"; Break}
		"HyperVWin2012R2Srv"				{$VDIType = "HyperV on Windows Server 2012 R2"; Break}
		"HyperVWin2016Std"					{$VDIType = "HyperV on Windows Server 2016 Standard Edition"; Break}
		"HyperVWin2016Dtc"					{$VDIType = "HyperV on Windows Server 2016 Datacenter Edition"; Break}
		"HyperVWin2016Srv"					{$VDIType = "HyperV on Windows Server 2016"; Break}
		"HyperVWin2019Std"					{$VDIType = "HyperV on Windows Server 2019 Standard Edition"; Break}
		"HyperVWin2019Dtc"					{$VDIType = "HyperV on Windows Server 2019 Datacenter Edition"; Break}
		"HyperVWin2019Srv"					{$VDIType = "HyperV on Windows Server 2019"; Break}
		"HyperVFailoverClusterUnknown"		{$VDIType = "HyperV Failover Cluster on Unknown Server"; Break}
		"HyperVFailoverClusterEnt"			{$VDIType = "HyperV Failover Cluster Enterprise Edition"; Break}
		"HyperVFailoverClusterDtc"			{$VDIType = "HyperV Failover Cluster Datacenter Edition"; Break}
		"HyperVFailoverClusterWin2012"		{$VDIType = "HyperV Failover Cluster on Windows Server 2012"; Break}
		"HyperVFailoverClusterWin2012R2"	{$VDIType = "HyperV Failover Cluster on Windows Server 2012 R2"; Break}
		"HyperVFailoverClusterWin2016"		{$VDIType = "HyperV Failover Cluster on Windows Server 2016"; Break}
		"HyperVFailoverClusterWin2019"		{$VDIType = "HyperV Failover Cluster on Windows Server 2019"; Break}
		"VmwareESXUnknown"					{$VDIType = "Vmware ESXi"; Break}
		"VmwareESXi4_0"						{$VDIType = "Vmware ESXi 4.0"; Break}
		"VmwareESX4_0"						{$VDIType = "Vmware ESX 4.0"; Break}
		"VmwareESXi4_1"						{$VDIType = "Vmware ESXi 4.1"; Break}
		"VmwareESX4_1"						{$VDIType = "Vmware ESX 4.1"; Break}
		"VmwareESXi5_0"						{$VDIType = "Vmware ESXi 5.0"; Break}
		"VmwareESXi5_1"						{$VDIType = "Vmware ESXi 5.1"; Break}
		"VmwareESXi5_5"						{$VDIType = "Vmware ESXi 5.5"; Break}
		"VmwareESXi6_0"						{$VDIType = "Vmware ESXi 6.0"; Break}
		"VmwareESXi6_5"						{$VDIType = "Vmware ESXi 6.5"; Break}
		"VmwareESXi6_7"						{$VDIType = "Vmware ESXi 6.7"; Break}
		"VmwareVCenterUnknown"				{$VDIType = "Vmware VCenter Server"; Break}
		"VmwareVCenter4_0"					{$VDIType = "Vmware VCenter Server 4.0"; Break}
		"VmwareVCenter4_1"					{$VDIType = "Vmware VCenter Server 4.1"; Break}
		"VmwareVCenter5_0"					{$VDIType = "Vmware VCenter Server 5.0"; Break}
		"VmwareVCenter5_1"					{$VDIType = "Vmware VCenter Server 5.1"; Break}
		"VmwareVCenter5_5"					{$VDIType = "Vmware VCenter Server 5.5"; Break}
		"VmwareVCenter6_0"					{$VDIType = "Vmware VCenter Server 6.0"; Break}
		"VmwareVCenter6_5"					{$VDIType = "Vmware VCenter Server 6.5"; Break}
		"VmwareVCenter6_7"					{$VDIType = "Vmware VCenter Server 6.7"; Break}
		"NutanixUnknown"					{$VDIType = "Nutanix unknown"; Break}
		"Nutanix5_0"						{$VDIType = "Nutanix 5.0"; Break}
		"RemotePCUnknown"					{$VDIType = "Remote PC Unknown"; Break}
		"RemotePCStatic"					{$VDIType = "Remote PC static"; Break}
		"RemotePCDynamic"					{$VDIType = "Remote PC dynamic"; Break}
		"ScaleUnknown"						{$VDIType = "Scale unknown"; Break}
		"Scale7_4"							{$VDIType = "Scale 7.4"; Break}
		"Azure"								{$VDIType = "Azure"; Break}					
		Default								{$VDIType = "Unable to determine VDI Host Type: $($VDIHost.VDIType)"; Break}
	}
	
	Return $VDIType
}
Function OutputSite
{
	Param([object]$Site)
	
	Write-Verbose "$(Get-Date): `tOutput Site $($Site.Name)"
	$RDSHosts = Get-RDS -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve RD Session Hosts for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve RD Session Hosts for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve RD Session Hosts for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve RD Session Hosts for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No RD Session Hosts retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No RD Session Hosts retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No RD Session Hosts retrieved for Site $($Site.Name)e"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No RD Session Hosts retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "Site - $($Site.Name)"
			WriteWordLine 2 0 "RD Session Hosts"
		}
		If($Text)
		{
			Line 0 "Site - $($Site.Name)"
			Line 1 "RD Session Hosts"
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "Site - $($Site.Name)"
			WriteHTMLLine 2 0 "RD Session Hosts"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Site RD Session Hosts"
		ForEach($RDSHost in $RDSHosts)
		{
			$RDSStatus = Get-RDSStatus -Id $RDSHost.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)"
				}
			}
			ElseIf($? -and $Null -eq $RDSStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)"
				}
				If($Text)
				{
					Line 0 "No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)"
				}
			}
			Else
			{
				$Sessions = "$($RDSStatus.ActiveSessions)/$($RDSHost.MaxSessions)"
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Address"; Value = $RDSHost.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $RDSStatus.AgentState; }) > $Null
					$ScriptInformation.Add(@{Data = "CPU"; Value = "$($RDSStatus.CPULoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "RAM"; Value = "$($RDSStatus.MemLoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk read time"; Value = "$($RDSStatus.DiskRead)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk write time"; Value = "$($RDSStatus.DiskWrite)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Sessions"; Value = $Sessions; }) > $Null
					$ScriptInformation.Add(@{Data = "Preferred PA"; Value = $RDSStatus.PreferredPA; }) > $Null
					$ScriptInformation.Add(@{Data = "Operating system"; Value = $RDSStatus.ServerOS; }) > $Null
					$ScriptInformation.Add(@{Data = "Agent version"; Value = $RDSStatus.AgentVer; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Address`t`t: " $RDSHost.Server
					Line 2 "Status`t`t: " $RDSStatus.AgentState
					Line 2 "CPU`t`t: " "$($RDSStatus.CPULoad)%"
					Line 2 "RAM`t`t: " "$($RDSStatus.MemLoad)%"
					Line 2 "Disk read time`t: " "$($RDSStatus.DiskRead)%"
					Line 2 "Disk write time`t: " "$($RDSStatus.DiskWrite)%"
					Line 2 "Sessions`t: " $Sessions
					Line 2 "Preferred PA`t: " $RDSStatus.PreferredPA
					Line 2 "Operating system: " $RDSStatus.ServerOS
					Line 2 "Agent version`t: " $RDSStatus.AgentVer
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Address",($Script:htmlsb),$RDSHost.Server,$htmlwhite)
					$rowdata += @(,( "Status",($Script:htmlsb),$RDSStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,( "CPU",($Script:htmlsb),"$($RDSStatus.CPULoad)%",$htmlwhite))
					$rowdata += @(,( "RAM",($Script:htmlsb),"$($RDSStatus.MemLoad)%",$htmlwhite))
					$rowdata += @(,( "Disk read time",($Script:htmlsb),"$($RDSStatus.DiskRead)%",$htmlwhite))
					$rowdata += @(,( "Disk write time",($Script:htmlsb),"$($RDSStatus.DiskWrite)%",$htmlwhite))
					$rowdata += @(,( "Sessions",($Script:htmlsb),$Sessions,$htmlwhite))
					$rowdata += @(,( "Preferred PA",($Script:htmlsb),$RDSStatus.PreferredPA,$htmlwhite))
					$rowdata += @(,( "Operating system",($Script:htmlsb),$RDSStatus.ServerOS,$htmlwhite))
					$rowdata += @(,( "Agent version",($Script:htmlsb),$RDSStatus.AgentVer,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
		}
	}

	$VDIHosts = Get-VDIHost -SiteId $Site.Id -EA 0 4>$Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve VDI for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve VDI for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve VDI for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve VDI for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No VDI retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No VDI retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No VDI retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No VDI retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "VDI Providers"
		}
		If($Text)
		{
			Line 1 "VDI Providers"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "VDI Providers"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Site VDI Providers"
		ForEach($VDIHost in $VDIHosts)
		{
			$VDIHostStatus = Get-VDIHostStatus -Id $VDIHost.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)"
				}
			}
			ElseIf($? -and $Null -eq $VDIHostStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No VDI Host Status retrieved for VDI Host $($VDIHost.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No VDI Host Status retrieved for VDI Host $($VDIHost.Id)"
				}
				If($Text)
				{
					Line 0 "No VDI Host Status retrieved for VDI Host $($VDIHost.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No VDI Host Status retrieved for VDI Host $($VDIHost.Id)"
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Address"; Value = $VDIHost.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $VDIHostStatus.AgentState.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "CPU"; Value = "$($VDIHostStatus.CPULoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "RAM"; Value = "$($VDIHostStatus.MemLoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk read time"; Value = "$($VDIHostStatus.DiskRead)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk write time"; Value = "$($VDIHostStatus.DiskWrite)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Preferred PA"; Value = $VDIHostStatus.PreferredPA; }) > $Null
					$ScriptInformation.Add(@{Data = "Operating system"; Value = $VDIHostStatus.ServerOS; }) > $Null
					$ScriptInformation.Add(@{Data = "Agent version"; Value = $VDIHostStatus.AgentVer; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Address`t`t: " $VDIHost.Server
					Line 2 "Status`t`t: " $VDIHostStatus.AgentState.ToString()
					Line 2 "CPU`t`t: " "$($VDIHostStatus.CPULoad)%"
					Line 2 "RAM`t`t: " "$($VDIHostStatus.MemLoad)%"
					Line 2 "Disk read time`t: " "$($VDIHostStatus.DiskRead)%"
					Line 2 "Disk write time`t: " "$($VDIHostStatus.DiskWrite)%"
					Line 2 "Preferred PA`t: " $VDIHostStatus.PreferredPA
					Line 2 "Operating system: " $VDIHostStatus.ServerOS
					Line 2 "Agent version`t: " $VDIHostStatus.AgentVer
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Address",($Script:htmlsb),$VDIHost.Server,$htmlwhite)
					$rowdata += @(,("Status",($Script:htmlsb),$VDIHostStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,("CPU",($Script:htmlsb),"$($VDIHostStatus.CPULoad)%",$htmlwhite))
					$rowdata += @(,("RAM",($Script:htmlsb),"$($VDIHostStatus.MemLoad)%",$htmlwhite))
					$rowdata += @(,("Disk read time",($Script:htmlsb),"$($VDIHostStatus.DiskRead)%",$htmlwhite))
					$rowdata += @(,("Disk write time",($Script:htmlsb),"$($VDIHostStatus.DiskWrite)%",$htmlwhite))
					$rowdata += @(,("Preferred PA",($Script:htmlsb),$VDIHostStatus.PreferredPA,$htmlwhite))
					$rowdata += @(,("Operating system",($Script:htmlsb),$VDIHostStatus.ServerOS,$htmlwhite))
					$rowdata += @(,("Agent version",($Script:htmlsb),$VDIHostStatus.AgentVer,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
		}
	}
	
	$GWs = Get-GW -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Gateways for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Gateways for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Gateways for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Gateways for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Gateways retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Gateways retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No Gateways retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Gateways retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Gateways"
		}
		If($Text)
		{
			Line 1 "Gateways"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Gateways"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Site Gateways"
		ForEach($GW in $GWs)
		{
			$GWStatus = Get-GWStatus -Id $GW.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve Gateway Status for Gateway $($GW.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve Gateway Status for Gateway $($GW.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve Gateway Status for Gateway $($GW.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve Gateway Status for Gateway $($GW.Id)"
				}
			}
			ElseIf($? -and $Null -eq $GWStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No Gateway Status retrieved for Gateway $($GW.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Gateway Status retrieved for Gateway $($GW.Id)"
				}
				If($Text)
				{
					Line 0 "No Gateway Status retrieved for Gateway $($GW.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Gateway Status retrieved for Gateway $($GW.Id)"
				}
			}
			Else
			{
				$Sessions = ($GWStatus.ActiveRDPSessions + $GWStatus.ActiveRDPSSLSessions)
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Address"; Value = $GW.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $GWStatus.AgentState; }) > $Null
					$ScriptInformation.Add(@{Data = "CPU"; Value = "$($GWStatus.CPULoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "RAM"; Value = "$($GWStatus.MemLoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk read time"; Value = "$($GWStatus.DiskRead)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk write time"; Value = "$($GWStatus.DiskWrite)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Sessions"; Value = $Sessions.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Preferred PA"; Value = $GWStatus.PreferredPA; }) > $Null
					$ScriptInformation.Add(@{Data = "Operating system"; Value = $GWStatus.ServerOS; }) > $Null
					$ScriptInformation.Add(@{Data = "Agent version"; Value = $GWStatus.AgentVer; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Address`t`t: " $GW.Server
					Line 2 "Status`t`t: " $GWStatus.AgentState
					Line 2 "CPU`t`t: " "$($GWStatus.CPULoad)%"
					Line 2 "RAM`t`t: " "$($GWStatus.MemLoad)%"
					Line 2 "Disk read time`t: " "$($GWStatus.DiskRead)%"
					Line 2 "Disk write time`t: " "$($GWStatus.DiskWrite)%"
					Line 2 "Sessions`t: " $Sessions.ToString()
					Line 2 "Preferred PA`t: " $GWStatus.PreferredPA
					Line 2 "Operating system: " $GWStatus.ServerOS
					Line 2 "Agent version`t: " $GWStatus.AgentVer
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Address",($Script:htmlsb),$GW.Server,$htmlwhite)
					$rowdata += @(,( "Status",($Script:htmlsb),$GWStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,( "CPU",($Script:htmlsb),"$($GWStatus.CPULoad)%",$htmlwhite))
					$rowdata += @(,( "RAM",($Script:htmlsb),"$($GWStatus.MemLoad)%",$htmlwhite))
					$rowdata += @(,( "Disk read time",($Script:htmlsb),"$($GWStatus.DiskRead)%",$htmlwhite))
					$rowdata += @(,( "Disk write time",($Script:htmlsb),"$($GWStatus.DiskWrite)%",$htmlwhite))
					$rowdata += @(,( "Sessions",($Script:htmlsb),$Sessions.ToString(),$htmlwhite))
					$rowdata += @(,( "Preferred PA",($Script:htmlsb),$GWStatus.PreferredPA,$htmlwhite))
					$rowdata += @(,( "Operating system",($Script:htmlsb),$GWStatus.ServerOS,$htmlwhite))
					$rowdata += @(,( "Agent version",($Script:htmlsb),$GWStatus.AgentVer,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
		}
	}
	
	$PAs = Get-PA -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Publishing Agents for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Publishing Agents for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Publishing Agents for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Publishing Agents for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Publishing Agents retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Publishing Agents retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No Publishing Agents retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Publishing Agents retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Publishing Agents"
		}
		If($Text)
		{
			Line 1 "Publishing Agents"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Publishing Agents"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Site Publishing Agents"
		ForEach($PA in $PAs)
		{
			$PAStatus = Get-PAStatus -Id $PA.Id -EA 0 4>$Null
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve Publishing Agents Status for Publishing Agents $($PA.Id)"
				}
			}
			ElseIf($? -and $Null -eq $GWStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)"
				}
				If($Text)
				{
					Line 0 "No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Publishing Agents Status retrieved for Publishing Agents $($PA.Id)"
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Address"; Value = $PA.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $PAStatus.AgentState; }) > $Null
					$ScriptInformation.Add(@{Data = "CPU"; Value = "$($PAStatus.CPULoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "RAM"; Value = "$($PAStatus.MemLoad)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk read time"; Value = "$($PAStatus.DiskRead)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Disk write time"; Value = "$($PAStatus.DiskWrite)%"; }) > $Null
					$ScriptInformation.Add(@{Data = "Operating system"; Value = $PAStatus.ServerOS; }) > $Null
					$ScriptInformation.Add(@{Data = "Agent version"; Value = $PAStatus.AgentVer; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Address`t`t: " $PA.Server
					Line 2 "Status`t`t: " $PAStatus.AgentState
					Line 2 "CPU`t`t: " "$($PAStatus.CPULoad)%"
					Line 2 "RAM`t`t: " "$($PAStatus.MemLoad)%"
					Line 2 "Disk read time`t: " "$($PAStatus.DiskRead)%"
					Line 2 "Disk write time`t: " "$($PAStatus.DiskWrite)%"
					Line 2 "Operating system: " $PAStatus.ServerOS
					Line 2 "Agent version`t: " $PAStatus.AgentVer
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Address",($Script:htmlsb),$PA.Server,$htmlwhite)
					$rowdata += @(,( "Status",($Script:htmlsb),$PAStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,( "CPU",($Script:htmlsb),"$($PAStatus.CPULoad)%",$htmlwhite))
					$rowdata += @(,( "RAM",($Script:htmlsb),"$($PAStatus.MemLoad)%",$htmlwhite))
					$rowdata += @(,( "Disk read time",($Script:htmlsb),"$($PAStatus.DiskRead)%",$htmlwhite))
					$rowdata += @(,( "Disk write time",($Script:htmlsb),"$($PAStatus.DiskWrite)%",$htmlwhite))
					$rowdata += @(,( "Operating system",($Script:htmlsb),$PAStatus.ServerOS,$htmlwhite))
					$rowdata += @(,( "Agent version",($Script:htmlsb),$PAStatus.AgentVer,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
		}
	}
	
	Write-Verbose "$(Get-Date): `t`tOutput Site RD Session Hosts for Site $($Site.Name)"
	$RDSHosts = Get-RDS -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve RD Session Hosts for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve RD Session Hosts for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve RD Session Hosts for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve RD Session Hosts for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No RD Session Hosts retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No RD Session Hosts retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No RD Session Hosts retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No RD Session Hosts retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "RD Session Hosts"
		}
		If($Text)
		{
			Line 1 "RD Session Hosts"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "RD Session Hosts"
		}

		ForEach($RDSHost in $RDSHosts)
		{
			Write-Verbose "$(Get-Date): `t`t`tOutput Site RD Session Host $($RDSHost.Server)"
			$RDSStatus = Get-RDSStatus -Id $RDSHost.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve RD Session Hosts Status for RD Session Hosts $($RDSHost.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve RD Session Hosts Status for RD Session Hosts $($RDSHost.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve RD Session Hosts Status for RD Session Hosts $($RDSHost.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve RD Session Hosts Status for RD Session Hosts $($RDSHost.Id)"
				}
			}
			ElseIf($? -and $Null -eq $RDSStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No RD Session Hosts Status retrieved for RD Session Hosts $($RDSHost.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No RD Session Hosts Status retrieved for RD Session Hosts $($RDSHost.Id)"
				}
				If($Text)
				{
					Line 0 "No RD Session Hosts Status retrieved for RD Session Hosts $($RDSHost.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No RD Session Hosts Status retrieved for RD Session Hosts $($RDSHost.Id)"
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Server $($RDSHost.Server)"
				}
				If($Text)
				{
					Line 2 "Server $($RDSHost.Server)"
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Server $($RDSHost.Server)"
				}

				$LogonStatus = "N/A"
				$Status = "N/A"
				Switch ($RDSStatus.AgentState)
				{
					"OK"						{$Status = "OK"; Break}
					"LogonDrainUntilRestart"	{$Status = "Logons disabled until reboot"; Break}
					"LogonDisabled"				{$Status = "New logons and reconnections disabled"; Break}
					Default						{$Status = "Unale to determine RDS Status Agent State: $($RDSStatus.AgentState)";Break}
				}
				
				Switch($RDSStatus.LoginStatus)
				{
					"Enabled"	{$LogonStatus = "Enabled"; Break}
					"Disabled"	{$LogonStatus = "Disabled"; Break}
					"DrainMode"	{$LogonStatus = "Drain Mode"; Break}
					Default		{$LogonStatus = "Unable to determine RDS Status Logon status: $($RDSStatus.LoginStatus)"; Break}
				}
				
				$UPDStatus = "Unknown"
				If($RDSStatus.UPDStatus -eq "NotSupported")
				{
					$UPDDefault = (Get-RDSDefaultSettings -SiteId $Site.Id -Ea 0).UPDMode
					
					If($UPDDefault -eq "DoNotChange")
					{
						$UPDStatus = "Disabled"
					}
					Else
					{
						$UPDStatus = $UPDDefault
					}
				}
				Else
				{
					$UPDStatus = $RDSStatus.UPDStatus
				}
				
				$RDSGroup = @()
				$Results = Get-RDSGroup -SiteId $Site.Id -EA 0 4>$Null
				If( $Results.RDSIds -Contains $RDSHost.Id )
				{
					$RDSGroup += $Results.Name
				}
				Else
				{
					$RDSGroup += ""
				}
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Server"; Value = $RDSHost.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $Status; }) > $Null
					$ScriptInformation.Add(@{Data = "Logon status"; Value = $LogonStatus; }) > $Null
					$ScriptInformation.Add(@{Data = "Group"; Value = $RDSGroup[0]; }) > $Null
					$ScriptInformation.Add(@{Data = "Direct address"; Value = $RDSHost.DirectAddress; }) > $Null
					$ScriptInformation.Add(@{Data = "Description"; Value = $RDSHost.Description; }) > $Null
					$ScriptInformation.Add(@{Data = "UPD"; Value = $UPDStatus; }) > $Null
					$ScriptInformation.Add(@{Data = "Log level"; Value = $RDSStatus.LogLevel; }) > $Null
					$ScriptInformation.Add(@{Data = "Agent version"; Value = $RDSStatus.AgentVer; }) > $Null
					$ScriptInformation.Add(@{Data = "ID"; Value = $RDSHost.Id; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 3 "Server`t`t: " $RDSHost.Server
					Line 3 "Status`t`t: " $Status
					Line 3 "Logon status`t: " $LogonStatus
					Line 3 "Group`t`t: " $RDSGroup[0]
					Line 3 "Direct address`t: " $RDSHost.DirectAddress
					Line 3 "Description`t: " $RDSHost.Description
					Line 3 "UPD`t`t: " $UPDStatus
					Line 3 "Log level`t: " $RDSStatus.LogLevel
					Line 3 "Agent version`t: " $RDSStatus.AgentVer
					Line 3 "ID`t`t: " $RDSHost.Id
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Server",($Script:htmlsb),$RDSHost.Server,$htmlwhite)
					$rowdata += @(,( "Status",($Script:htmlsb),$Status,$htmlwhite))
					$rowdata += @(,( "Logon status",($Script:htmlsb),$LogonStatus,$htmlwhite))
					$rowdata += @(,( "Group",($Script:htmlsb),$RDSGroup[0],$htmlwhite))
					$rowdata += @(,( "Direct address",($Script:htmlsb),$RDSHost.DirectAddress,$htmlwhite))
					$rowdata += @(,( "Description",($Script:htmlsb),$RDSHost.Description,$htmlwhite))
					$rowdata += @(,( "UPD",($Script:htmlsb),$UPDStatus,$htmlwhite))
					$rowdata += @(,( "Log level",($Script:htmlsb),$RDSStatus.LogLevel,$htmlwhite))
					$rowdata += @(,( "Agent version",($Script:htmlsb),$RDSStatus.AgentVer,$htmlwhite))
					$rowdata += @(,( "ID",($Script:htmlsb),$RDSHost.Id,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
			
			#Properties
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Properties"
			}
			If($Text)
			{
				Line 3 "Properties"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Properties"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Enable server in site"; Value = $RDSHost.Enabled.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Server"; Value = $RDSHost.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $RDSHost.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Direct Address"; Value = $RDSHost.DirectAddress; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Enable server in site`t`t`t`t`t: " $RDSHost.Enabled.ToString()
				Line 4 "Server`t`t`t`t`t`t`t: " $RDSHost.Server
				Line 4 "Description`t`t`t`t`t`t: " $RDSHost.Description
				Line 4 "Direct Address`t`t`t`t`t`t: " $RDSHost.DirectAddress
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable server in site",($Script:htmlsb),$RDSHost.Enabled.ToString(),$htmlwhite)
				$rowdata += @(,( "Server",($Script:htmlsb),$RDSHost.Server,$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$RDSHost.Description,$htmlwhite))
				$rowdata += @(,( "Direct Address",($Script:htmlsb),$RDSHost.DirectAddress,$htmlwhite))

				$msg = "Properties"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Agent Settings
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Agent Settings"
			}
			If($Text)
			{
				Line 3 "Agent Settings"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Agent Settings"
			}
			
			If($RDSHost.InheritDefaultAgentSettings)
			{
				#do we inherit group or site defaults?
				#is this RDS host in a group?
				$Results = Get-RDSGroup -SiteId $Site.Id -EA 0 4>$Null
				If( $Results.RDSIds -Contains $RDSHost.Id )
				{
					#does this group inherit default settings?
					If($Results.InheritDefaultAgentSettings -eq $False)
					{
						#no we don't, so get the default settings for the group
						$GroupDefaults = $Results.RDSDefSettings

						$RDSPort = $GroupDefaults.Port.ToString()
						$RDSMaxSessions = $GroupDefaults.MaxSessions.ToString()
						
						Switch ($GroupDefaults.SessionTimeout)
						{
							0		{$RDSPublishingSessionDisconnectTimeout = "Never"; Break}
							25		{$RDSPublishingSessionDisconnectTimeout = "25 seconds"; Break}
							60		{$RDSPublishingSessionDisconnectTimeout = "1 minute"; Break}
							300		{$RDSPublishingSessionDisconnectTimeout = "5 minutes"; Break}
							3600	{$RDSPublishingSessionDisconnectTimeout = "1 hour"; Break}
							Default	{$RDSPublishingSessionDisconnectTimeout = "Unable to determine Publishing Session Disconnect Timeout: $($GroupDefaults.SessionTimeout)"; Break}
						}
						
						Switch ($GroupDefaults.SessionLogoffTimeout)
						{
							0		{$RDSPublishingSessionResetTime = "Never"; Break}
							1		{$RDSPublishingSessionResetTime = "Immediate"; Break}
							25		{$RDSPublishingSessionResetTime = "25 seconds"; Break}
							60		{$RDSPublishingSessionResetTime = "1 minute"; Break}
							300		{$RDSPublishingSessionResetTime = "5 minutes"; Break}
							3600	{$RDSPublishingSessionResetTime = "1 hour"; Break}
							Default	{$RDSPublishingSessionResetTime = "Unable to determine Publishing Session Reset Timeout: $($GroupDefaults.SessionLogoffTimeout)"; Break}
						}
						
						Switch($GroupDefaults.AllowURLAndMailRedirection)
						{
							"Enabled"						{$RDSAllowClientURLMailRedirection = "Enabled"; Break}
							"Disbaled"						{$RDSAllowClientURLMailRedirection = "Disabled"; Break}
							"EnabledWithAppRegistration"	{$RDSAllowClientURLMailRedirection = "Enabled (Replace Registered Application)"; Break}
							Default 						{$RDSAllowClientURLMailRedirection = "Unable to determine Allow CLient URL/Mail Redirection: $($GroupDefaults.AllowURLAndMailRedirection)"; Break}
						}
						
						$RDSSupportShellURLNamespaceObject = $GroupDefaults.SupportShellURLNamespaceObjects.ToString()
						
						Switch ($GroupDefaults.DragAndDropMode)
						{
							"Bidirectional"		{$RDSDragAndDrop = "Bidirectional"; Break}
							"Disabled"			{$RDSDragAndDrop = "Disabled"; Break}
							"ClientToServer"	{$RDSDragAndDrop = "Client to server only"; Break}
							"ServerToClient"	{$RDSDragAndDrop = "Server to client only"; Break}
							Default				{$RDSDragAndDrop = "Unable to determine Drag and drop: $($GroupDefaults.DragAndDropMode)"; Break}
						}
						
						If($GroupDefaults.PreferredPAId -eq 0)
						{
							$RDSPreferredPublishingAgent = "Automatically"
						}
						Else
						{
							$RDSPreferredPublishingAgent = (Get-PA -Id $GroupDefaults.PreferredPAId -EA 0 4>$Null).Server
						}
					}
					Else
					{
						#yes we do, get the default settings for the Site
						#use the Site default settings

						$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
						
						If($? -and $Null -ne $RDSDefaults)
						{
							$RDSPort = $RDSDefaults.Port.ToString()
							$RDSMaxSessions = $RDSDefaults.MaxSessions.ToString()
							
							Switch ($RDSDefaults.SessionTimeout)
							{
								0		{$RDSPublishingSessionDisconnectTimeout = "Never"; Break}
								25		{$RDSPublishingSessionDisconnectTimeout = "25 seconds"; Break}
								60		{$RDSPublishingSessionDisconnectTimeout = "1 minute"; Break}
								300		{$RDSPublishingSessionDisconnectTimeout = "5 minutes"; Break}
								3600	{$RDSPublishingSessionDisconnectTimeout = "1 hour"; Break}
								Default	{$RDSPublishingSessionDisconnectTimeout = "Unable to determine Publishing Session Disconnect Timeout: $($RDSDefaults.SessionTimeout)"; Break}
							}
							
							Switch ($RDSDefaults.SessionLogoffTimeout)
							{
								0		{$RDSPublishingSessionResetTime = "Never"; Break}
								1		{$RDSPublishingSessionResetTime = "Immediate"; Break}
								25		{$RDSPublishingSessionResetTime = "25 seconds"; Break}
								60		{$RDSPublishingSessionResetTime = "1 minute"; Break}
								300		{$RDSPublishingSessionResetTime = "5 minutes"; Break}
								3600	{$RDSPublishingSessionResetTime = "1 hour"; Break}
								Default	{$RDSPublishingSessionResetTime = "Unable to determine Publishing Session Reset Timeout: $($RDSDefaults.SessionLogoffTimeout)"; Break}
							}
							
							Switch($RDSDefaults.AllowURLAndMailRedirection)
							{
								"Enabled"						{$RDSAllowClientURLMailRedirection = "Enabled"; Break}
								"Disbaled"						{$RDSAllowClientURLMailRedirection = "Disabled"; Break}
								"EnabledWithAppRegistration"	{$RDSAllowClientURLMailRedirection = "Enabled (Replace Registered Application)"; Break}
								Default 						{$RDSAllowClientURLMailRedirection = "Unable to determine Allow CLient URL/Mail Redirection: $($RDSDefaults.AllowURLAndMailRedirection)"; Break}
							}
							
							$RDSSupportShellURLNamespaceObject = $RDSDefaults.SupportShellURLNamespaceObjects.ToString()
							
							Switch ($RDSDefaults.DragAndDropMode)
							{
								"Bidirectional"		{$RDSDragAndDrop = "Bidirectional"; Break}
								"Disabled"			{$RDSDragAndDrop = "Disabled"; Break}
								"ClientToServer"	{$RDSDragAndDrop = "Client to server only"; Break}
								"ServerToClient"	{$RDSDragAndDrop = "Server to client only"; Break}
								Default				{$RDSDragAndDrop = "Unable to determine Drag and drop: $($RDSDefaults.DragAndDropMode)"; Break}
							}
							
							If($RDSDefaults.PreferredPAId -eq 0)
							{
								$RDSPreferredPublishingAgent = "Automatically"
							}
							Else
							{
								$RDSPreferredPublishingAgent = (Get-PA -Id $RDSDefaults.PreferredPAId -EA 0 4>$Null).Server
							}
							$RDSAllowRemoteExec = $RDSDefaults.AllowRemoteExec.ToString()
							$RDSUseRemoteApps = $RDSDefaults.UseRemoteApps.ToString()
							$RDSEnableAppMonitoring = $RDSDefaults.EnableAppMonitoring.ToString()
							$RDSAllowFileTransfer = $RDSDefaults.AllowFileTransfer.ToString()
						}
						Else
						{
							#unable to retrieve default, use built-in default values
							$RDSPort = "3389"
							$RDSMaxSessions = "250"
							$RDSPublishingSessionDisconnectTimeout = "25 seconds"
							$RDSPublishingSessionResetTime = "Immediate"
							$RDSAllowClientURLMailRedirection = "Enabled"
							$RDSSupportShellURLNamespaceObject = "True"
							$RDSDragAndDrop = "Bidirectional"
							$RDSPreferredPublishingAgent = "Automatically"
							$RDSAllowRemoteExec = "True"
							$RDSUseRemoteApps = "False"
							$RDSEnableAppMonitoring = "True"
							$RDSAllowFileTransfer = "True"
						}
					}
				}
				Else
				{
					#server is not in an RDS group
					#get the settings configured for this RDS host
					$RDSPort = $RDSHost.Port.ToString()
					$RDSMaxSessions = $RDSHost.MaxSessions.ToString()
					
					Switch ($RDSHost.SessionTimeout)
					{
						0		{$RDSPublishingSessionDisconnectTimeout = "Never"; Break}
						25		{$RDSPublishingSessionDisconnectTimeout = "25 seconds"; Break}
						60		{$RDSPublishingSessionDisconnectTimeout = "1 minute"; Break}
						300		{$RDSPublishingSessionDisconnectTimeout = "5 minutes"; Break}
						3600	{$RDSPublishingSessionDisconnectTimeout = "1 hour"; Break}
						Default	{$RDSPublishingSessionDisconnectTimeout = "Unable to determine Publishing Session Disconnect Timeout: $($RDSHost.SessionTimeout)"; Break}
					}
					
					Switch ($RDSHost.SessionLogoffTimeout)
					{
						0		{$RDSPublishingSessionResetTime = "Never"; Break}
						1		{$RDSPublishingSessionResetTime = "Immediate"; Break}
						25		{$RDSPublishingSessionResetTime = "25 seconds"; Break}
						60		{$RDSPublishingSessionResetTime = "1 minute"; Break}
						300		{$RDSPublishingSessionResetTime = "5 minutes"; Break}
						3600	{$RDSPublishingSessionResetTime = "1 hour"; Break}
						Default	{$RDSPublishingSessionResetTime = "Unable to determine Publishing Session Reset Timeout: $($RDSHost.SessionLogoffTimeout)"; Break}
					}
					
					Switch($RDSHost.AllowURLAndMailRedirection)
					{
						"Enabled"						{$RDSAllowClientURLMailRedirection = "Enabled"; Break}
						"Disbaled"						{$RDSAllowClientURLMailRedirection = "Disabled"; Break}
						"EnabledWithAppRegistration"	{$RDSAllowClientURLMailRedirection = "Enabled (Replace Registered Application)"; Break}
						Default 						{$RDSAllowClientURLMailRedirection = "Unable to determine Allow CLient URL/Mail Redirection: $($RDSHost.AllowURLAndMailRedirection)"; Break}
					}
					
					$RDSSupportShellURLNamespaceObject = $RDSHost.SupportShellURLNamespaceObjects.ToString()
					
					Switch ($RDSHost.DragAndDropMode)
					{
						"Bidirectional"		{$RDSDragAndDrop = "Bidirectional"; Break}
						"Disabled"			{$RDSDragAndDrop = "Disabled"; Break}
						"ClientToServer"	{$RDSDragAndDrop = "Client to server only"; Break}
						"ServerToClient"	{$RDSDragAndDrop = "Server to client only"; Break}
						Default				{$RDSDragAndDrop = "Unable to determine Drag and drop: $($RDSHost.DragAndDropMode)"; Break}
					}
					
					If($RDSHost.PreferredPAId -eq 0)
					{
						$RDSPreferredPublishingAgent = "Automatically"
					}
					Else
					{
						$RDSPreferredPublishingAgent = (Get-PA -Id $RDSHost.PreferredPAId -EA 0 4>$Null).Server
					}
					$RDSAllowRemoteExec = $RDSHost.AllowRemoteExec.ToString()
					$RDSUseRemoteApps = $RDSHost.UseRemoteApps.ToString()
					$RDSEnableAppMonitoring = $RDSHost.EnableAppMonitoring.ToString()
					$RDSAllowFileTransfer = $RDSHost.AllowFileTransfer.ToString()
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this RDS host
				$RDSPort = $RDSHost.Port.ToString()
				$RDSMaxSessions = $RDSHost.MaxSessions.ToString()
				
				Switch ($RDSHost.SessionTimeout)
				{
					0		{$RDSPublishingSessionDisconnectTimeout = "Never"; Break}
					25		{$RDSPublishingSessionDisconnectTimeout = "25 seconds"; Break}
					60		{$RDSPublishingSessionDisconnectTimeout = "1 minute"; Break}
					300		{$RDSPublishingSessionDisconnectTimeout = "5 minutes"; Break}
					3600	{$RDSPublishingSessionDisconnectTimeout = "1 hour"; Break}
					Default	{$RDSPublishingSessionDisconnectTimeout = "Unable to determine Publishing Session Disconnect Timeout: $($RDSHost.SessionTimeout)"; Break}
				}
				
				Switch ($RDSHost.SessionLogoffTimeout)
				{
					0		{$RDSPublishingSessionResetTime = "Never"; Break}
					1		{$RDSPublishingSessionResetTime = "Immediate"; Break}
					25		{$RDSPublishingSessionResetTime = "25 seconds"; Break}
					60		{$RDSPublishingSessionResetTime = "1 minute"; Break}
					300		{$RDSPublishingSessionResetTime = "5 minutes"; Break}
					3600	{$RDSPublishingSessionResetTime = "1 hour"; Break}
					Default	{$RDSPublishingSessionResetTime = "Unable to determine Publishing Session Reset Timeout: $($RDSHost.SessionLogoffTimeout)"; Break}
				}
				
				Switch($RDSHost.AllowURLAndMailRedirection)
				{
					"Enabled"						{$RDSAllowClientURLMailRedirection = "Enabled"; Break}
					"Disbaled"						{$RDSAllowClientURLMailRedirection = "Disabled"; Break}
					"EnabledWithAppRegistration"	{$RDSAllowClientURLMailRedirection = "Enabled (Replace Registered Application)"; Break}
					Default 						{$RDSAllowClientURLMailRedirection = "Unable to determine Allow CLient URL/Mail Redirection: $($RDSHost.AllowURLAndMailRedirection)"; Break}
				}
				
				$RDSSupportShellURLNamespaceObject = $RDSHost.SupportShellURLNamespaceObjects.ToString()
				
				Switch ($RDSHost.DragAndDropMode)
				{
					"Bidirectional"		{$RDSDragAndDrop = "Bidirectional"; Break}
					"Disabled"			{$RDSDragAndDrop = "Disabled"; Break}
					"ClientToServer"	{$RDSDragAndDrop = "Client to server only"; Break}
					"ServerToClient"	{$RDSDragAndDrop = "Server to client only"; Break}
					Default				{$RDSDragAndDrop = "Unable to determine Drag and drop: $($RDSHost.DragAndDropMode)"; Break}
				}
				
				If($RDSHost.PreferredPAId -eq 0)
				{
					$RDSPreferredPublishingAgent = "Automatically"
				}
				Else
				{
					$RDSPreferredPublishingAgent = (Get-PA -Id $RDSHost.PreferredPAId -EA 0 4>$Null).Server
				}
				$RDSAllowRemoteExec = $RDSHost.AllowRemoteExec.ToString()
				$RDSUseRemoteApps = $RDSHost.UseRemoteApps.ToString()
				$RDSEnableAppMonitoring = $RDSHost.EnableAppMonitoring.ToString()
				$RDSAllowFileTransfer = $RDSHost.AllowFileTransfer.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSHost.InheritDefaultAgentSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Port"; Value = $RDSPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Max Sessions"; Value = $RDSMaxSessions; }) > $Null
				$ScriptInformation.Add(@{Data = "Publishing Session Disconnect Timeout"; Value = $RDSPublishingSessionDisconnectTimeout; }) > $Null
				$ScriptInformation.Add(@{Data = "Publishing Session Reset Timeout"; Value = $RDSPublishingSessionResetTime; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow Client URL/Mail Redirection"; Value = $RDSAllowClientURLMailRedirection; }) > $Null
				$ScriptInformation.Add(@{Data = "Support Shell URL namespace objects"; Value = $RDSSupportShellURLNamespaceObject; }) > $Null
				$ScriptInformation.Add(@{Data = "Drag and drop"; Value = $RDSDragAndDrop; }) > $Null
				$ScriptInformation.Add(@{Data = "Preferred Publishing Agent"; Value = $RDSPreferredPublishingAgent; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow 2xRemoteExec to send command to the client"; Value = $RDSAllowRemoteExec; }) > $Null
				$ScriptInformation.Add(@{Data = "Use RemoteApp if available"; Value = $RDSUseRemoteApps; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable applications monitoring"; Value = $RDSEnableAppMonitoring; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow file transfer command (HTML5 and Chrome clients)"; Value = $RDSAllowFileTransfer; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Inherit default settings`t`t`t`t: " $RDSHost.InheritDefaultAgentSettings.ToString()
				Line 4 "Port`t`t`t`t`t`t`t: " $RDSPort
				Line 4 "Max Sessions`t`t`t`t`t`t: " $RDSMaxSessions
				Line 4 "Publishing Session Disconnect Timeout`t`t`t: " $RDSPublishingSessionDisconnectTimeout
				Line 4 "Publishing Session Reset Timeout`t`t`t: " $RDSPublishingSessionResetTime
				Line 4 "Allow Client URL/Mail Redirection`t`t`t: " $RDSAllowClientURLMailRedirection
				Line 4 "Support Shell URL namespace objects`t`t`t: " $RDSSupportShellURLNamespaceObject
				Line 4 "Drag and drop`t`t`t`t`t`t: " $RDSDragAndDrop
				Line 4 "Preferred Publishing Agent`t`t`t`t: " $RDSPreferredPublishingAgent
				Line 4 "Allow 2xRemoteExec to send command to the client`t: " $RDSAllowRemoteExec
				Line 4 "Use RemoteApp if available`t`t`t`t: " $RDSUseRemoteApps
				Line 4 "Enable applications monitoring`t`t`t`t: " $RDSEnableAppMonitoring
				Line 4 "Allow file transfer command (HTML5 and Chrome clients)`t: " $RDSAllowFileTransfer
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSHost.InheritDefaultAgentSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Port",($Script:htmlsb),$RDSPort,$htmlwhite))
				$rowdata += @(,( "Max Sessions",($Script:htmlsb),$RDSMaxSessions,$htmlwhite))
				$rowdata += @(,( "Publishing Session Disconnect Timeout",($Script:htmlsb),$RDSPublishingSessionDisconnectTimeout,$htmlwhite))
				$rowdata += @(,( "Publishing Session Reset Timeout",($Script:htmlsb),$RDSPublishingSessionResetTime,$htmlwhite))
				$rowdata += @(,( "Allow Client URL/Mail Redirection",($Script:htmlsb),$RDSAllowClientURLMailRedirection,$htmlwhite))
				$rowdata += @(,( "Support Shell URL namespace objects",($Script:htmlsb),$RDSSupportShellURLNamespaceObject,$htmlwhite))
				$rowdata += @(,( "Drag and drop",($Script:htmlsb),$RDSDragAndDrop,$htmlwhite))
				$rowdata += @(,( "Preferred Publishing Agent",($Script:htmlsb),$RDSPreferredPublishingAgent,$htmlwhite))
				$rowdata += @(,( "Allow 2xRemoteExec to send command to the client",($Script:htmlsb),$RDSAllowRemoteExec,$htmlwhite))
				$rowdata += @(,( "Use RemoteApp if available",($Script:htmlsb),$RDSUseRemoteApps,$htmlwhite))
				$rowdata += @(,( "Enable applications monitoring",($Script:htmlsb),$RDSEnableAppMonitoring,$htmlwhite))
				$rowdata += @(,( "Allow file transfer command (HTML5 and Chrome clients)",($Script:htmlsb),$RDSAllowFileTransfer,$htmlwhite))

				$msg = "Agent Settings"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#User Profile Disks
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "User Profile Disks"
			}
			If($Text)
			{
				Line 3 "User Profile Disks"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "User Profile Disks"
			}
			
			If($RDSHost.InheritDefaultUPDSettings)
			{
				#do we inherit group or site defaults?
				#is this RDS host in a group?
				$Results = Get-RDSGroup -SiteId $Site.Id -EA 0 4>$Null
				If( $Results.RDSIds -Contains $RDSHost.Id )
				{
					#does this group inherit default settings?
					If($Results.InheritDefaultUPDSettings -eq $False)
					{
						#no we don't, so get the default settings for the group
						$GroupDefaults = $Results.RDSDefSettings

						Switch ($GroupDefaults.UPDMode)
						{
							"DoNotChange"	{$RDSUPDState = "Do not change"; Break}
							"Enabled"		{$RDSUPDState = "Enabled"; Break}
							"Disabled"		{$RDSUPDState = "Disabled"; Break}
							Default			{$RDSUPDState = "Unable to determine Current UPD State: $($GroupDefaults.UPDMode)"; Break}
						}
						
						$RDSUPDLocation = $GroupDefaults.DiskPath
						$RDSUPDSize = $GroupDefaults.MaxUserProfileDiskSizeGB
					}
					Else
					{
						#yes we do, get the default settings for the Site
						#use the Site default settings
						$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
						
						If($? -and $Null -ne $RDSDefaults)
						{
							Switch ($RDSDefaults.UPDMode)
							{
								"DoNotChange"	{$RDSUPDState = "Do not change"; Break}
								"Enabled"		{$RDSUPDState = "Enabled"; Break}
								"Disabled"		{$RDSUPDState = "Disabled"; Break}
								Default			{$RDSUPDState = "Unable to determine Current UPD State: $($RDSDefaults.UPDMode)"; Break}
							}
							
							$RDSUPDLocation = $RDSDefaults.DiskPath
							$RDSUPDSize = $RDSDefaults.MaxUserProfileDiskSizeGB
						}
						Else
						{
							#unable to retrieve default, use built-in default values
							$RDSUPDState = "Do not change"
							$RDSUPDLocation = ""
							$RDSUPDSize = "20"
						}
					}
				}
				Else
				{
					#RDS Host is not in a group
					#get the settings for the host
					Switch ($RDSHost.UPDMode)
					{
						"DoNotChange"	{$RDSUPDState = "Do not change"; Break}
						"Enabled"		{$RDSUPDState = "Enabled"; Break}
						"Disabled"		{$RDSUPDState = "Disabled"; Break}
						Default			{$RDSUPDState = "Unable to determine Current UPD State: $($RDSHost.UPDMode)"; Break}
					}
					
					$RDSUPDLocation = $RDSHost.DiskPath
					$RDSUPDSize = $RDSHost.MaxUserProfileDiskSizeGB.ToString()
				}
			}
			Else
			{
				#we don't inherit
				#get the settings for the host
				Switch ($RDSHost.UPDMode)
				{
					"DoNotChange"	{$RDSUPDState = "Do not change"; Break}
					"Enabled"		{$RDSUPDState = "Enabled"; Break}
					"Disabled"		{$RDSUPDState = "Disabled"; Break}
					Default			{$RDSUPDState = "Unable to determine Current UPD State: $($RDSHost.UPDMode)"; Break}
				}
				
				$RDSUPDLocation = $RDSHost.DiskPath
				$RDSUPDSize = $RDSHost.MaxUserProfileDiskSizeGB.ToString()
			}
				
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSHost.InheritDefaultUPDSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "UPD State"; Value = $RDSUPDState; }) > $Null
				$ScriptInformation.Add(@{Data = "Location of user profile disks"; Value = $RDSUPDLocation; }) > $Null
				$ScriptInformation.Add(@{Data = "Maximum size (in GB)"; Value = $RDSUPDSize; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Inherit default settings`t`t`t`t: " $RDSHost.InheritDefaultUPDSettings.ToString()
				Line 4 "UPD State`t`t`t`t`t`t: " $RDSUPDState
				Line 4 "Location of user profile disks`t`t`t`t: " $RDSUPDLocation
				Line 4 "Maximum size (in GB)`t`t`t`t`t: " $RDSUPDSize
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSHost.InheritDefaultUPDSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "UPD State",($Script:htmlsb),$RDSUPDState,$htmlwhite))
				$rowdata += @(,( "Location of user profile disks",($Script:htmlsb),$RDSUPDLocation,$htmlwhite))
				$rowdata += @(,( "Maximum size (in GB)",($Script:htmlsb),$RDSUPDSize,$htmlwhite))

				$msg = "User Profile Disks"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Desktop Access
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Desktop Access"
			}
			If($Text)
			{
				Line 3 "Desktop Access"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Desktop Access"
			}
			
			If($RDSHost.InheritDefaultDesktopAccessSettings)
			{
				#do we inherit group or site defaults?
				#is this RDS host in a group?
				#http://woshub.com/hot-to-convert-sid-to-username-and-vice-versa/
				#for translating the User SID to the AD user name
				$Results = Get-RDSGroup -SiteId $Site.Id -EA 0 4>$Null
				If( $Results.RDSIds -Contains $RDSHost.Id )
				{
					#does this group inherit default settings?
					If($Results.InheritDefaultDesktopAccessSettings -eq $False)
					{
						#no we don't, so get the default settings for the group
						$GroupDefaults = $Results.RDSDefSettings

						$RDSRestrictDesktopAccess = $GroupDefaults.RestrictDesktopAccess.ToString()
						$RDSRestrictedUsers = @()
						
						ForEach($User in $GroupDefaults.RestrictedUsers)
						{
							$objSID = New-Object System.Security.Principal.SecurityIdentifier ($User)
							$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
							
							$RDSRestrictedUsers += $objUser.Value
						}
					}
					Else
					{
						#yes we do, get the default settings for the Site
						#use the Site default settings
						$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
						
						If($? -and $Null -ne $RDSDefaults)
						{
							$RDSRestrictDesktopAccess = $RDSDefaults.RestrictDesktopAccess.ToString()
							$RDSRestrictedUsers = @()
							
							ForEach($User in $RDSDefaults.RestrictedUsers)
							{
								$objSID = New-Object System.Security.Principal.SecurityIdentifier ($User)
								$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
								
								$RDSRestrictedUsers += $objUser.Value
							}
						}
						Else
						{
							#unable to retrieve default, use built-in default values
							$RDSRestrictDesktopAccess = "False"
							$RDSRestrictedUsers = @()
						}
					}
				}
				Else
				{
					#RDS Host is not in a group
					#get the settings for the host
					$RDSRestrictDesktopAccess = $RDSHost.RestrictDesktopAccess.ToString()
					$RDSRestrictedUsers = @()
					
					ForEach($User in $RDSHost.RestrictedUsers)
					{
						$objSID = New-Object System.Security.Principal.SecurityIdentifier ($User)
						$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
						
						$RDSRestrictedUsers += $objUser.Value
					}
				}
			}
			Else
			{
				#we don't inherit
				#get the settings for the host
				$RDSRestrictDesktopAccess = $RDSHost.RestrictDesktopAccess.ToString()
				$RDSRestrictedUsers = @()
				
				ForEach($User in $RDSHost.RestrictedUsers)
				{
					$objSID = New-Object System.Security.Principal.SecurityIdentifier ($User)
					$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
					
					$RDSRestrictedUsers += $objUser.Value
				}
			}
				
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSHost.InheritDefaultDesktopAccessSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Restrict direct desktop access to the following users"; Value = $RDSRestrictDesktopAccess; }) > $Null
				
				$cnt = -1
				ForEach($Item in $RDSRestrictedUsers)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Users"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Inherit default settings`t`t`t`t: " $RDSHost.InheritDefaultDesktopAccessSettings.ToString()
				Line 4 "Restrict direct desktop access to the following users`t: " $RDSRestrictDesktopAccess
				$cnt = -1
				ForEach($Item in $RDSRestrictedUsers)
				{
					$cnt++
					If($cnt -eq 0)
					{
						Line 10 "Users`t: " $Item
					}
					Else
					{
						Line 11 "  " $Item
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSHost.InheritDefaultDesktopAccessSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Restrict direct desktop access to the following users",($Script:htmlsb),$RDSRestrictDesktopAccess,$htmlwhite))
				
				$cnt = -1
				ForEach($Item in $RDSRestrictedUsers)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$rowdata += @(,( "Users",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,( "",($Script:htmlsb),$Item,$htmlwhite))
					}
				}

				$msg = "Desktop Access"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#RDP Printer
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "RDP Printer"
			}
			If($Text)
			{
				Line 3 "RDP Printer"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "RDP Printer"
			}
			
			If($RDSHost.InheritDefaultPrinterSettings)
			{
				#do we inherit group or site defaults?
				#is this RDS host in a group?
				$Results = Get-RDSGroup -SiteId $Site.Id -EA 0 4>$Null
				If( $Results.RDSIds -Contains $RDSHost.Id )
				{
					#does this group inherit default settings?
					If($Results.InheritDefaultPrinterSettings -eq $False)
					{
						#no we don't, so get the default settings for the group
						$RDSDefaults = $Results.RDSDefSettings

						Switch ($RDSDefaults.PrinterNameFormat)
						{
							"PrnFormat_PRN_CMP_SES"	{$RDSPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
							"PrnFormat_SES_CMP_PRN"	{$RDSPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
							"PrnFormat_PRN_REDSES"	{$RDSPrinterNameFormat = "Printername (redirected Session no.)"; Break}
							Default					{$RDSPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($RDSDefaults.PrinterNameFormat)"; Break}
						}
						
						$RDSRemoveSessionNumberFromPrinter = $GroupDefaults.RemoveSessionNumberFromPrinterName.ToString()
					}
					Else
					{
						#yes we do, get the default settings for the Site
						#use the Site default settings
						$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
						
						If($? -and $Null -ne $RDSDefaults)
						{
							Switch ($RDSDefaults.PrinterNameFormat)
							{
								"PrnFormat_PRN_CMP_SES"	{$RDSPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
								"PrnFormat_SES_CMP_PRN"	{$RDSPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
								"PrnFormat_PRN_REDSES"	{$RDSPrinterNameFormat = "Printername (redirected Session no.)"; Break}
								Default					{$RDSPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($RDSDefaults.PrinterNameFormat)"; Break}
							}
							
							$RDSRemoveSessionNumberFromPrinter = $RDSDefaults.RemoveSessionNumberFromPrinterName.ToString()
						}
						Else
						{
							#unable to retrieve default, use built-in default values
							$RDSPrinterNameFormat = "Printername (from Computername) in Session no."
							$RDSRemoveSessionNumberFromPrinter = "False"
						}
					}
				}
				Else
				{
					#RDS Host is not in a group
					#get the settings for the host
					Switch ($RDSHost.PrinterNameFormat)
					{
						"PrnFormat_PRN_CMP_SES"	{$RDSPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
						"PrnFormat_SES_CMP_PRN"	{$RDSPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
						"PrnFormat_PRN_REDSES"	{$RDSPrinterNameFormat = "Printername (redirected Session no.)"; Break}
						Default					{$RDSPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($RDSHost.PrinterNameFormat)"; Break}
					}
					
					$RDSRemoveSessionNumberFromPrinter = $RDSHost.RemoveSessionNumberFromPrinterName.ToString()
				}
			}
			Else
			{
				#we don't inherit
				#get the settings for the host
				Switch ($RDSHost.PrinterNameFormat)
				{
					"PrnFormat_PRN_CMP_SES"	{$RDSPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
					"PrnFormat_SES_CMP_PRN"	{$RDSPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
					"PrnFormat_PRN_REDSES"	{$RDSPrinterNameFormat = "Printername (redirected Session no.)"; Break}
					Default					{$RDSPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($RDSHost.PrinterNameFormat)"; Break}
				}
				
				$RDSRemoveSessionNumberFromPrinter = $RDSHost.RemoveSessionNumberFromPrinterName.ToString()
			}

			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSHost.InheritDefaultPrinterSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "RDP Printer Name Format"; Value = $RDSPrinterNameFormat; }) > $Null
				$ScriptInformation.Add(@{Data = "Remove session number from printer name"; Value = $RDSRemoveSessionNumberFromPrinter; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Inherit default settings`t`t`t`t: " $RDSHost.InheritDefaultPrinterSettings.ToString()
				Line 4 "RDP Printer Name Format`t`t`t`t`t: " $RDSPrinterNameFormat
				Line 4 "Remove session number from printer name`t`t`t: " $RDSRemoveSessionNumberFromPrinter
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSHost.InheritDefaultPrinterSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "RDP Printer Name Format",($Script:htmlsb),$RDSPrinterNameFormat,$htmlwhite))
				$rowdata += @(,( "Remove session number from printer name",($Script:htmlsb),$RDSRemoveSessionNumberFromPrinter,$htmlwhite))

				$msg = "RDP Printer"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}

	Write-Verbose "$(Get-Date): `t`tOutput Site RD Session Host Groups for Site $($Site.Name)"
	$RDSGroups = Get-RDSGroup -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve RD Session Host Groups for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve RD Session Host Groups for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve RD Session Host Groups for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve RD Session Host Groups for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No RD Session Host Groups retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No RD Session Host Groups retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No RD Session Host Groups retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No RD Session Host Groups retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Groups"
		}
		If($Text)
		{
			Line 1 "Groups"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Groups"
		}

		ForEach($RDSGroup in $RDSGroups)
		{
			Write-Verbose "$(Get-Date): `t`t`tOutput Site RD Session Host Group $($RDSGroup.Name)"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Group $($RDSGroup.Name)"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $RDSGroup.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Members"; Value = $RDSGroup.RDSIds.Count.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $RDSGroup.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "ID"; Value = $RDSGroup.Id; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Name`t`t: " $RDSGroup.Name
				Line 2 "Members`t`t: " $RDSGroup.RDSIds.Count.ToString()
				Line 2 "Description`t: " $RDSGroup.Description
				Line 2 "ID`t`t: " $RDSGroup.Id
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Group $($RDSGroup.Name)"
				$rowdata = @()
				$columnHeaders = @("Server",($Script:htmlsb),$RDSGroup.Name,$htmlwhite)
				$rowdata += @(,( "Members",($Script:htmlsb),$RDSGroup.RDSIds.Count.ToString(),$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$RDSGroup.Description,$htmlwhite))
				$rowdata += @(,( "ID",($Script:htmlsb),$RDSGroup.Id,$htmlwhite))

				$msg = ""
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#General
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "General"
			}
			If($Text)
			{
				Line 2 "General"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "General"
			}
			
			#get any group members
			$RDSGroupMembers = @(Get-RDSGroupMember -GroupId $RDSGroup.Id -EA 0 4>$Null)
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Enable Group in site"; Value = $RDSGroup.Enabled.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Name"; Value = $RDSGroup.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $RDSGroup.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "RD session hosts based on a template"; Value = $RDSGroup.UseRASTemplate; }) > $Null
				$ScriptInformation.Add(@{Data = "RAS template"; Value = $RDSGroup.RASTemplateId; }) > $Null
				$ScriptInformation.Add(@{Data = "Group Members"; Value = ""; }) > $Null
				If($RDSGroupMembers.Count -gt 0)
				{
					$cnt=-1
					ForEach($RDSGroupMember in $RDSGroupMembers)
					{
						$cnt++
						$ScriptInformation.Add(@{Data = "     Member"; Value = $RDSGroupMember.Server; }) > $Null
						$ScriptInformation.Add(@{Data = "     Logon status"; Value = $RDSGroupMember.Enabled.ToString(); }) > $Null
						$ScriptInformation.Add(@{Data = "     Type"; Value = "Server"; }) > $Null
						$ScriptInformation.Add(@{Data = "     Description"; Value = $RDSGroupMember.Description; }) > $Null
						$ScriptInformation.Add(@{Data = ""; Value = ""; }) > $Null
					}
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "No Group Members Found"; Value = ""; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Enable Group in site`t`t`t: " $RDSGroup.Enabled.ToString()
				Line 3 "Name`t`t`t`t`t: " $RDSGroup.Name
				Line 3 "Description`t`t`t`t: " $RDSGroup.Description
				Line 3 "RD session hosts based on a template`t: " $RDSGroup.UseRASTemplate
				Line 3 "RAS template`t`t`t`t: " $RDSGroup.RASTemplateId
				Line 3 "Group Members" ""
				If($RDSGroupMembers.Count -gt 0)
				{
					$cnt=-1
					ForEach($RDSGroupMember in $RDSGroupMembers)
					{
						$cnt++
						Line 4 "Member`t`t: " $RDSGroupMember.Server
						Line 4 "Logon status`t: " $RDSGroupMember.Enabled.ToString()
						Line 4 "Type`t`t: " "Server"
						Line 4 "Description`t: " $RDSGroupMember.Description
						Line 4 ""
					}
				}
				Else
				{
					Line 4 "No Group Members Found" ""
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable Group in site",($Script:htmlsb),$RDSGroup.Enabled.ToString(),$htmlwhite)
				$rowdata += @(,( "Name",($Script:htmlsb),$RDSGroup.Name,$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$RDSGroup.Description,$htmlwhite))
				$rowdata += @(,( "RD session hosts based on a template",($Script:htmlsb),$RDSGroup.UseRASTemplate.ToString(),$htmlwhite))
				$rowdata += @(,( "RAS template",($Script:htmlsb),$RDSGroup.RASTemplateId.ToString(),$htmlwhite))
				$rowdata += @(,( "Group Members",($Script:htmlsb),"",$htmlwhite))
				If($RDSGroupMembers.Count -gt 0)
				{
					$cnt=-1
					ForEach($RDSGroupMember in $RDSGroupMembers)
					{
						$cnt++
						$rowdata += @(,( "     Member",($Script:htmlsb),$RDSGroupMember.Server,$htmlwhite))
						$rowdata += @(,( "     Logon status",($Script:htmlsb),$RDSGroupMember.Enabled.ToString(),$htmlwhite))
						$rowdata += @(,( "     Type",($Script:htmlsb),"Server",$htmlwhite))
						$rowdata += @(,( "     Description",($Script:htmlsb),$RDSGroupMember.Description,$htmlwhite))
						$rowdata += @(,( "",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					$rowdata += @(,( "No Group Members Found",($Script:htmlsb),"",$htmlwhite))
				}

				$msg = "General"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Agent Settings
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Agent Settings"
			}
			If($Text)
			{
				Line 2 "Agent Settings"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Agent Settings"
			}
			
			If($RDSGroup.InheritDefaultAgentSettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $RDSDefaults)
				{
					$RDSPort = $RDSDefaults.Port.ToString()
					$RDSMaxSessions = $RDSDefaults.MaxSessions.ToString()
					
					Switch ($RDSDefaults.SessionTimeout)
					{
						0		{$RDSPublishingSessionDisconnectTimeout = "Never"; Break}
						25		{$RDSPublishingSessionDisconnectTimeout = "25 seconds"; Break}
						60		{$RDSPublishingSessionDisconnectTimeout = "1 minute"; Break}
						300		{$RDSPublishingSessionDisconnectTimeout = "5 minutes"; Break}
						3600	{$RDSPublishingSessionDisconnectTimeout = "1 hour"; Break}
						Default	{$RDSPublishingSessionDisconnectTimeout = "Unable to determine Publishing Session Disconnect Timeout: $($RDSDefaults.SessionTimeout)"; Break}
					}
					
					Switch ($RDSDefaults.SessionLogoffTimeout)
					{
						0		{$RDSPublishingSessionResetTime = "Never"; Break}
						1		{$RDSPublishingSessionResetTime = "Immediate"; Break}
						25		{$RDSPublishingSessionResetTime = "25 seconds"; Break}
						60		{$RDSPublishingSessionResetTime = "1 minute"; Break}
						300		{$RDSPublishingSessionResetTime = "5 minutes"; Break}
						3600	{$RDSPublishingSessionResetTime = "1 hour"; Break}
						Default	{$RDSPublishingSessionResetTime = "Unable to determine Publishing Session Reset Timeout: $($RDSDefaults.SessionLogoffTimeout)"; Break}
					}
					
					Switch($RDSDefaults.AllowURLAndMailRedirection)
					{
						"Enabled"						{$RDSAllowClientURLMailRedirection = "Enabled"; Break}
						"Disbaled"						{$RDSAllowClientURLMailRedirection = "Disabled"; Break}
						"EnabledWithAppRegistration"	{$RDSAllowClientURLMailRedirection = "Enabled (Replace Registered Application)"; Break}
						Default 						{$RDSAllowClientURLMailRedirection = "Unable to determine Allow CLient URL/Mail Redirection: $($RDSDefaults.AllowURLAndMailRedirection)"; Break}
					}
					
					$RDSSupportShellURLNamespaceObject = $RDSDefaults.SupportShellURLNamespaceObjects.ToString()
					
					Switch ($RDSDefaults.DragAndDropMode)
					{
						"Bidirectional"		{$RDSDragAndDrop = "Bidirectional"; Break}
						"Disabled"			{$RDSDragAndDrop = "Disabled"; Break}
						"ClientToServer"	{$RDSDragAndDrop = "Client to server only"; Break}
						"ServerToClient"	{$RDSDragAndDrop = "Server to client only"; Break}
						Default				{$RDSDragAndDrop = "Unable to determine Drag and drop: $($RDSDefaults.DragAndDropMode)"; Break}
					}
					
					If($RDSDefaults.PreferredPAId -eq 0)
					{
						$RDSPreferredPublishingAgent = "Automatically"
					}
					Else
					{
						$RDSPreferredPublishingAgent = (Get-PA -Id $RDSDefaults.PreferredPAId -EA 0 4>$Null).Server
					}
					$RDSAllowRemoteExec = $RDSDefaults.AllowRemoteExec.ToString()
					$RDSUseRemoteApps = $RDSDefaults.UseRemoteApps.ToString()
					$RDSEnableAppMonitoring = $RDSDefaults.EnableAppMonitoring.ToString()
					$RDSAllowFileTransfer = $RDSDefaults.AllowFileTransfer.ToString()
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$RDSPort = "3389"
					$RDSMaxSessions = "250"
					$RDSPublishingSessionDisconnectTimeout = "25 seconds"
					$RDSPublishingSessionResetTime = "Immediate"
					$RDSAllowClientURLMailRedirection = "Enabled"
					$RDSSupportShellURLNamespaceObject = "True"
					$RDSDragAndDrop = "Bidirectional"
					$RDSPreferredPublishingAgent = "Automatically"
					$RDSAllowRemoteExec = "True"
					$RDSUseRemoteApps = "False"
					$RDSEnableAppMonitoring = "True"
					$RDSAllowFileTransfer = "True"
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this RDS group
				$RDSGroupDefaults = $RDSGroup.RDSDefSettings
				$RDSPort = $RDSGroupDefaults.Port.ToString()
				$RDSMaxSessions = $RDSGroupDefaults.MaxSessions.ToString()
				
				Switch ($RDSGroupDefaults.SessionTimeout)
				{
					0		{$RDSPublishingSessionDisconnectTimeout = "Never"; Break}
					25		{$RDSPublishingSessionDisconnectTimeout = "25 seconds"; Break}
					60		{$RDSPublishingSessionDisconnectTimeout = "1 minute"; Break}
					300		{$RDSPublishingSessionDisconnectTimeout = "5 minutes"; Break}
					3600	{$RDSPublishingSessionDisconnectTimeout = "1 hour"; Break}
					Default	{$RDSPublishingSessionDisconnectTimeout = "Unable to determine Publishing Session Disconnect Timeout: $($RDSGroup.SessionTimeout)"; Break}
				}
				
				Switch ($RDSGroupDefaults.SessionLogoffTimeout)
				{
					0		{$RDSPublishingSessionResetTime = "Never"; Break}
					1		{$RDSPublishingSessionResetTime = "Immediate"; Break}
					25		{$RDSPublishingSessionResetTime = "25 seconds"; Break}
					60		{$RDSPublishingSessionResetTime = "1 minute"; Break}
					300		{$RDSPublishingSessionResetTime = "5 minutes"; Break}
					3600	{$RDSPublishingSessionResetTime = "1 hour"; Break}
					Default	{$RDSPublishingSessionResetTime = "Unable to determine Publishing Session Reset Timeout: $($RDSGroup.SessionLogoffTimeout)"; Break}
				}
				
				Switch($RDSGroupDefaults.AllowURLAndMailRedirection)
				{
					"Enabled"						{$RDSAllowClientURLMailRedirection = "Enabled"; Break}
					"Disbaled"						{$RDSAllowClientURLMailRedirection = "Disabled"; Break}
					"EnabledWithAppRegistration"	{$RDSAllowClientURLMailRedirection = "Enabled (Replace Registered Application)"; Break}
					Default 						{$RDSAllowClientURLMailRedirection = "Unable to determine Allow CLient URL/Mail Redirection: $($RDSGroup.AllowURLAndMailRedirection)"; Break}
				}
				
				$RDSSupportShellURLNamespaceObject = $RDSGroupDefaults.SupportShellURLNamespaceObjects.ToString()
				
				Switch ($RDSGroupDefaults.DragAndDropMode)
				{
					"Bidirectional"		{$RDSDragAndDrop = "Bidirectional"; Break}
					"Disabled"			{$RDSDragAndDrop = "Disabled"; Break}
					"ClientToServer"	{$RDSDragAndDrop = "Client to server only"; Break}
					"ServerToClient"	{$RDSDragAndDrop = "Server to client only"; Break}
					Default				{$RDSDragAndDrop = "Unable to determine Drag and drop: $($RDSGroup.DragAndDropMode)"; Break}
				}
				
				If($RDSGroupDefaults.PreferredPAId -eq 0)
				{
					$RDSPreferredPublishingAgent = "Automatically"
				}
				Else
				{
					$RDSPreferredPublishingAgent = (Get-PA -Id $RDSGroupDefaults.PreferredPAId -EA 0 4>$Null).Server
				}
				$RDSAllowRemoteExec = $RDSGroupDefaults.AllowRemoteExec.ToString()
				$RDSUseRemoteApps = $RDSGroupDefaults.UseRemoteApps.ToString()
				$RDSEnableAppMonitoring = $RDSGroupDefaults.EnableAppMonitoring.ToString()
				$RDSAllowFileTransfer = $RDSGroupDefaults.AllowFileTransfer.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSGroup.InheritDefaultAgentSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Port"; Value = $RDSPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Max Sessions"; Value = $RDSMaxSessions; }) > $Null
				$ScriptInformation.Add(@{Data = "Publishing Session Disconnect Timeout"; Value = $RDSPublishingSessionDisconnectTimeout; }) > $Null
				$ScriptInformation.Add(@{Data = "Publishing Session Reset Timeout"; Value = $RDSPublishingSessionResetTime; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow Client URL/Mail Redirection"; Value = $RDSAllowClientURLMailRedirection; }) > $Null
				$ScriptInformation.Add(@{Data = "Support Shell URL namespace objects"; Value = $RDSSupportShellURLNamespaceObject; }) > $Null
				$ScriptInformation.Add(@{Data = "Drag and drop"; Value = $RDSDragAndDrop; }) > $Null
				$ScriptInformation.Add(@{Data = "Preferred Publishing Agent"; Value = $RDSPreferredPublishingAgent; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow 2xRemoteExec to send command to the client"; Value = $RDSAllowRemoteExec; }) > $Null
				$ScriptInformation.Add(@{Data = "Use RemoteApp if available"; Value = $RDSUseRemoteApps; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable applications monitoring"; Value = $RDSEnableAppMonitoring; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow file transfer command (HTML5 and Chrome clients)"; Value = $RDSAllowFileTransfer; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings`t`t`t`t: " $RDSGroup.InheritDefaultAgentSettings.ToString()
				Line 3 "Port`t`t`t`t`t`t`t: " $RDSPort
				Line 3 "Max Sessions`t`t`t`t`t`t: " $RDSMaxSessions
				Line 3 "Publishing Session Disconnect Timeout`t`t`t: " $RDSPublishingSessionDisconnectTimeout
				Line 3 "Publishing Session Reset Timeout`t`t`t: " $RDSPublishingSessionResetTime
				Line 3 "Allow Client URL/Mail Redirection`t`t`t: " $RDSAllowClientURLMailRedirection
				Line 3 "Support Shell URL namespace objects`t`t`t: " $RDSSupportShellURLNamespaceObject
				Line 3 "Drag and drop`t`t`t`t`t`t: " $RDSDragAndDrop
				Line 3 "Preferred Publishing Agent`t`t`t`t: " $RDSPreferredPublishingAgent
				Line 3 "Allow 2xRemoteExec to send command to the client`t: " $RDSAllowRemoteExec
				Line 3 "Use RemoteApp if available`t`t`t`t: " $RDSUseRemoteApps
				Line 3 "Enable applications monitoring`t`t`t`t: " $RDSEnableAppMonitoring
				Line 3 "Allow file transfer command (HTML5 and Chrome clients)`t: " $RDSAllowFileTransfer
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSGroup.InheritDefaultAgentSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Port",($Script:htmlsb),$RDSPort,$htmlwhite))
				$rowdata += @(,( "Max Sessions",($Script:htmlsb),$RDSMaxSessions,$htmlwhite))
				$rowdata += @(,( "Publishing Session Disconnect Timeout",($Script:htmlsb),$RDSPublishingSessionDisconnectTimeout,$htmlwhite))
				$rowdata += @(,( "Publishing Session Reset Timeout",($Script:htmlsb),$RDSPublishingSessionResetTime,$htmlwhite))
				$rowdata += @(,( "Allow Client URL/Mail Redirection",($Script:htmlsb),$RDSAllowClientURLMailRedirection,$htmlwhite))
				$rowdata += @(,( "Support Shell URL namespace objects",($Script:htmlsb),$RDSSupportShellURLNamespaceObject,$htmlwhite))
				$rowdata += @(,( "Drag and drop",($Script:htmlsb),$RDSDragAndDrop,$htmlwhite))
				$rowdata += @(,( "Preferred Publishing Agent",($Script:htmlsb),$RDSPreferredPublishingAgent,$htmlwhite))
				$rowdata += @(,( "Allow 2xRemoteExec to send command to the client",($Script:htmlsb),$RDSAllowRemoteExec,$htmlwhite))
				$rowdata += @(,( "Use RemoteApp if available",($Script:htmlsb),$RDSUseRemoteApps,$htmlwhite))
				$rowdata += @(,( "Enable applications monitoring",($Script:htmlsb),$RDSEnableAppMonitoring,$htmlwhite))
				$rowdata += @(,( "Allow file transfer command (HTML5 and Chrome clients)",($Script:htmlsb),$RDSAllowFileTransfer,$htmlwhite))

				$msg = "Agent Settings"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#User Profile Disks
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "User Profile Disks"
			}
			If($Text)
			{
				Line 2 "User Profile Disks"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "User Profile Disks"
			}
			
			If($RDSGroup.InheritDefaultUPDSettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings
				$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $RDSDefaults)
				{
					Switch ($RDSDefaults.UPDMode)
					{
						"DoNotChange"	{$RDSUPDState = "Do not change"; Break}
						"Enabled"		{$RDSUPDState = "Enabled"; Break}
						"Disabled"		{$RDSUPDState = "Disabled"; Break}
						Default			{$RDSUPDState = "Unable to determine Current UPD State: $($RDSDefaults.UPDMode)"; Break}
					}
					
					$RDSUPDLocation = $RDSDefaults.DiskPath
					$RDSUPDSize = $RDSDefaults.MaxUserProfileDiskSizeGB
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$RDSUPDState = "Do not change"
					$RDSUPDLocation = ""
					$RDSUPDSize = "20"
				}
			}
			Else
			{
				#we don't inherit
				#get the settings for the group
				$RDSGroupDefaults = $RDSGroup.RDSDefSettings
				Switch ($RDSGroupDefaults.UPDMode)
				{
					"DoNotChange"	{$RDSUPDState = "Do not change"; Break}
					"Enabled"		{$RDSUPDState = "Enabled"; Break}
					"Disabled"		{$RDSUPDState = "Disabled"; Break}
					Default			{$RDSUPDState = "Unable to determine Current UPD State: $($RDSGroup.UPDMode)"; Break}
				}
				
				$RDSUPDLocation = $RDSGroupDefaults.DiskPath
				$RDSUPDSize = $RDSGroupDefaults.MaxUserProfileDiskSizeGB.ToString()
			}
				
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSGroup.InheritDefaultUPDSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "UPD State"; Value = $RDSUPDState; }) > $Null
				$ScriptInformation.Add(@{Data = "Location of user profile disks"; Value = $RDSUPDLocation; }) > $Null
				$ScriptInformation.Add(@{Data = "Maximum size (in GB)"; Value = $RDSUPDSize; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings`t`t`t`t: " $RDSGroup.InheritDefaultUPDSettings.ToString()
				Line 3 "UPD State`t`t`t`t`t`t: " $RDSUPDState
				Line 3 "Location of user profile disks`t`t`t`t: " $RDSUPDLocation
				Line 3 "Maximum size (in GB)`t`t`t`t`t: " $RDSUPDSize
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSGroup.InheritDefaultUPDSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "UPD State",($Script:htmlsb),$RDSUPDState,$htmlwhite))
				$rowdata += @(,( "Location of user profile disks",($Script:htmlsb),$RDSUPDLocation,$htmlwhite))
				$rowdata += @(,( "Maximum size (in GB)",($Script:htmlsb),$RDSUPDSize,$htmlwhite))

				$msg = "User Profile Disks"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Desktop Access
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Desktop Access"
			}
			If($Text)
			{
				Line 2 "Desktop Access"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Desktop Access"
			}
			
			If($RDSGroup.InheritDefaultDesktopAccessSettings)
			{
				#do we inherit group or site defaults?
				#http://woshub.com/hot-to-convert-sid-to-username-and-vice-versa/
				#for translating the User SID to the AD user name
				#yes we do, get the default settings for the Site
				#use the Site default settings
				$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $RDSDefaults)
				{
					$RDSRestrictDesktopAccess = $RDSDefaults.RestrictDesktopAccess.ToString()
					$RDSRestrictedUsers = @()
					
					ForEach($User in $RDSDefaults.RestrictedUsers)
					{
						$objSID = New-Object System.Security.Principal.SecurityIdentifier ($User)
						$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
						
						$RDSRestrictedUsers += $objUser.Value
					}
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$RDSRestrictDesktopAccess = "False"
					$RDSRestrictedUsers = @()
				}
			}
			Else
			{
				#we don't inherit
				#get the settings for the group
				$RDSGroupDefaults = $RDSGroup.RDSDefSettings
				$RDSRestrictDesktopAccess = $RDSGroupDefaults.RestrictDesktopAccess.ToString()
				$RDSRestrictedUsers = @()
				
				ForEach($User in $RDSGroupDefaults.RestrictedUsers)
				{
					$objSID = New-Object System.Security.Principal.SecurityIdentifier ($User)
					$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
					
					$RDSRestrictedUsers += $objUser.Value
				}
			}
				
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSGroup.InheritDefaultDesktopAccessSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Restrict direct desktop access to the following users"; Value = $RDSRestrictDesktopAccess; }) > $Null
				
				$cnt = -1
				ForEach($Item in $RDSRestrictedUsers)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Users"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings`t`t`t`t: " $RDSGroup.InheritDefaultDesktopAccessSettings.ToString()
				Line 3 "Restrict direct desktop access to the following users`t: " $RDSRestrictDesktopAccess
				$cnt = -1
				ForEach($Item in $RDSRestrictedUsers)
				{
					$cnt++
					If($cnt -eq 0)
					{
						Line 9 "Users`t: " $Item
					}
					Else
					{
						Line 10 "  " $Item
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSGroup.InheritDefaultDesktopAccessSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Restrict direct desktop access to the following users",($Script:htmlsb),$RDSRestrictDesktopAccess,$htmlwhite))
				
				$cnt = -1
				ForEach($Item in $RDSRestrictedUsers)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$rowdata += @(,( "Users",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,( "",($Script:htmlsb),$Item,$htmlwhite))
					}
				}

				$msg = "Desktop Access"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#RDP Printer
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "RDP Printer"
			}
			If($Text)
			{
				Line 2 "RDP Printer"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "RDP Printer"
			}
			
			If($RDSGroup.InheritDefaultPrinterSettings)
			{
				#do we inherit group or site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings
				$RDSDefaults = Get-RDSDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $RDSDefaults)
				{
					Switch ($RDSDefaults.PrinterNameFormat)
					{
						"PrnFormat_PRN_CMP_SES"	{$RDSPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
						"PrnFormat_SES_CMP_PRN"	{$RDSPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
						"PrnFormat_PRN_REDSES"	{$RDSPrinterNameFormat = "Printername (redirected Session no.)"; Break}
						Default					{$RDSPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($RDSFefaults.PrinterNameFormat)"; Break}
					}
					
					$RDSRemoveSessionNumberFromPrinter = $RDSDefaults.RemoveSessionNumberFromPrinterName.ToString()
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$RDSPrinterNameFormat = "Printername (from Computername) in Session no."
					$RDSRemoveSessionNumberFromPrinter = "False"
				}
			}
			Else
			{
				#we don't inherit
				#get the settings for the group
				$RDSGroupDefaults = $RDSGroup.RDSDefSettings
				Switch ($RDSGroupDefaults.PrinterNameFormat)
				{
					"PrnFormat_PRN_CMP_SES"	{$RDSPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
					"PrnFormat_SES_CMP_PRN"	{$RDSPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
					"PrnFormat_PRN_REDSES"	{$RDSPrinterNameFormat = "Printername (redirected Session no.)"; Break}
					Default					{$RDSPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($RDSFefaults.PrinterNameFormat)"; Break}
				}
				
				$RDSRemoveSessionNumberFromPrinter = $RDSGroupDefaults.RemoveSessionNumberFromPrinterName.ToString()
			}

			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $RDSGroup.InheritDefaultPrinterSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "RDP Printer Name Format"; Value = $RDSPrinterNameFormat; }) > $Null
				$ScriptInformation.Add(@{Data = "Remove session number from printer name"; Value = $RDSRemoveSessionNumberFromPrinter; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings`t`t`t`t: " $RDSGroup.InheritDefaultPrinterSettings.ToString()
				Line 3 "RDP Printer Name Format`t`t`t`t`t: " $RDSPrinterNameFormat
				Line 3 "Remove session number from printer name`t`t`t: " $RDSRemoveSessionNumberFromPrinter
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$RDSGroup.InheritDefaultPrinterSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "RDP Printer Name Format",($Script:htmlsb),$RDSPrinterNameFormat,$htmlwhite))
				$rowdata += @(,( "Remove session number from printer name",($Script:htmlsb),$RDSRemoveSessionNumberFromPrinter,$htmlwhite))

				$msg = "RDP Printer"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}

	Write-Verbose "$(Get-Date): `t`tOutput Site RD Session Host Scheduler for Site $($Site.Name)"
	$RDSSchedules = Get-RDSSchedule -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve RD Session Host Scheduler for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve RD Session Host Scheduler for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve RD Session Host Scheduler for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve RD Session Host Scheduler for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No RD Session Host Scheduler retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No RD Session Host Scheduler retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No RD Session Host Scheduler retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No RD Session Host Scheduler retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Scheduler"
		}
		If($Text)
		{
			Line 1 "Scheduler"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Scheduler"
		}

		ForEach($RDSSchedule in $RDSSchedules)
		{
			Write-Verbose "$(Get-Date): `t`t`tOutput Site RD Session Host Scheduler $($RDSSchedule.Name)"
			$Action = $RDSSchedule.Action
			If($RDSSChedule.Action -eq "Reboot")
			{
				If($RDSSchedule.DrainMode)
				{
					$Action = "Reboot - Drain Mode"
				}
				Else
				{
					$Action = "Reboot"
				}
			}

			If($Action -eq "Reboot - Drain Mode")
			{
				Switch ($RDSSchedule.CompleteRebootInSecs)
				{
					600		{$TimeDuration = "10 minutes"; Break}
					900		{$TimeDuration = "15 minutes"; Break}
					1800	{$TimeDuration = "30 minutes"; Break}
					2700	{$TimeDuration = "45 minutes"; Break}
					3600	{$TimeDuration = "1 hour"; Break}
					7200	{$TimeDuration = "2 hours"; Break}
					10800	{$TimeDuration = "3 hours"; Break}
					Default	{$TimeDuration = "Unable to determine Complete in seconds: $($RDSSchedule.CompleteRebootInSecs)"; Break}
				}
				
				Switch ($RDSSchedule.ForceRebootAfterSecs)
				{
					900		{$ForceRebootTime = "15 minutes"; Break}
					1800	{$ForceRebootTime = "30 minutes"; Break}
					2700	{$ForceRebootTime = "45 minutes"; Break}
					3600	{$ForceRebootTime = "1 hour"; Break}
					7200	{$ForceRebootTime = "2 hours"; Break}
					10800	{$ForceRebootTime = "3 hours"; Break}
					21600	{$ForceRebootTime = "6 hours"; Break}
					43200	{$ForceRebootTime = "12 hours"; Break}
					86400	{$ForceRebootTime = "1 day"; Break}
					Default	{$ForceRebootTime = "Unable to determine Force reboot after seconds: $($RDSSchedule.ForceRebootAfterSecs)"; Break}
				}
			}
			ElseIf($Action -eq "Disable")
			{
				Switch ($RDSSchedule.DurationInSecs)
				{
					600		{$TimeDuration = "10 minutes"; Break}
					900		{$TimeDuration = "15 minutes"; Break}
					1800	{$TimeDuration = "30 minutes"; Break}
					2700	{$TimeDuration = "45 minutes"; Break}
					3600	{$TimeDuration = "1 hour"; Break}
					7200	{$TimeDuration = "2 hours"; Break}
					10800	{$TimeDuration = "3 hours"; Break}
					21600	{$TimeDuration = "6 hours"; Break}
					43200	{$TimeDuration = "12 hours"; Break}
					86400	{$TimeDuration = "1 day"; Break}
					Default	{$TimeDuration = "Unable to determine Duration in seconds: $($RDSSchedule.DurationInSecs)"; Break}
				}
				
				Switch ($RDSSchedule.DisableAction)
				{
					"KeepSessionState"			{$OnDisable = "Keep current sessions state"; Break}
					"DisconnectActiveSessions"	{$OnDisable = "Disconnect any active session"; Break}
					"ResetAllSessions"			{$OnDisable = "Reset all sessions"; Break}
					Default						{$OnDisable = "Unable to determine On disable: $($RDSSchedule.DisableAction)"; Break}
				}
			}
			
			Switch ($RDSSchedule.Repeat)
			{
				Never			{$Repeat = "Never "; Break}
				EveryDay		{$Repeat = "Every day"; Break}
				EveryWeek		{$Repeat = "Every week"; Break}
				Every2Weeks		{$Repeat = "Every 2 weeks"; Break}
				EveryMonth		{$Repeat = "Every month"; Break}
				EveryYear		{$Repeat = "Every year"; Break}
				SpecificDays	{$Repeat = "Every $($RDSSchedule.SpecificDays)"; Break}
				Default			{$Repeat = "Unable to determine the Repeat: $($RDSSchedule.Repeat)"; Break}
			}
			
			$Target = @()
			If($RDSSchedule.TargetType -eq "Server")
			{
				ForEach($Item in $RDSSchedule.TargetIds)
				{
					$Result = Get-RDS -Id $Item -EA 0 4>$Null
					
					If($? -and $Null -ne $Result)
					{
						$Target += $Result.Server
					}
					Else
					{
						Target += "Unable to find RDS Server for ID $($Item)"
					}
				}
			}
			ElseIf($RDSSchedule.TargetType -eq "ServerGroup")
			{
				ForEach($Item in $RDSSchedule.TargetIds)
				{
					$Result = Get-RDSGroup -Id $Item -EA 0 4>$Null
					
					If($? -and $Null -ne $Result)
					{
						$Target += $Result.Name
					}
					Else
					{
						Target += "Unable to find RDS Server Group for ID $($Item)"
					}
				}
			}
			Else
			{
				Target += "Unable to determine Target for TargetType: $($RDSSchedule.TargetType)"
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Schedule Name $($RDSSchedule.Name)"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $RDSSchedule.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Action"; Value = $Action; }) > $Null
				
				$cnt=-1
				ForEach($Item in $Target)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Target"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				
				$ScriptInformation.Add(@{Data = "Start"; Value = $RDSSchedule.StartDateTime; }) > $Null
				$ScriptInformation.Add(@{Data = "Repeat"; Value = $Repeat; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $RDSSchedule.Description; }) > $Null
				#$ScriptInformation.Add(@{Data = "Last modification by"; Value = ""; }) > $Null
				#$ScriptInformation.Add(@{Data = "Modified on"; Value = ""; }) > $Null
				#$ScriptInformation.Add(@{Data = "Created by"; Value = ""; }) > $Null
				#$ScriptInformation.Add(@{Data = "Created on"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "ID"; Value = $RDSSchedule.Id; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Name`t`t`t`t: " $RDSSchedule.Name
				Line 2 "Action`t`t`t`t: " $Action
				
				$cnt=-1
				ForEach($Item in $Target)
				{
					$cnt++
					If($cnt -eq 0)
					{
						Line 2 "Target`t`t`t`t: " $Item
					}
					Else
					{
						Line 6 "  " $Item
					}
				}
				
				Line 2 "Start`t`t`t`t: " $RDSSchedule.StartDateTime
				Line 2 "Repeat`t`t`t`t: " $Repeat
				Line 2 "Description`t`t`t: " $RDSSchedule.Description
				#Line 2 "Last modification by: " ""
				#Line 2 "Modified on: " ""
				#Line 2 "Created by: " ""
				#Line 2 "Created on: " ""
				Line 2 "ID`t`t`t`t: " $RDSSchedule.Id
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Schedule Name $($RDSSchedule.Name)"
				$rowdata = @()
				$columnHeaders = @("Name",($Script:htmlsb),$RDSSchedule.Name,$htmlwhite)
				$rowdata += @(,( "Action",($Script:htmlsb),$Action,$htmlwhite))
				
				$cnt=-1
				ForEach($Item in $Target)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$rowdata += @(,( "Target",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,( "",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				
				$rowdata += @(,( "Start",($Script:htmlsb),$RDSSchedule.StartDateTime,$htmlwhite))
				$rowdata += @(,( "Repeat",($Script:htmlsb),$Repeat,$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$RDSSchedule.Description,$htmlwhite))
				#$rowdata += @(,( "Last modification by",($Script:htmlsb),"",$htmlwhite))
				#$rowdata += @(,( "Modified on",($Script:htmlsb),"",$htmlwhite))
				#$rowdata += @(,( "Created by",($Script:htmlsb),"",$htmlwhite))
				#$rowdata += @(,( "Created on",($Script:htmlsb),"",$htmlwhite))
				$rowdata += @(,( "ID",($Script:htmlsb),$RDSSchedule.Id,$htmlwhite))

				$msg = ""
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
			
			#Properties
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Properties"
			}
			If($Text)
			{
				Line 2 "Properties"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Properties"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Enable Schedule"; Value = $RDSSchedule.Enabled.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Name"; Value = $RDSSchedule.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Action"; Value = $Action; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $RDSSchedule.Description; }) > $Null
				
				$cnt=-1
				ForEach($Item in $Target)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Target"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Enable Schedule`t`t`t: " $RDSSchedule.Enabled.ToString()
				Line 3 "Name`t`t`t`t: " $RDSSchedule.Name
				Line 3 "Action`t`t`t`t: " $Action
				Line 3 "Description`t`t`t: " $RDSSchedule.Description
				
				$cnt=-1
				ForEach($Item in $Target)
				{
					$cnt++
					If($cnt -eq 0)
					{
						Line 3 "Target`t`t`t`t: " $Item
					}
					Else
					{
						Line 7 "  " $Item
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable schedule",($Script:htmlsb),$RDSSchedule.Enabled.ToString(),$htmlwhite)
				$rowdata += @(,( "Name",($Script:htmlsb),$RDSSchedule.Name,$htmlwhite))
				$rowdata += @(,( "Action",($Script:htmlsb),$Action,$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$RDSSchedule.Description,$htmlwhite))
				
				$cnt=-1
				ForEach($Item in $Target)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$rowdata += @(,( "Target",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,( "",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				
				$msg = "Properties"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Trigger
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Trigger"
			}
			If($Text)
			{
				Line 2 "Trigger"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Trigger"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Date"; Value = $RDSSchedule.StartDateTime.ToShortDateString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Start"; Value = $RDSSchedule.StartDateTime.ToShortTimeString(); }) > $Null
				If($Action -eq "Reboot - Drain Mode")
				{
					$ScriptInformation.Add(@{Data = "Complete in"; Value = $TimeDuration  ; }) > $Null
				}
				ElseIf($Action -eq "Disable")
				{
					$ScriptInformation.Add(@{Data = "Duration"; Value = $TimeDuration; }) > $Null
				}
				$ScriptInformation.Add(@{Data = "Repeat"; Value = $Repeat; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Date`t`t`t`t: " $RDSSchedule.StartDateTime.ToShortDateString()
				Line 3 "Start`t`t`t`t: " $RDSSchedule.StartDateTime.ToShortTimeString()
				If($Action -eq "Reboot - Drain Mode")
				{
					Line 3 "Complete in`t`t`t: " $TimeDuration
				}
				ElseIf($Action -eq "Disable")
				{
					Line 3 "Duration`t`t`t: " $TimeDuration
				}
				Line 3 "Repeat`t`t`t`t: " $Repeat
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Date",($Script:htmlsb),$RDSSchedule.StartDateTime.ToShortDateString(),$htmlwhite)
				$rowdata += @(,( "Start",($Script:htmlsb),$RDSSchedule.StartDateTime.ToShortTimeString(),$htmlwhite))
				If($Action -eq "Reboot - Drain Mode")
				{
					$rowdata += @(,( "Complete in",($Script:htmlsb),$TimeDuration,$htmlwhite))
				}
				ElseIf($Action -eq "Disable")
				{
					$rowdata += @(,( "Duration",($Script:htmlsb),$TimeDuration,$htmlwhite))
				}
				$rowdata += @(,( "Repeat",($Script:htmlsb),$Repeat,$htmlwhite))

				$msg = "Trigger"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Options
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Options"
			}
			If($Text)
			{
				Line 2 "Options"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Options"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Send message before schedule is triggered"; Value = ""; }) > $Null

				If($RDSSChedule.Messages.Count -gt 0)
				{
					ForEach($Item in $RDSSChedule.Messages)
					{
						Switch ($Item.SendMsgSecs)
						{
							900		{$MsgTime = "15 minutes $($Item.SendMsgWhen)"; Break}
							1800	{$MsgTime = "30 minutes $($Item.SendMsgWhen)"; Break}
							2700	{$MsgTime = "45 minutes $($Item.SendMsgWhen)"; Break}
							3600	{$MsgTime = "1 hour $($Item.SendMsgWhen)"; Break}
							7200	{$MsgTime = "2 hours $($Item.SendMsgWhen)"; Break}
							10800	{$MsgTime = "3 hours $($Item.SendMsgWhen)"; Break}
							Default	{$MsgTime = "Unable to determine scheduled message Time: $($Item.SendMsgSecs)"; Break}
						}
						
						$ScriptInformation.Add(@{Data = "Enabled"; Value = $Item.Enabled.ToString(); }) > $Null
						$ScriptInformation.Add(@{Data = "Body"; Value = $Item.Message; }) > $Null
						$ScriptInformation.Add(@{Data = "Title"; Value = $Item.MessageTitle; }) > $Null
						$ScriptInformation.Add(@{Data = "Time"; Value = $MsgTime; }) > $Null
					}
				}

				If($Action -ne "Disable")
				{
					$ScriptInformation.Add(@{Data = "Enable Drain Mode"; Value = $RDSSchedule.DrainMode.ToString(); }) > $Null
				}
				If($Action -eq "Reboot - Drain Mode")
				{
					$ScriptInformation.Add(@{Data = "Force server reboot after"; Value = $ForceRebootTime; }) > $Null
				}
				If($Action -like "Reboot*")
				{
					$ScriptInformation.Add(@{Data = "Enforce schedule for currently inactive RD Session Host"; Value = $RDSSchedule.EnforceOnInactive.ToString(); }) > $Null
				}
				If($Action -eq "Disable")
				{
					$ScriptInformation.Add(@{Data = "On disable"; Value = $OnDisable; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Send message before schedule is triggered"
				
				If($RDSSChedule.Messages.Count -gt 0)
				{
					ForEach($Item in $RDSSChedule.Messages)
					{
						Switch ($Item.SendMsgSecs)
						{
							900		{$MsgTime = "15 minutes $($Item.SendMsgWhen)"; Break}
							1800	{$MsgTime = "30 minutes $($Item.SendMsgWhen)"; Break}
							2700	{$MsgTime = "45 minutes $($Item.SendMsgWhen)"; Break}
							3600	{$MsgTime = "1 hour $($Item.SendMsgWhen)"; Break}
							7200	{$MsgTime = "2 hours $($Item.SendMsgWhen)"; Break}
							10800	{$MsgTime = "3 hours $($Item.SendMsgWhen)"; Break}
							Default	{$MsgTime = "Unable to determine scheduled message Time: $($Item.SendMsgSecs)"; Break}
						}
						
						Line 3 "Enabled`t: " $Item.Enabled.ToString()
						Line 3 "Body`t: " $Item.Message
						Line 3 "Title`t: " $Item.MessageTitle
						Line 3 "Time`t: " $MsgTime
					}
				}

				If($Action -ne "Disable")
				{
					Line 3 "Enable Drain Mode`t`t: " $RDSSchedule.DrainMode.ToString()
				}
				If($Action -eq "Reboot - Drain Mode")
				{
					Line 3 "Force server reboot after`t: " $ForceRebootTime
				}
				If($Action -like "Reboot*")
				{
					Line 3 "Enforce schedule for currently"
					Line 3 "inactive RD Session Host`t: " $RDSSchedule.EnforceOnInactive.ToString()
				}
				If($Action -eq "Disable")
				{
					Line 3 "On disable`t`t`t: " $OnDisable
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Send message before schedule is triggered",($Script:htmlsb),"",$htmlwhite)
				
				If($RDSSChedule.Messages.Count -gt 0)
				{
					ForEach($Item in $RDSSChedule.Messages)
					{
						Switch ($Item.SendMsgSecs)
						{
							900		{$MsgTime = "15 minutes $($Item.SendMsgWhen)"; Break}
							1800	{$MsgTime = "30 minutes $($Item.SendMsgWhen)"; Break}
							2700	{$MsgTime = "45 minutes $($Item.SendMsgWhen)"; Break}
							3600	{$MsgTime = "1 hour $($Item.SendMsgWhen)"; Break}
							7200	{$MsgTime = "2 hours $($Item.SendMsgWhen)"; Break}
							10800	{$MsgTime = "3 hours $($Item.SendMsgWhen)"; Break}
							Default	{$MsgTime = "Unable to determine scheduled message Time: $($Item.SendMsgSecs)"; Break}
						}
						
						$rowdata += @(,( "Enabled",($Script:htmlsb),$Item.Enabled.ToString(),$htmlwhite))
						$rowdata += @(,( "Body",($Script:htmlsb),$Item.Message,$htmlwhite))
						$rowdata += @(,( "Title",($Script:htmlsb),$Item.MessageTitle,$htmlwhite))
						$rowdata += @(,( "Time",($Script:htmlsb),$MsgTime,$htmlwhite))
					}
				}

				If($Action -ne "Disable")
				{
					$rowdata += @(,( "Enable Drain Mode",($Script:htmlsb),$RDSSchedule.DrainMode.ToString(),$htmlwhite))
				}
				If($Action -eq "Reboot - Drain Mode")
				{
					$rowdata += @(,( "Force server reboot after",($Script:htmlsb),$ForceRebootTime,$htmlwhite))
				}
				If($Action -like "Reboot*")
				{
					$rowdata += @(,( "Enforce schedule for currently inactive RD Session Host",($Script:htmlsb),$RDSSchedule.EnforceOnInactive.ToString(),$htmlwhite))
				}
				If($Action -eq "Disable")
				{
					$rowdata += @(,( "On disable",($Script:htmlsb),$OnDisable,$htmlwhite))
				}

				$msg = "Options"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}

	$VDIHosts = Get-VDIHost -SiteId $Site.Id -EA 0 4>$Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve VDI for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve VDI for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve VDI for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve VDI for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No VDI retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No VDI retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No VDI retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No VDI retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "VDI"
		}
		If($Text)
		{
			Line 1 "VDI"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "VDI"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Site VDI"
		ForEach($VDIHost in $VDIHosts)
		{
			$VDIHostStatus = Get-VDIHostStatus -Id $VDIHost.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve VDI Host Status for VDI Host $($VDIHost.Id)"
				}
			}
			ElseIf($? -and $Null -eq $VDIHostStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No VDI Host Status retrieved for VDI Host $($VDIHost.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No VDI Host Status retrieved for VDI Host $($VDIHost.Id)"
				}
				If($Text)
				{
					Line 0 "No VDI Host Status retrieved for VDI Host $($VDIHost.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No VDI Host Status retrieved for VDI Host $($VDIHost.Id)"
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Providers $($VDIHost.Server)"
				}
				If($Text)
				{
					Line 2 "Providers"
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Providers $($VDIHost.Server)"
				}

				$VDIType = GetVDIType $VDIHost.VDIType
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Provider"; Value = $VDIHost.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Type"; Value = $VDIType; }) > $Null
					$ScriptInformation.Add(@{Data = "VDI Agent"; Value = $VDIHost.VDIAgent; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $VDIHostStatus.AgentState.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Description"; Value = $VDIHost.Description; }) > $Null
					$ScriptInformation.Add(@{Data = "Log level"; Value = $VDIHostStatus.LogLevel; }) > $Null
					$ScriptInformation.Add(@{Data = "ID"; Value = $VDIHost.Id; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 3 "Provider`t: " $VDIHost.Server
					Line 3 "Type`t`t: " $VDIType
					Line 3 "VDI Agent`t: " $VDIHost.VDIAgent
					Line 3 "Status`t`t: " $VDIHostStatus.AgentState.ToString()
					Line 3 "Description`t: " $VDIHost.Description
					Line 3 "Log level`t: " $VDIHostStatus.LogLevel
					Line 3 "ID`t`t: " $VDIHost.Id
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Provider",($Script:htmlsb),$VDIHost.Server,$htmlwhite)
					$rowdata += @(,( "Type",($Script:htmlsb),$VDIType,$htmlwhite))
					$rowdata += @(,( "VDI Agent",($Script:htmlsb),$VDIHost.VDIAgent,$htmlwhite))
					$rowdata += @(,( "Status",($Script:htmlsb),$VDIHostStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,( "Description",($Script:htmlsb),$VDIHost.Description,$htmlwhite))
					$rowdata += @(,( "Log level",($Script:htmlsb),$VDIHostStatus.LogLevel,$htmlwhite))
					$rowdata += @(,( "Id",($Script:htmlsb),$VDIHost.Id,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
			
			#Properties
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Properties"
			}
			If($Text)
			{
				Line 3 "Properties"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Properties"
			}
			
			$HostPA = Get-PA -Id $VDIHost.PreferredPAId -EA 0 4>$Null
			
			If($? -and -$Null -ne $HostPA)
			{
				If($VDIHost.VDIAgent -eq $HostPa.Server)
				{
					$DedicatedVDIAgent = $False
				}
				Else
				{
					$DedicatedVDIAgent = $True
				}
			}
			ElseIf($? -and $Null -eq $HostPA)
			{
				$DedicatedVDIAgent = $False
			}
			Else
			{
				$DedicatedVDIAgent = $False
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Enable provider in site"; Value = $VDIHost.Enabled.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Type"; Value = $VDIType; }) > $Null
				$ScriptInformation.Add(@{Data = "Host"; Value = $VDIHost.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "Port"; Value = $VDIHost.VDIPort.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $VDIHost.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Dedicated VDI Agent"; Value = $DedicatedVDIAgent.ToString(); }) > $Null
				If($DedicatedVDIAgent)
				{
					$ScriptInformation.Add(@{Data = "Agent address"; Value = $VDIHost.VDIAgent; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Enable provider in site:`t" $VDIHost.Enabled.ToString()
				Line 4 "Type:`t`t`t`t" $VDIType
				Line 4 "Host:`t`t`t`t" $VDIHost.Server
				Line 4 "Port:`t`t`t`t" $VDIHost.VDIPort.ToString()
				Line 4 "Description:`t`t`t" $VDIHost.Description
				Line 4 "Dedicated VDI Agent:`t`t" $DedicatedVDIAgent.ToString()
				If($DedicatedVDIAgent)
				{
					Line 4 "Agent address:`t`t`t" $VDIHost.VDIAgent
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable provider in site",($Script:htmlsb),$VDIHost.Enabled.ToString(),$htmlwhite)
				$rowdata += @(,("Type",($Script:htmlsb),$VDIType,$htmlwhite))
				$rowdata += @(,("Host",($Script:htmlsb),$VDIHost.Server,$htmlwhite))
				$rowdata += @(,("Port",($Script:htmlsb),$VDIHost.VDIPort.ToString(),$htmlwhite))
				$rowdata += @(,("Description",($Script:htmlsb),$VDIHost.Description,$htmlwhite))
				$rowdata += @(,("Dedicated VDI Agent",($Script:htmlsb),$DedicatedVDIAgent.ToString(),$htmlwhite))
				If($DedicatedVDIAgent)
				{
					$rowdata += @(,("Agent address",($Script:htmlsb),$VDIHost.VDIAgent,$htmlwhite))
				}

				$msg = "Properties"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
			
			#Credentials
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Credentials"
			}
			If($Text)
			{
				Line 3 "Credentials"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Credentials"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Username"; Value = $VDIHost.VDIUsername; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Username:`t`t`t" $VDIHost.VDIUsername
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Username",($Script:htmlsb),$VDIHost.VDIUsername,$htmlwhite)

				$msg = "Credentials"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Agent Settings
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Agent Settings"
			}
			If($Text)
			{
				Line 3 "Agent Settings"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Agent Settings"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Max connections"; Value = $VDIHost.MaxGuests.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Publishing Session Timeout"; Value = $VDIHost.SessionTimeout.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Allow Client URL/Mail Redirection"; Value = $VDIHost.AllowURLAndMailRedirection.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Support Shell URL namespace objects"; Value = $VDIHost.SupportShellURLNamespaceObjects.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Preferred Publishing Agent"; Value = $VDIHostStatus.PreferredPA; }) > $Null
				$ScriptInformation.Add(@{Data = "Allow file transfer command (HTML5 and Chrome clients)"; Value = $VDIHost.AllowFileTransfer.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "Max connections:`t`t`t`t`t" $VDIHost.MaxGuests.ToString()
				Line 4 "Publishing Session Timeout:`t`t`t`t" $VDIHost.SessionTimeout.ToString()
				Line 4 "Allow Client URL/Mail Redirection:`t`t`t" $VDIHost.AllowURLAndMailRedirection.ToString()
				Line 4 "Support Shell URL namespace objects:`t`t`t" $VDIHost.SupportShellURLNamespaceObjects.ToString()
				Line 4 "Preferred Publishing Agent:`t`t`t`t" $VDIHostStatus.PreferredPA
				Line 4 "Allow file transfer command (HTML5 and Chrome clients):`t" $VDIHost.AllowFileTransfer.ToString()
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Max connections",($Script:htmlsb),$VDIHost.MaxGuests.ToString(),$htmlwhite)
				$rowdata += @(,("Publishing Session Timeout",($Script:htmlsb),$VDIHost.SessionTimeout.ToString(),$htmlwhite))
				$rowdata += @(,("Allow Client URL/Mail Redirection",($Script:htmlsb),$VDIHost.AllowURLAndMailRedirection.ToString(),$htmlwhite))
				$rowdata += @(,("Support Shell URL namespace objects",($Script:htmlsb),$VDIHost.SupportShellURLNamespaceObjects.ToString(),$htmlwhite))
				$rowdata += @(,("Preferred Publishing Agent",($Script:htmlsb),$VDIHostStatus.PreferredPA,$htmlwhite))
				$rowdata += @(,("Allow file transfer command (HTML5 and Chrome clients)",($Script:htmlsb),$VDIHost.AllowFileTransfer.ToString(),$htmlwhite))

				$msg = "Agent Settings"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#RDP Printer
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "RDP Printer"
			}
			If($Text)
			{
				Line 3 "RDP Printer"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "RDP Printer"
			}
			
			Switch ($VDIHost.PrinterNameFormat)
			{
				"PrnFormat_PRN_CMP_SES"	{$VDIPrinterNameFormat = "Printername (from Computername) in Session no."; Break}
				"PrnFormat_SES_CMP_PRN"	{$VDIPrinterNameFormat = "Session no. (Computername from) Printername"; Break}
				"PrnFormat_PRN_REDSES"	{$VDIPrinterNameFormat = "Printername (redirected Session no.)"; Break}
				Default					{$VDIPrinterNameFormat = "Unable to determine RDP Printer Name Format: $($VDIHost.PrinterNameFormat)"; Break}
			}
			
			$VDIRemoveSessionNumberFromPrinter = $VDIHost.RemoveSessionNumberFromPrinterName.ToString()

			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "RDP Printer Name Format"; Value = $VDIPrinterNameFormat; }) > $Null
				$ScriptInformation.Add(@{Data = "Remove session number from printer name"; Value = $VDIRemoveSessionNumberFromPrinter; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 4 "RDP Printer Name Format:`t`t`t`t" $VDIPrinterNameFormat
				Line 4 "Remove session number from printer name:`t`t" $VDIRemoveSessionNumberFromPrinter
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("RDP Printer Name Format",($Script:htmlsb),$VDIPrinterNameFormat,$htmlwhite)
				$rowdata += @(,("Remove session number from printer name",($Script:htmlsb),$VDIRemoveSessionNumberFromPrinter,$htmlwhite))

				$msg = "RDP Printer"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
		
		#Pools
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 "Pools"
		}
		If($Text)
		{
			Line 2 "Pools"
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 "Pools"
		}
		
		If($MSWord -or $PDF)
		{
			$VDIPools = Get-VDIPool -SiteId $Site.Id -EA 0 4>$Null
			If($? -and $Null -ne $VDIPools)
			{
				ForEach($VDIPool in $VDIPools)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Pools"; Value = ""; }) > $Null
					$ScriptInformation.Add(@{Data = "  Name"; Value = $VDIPool.Name; }) > $Null
					
					ForEach($Item in $VDIPool.Members)
					{
						$VDIPoolMembers = Get-VDIPoolMember -VDIPoolName $VDIPool.Name -EA 0 4>$Null
						
						If($? -and $Null -ne $VDIPoolMembers)
						{
							$cnt = -1
							ForEach($VDIPoolMember in $VDIPoolMembers)
							{
								Switch($VDIPoolMember.Type)
								{
									"ALLGUESTINHOST"	{$MemberType = "All Guest VMs in Host"; Break}
									"GUEST"				{$MemberType = "Guest VM"; Break}
									"NATIVEPOOL"		{$MemberType = "Native Pool"; Break}
									"TEMPLATEGUEST"		{$MemberType = "Template"; Break}
									"UNKNOWN"			{$MemberType = "Unknown"; Break}
									Default				{$MemberType = "Unable to determine Pool Member Type: $($VDIPoolMember.Type)"; Break}
								}
								$cnt++
								If($cnt -eq 0)
								{
									$ScriptInformation.Add(@{Data = "    Members"; Value = "Name: $($VDIPoolMember.Name) Type: $MemberType"; }) > $Null
								}
								Else
								{
									$ScriptInformation.Add(@{Data = ""; Value = "Name: $($VDIPoolMember.Name) Type: $MemberType"; }) > $Null
								}
							}
						}
						ElseIf($? -and $Null -eq $VDIPoolMembers)
						{
							$ScriptInformation.Add(@{Data = "    Members"; Value = "None found"; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = "    Members"; Value = "Unable to retrieve"; }) > $Null
						}
					}
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
			}
			ElseIf($? -and $Null -eq $VDIPools)
			{
				WriteWordLine 0 0 "No VDI Pools found for Site $($Site.Name)"
			}
			Else
			{
				WriteWordLine 0 0 "Unable to retrieve VDI Pools for Site $($Site.Name)"
			}
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			$VDIPools = Get-VDIPool -SiteId $Site.Id -EA 0 4>$Null
			If($? -and $Null -ne $VDIPools)
			{
				ForEach($VDIPool in $VDIPools)
				{
					Line 3 "Pools"
					Line 4 "Name:`t`t" $VDIPool.Name
					
					ForEach($Item in $VDIPool.Members)
					{
						$VDIPoolMembers = Get-VDIPoolMember -VDIPoolName $VDIPool.Name -EA 0 4>$Null
						
						If($? -and $Null -ne $VDIPoolMembers)
						{
							$cnt = -1
							ForEach($VDIPoolMember in $VDIPoolMembers)
							{
								Switch($VDIPoolMember.Type)
								{
									"ALLGUESTINHOST"	{$MemberType = "All Guest VMs in Host"; Break}
									"GUEST"				{$MemberType = "Guest VM"; Break}
									"NATIVEPOOL"		{$MemberType = "Native Pool"; Break}
									"TEMPLATEGUEST"		{$MemberType = "Template"; Break}
									"UNKNOWN"			{$MemberType = "Unknown"; Break}
									Default				{$MemberType = "Unable to determine Pool Member Type: $($VDIPoolMember.Type)"; Break}
								}
								$cnt++
								If($cnt -eq 0)
								{
									Line 4 "Members:`t" "Name: $($VDIPoolMember.Name) Type: $MemberType"
								}
								Else
								{
									Line 6 "Name: $($VDIPoolMember.Name) Type: $MemberType"
								}
							}
						}
						ElseIf($? -and $Null -eq $VDIPoolMembers)
						{
							Line 4 "Members:`t" "None found"
						}
						Else
						{
							Line 4 "Members:`t" "Unable to retrieve"
						}
					}
				}
				Line 0 ""
			}
			ElseIf($? -and $Null -eq $VDIPools)
			{
				Line 0 "No VDI Pools found for Site $($Site.Name)"
			}
			Else
			{
				Line 0 "Unable to retrieve VDI Pools for Site $($Site.Name)"
			}
			Line 0 ""
		}
		If($HTML)
		{
			$VDIPools = Get-VDIPool -SiteId $Site.Id -EA 0 4>$Null
			If($? -and $Null -ne $VDIPools)
			{
				ForEach($VDIPool in $VDIPools)
				{
					$rowdata = @()
					$columnHeaders = @("Pools",($Script:htmlsb),"",$htmlwhite)
					$rowdata += @(,("  Name",($Script:htmlsb),$VDIPool.Name,$htmlwhite))
					
					ForEach($Item in $VDIPool.Members)
					{
						$VDIPoolMembers = Get-VDIPoolMember -VDIPoolName $VDIPool.Name -EA 0 4>$Null
						
						If($? -and $Null -ne $VDIPoolMembers)
						{
							$cnt = -1
							ForEach($VDIPoolMember in $VDIPoolMembers)
							{
								Switch($VDIPoolMember.Type)
								{
									"ALLGUESTINHOST"	{$MemberType = "All Guest VMs in Host"; Break}
									"GUEST"				{$MemberType = "Guest VM"; Break}
									"NATIVEPOOL"		{$MemberType = "Native Pool"; Break}
									"TEMPLATEGUEST"		{$MemberType = "Template"; Break}
									"UNKNOWN"			{$MemberType = "Unknown"; Break}
									Default				{$MemberType = "Unable to determine Pool Member Type: $($VDIPoolMember.Type)"; Break}
								}
								$cnt++
								If($cnt -eq 0)
								{
									$rowdata += @(,("    Members",($Script:htmlsb),"Name: $($VDIPoolMember.Name) Type: $MemberType",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,("",($Script:htmlsb),"Name: $($VDIPoolMember.Name) Type: $MemberType",$htmlwhite))
								}
							}
						}
						ElseIf($? -and $Null -eq $VDIPoolMembers)
						{
							$rowdata += @(,("    Members",($Script:htmlsb),"None found",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("    Members",($Script:htmlsb),"Unable to retrieve",$htmlwhite))
						}
					}

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
			ElseIf($? -and $Null -eq $VDIPools)
			{
				WriteHTMLLine 0 0 "No VDI Pools found for Site $($Site.Name)"
			}
			Else
			{
				WriteHTMLLine 0 0 "Unable to retrieve VDI Pools for Site $($Site.Name)"
			}
			WriteHTMLLine 0 0 ""
		}
		
		#Templates
		
		$VDITemplates = Get-VDITemplate -SiteId $Site.Id -EA 0 4>$Null
		If($? -and $Null -ne $VDITemplates)
		{
			ForEach($VDITemplate in $VDITemplates)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Templates $($VDITemplate.Name)"
				}
				If($Text)
				{
					Line 2 "Templates $($VDITemplate.Name)"
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Templates $($VDITemplate.Name)"
				}
		
				$VDIHost = Get-VDIHost -Id $VDITemplate.VDIHostId -EA 0 4>$Null
				
				If($? -and $null -ne $VDIHost)
				{
					$Provider = $VDIHost.Server
					$ProviderType = GetVDIType $VDIHost.VDIType
					
					$ProviderStatus = (Get-VDIHostStatus -Id $VDITemplate.VDIHostId -EA 0 4>$Null).AgentState.ToString() 
				}
				ElseIf($? -$Null -eq $VDIHost)
				{
					$Provider  = "No Provider found"
					$ProviderType = "Unknown"
					$ProviderStatus = "Unknown"
				}
				Else
				{
					$Provider  = "Unable to retrieve Provider with an ID of $($VDITemplate.VDIHostId)"
					$ProviderType = "Unknown"
					$ProviderStatus = "Unknown"
				}
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Name"; Value = $VDITemplate.Name; }) > $Null
					$ScriptInformation.Add(@{Data = "Type"; Value = $VDITemplate.TemplateType.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Provider"; Value = $Provider; }) > $Null
					$ScriptInformation.Add(@{Data = "Provider Type"; Value = $ProviderType; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $ProviderStatus; }) > $Null
					$ScriptInformation.Add(@{Data = "ID"; Value = $VDITemplate.Id.ToString(); }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 3 "Name:`t`t`t" $VDITemplate.Name
					Line 3 "Type:`t`t`t" $VDITemplate.TemplateType.ToString()
					Line 3 "Provider:`t`t" $Provider
					Line 3 "Provider Type:`t`t" $ProviderType
					Line 3 "Status:`t`t`t" $ProviderStatus
					Line 3 "ID:`t`t`t" $VDITemplate.Id.ToString()
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Name",($Script:htmlsb),$VDITemplate.Name,$htmlwhite)
					$rowdata += @(,( "Type",($Script:htmlsb),$VDITemplate.TemplateType.ToString(),$htmlwhite))
					$rowdata += @(,( "Provider",($Script:htmlsb),$Provider,$htmlwhite))
					$rowdata += @(,( "Provider type",($Script:htmlsb),$ProviderType,$htmlwhite))
					$rowdata += @(,( "Status",($Script:htmlsb),$ProviderStatus,$htmlwhite))
					$rowdata += @(,( "Id",($Script:htmlsb),$VDITemplate.Id.ToString(),$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}

				#Properties
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Properties"
				}
				If($Text)
				{
					Line 3 "Properties"
				}
				If($HTML)
				{
					#WriteHTMLLine 4 0 "Properties"
				}
				
				Switch ($VDITemplate.UnusedGuestDurationMins)
				{
					0		{$DeleteVMsTime = "Never"; Break}
					1440	{$DeleteVMsTime = "1 day"; Break}
					10080	{$DeleteVMsTime = "1 week"; Break}
					43200	{$DeleteVMsTime = "30 days"; Break}
					Default	{$DeleteVMsTime = "Unable to determine Delete unused guest VMs after: $($VDITemplate.UnusedGuestDurationMins)"; Break}
				}
				
				Switch($VDITemplate.CloneMethod)
				{
					"LinkedClone"	{$CloneMethod = "Create a linked clone"; Break}
					"FullClone"		{$CloneMethod = "Create a full clone"; Break}
					Default			{$CloneMethod = "Unable to determine Clone method: $($VDITemplate.CloneMethod)"; Break}
				}

				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Template name"; Value = $VDITemplate.Name; }) > $Null
					$ScriptInformation.Add(@{Data = "Maximum guest VMs"; Value = $VDITemplate.MaxGuests.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Keep available buffer"; Value = $VDITemplate.PreCreatedGuests.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Guest VM name"; Value = $VDITemplate.GuestNameFormat; }) > $Null
					$ScriptInformation.Add(@{Data = "Delete unused guest VMs after"; Value = $DeleteVMsTime; }) > $Null
					$ScriptInformation.Add(@{Data = "Clone method"; Value = $CloneMethod; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 4 "Template name:`t`t`t" $VDITemplate.Name
					Line 4 "Maximum guest VMs:`t`t" $VDITemplate.MaxGuests.ToString()
					Line 4 "Keep available buffer:`t`t" $VDITemplate.PreCreatedGuests.ToString()
					Line 4 "Guest VM name:`t`t`t" $VDITemplate.GuestNameFormat
					Line 4 "Delete unused guest VMs after:`t" $DeleteVMsTime
					Line 4 "Clone method:`t`t`t" $CloneMethod
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Template name",($Script:htmlsb),$VDITemplate.Name,$htmlwhite)
					$rowdata += @(,("Maximum guest VMs",($Script:htmlsb),$VDITemplate.MaxGuests.ToString(),$htmlwhite))
					$rowdata += @(,("Keep available buffer",($Script:htmlsb),$VDITemplate.PreCreatedGuests.ToString(),$htmlwhite))
					$rowdata += @(,("Guest VM name",($Script:htmlsb),$VDITemplate.GuestNameFormat,$htmlwhite))
					$rowdata += @(,("Delete unused guest VMs after",($Script:htmlsb),$DeleteVMsTime,$htmlwhite))
					$rowdata += @(,("Clone method",($Script:htmlsb),$CloneMethod,$htmlwhite))

					$msg = "Properties"
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
				
				#Advanced
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Advanced"
				}
				If($Text)
				{
					Line 3 "Advanced"
				}
				If($HTML)
				{
					#WriteHTMLLine 4 0 "Advanced"
				}
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Folder"; Value = $VDITemplate.FolderName; }) > $Null
					$ScriptInformation.Add(@{Data = "Native Pool"; Value = $VDITemplate.NativePoolName; }) > $Null
					$ScriptInformation.Add(@{Data = "Physical Host"; Value = $VDITemplate.PhysicalHostName; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 4 "Folder:`t`t`t`t" $VDITemplate.FolderName
					Line 4 "Native Pool:`t`t`t" $VDITemplate.NativePoolName
					Line 4 "Physical Host:`t`t`t" $VDITemplate.PhysicalHostName
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Folder",($Script:htmlsb),$VDITemplate.FolderName,$htmlwhite)
					$rowdata += @(,("Native Pool",($Script:htmlsb),$VDITemplate.NativePoolName,$htmlwhite))
					$rowdata += @(,("Physical Host",($Script:htmlsb),$VDITemplate.PhysicalHostName,$htmlwhite))

					$msg = "Advanced"
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
				
				#Preparation
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Preparation"
				}
				If($Text)
				{
					Line 3 "Preparation"
				}
				If($HTML)
				{
					#WriteHTMLLine 4 0 "Preparation"
				}
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Image preparation tool"; Value = $VDITemplate.ImagePrepTool.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Computer name"; Value = $VDITemplate.ComputerName; }) > $Null
					$ScriptInformation.Add(@{Data = "Owner name"; Value = $VDITemplate.OwnerName; }) > $Null
					$ScriptInformation.Add(@{Data = "Organization"; Value = $VDITemplate.Organization; }) > $Null
					$ScriptInformation.Add(@{Data = "Join domain"; Value = $VDITemplate.Domain; }) > $Null
					$ScriptInformation.Add(@{Data = "Administrator"; Value = $VDITemplate.Administrator; }) > $Null
					$ScriptInformation.Add(@{Data = "Target OU"; Value = $VDITemplate.DomainOrgUnit; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 4 "Image preparation tool:`t`t" $VDITemplate.ImagePrepTool.ToString()
					Line 4 "Computer name:`t`t`t" $VDITemplate.ComputerName
					Line 4 "Owner name:`t`t`t" $VDITemplate.OwnerName
					Line 4 "Organization:`t`t`t" $VDITemplate.Organization
					Line 4 "Join domain:`t`t`t" $VDITemplate.Domain
					Line 4 "Administrator:`t`t`t" $VDITemplate.Administrator
					Line 4 "Target OU:`t`t`t" $VDITemplate.DomainOrgUnit
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Image preparation tool",($Script:htmlsb),$VDITemplate.ImagePrepTool.ToString(),$htmlwhite)
					$rowdata += @(,("Computer name",($Script:htmlsb),$VDITemplate.ComputerName,$htmlwhite))
					$rowdata += @(,("Owner name",($Script:htmlsb),$VDITemplate.OwnerName,$htmlwhite))
					$rowdata += @(,("Organization",($Script:htmlsb),$VDITemplate.Organization,$htmlwhite))
					$rowdata += @(,("Join domain",($Script:htmlsb),$VDITemplate.Domain,$htmlwhite))
					$rowdata += @(,("Administrator",($Script:htmlsb),$VDITemplate.Administrator,$htmlwhite))
					$rowdata += @(,("Target OU",($Script:htmlsb),$VDITemplate.DomainOrgUnit,$htmlwhite))

					$msg = "Preparation"
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
				
				#License Keys
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "License Keys"
				}
				If($Text)
				{
					Line 3 "License Keys"
				}
				If($HTML)
				{
					#WriteHTMLLine 4 0 "License Keys"
				}
				
				If($MSWord -or $PDF)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					
					If($VDITemplate.LicenseKeyType.ToString() -eq "KMS")
					{
						$ScriptInformation.Add(@{Data = "License key management type"; Value = "Key Management Service (KMS)"; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = "License key management type"; Value = "Multple Activation Key (MAK)"; }) > $Null
						
						$LicenseKeys = Get-VDITemplateLicenseKey -Id $VDITemplate.Id -EA 0 4>$Null
						
						ForEach($Item in $LicenseKeys)
						{
							$ScriptInformation.Add(@{Data = "License Key"; Value = $Item.LicenseKey; }) > $Null
							$ScriptInformation.Add(@{Data = "Key Limit"; Value = $Item.KeyLimit; }) > $Null
						}
					}

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					If($VDITemplate.LicenseKeyType.ToString() -eq "KMS")
					{
						Line 4 "License key management type:`t" "Key Management Service (KMS)"
					}
					Else
					{
						Line 4 "License key management type:`t`t`t" "Multple Activation Key (MAK)"
						
						$LicenseKeys = Get-VDITemplateLicenseKey -Id $VDITemplate.Id -EA 0 4>$Null

						ForEach($Item in $LicenseKeys)
						{
							Line 4 "License Key:`t" $Item.LicenseKey
							Line 4 "Key Limit:`t" $Item.KeyLimit
						}
					}
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					If($VDITemplate.LicenseKeyType.ToString() -eq "KMS")
					{
						$columnHeaders = @("License key management type",($Script:htmlsb),"Key Management Service (KMS)",$htmlwhite)
					}
					Else
					{
						$columnHeaders = @("License key management type",($Script:htmlsb),"Multple Activation Key (MAK)",$htmlwhite)
						
						$LicenseKeys = Get-VDITemplateLicenseKey -Id $VDITemplate.Id -EA 0 4>$Null

						$cnt = -1
						ForEach($Item in $LicenseKeys)
						{
							$rowdata += @(,("License Key",($Script:htmlsb),$Item.LicenseKey,$htmlwhite))
							$rowdata += @(,("Key Limit",($Script:htmlsb),$Item.KeyLimit,$htmlwhite))
						}
					}

					$msg = "License Keys"
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
		}
		ElseIf($? -and $Null -eq $VDITemplates)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "No VDI Templates found for Site $($Site.Name)"
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 "No VDI Templates found for Site $($Site.Name)"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "No VDI Templates found for Site $($Site.Name)"
				WriteHTMLLine 0 0 ""
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Unable to retrieve VDI Templates for Site $($Site.Name)"
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 "Unable to retrieve VDI Templates for Site $($Site.Name)"
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Unable to retrieve VDI Templates for Site $($Site.Name)"
				WriteHTMLLine 0 0 ""
			}
		}
		
		#Desktops
		#can't find this
	}
	
	$GWs = Get-GW -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Gateways for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Gateways for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Gateways for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Gateways for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Gateways retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Gateways retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No Gateways retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Gateways retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Gateways"
		}
		If($Text)
		{
			Line 1 "Gateways"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Gateways"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Gateways"
		ForEach($GW in $GWs)
		{
			$GWStatus = Get-GWStatus -Id $GW.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve Gateway Status for Gateway $($GW.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve Gateway Status for Gateway $($GW.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve Gateway Status for Gateway $($GW.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve Gateway Status for Gateway $($GW.Id)"
				}
			}
			ElseIf($? -and $Null -eq $GWStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No Gateway Status retrieved for Gateway $($GW.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Gateway Status retrieved for Gateway $($GW.Id)"
				}
				If($Text)
				{
					Line 0 "No Gateway Status retrieved for Gateway $($GW.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Gateway Status retrieved for Gateway $($GW.Id)"
				}
			}
			Else
			{
				If($GW.InheritDefaultSslTlsSettings)
				{
					#do we inherit site defaults?
					#yes we do, get the default settings for the Site
					#use the Site default settings

					$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
					
					If($? -and $Null -ne $GWDefaults)
					{
						If($GWDefaults.EnableHSTS)
						{
							$GWEnableHSTS = $GWDefaults.EnableHSTS.ToString()
							$GWHSTSMaxage = $GWDefaults.HSTSMaxAge.ToString()
							$GWHSTSIncludeSubdomains = $GWDefaults.HSTSIncludeSubdomains.ToString()
							$GWHSTSPreload = $GWDefaults.HSTSPreload.ToString()
						}
						Else
						{
							$GWEnableHSTS = $GWDefaults.EnableHSTS.ToString()
						}
						$GWEnableSSL = $GWDefaults.EnableSSL.ToString()
						$GWEnableSSLOnPort = $GWDefaults.SSLPort.ToString()
						
						Switch ($GWDefaults.MinSSLVersion)
						{
							"SSLv2"		{$GWAcceptedSSLVersions = "SSL v2 - TLS v1.2 (Weak)"; Break}
							"SSLv3"		{$GWAcceptedSSLVersions = "SSL v3 - TLS v1.2"; Break}
							"TLSv1"		{$GWAcceptedSSLVersions = "TLS v1 - TLS v1.2"; Break}
							"TLSv1_1"	{$GWAcceptedSSLVersions = "TLS v1.1 - TLS v1.2"; Break}
							"TLSv1_2"	{$GWAcceptedSSLVersions = "TLS v1.2 only (Strong)"; Break}
							Default		{$GWAcceptedSSLVersions = "Unable to determine Minimum SSL Version: $($GWDefaults.MinSSLVersion)"; Break}
						}
						
						$GWCipherStrength = $GWDefaults.CipherStrength.ToString()
						$GWCipher = $GWDefaults.Cipher
						
						If($GWDefaults.CertificateId -eq 0)
						{
							$GWCertificates = "All matching usage"
						}
						Else
						{
							$Results = Get-RASCertificate -Id $GWDefaults.CertificateId -EA 0 4>$Null
							
							If($? -and $Null -ne $Results)
							{
								$GWCertificates = $Results.Name
							}
							Else
							{
								$GWCertificates = "Unable to determin Gateway Certificate: $($GWDefaults.CertificateId)"
							}
							
						}
					}
					Else
					{
						#unable to retrieve default, use built-in default values
						$GWEnableHSTS = "False"
						$GWEnableSSL = "True"
						$GWEnableSSLOnPort = "443"
						$GWAcceptedSSLVersions = "TLS v1 - TLS v1.2"
						$GWCipherStrength = "High"
						$GWCipher = "EECDH:!SSLv2:!SSLv3:!aNULL:!RC4:!ADH:!eNULL:!LOW:!MEDIUM:!EXP:+HIGH"
						$GWCertificates = "All matching usage"
					}
				}
				Else
				{
					#we don't inherit settings
					#get the settings configured for this GW
					If($GW.EnableHSTS)
					{
						$GWEnableHSTS = $GW.EnableHSTS.ToString()
						$GWHSTSMaxage = $GW.HSTSMaxAge.ToString()
						$GWHSTSIncludeSubdomains = $GW.HSTSIncludeSubdomains.ToString()
						$GWHSTSPreload = $GW.HSTSPreload.ToString()
					}
					Else
					{
						$GWEnableHSTS = $GW.EnableHSTS.ToString()
					}
					$GWEnableSSL = $GW.EnableSSL.ToString()
					$GWEnableSSLOnPort = $GW.SSLPort.ToString()
					
					Switch ($GW.MinSSLVersion)
					{
						"SSLv2"		{$GWAcceptedSSLVersions = "SSL v2 - TLS v1.2 (Weak)"; Break}
						"SSLv3"		{$GWAcceptedSSLVersions = "SSL v3 - TLS v1.2"; Break}
						"TLSv1"		{$GWAcceptedSSLVersions = "TLS v1 - TLS v1.2"; Break}
						"TLSv1_1"	{$GWAcceptedSSLVersions = "TLS v1.1 - TLS v1.2"; Break}
						"TLSv1_2"	{$GWAcceptedSSLVersions = "TLS v1.2 only (Strong)"; Break}
						Default		{$GWAcceptedSSLVersions = "Unable to determine Minimum SSL Version: $($GW.MinSSLVersion)"; Break}
					}
					
					$GWCipherStrength = $GW.CipherStrength.ToString()
					$GWCipher = $GW.Cipher
					
					If($GW.CertificateId -eq 0)
					{
						$GWCertificates = "All matching usage"
					}
					Else
					{
						$Results = Get-RASCertificate -Id $GW.CertificateId -EA 0 4>$Null
						
						If($? -and $Null -ne $Results)
						{
							#double replace to remove the < and > from the cert name so it doesn't mess up HTML output
							$GWCertificates = $Results.Name.Replace("<","").Replace(">","")	
						}
						Else
						{
							$GWCertificates = "Unable to determin Gateway Certificate: $($GW.CertificateId)"
						}
					}
				}
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Gateways $($GW.Server)"
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Server"; Value = $GW.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Mode"; Value = $GW.GWMode; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $GWStatus.AgentState; }) > $Null
					$ScriptInformation.Add(@{Data = "Description"; Value = $GW.Description; }) > $Null
					$ScriptInformation.Add(@{Data = "Certificate"; Value = $GWCertificates; }) > $Null
					$ScriptInformation.Add(@{Data = "Log level"; Value = $GWStatus.LogLevel; }) > $Null
					$ScriptInformation.Add(@{Data = "ID"; Value = $GW.Id; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Gateways $($GW.Server)"
					Line 3 "Server`t`t: " $GW.Server
					Line 3 "Mode`t`t: " $GW.GWMode
					Line 3 "Status`t`t: " $GWStatus.AgentState
					Line 3 "Description`t: " $GW.Description
					Line 3 "Certificate`t: " $GWCertificates
					Line 3 "Log level`t: " $GWStatus.LogLevel
					Line 3 "ID`t`t: " $GW.Id
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Gateways $($GW.Server)"
					$rowdata = @()
					$columnHeaders = @("Server",($Script:htmlsb),$GW.Server,$htmlwhite)
					$rowdata += @(,( "Mode",($Script:htmlsb),$GW.GWMode.ToString(),$htmlwhite))
					$rowdata += @(,( "Status",($Script:htmlsb),$GWStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,( "Description",($Script:htmlsb),$GW.Description,$htmlwhite))
					$rowdata += @(,( "Certificate",($Script:htmlsb),$GWCertificates,$htmlwhite))
					$rowdata += @(,( "Log level",($Script:htmlsb),$GWStatus.LogLevel,$htmlwhite))
					$rowdata += @(,( "ID",($Script:htmlsb),$GW.Id,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Properties"
			}
			If($Text)
			{
				Line 2 "Properties"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Properties"
			}
			
			Switch ($GW.IPVersion)
			{
				"Version4"		{$IPVersion = "Version 4"; Break}
				"Version6"		{$IPVersion = "Version 6"; Break}
				"BothVersions"	{$IPVersion = "Both version 4 & 6"; Break}
				Default			{$IPVersion = "Unable to determine IP version: $($GW.IPVersion)"; Break}
			}
			
			$GWIPs = $GW.IPs.Split(";")
			
			If($GW.BindV4Addresses -eq "")
			{
				$GWBindV4Addresses = "All available addresses"
			}
			Else
			{
				$GWBindV4Addresses = $GW.BindV4Addresses
			}

			If($GW.OptimizeConnectionIPv4 -eq "<All>")
			{
				$GWOptimizeV4 = "All available addresses"
			}
			ElseIf($GW.OptimizeConnectionIPv4 -eq "<None>")
			{
				$GWOptimizeV4 = "None from the available"
			}
			Else
			{
				$GWOptimizeV4 = $GW.OptimizeConnectionIPv4
			}

			If($GW.BindV6Addresses -eq "")
			{
				$GWBindV6Addresses = "All available addresses"
			}
			Else
			{
				$GWBindV6Addresses = $GW.BindV6Addresses
			}

			If($GW.OptimizeConnectionIPv6 -eq "<All>")
			{
				$GWOptimizeV6 = "All available addresses"
			}
			ElseIf($GW.OptimizeConnectionIPv6 -eq "<None>")
			{
				$GWOptimizeV6 = "None from the available"
			}
			Else
			{
				$GWOptimizeV6 = $GW.OptimizeConnectionIPv6
			}

			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Enable RAS Secure Client Gateway in Site"; Value = $GW.Enabled.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Server"; Value = $GW.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $GW.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "IP version"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "     Use IP version"; Value = $IPVersion; }) > $Null
				
				$cnt = -1
				ForEach($Item in $GWIPs)
				{
					$cnt++
					If($cnt -eq 0 )
					{
						$ScriptInformation.Add(@{Data = "     IP(s)"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				
				$ScriptInformation.Add(@{Data = "Bind to IP"; Value = ""; }) > $Null
				If($GW.IPVersion -ne "Version6")
				{
					$ScriptInformation.Add(@{Data = "     Bind to the following IPv4"; Value = $GWBindV4Addresses; }) > $Null
					$ScriptInformation.Add(@{Data = "     Remove system buffers for"; Value = $GWOptimizeV4; }) > $Null
				}
				If($GW.IPVersion -ne "Version4")
				{
					$ScriptInformation.Add(@{Data = "     Bind to the following IPv6"; Value = $GWBindV6Addresses; }) > $Null
					$ScriptInformation.Add(@{Data = "     Remove system buffers for"; Value = $GWOptimizeV6; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Enable RAS Secure Client Gateway in Site: " $GW.Enabled.ToString()
				Line 3 "Server:`t`t`t`t`t" $GW.Server
				Line 3 "Description:`t`t`t`t" $GW.Description
				Line 3 "IP version" ""
				Line 4 "Use IP version:`t`t`t" $IPVersion
				
				$cnt = -1
				ForEach($Item in $GWIPs)
				{
					$cnt++
					If($cnt -eq 0 )
					{
						Line 4 "IP(s):`t`t`t`t" $Item
					}
					Else
					{
						Line 8 "  " $Item
					}
				}
				
				Line 3 "Bind to IP" ""
				If($GW.IPVersion -ne "Version6")
				{
					Line 4 "Bind to the following IPv4:`t" $GWBindV4Addresses
					Line 4 "Remove system buffers for:`t" $GWOptimizeV4
				}
				If($GW.IPVersion -ne "Version4")
				{
					Line 4 "Bind to the following IPv6:`t" $GWBindV6Addresses
					Line 4 "Remove system buffers for:`t" $GWOptimizeV6
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable RAS Secure Client Gateway in Site",($Script:htmlsb),$GW.Enabled.ToString(),$htmlwhite)
				$rowdata += @(,( "Server",($Script:htmlsb),$GW.Server,$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$GW.Description,$htmlwhite))
				$rowdata += @(,( "IP version",($Script:htmlsb),"",$htmlwhite))
				$rowdata += @(,( "     Use IP version",($Script:htmlsb),$IPVersion,$htmlwhite))
				
				$cnt = -1
				ForEach($Item in $GWIPs)
				{
					$cnt++
					If($cnt -eq 0 )
					{
						$rowdata += @(,( "     IP(s)",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,( "",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				
				$rowdata += @(,( "Bind to IP",($Script:htmlsb),"",$htmlwhite))
				If($GW.IPVersion -ne "Version6")
				{
					$rowdata += @(,( "     Bind to the following IPv4",($Script:htmlsb),$GWBindV4Addresses,$htmlwhite))
					$rowdata += @(,( "     Remove system buffers for",($Script:htmlsb),$GWOptimizeV4,$htmlwhite))
				}
				If($GW.IPVersion -ne "Version4")
				{
					$rowdata += @(,( "     Bind to the following IPv6",($Script:htmlsb),$GWBindV6Addresses,$htmlwhite))
					$rowdata += @(,( "     Remove system buffers for",($Script:htmlsb),$GWOptimizeV6,$htmlwhite))
				}

				$msg = "Properties"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
			
			#Mode
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Mode"
			}
			If($Text)
			{
				Line 2 "Mode"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Mode"
			}
			
			If($GW.InheritDefaultModeSettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $GWDefaults)
				{
					$GWMode = $GWDefaults.GWMode.ToString()
					$GWForwardRequests = $GWDefaults.NormalModeForwarding.ToString()
					If($GWMode -eq "Normal")
					{
						$GWServers = $GWDefaults.ForwardHttpServers.Split(";")
					}
					Else
					{
						$GWServers = $GWDefaults.ForwardGatewayServers.Split(";")
					}
				
					If($GWDefaults.PreferredPAId -eq 0)
					{
						$GWPreferredPublishingAgent = "Automatically"
					}
					Else
					{
						$GWPreferredPublishingAgent = (Get-PA -Id $GWDefaults.PreferredPAId -EA 0 4>$Null).Server
					}
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$GWMode = "Normal"
					$GWForwardRequests = "False"
					$GWServers = @("localhost:81")
					$GWPreferredPublishingAgent = "Automatically"
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this GW
				$GWMode = $GW.GWMode.ToString()
				$GWForwardRequests = $GW.NormalModeForwarding.ToString()
				If($GWMode -eq "Normal")
				{
					$GWServers = $GW.ForwardHttpServers.Split(";")
				}
				Else
				{
					$GWServers = $GW.ForwardGatewayServers.Split(";")
				}
			
				If($GW.PreferredPAId -eq 0)
				{
					$GWPreferredPublishingAgent = "Automatically"
				}
				Else
				{
					$GWPreferredPublishingAgent = (Get-PA -Id $GW.PreferredPAId -EA 0 4>$Null).Server
				}
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultModeSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Gateway mode"; Value = $GWMode; }) > $Null
				
				If($GWMode -eq "Normal")
				{
					If($GWForwardRequests -eq "True")
					{
						$ScriptInformation.Add(@{Data = "Forward requests to HTTP Server"; Value = $GWForwardRequests; }) > $Null
						
						$cnt = -1
						ForEach($Item in $GWServers)
						{
							$cnt++
							$tmparray = $Item.Split(":")
							$tmpserver = $tmparray[0]
							$tmpport = $tmparray[1]
							If($cnt -eq 0)
							{
								$ScriptInformation.Add(@{Data = "Server(s)"; Value = "Server: $($tmpserver)  Port: $($tmpport)"; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "Server: $($tmpserver)  Port: $($tmpport)"; }) > $Null
							}
						}
					}
					
					$ScriptInformation.Add(@{Data = "Preferred Publishing Agent"; Value = $GWPreferredPublishingAgent; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Forward requests to next RAS Secure Client Gateway in chain (cascaded firewall)"; Value = $GWForwardRequests; }) > $Null
					
					If($GWForwardRequests -eq "True")
					{
						$cnt = -1
						ForEach($Item in $GWServers)
						{
							$cnt++
							$tmparray = $Item.Split(":")
							$tmpserver = $tmparray[0]
							$tmpport = $tmparray[1]
							If($cnt -eq 0)
							{
								$ScriptInformation.Add(@{Data = "Server(s)"; Value = "Server: $($tmpserver)  Port: $($tmpport)"; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "Server: $($tmpserver)  Port: $($tmpport)"; }) > $Null
							}
						}
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultModeSettings.ToString()
				Line 3 "Gateway mode:`t`t`t`t" $GWMode
				If($GWMode -eq "Normal")
				{
					Line 3 "Forward requests to HTTP Server:`t" $GWForwardRequests
					
					$cnt = -1
					ForEach($Item in $GWServers)
					{
						$cnt++
						$tmparray = $Item.Split(":")
						$tmpserver = $tmparray[0]
						$tmpport = $tmparray[1]
						If($cnt -eq 0)
						{
							Line 3 "Server(s):`t`t`t`t" "Server: $($tmpserver)`tPort: $($tmpport)"
						}
						Else
						{
							Line 8 "  " "Server: $($tmpserver)`tPort: $($tmpport)"
						}
					}
					
					Line 3 "Preferred Publishing Agent:`t`t" $GWPreferredPublishingAgent
				}
				Else
				{
					Line 3 "Forward requests to next "
					Line 3 "RAS Secure Client Gateway"
					Line 3 "in chain (cascaded firewall):`t`t" $GWForwardRequests
					
					$cnt = -1
					ForEach($Item in $GWServers)
					{
						$cnt++
						$tmparray = $Item.Split(":")
						$tmpserver = $tmparray[0]
						$tmpport = $tmparray[1]
						If($cnt -eq 0)
						{
							Line 3 "Server(s):`t`t`t`t" "Server: $($tmpserver)`tPort: $($tmpport)"
						}
						Else
						{
							Line 8 "  " "Server: $($tmpserver)`tPort: $($tmpport)"
						}
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultModeSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Gateway mode",($Script:htmlsb),$GWMode,$htmlwhite))
				If($GWMode -eq "Normal")
				{
					$rowdata += @(,( "Forward requests to HTTP Server",($Script:htmlsb),$GWForwardRequests,$htmlwhite))
					
					$cnt = -1
					ForEach($Item in $GWServers)
					{
						$cnt++
						$tmparray = $Item.Split(":")
						$tmpserver = $tmparray[0]
						$tmpport = $tmparray[1]
						If($cnt -eq 0)
						{
							$rowdata += @(,( "Server(s)",($Script:htmlsb),"Server: $($tmpserver)  Port: $($tmpport)",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,( "",($Script:htmlsb),"Server: $($tmpserver)  Port: $($tmpport)",$htmlwhite))
						}
					}
					
					$rowdata += @(,( "Preferred Publishing Agent",($Script:htmlsb),$GWPreferredPublishingAgent,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,( "Forward requests to next RAS Secure Client Gateway in chain (cascaded firewall)",($Script:htmlsb),$GWForwardRequests,$htmlwhite))
					
					$cnt = -1
					ForEach($Item in $GWServers)
					{
						$cnt++
						$tmparray = $Item.Split(":")
						$tmpserver = $tmparray[0]
						$tmpport = $tmparray[1]
						If($cnt -eq 0)
						{
							$rowdata += @(,( "Server(s)",($Script:htmlsb),"Server: $($tmpserver)  Port: $($tmpport)",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,( "",($Script:htmlsb),"Server: $($tmpserver)  Port: $($tmpport)",$htmlwhite))
						}
					}
				}

				$msg = "Mode"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Network
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Network"
			}
			If($Text)
			{
				Line 2 "Network"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Network"
			}
			
			If($GW.InheritDefaultNetworkSettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $GWDefaults)
				{
					$GWEnableRASSecureClientGatewayPort = $GWDefaults.EnableGWPort.ToString()
					$GWRASSecureClientGatewayPort = $GWDefaults.GWPort.ToString()
					$GWEnableRDPPort = $GWDefaults.EnableRDP.ToString()
					$GWRDPPort = $GWDefaults.RDPPort.ToString()
					$GWEnableBroadcastRASSecureClientGatewayAddress = $GWDefaults.Broadcast.ToString()
					$GWEnableRDPUPDDataTunneling = $GWDefaults.EnableRDPUDP.ToString()
					$GWEnableClientManagerPort = $GWDefaults.EnableClientManagerPort.ToString()
					$GWClientManagerPort = "20009"
					$GWEnableRDPDOSAttackFilter = $GWDefaults.DOSPro.ToString()
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$GWEnableRASSecureClientGatewayPort = "True"
					$GWRASSecureClientGatewayPort = "80"
					$GWEnableRDPPort = "False"
					$GWEnableBroadcastRASSecureClientGatewayAddress = "True"
					$GWEnableRDPUPDDataTunneling = "True"
					$GWEnableClientManagerPort = "True"
					$GWClientManagerPort = "20009"
					$GWEnableRDPDOSAttackFilter = "True"
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this GW
				$GWEnableRASSecureClientGatewayPort = $GW.EnableGWPort.ToString()
				$GWRASSecureClientGatewayPort = $GW.GWPort.ToString()
				$GWEnableRDPPort = $GW.EnableRDP.ToString()
				$GWRDPPort = $GW.RDPPort.ToString()
				$GWEnableBroadcastRASSecureClientGatewayAddress = $GW.Broadcast.ToString()
				$GWEnableRDPUPDDataTunneling = $GW.EnableRDPUDP.ToString()
				$GWEnableClientManagerPort = $GW.EnableClientManagerPort.ToString()
				$GWClientManagerPort = "20009"
				$GWEnableRDPDOSAttackFilter = $GW.DOSPro.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultNetworkSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Enable RAS Secure Client Gateway Port"; Value = $GWEnableRASSecureClientGatewayPort; }) > $Null
				$ScriptInformation.Add(@{Data = "RAS Secure Client Gateway Port"; Value = $GWRASSecureClientGatewayPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable RDP Port"; Value = $GWEnableRDPPort; }) > $Null
				$ScriptInformation.Add(@{Data = "RDP Port"; Value = $GWRDPPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable Broadcast RAS Secure Client Gateway Address"; Value = $GWEnableBroadcastRASSecureClientGatewayAddress; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable RDP UDP Data Tunneling"; Value = $GWEnableRDPUPDDataTunneling; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable Client Manager Port"; Value = $GWEnableClientManagerPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Client Manager Port"; Value = $GWClientManagerPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Enable RDP DOS Attack Filter"; Value = $GWEnableRDPDOSAttackFilter; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultNetworkSettings.ToString()
				Line 3 "Enable RAS Secure Client Gateway Port:`t" $GWEnableRASSecureClientGatewayPort
				Line 3 "RAS Secure Client Gateway Port:`t`t" $GWRASSecureClientGatewayPort
				Line 3 "Enable RDP Port:`t`t`t" $GWEnableRDPPort
				Line 3 "RDP Port:`t`t`t`t" $GWRDPPort
				Line 3 "Enable Broadcast RAS Secure "
				Line 3 "Client Gateway Address:`t`t`t" $GWEnableBroadcastRASSecureClientGatewayAddress
				Line 3 "Enable RDP UDP Data Tunneling:`t`t" $GWEnableRDPUPDDataTunneling
				Line 3 "Enable Client Manager Port:`t`t" $GWEnableClientManagerPort
				Line 3 "Client Manager Port:`t`t`t" $GWClientManagerPort
				Line 3 "Enable RDP DOS Attack Filter:`t`t" $GWEnableRDPDOSAttackFilter
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultNetworkSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Enable RAS Secure Client Gateway Port",($Script:htmlsb),$GWEnableRASSecureClientGatewayPort,$htmlwhite))
				$rowdata += @(,( "RAS Secure Client Gateway Port",($Script:htmlsb),$GWRASSecureClientGatewayPort,$htmlwhite))
				$rowdata += @(,( "Enable RDP Port",($Script:htmlsb),$GWEnableRDPPort,$htmlwhite))
				$rowdata += @(,( "RDP Port",($Script:htmlsb),$GWRDPPort,$htmlwhite))
				$rowdata += @(,( "Enable Broadcast RAS Secure Client Gateway Address",($Script:htmlsb),$GWEnableBroadcastRASSecureClientGatewayAddress,$htmlwhite))
				$rowdata += @(,( "Enable RDP UDP Data Tunneling",($Script:htmlsb),$GWEnableRDPUPDDataTunneling,$htmlwhite))
				$rowdata += @(,( "Enable Client Manager Port",($Script:htmlsb),$GWEnableClientManagerPort,$htmlwhite))
				$rowdata += @(,( "Client Manager Port",($Script:htmlsb),$GWClientManagerPort,$htmlwhite))
				$rowdata += @(,( "Enable RDP DOS Attack Filter",($Script:htmlsb),$GWEnableRDPDOSAttackFilter,$htmlwhite))

				$msg = "Network"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#SSL/TLS
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "SSL/TLS"
			}
			If($Text)
			{
				Line 2 "SSL/TLS"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "SSL/TLS"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultSslTlsSettings.ToString(); }) > $Null

				If($GWEnableHSTS -eq "False")
				{
					$ScriptInformation.Add(@{Data = "HSTS is off"; Value = ""; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "HSTS is on"; Value = ""; }) > $Null
					$ScriptInformation.Add(@{Data = "Enforce HTTP strict transport security (HSTS)"; Value = $GWEnableHSTS; }) > $Null
					$ScriptInformation.Add(@{Data = "Max-age"; Value = "$GWHSTSMaxage months"; }) > $Null
					$ScriptInformation.Add(@{Data = "Include subdomains"; Value = $GWHSTSIncludeSubdomains; }) > $Null
					$ScriptInformation.Add(@{Data = "Preload"; Value = $GWHSTSPreload; }) > $Null
				}

				$ScriptInformation.Add(@{Data = "Enable SSL"; Value = $GWEnableSSL; }) > $Null
				$ScriptInformation.Add(@{Data = "on Port"; Value = $GWEnableSSLOnPort; }) > $Null
				$ScriptInformation.Add(@{Data = "Accepted SSL Versions"; Value = $GWAcceptedSSLVersions; }) > $Null
				$ScriptInformation.Add(@{Data = "Cipher Strength"; Value = $GWCipherStrength; }) > $Null
				$ScriptInformation.Add(@{Data = "Cipher"; Value = $GWCipher; }) > $Null
				$ScriptInformation.Add(@{Data = "Certificates"; Value = $GWCertificates; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultSslTlsSettings.ToString()

				If($GWEnableHSTS -eq "False")
				{
					Line 3 "HSTS is off"
				}
				Else
				{
					Line 3 "HSTS is on"
					Line 3 "Enforce HTTP strict transport security:`t" $GWEnableHSTS
					Line 3 "Max-age:`t`t`t`t" "$GWHSTSMaxage months"
					Line 3 "Include subdomains:`t`t`t" $GWHSTSIncludeSubdomains
					Line 3 "Preload:`t`t`t`t" $GWHSTSPreload
				}

				Line 3 "Enable SSL:`t`t`t`t" $GWEnableSSL
				Line 3 "on Port:`t`t`t`t" $GWEnableSSLOnPort
				Line 3 "Accepted SSL Versions:`t`t`t" $GWAcceptedSSLVersions
				Line 3 "Cipher Strength:`t`t`t" $GWCipherStrength
				Line 3 "Cipher:`t`t`t`t`t" $GWCipher
				Line 3 "Certificates:`t`t`t`t" $GWCertificates
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultSslTlsSettings.ToString(),$htmlwhite)

				If($GWEnableHSTS -eq "False")
				{
					$rowdata += @(,( "HSTS is off",($Script:htmlsb),"",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,( "HSTS is on",($Script:htmlsb),,$htmlwhite))
					$rowdata += @(,( "Enforce HTTP strict transport security (HSTS)",($Script:htmlsb),$GWEnableHSTS,$htmlwhite))
					$rowdata += @(,( "Max-age",($Script:htmlsb),"$GWHSTSMaxage months",$htmlwhite))
					$rowdata += @(,( "Include subdomains",($Script:htmlsb),$GWHSTSIncludeSubdomains,$htmlwhite))
					$rowdata += @(,( "Preload",($Script:htmlsb),$GWHSTSPreload,$htmlwhite))
				}

				$rowdata += @(,( "Enable SSL",($Script:htmlsb),$GWEnableSSL,$htmlwhite))
				$rowdata += @(,( "on Port",($Script:htmlsb),$GWEnableSSLOnPort,$htmlwhite))
				$rowdata += @(,( "Accepted SSL Versions",($Script:htmlsb),$GWAcceptedSSLVersions,$htmlwhite))
				$rowdata += @(,( "Cipher Strength",($Script:htmlsb),$GWCipherStrength,$htmlwhite))
				$rowdata += @(,( "Cipher",($Script:htmlsb),$GWCipher,$htmlwhite))
				$rowdata += @(,( "Certificates",($Script:htmlsb),$GWCertificates,$htmlwhite))

				$msg = "SSL/TLS"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#HTML5
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "HTML5"
			}
			If($Text)
			{
				Line 2 "HTML5"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "HTML5"
			}
			
			If($GW.InheritDefaultHTML5Settings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $GWDefaults)
				{
					$GWEnableHTML5Client = $GWDefaults.EnableHTML5.ToString()
					
					Switch($GWDefaults.LaunchMethod)
					{
						"ParallelsClientAndHTML5"	{$GWLaunchSessionsUsing = "Launch apps with Parallels Client & Fallback to HTML5"; Break}
						"ParallelsClient"			{$GWLaunchSessionsUsing = "Launch apps with Parallels Client"; Break}
						"HTML5"						{$GWLaunchSessionsUsing = "Launch apps in Browser only (HTML5 Only)"; Break}
						Default						{$GWLaunchSessionsUsing = "Unable to determine Launch sessions using: $($GWDefaults.LaunchMethod)"; Break}
					}
					
					$GWAllowLaunchMethod = $GWDefaults.AllowLaunchMethod.ToString()
					$GWAllowAppsInNewTab = $GWDefaults.AllowAppsInNewTab.ToString()
					$GWUsePreWin2000LoginFormat = $GWDefaults.UsePreWin2000LoginFormat.ToString()
					$GWAllowEmbed = $GWDefaults.AllowEmbed.ToString()
					$GWAllowFileTransfer = $GWDefaults.AllowFileTransfer.ToString()
					$GWAllowClipboard = $GWDefaults.AllowClipboard.ToString()
					$GWEnableAlternateNLBHostname = $GWDefaults.EnableAlternateNLBHost.ToString()
					$GWAlternameNLBHostname = $GWDefaults.AlternateNLBHost
					$GWEnableAlternateNLBPort = $GWDefaults.EnableAlternateNLBPort.ToString()
					$GWAlternateNLBPort = $GWDefaults.AlternateNLBPort.ToString()
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$GWEnableHTML5Client = "True"
					$GWLaunchSessionsUsing = "Launch apps with Parallels CLient & Fallback to HTML5"
					$GWAllowLaunchMethod = "True"
					$GWAllowAppsInNewTab = "False"
					$GWUsePreWin2000LoginFormat = "True"
					$GWAllowEmbed = "False"
					$GWAllowFileTransfer = "True"
					$GWAllowClipboard = "True"
					$GWEnableAlternateNLBHostname = "False"
					$GWAlternameNLBHostname = ""
					$GWEnableAlternateNLBPort = "False"
					$GWAlternateNLBPort = "8443"
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this GW
				$GWEnableHTML5Client = $GW.EnableHTML5.ToString()
				
				Switch($GW.LaunchMethod)
				{
					"ParallelsClientAndHTML5"	{$GWLaunchSessionsUsing = "Launch apps with Parallels Client & Fallback to HTML5"; Break}
					"ParallelsClient"			{$GWLaunchSessionsUsing = "Launch apps with Parallels Client"; Break}
					"HTML5"						{$GWLaunchSessionsUsing = "Launch apps in Browser only (HTML5 Only)"; Break}
					Default						{$GWLaunchSessionsUsing = "Unable to determine Launch sessions using: $($GW.LaunchMethod)"; Break}
				}
				
				$GWAllowLaunchMethod = $GW.AllowLaunchMethod.ToString()
				$GWAllowAppsInNewTab = $GW.AllowAppsInNewTab.ToString()
				$GWUsePreWin2000LoginFormat = $GW.UsePreWin2000LoginFormat.ToString()
				$GWAllowEmbed = $GW.AllowEmbed.ToString()
				$GWAllowFileTransfer = $GW.AllowFileTransfer.ToString()
				$GWAllowClipboard = $GW.AllowClipboard.ToString()
				$GWEnableAlternateNLBHostname = $GW.EnableAlternateNLBHost.ToString()
				$GWAlternameNLBHostname = $GW.AlternateNLBHost
				$GWEnableAlternateNLBPort = $GW.EnableAlternateNLBPort.ToString()
				$GWAlternateNLBPort = $GW.AlternateNLBPort.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultHTML5Settings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Enable HTML5 Client"; Value = $GWEnableHTML5Client; }) > $Null
				$ScriptInformation.Add(@{Data = "Client"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "     Launch sessions using"; Value = $GWLaunchSessionsUsing; }) > $Null
				$ScriptInformation.Add(@{Data = "     Allow user to select a launch method"; Value = $GWAllowLaunchMethod; }) > $Null
				$ScriptInformation.Add(@{Data = "     Allow opening applications in a new tab"; Value = $GWAllowAppsInNewTab; }) > $Null
				$ScriptInformation.Add(@{Data = "     Use Pre Windows 2000 login format"; Value = $GWUsePreWin2000LoginFormat; }) > $Null
				$ScriptInformation.Add(@{Data = "Restrictions"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "     Allow embedding of Parallels HTML5 Client into other web pages"; Value = $GWAllowEmbed; }) > $Null
				$ScriptInformation.Add(@{Data = "     Allow file transfer command"; Value = $GWAllowFileTransfer; }) > $Null
				$ScriptInformation.Add(@{Data = "     Allow clipboard command"; Value = $GWAllowClipboard; }) > $Null
				$ScriptInformation.Add(@{Data = "Network Load Balancer access"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "     Use alternate hostname"; Value = $GWEnableAlternateNLBHostname; }) > $Null
				If($GWEnableAlternateNLBHostname -eq "True")
				{
					$ScriptInformation.Add(@{Data = ""; Value = $GWAlternameNLBHostname; }) > $Null
				}
				$ScriptInformation.Add(@{Data = "     Use alternate port"; Value = $GWEnableAlternateNLBPort; }) > $Null
				If($GWEnableAlternateNLBPort -eq "True")
				{
					$ScriptInformation.Add(@{Data = ""; Value = $GWAlternateNLBPort; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultHTML5Settings.ToString()
				Line 3 "Enable HTML5 Client:`t`t`t" $GWEnableHTML5Client
				Line 3 "Client" ""
				Line 4 "Launch sessions using:`t`t" $GWLaunchSessionsUsing
				Line 4 "Allow user to select "
				Line 4 "a launch method:`t`t" $GWAllowLaunchMethod
				Line 4 "Allow opening applications "
				Line 4 "in a new tab:`t`t`t" $GWAllowAppsInNewTab
				Line 4 "Use Pre Windows 2000 "
				Line 4 "login format:`t`t`t" $GWUsePreWin2000LoginFormat
				Line 3 "Restrictions" ""
				Line 4 "Allow embedding of Parallels "
				Line 4 "HTML5 Client into other "
				Line 4 "web pages:`t`t`t" $GWAllowEmbed
				Line 4 "Allow file transfer command:`t" $GWAllowFileTransfer
				Line 4 "Allow clipboard command:`t" $GWAllowClipboard
				Line 3 "Network Load Balancer access" ""
				Line 4 "Use alternate hostname:`t`t" $GWEnableAlternateNLBHostname
				If($GWEnableAlternateNLBHostname -eq "True")
				{
					Line 8 $GWAlternameNLBHostname
				}
				Line 4 "Use alternate port:`t`t" $GWEnableAlternateNLBPort
				If($GWEnableAlternateNLBPort -eq "True")
				{
					Line 8 $GWAlternateNLBPort
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultHTML5Settings.ToString(),$htmlwhite)
				$rowdata += @(,( "Enable HTML5 Client",($Script:htmlsb),$GWEnableHTML5Client,$htmlwhite))
				$rowdata += @(,( "Client",($Script:htmlsb),"",$htmlwhite))
				$rowdata += @(,( "     Launch sessions using",($Script:htmlsb),$GWLaunchSessionsUsing,$htmlwhite))
				$rowdata += @(,( "     Allow user to select a launch method",($Script:htmlsb),$GWAllowLaunchMethod,$htmlwhite))
				$rowdata += @(,( "     Allow opening applications in a new tab",($Script:htmlsb),$GWAllowAppsInNewTab,$htmlwhite))
				$rowdata += @(,( "     Use Pre Windows 2000 login format",($Script:htmlsb),$GWUsePreWin2000LoginFormat,$htmlwhite))
				$rowdata += @(,( "Restrictions",($Script:htmlsb),"",$htmlwhite))
				$rowdata += @(,( "     Allow embedding of Parallels HTML5 Client into other web pages",($Script:htmlsb),$GWAllowEmbed,$htmlwhite))
				$rowdata += @(,( "     Allow file transfer command",($Script:htmlsb),$GWAllowFileTransfer,$htmlwhite))
				$rowdata += @(,( "     Allow clipboard command",($Script:htmlsb),$GWAllowClipboard,$htmlwhite))
				$rowdata += @(,( "Network Load Balancer access",($Script:htmlsb),"",$htmlwhite))
				$rowdata += @(,( "     Use alternate hostname",($Script:htmlsb),$GWEnableAlternateNLBHostname,$htmlwhite))
				If($GWEnableAlternateNLBHostname -eq "True")
				{
					$rowdata += @(,( "",($Script:htmlsb),$GWAlternameNLBHostname,$htmlwhite))
				}
				$rowdata += @(,( "     Use alternate port",($Script:htmlsb),$GWEnableAlternateNLBPort,$htmlwhite))
				If($GWEnableAlternateNLBPort -eq "True")
				{
					$rowdata += @(,( "",($Script:htmlsb),$GWAlternateNLBPort,$htmlwhite))
				}

				$msg = "HTML5"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Wyse
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Wyse"
			}
			If($Text)
			{
				Line 2 "Wyse"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Wyse"
			}
			
			If($GW.InheritDefaultWyseSettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $GWDefaults)
				{
					$GWEnableWyse = $GWDefaults.EnableWyseSupport.ToString()
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$GWEnableWyse = "True"
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this GW
				$GWEnableWyse = $GW.EnableWyseSupport.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultWyseSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Enable Wyse ThinOS Support"; Value = $GWEnableWyse; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultWyseSettings.ToString()
				Line 3 "Enable Wyse ThinOS Support:`t`t" $GWEnableWyse
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultWyseSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Enable Wyse ThinOS Support",($Script:htmlsb),$GWEnableWyse,$htmlwhite))

				$msg = "Wyse"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Security
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Security"
			}
			If($Text)
			{
				Line 2 "Security"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Security"
			}
			
			If($GW.InheritDefaultSecuritySettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $GWDefaults)
				{
					If($GWDefaults.SecurityMode -eq "AllowAllExcept")
					{
						$GWAllow = "AllowAllExcept"
						$MACAddresses = @()
						ForEach($Item in $GWDefaults.MACAllowExcept)
						{
							$MACAddresses += $Item
						}
					}
					Else
					{
						$GWAllow = "AllowOnly"
						$MACAddresses = @()
						ForEach($Item in $GWDefaults.MACAllowOnly)
						{
							$MACAddresses += $Item
						}
					}
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$GWAllow = "AllowAllExcept"
					$MACAddresses = @()
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this GW
				If($GW.SecurityMode -eq "AllowAllExcept")
				{
					$GWAllow = "AllowAllExcept"
					$MACAddresses = @()
					ForEach($Item in $GW.MACAllowExcept)
					{
						$MACAddresses += $Item
					}
				}
				Else
				{
					$GWAllow = "AllowOnly"
					$MACAddresses = @()
					ForEach($Item in $GW.MACAllowOnly)
					{
						$MACAddresses += $Item
					}
				}
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultSecuritySettings.ToString(); }) > $Null
				If($GWAllow -eq "AllowAllExcept")
				{
					$ScriptInformation.Add(@{Data = "Allow all except"; Value = ""; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Allow only"; Value = ""; }) > $Null
				}
				$cnt =-1
				ForEach($Item in $MACAddresses)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "MAC Address"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultSecuritySettings.ToString()
				If($GWAllow -eq "AllowAllExcept")
				{
					Line 3 "Allow all except"
				}
				Else
				{
					Line 3 "Allow only"
				}
				$cnt =-1
				ForEach($Item in $MACAddresses)
				{
					$cnt++
					If($cnt -eq 0)
					{
						Line 3 "MAC Address:`t`t`t`t" $Item
					}
					Else
					{
						Line 8 "  " $Item
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultSecuritySettings.ToString(),$htmlwhite)
				If($GWAllow -eq "AllowAllExcept")
				{
					$rowdata += @(,( "Allow all except",($Script:htmlsb),"",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,( "Allow only",($Script:htmlsb),"",$htmlwhite))
				}
				$cnt =-1
				ForEach($Item in $MACAddresses)
				{
					$cnt++
					If($cnt -eq 0)
					{
						$rowdata += @(,( "MAC Address",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,( "",($Script:htmlsb),$Item,$htmlwhite))
					}
				}

				$msg = "Security"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}

			#Web
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Web"
			}
			If($Text)
			{
				Line 2 "Web"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Web"
			}
			
			If($GW.InheritDefaultWebSettings)
			{
				#do we inherit site defaults?
				#yes we do, get the default settings for the Site
				#use the Site default settings

				$GWDefaults = Get-GWDefaultSettings -SiteId $Site.Id -EA 0 4>$Null
				
				If($? -and $Null -ne $GWDefaults)
				{
					$GWDefaultURL = $GWDefaults.WebRequestsURL
					$GWWebCookie = $GWDefaults.WebCookie
				}
				Else
				{
					#unable to retrieve default, use built-in default values
					$GWDefaultURL = "https://%hostname%/RASHTML5Gateway"
					$GWWebCookie = "ASP.NET_SessionId"
				}
			}
			Else
			{
				#we don't inherit settings
				#get the settings configured for this GW
				$GWDefaultURL = $GW.WebRequestsURL
				$GWWebCookie = $GW.WebCookie
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $GW.InheritDefaultWebSettings.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Default URL"; Value = $GWDefaultURL; }) > $Null
				$ScriptInformation.Add(@{Data = "Web cookie"; Value = $GWWebCookie; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Inherit default settings:`t`t" $GW.InheritDefaultWebSettings.ToString()
				Line 3 "Default URL:`t`t`t`t" $GWDefaultURL
				Line 3 "Web cookie:`t`t`t`t" $GWWebCookie
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$GW.InheritDefaultWebSettings.ToString(),$htmlwhite)
				$rowdata += @(,( "Default URL",($Script:htmlsb),$GWDefaultURL,$htmlwhite))
				$rowdata += @(,( "Web cookie",($Script:htmlsb),$GWWebCookie,$htmlwhite))

				$msg = "Web"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}

	$PAs = Get-PA -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Publishing Agents for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Publishing Agents for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Publishing Agents for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Publishing Agents for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Publishing Agents retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Publishing Agents retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No Publishing Agents retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Publishing Agents retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Publishing Agents"
		}
		If($Text)
		{
			Line 1 "Publishing Agents"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Publishing Agents"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Publishing Agents"
		ForEach($PA in $PAs)
		{
			$PAStatus = Get-PAStatus -Id $PA.Id -EA 0 4>$Null
			
			If(!$?)
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve Publishing Agent Status for Publishing Agent $($PA.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve Publishing Agent Status for Publishing Agent $($PA.Id)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve Publishing Agent Status for Publishing Agent $($PA.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve Publishing Agent Status for Publishing Agent $($PA.Id)"
				}
			}
			ElseIf($? -and $Null -eq $GWStatus)
			{
				Write-Warning "
				`n`n
				`t`t
				No Publishing Agent Status retrieved for Publishing Agent $($PA.Id)`
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "No Publishing Agent Status retrieved for Publishing Agent $($PA.Id)"
				}
				If($Text)
				{
					Line 0 "No Publishing Agent Status retrieved for Publishing Agent $($PA.Id)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "No Publishing Agent Status retrieved for Publishing Agent $($PA.Id)"
				}
			}
			Else
			{
				If($PA.Standby -eq $False)
				{
					$PAPriority = "Master"
				}
				Else
				{
					$PAPriority = "Standby"
				}

				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "Publishing Agents $($PA.Server)"
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Server"; Value = $PA.Server; }) > $Null
					$ScriptInformation.Add(@{Data = "Priority"; Value = $PAPriority; }) > $Null
					$ScriptInformation.Add(@{Data = "Status"; Value = $PAStatus.AgentState; }) > $Null
					$ScriptInformation.Add(@{Data = "Log level"; Value = $PAStatus.LogLevel; }) > $Null
					#$ScriptInformation.Add(@{Data = "Last modification by"; Value = ""; }) > $Null
					#$ScriptInformation.Add(@{Data = "Modified on"; Value = ""; }) > $Null
					#$ScriptInformation.Add(@{Data = "Created by"; Value = ""; }) > $Null
					#$ScriptInformation.Add(@{Data = "Created on"; Value = ""; }) > $Null
					$ScriptInformation.Add(@{Data = "ID"; Value = $PA.Id; }) > $Null

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Publishing Agents $($PA.Server)"
					Line 3 "Server`t`t: " $PA.Server
					Line 3 "Priority`t: " $PAPriority
					Line 3 "Status`t`t: " $PAStatus.AgentState
					Line 3 "Log level`t: " $PAStatus.LogLevel
					#Line 3 "Last modification by: " ""
					#Line 3 "Modified on: " ""
					#Line 3 "Created by: " ""
					#Line 3 "Created on: " ""
					Line 3 "ID`t`t: " $PA.Id
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Publishing Agents $($PA.Server)"
					$rowdata = @()
					$columnHeaders = @("Server",($Script:htmlsb),$PA.Server,$htmlwhite)
					$rowdata += @(,( "Priority",($Script:htmlsb),$PAPriority,$htmlwhite))
					$rowdata += @(,( "Status",($Script:htmlsb),$PAStatus.AgentState.ToString(),$htmlwhite))
					$rowdata += @(,( "Log level",($Script:htmlsb),$PAStatus.LogLevel,$htmlwhite))
					#$rowdata += @(,( "Last modification by",($Script:htmlsb),"",$htmlwhite))
					#$rowdata += @(,( "Modified on",($Script:htmlsb),"",$htmlwhite))
					#$rowdata += @(,( "Created by",($Script:htmlsb),"",$htmlwhite))
					#$rowdata += @(,( "Created on",($Script:htmlsb),"",$htmlwhite))
					$rowdata += @(,( "ID",($Script:htmlsb),$PA.Id,$htmlwhite))

					$msg = ""
					$columnWidths = @("200","275")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
					WriteHTMLLine 0 0 ""
				}
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Properties"
			}
			If($Text)
			{
				Line 2 "Properties"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Properties"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Enable Server in Site"; Value = $PA.Enabled.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Server"; Value = $PA.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "IP"; Value = $PA.IP; }) > $Null
				$ScriptInformation.Add(@{Data = "Alternative IPs"; Value = $PA.AlternativeIPs; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PA.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Standby"; Value = $PA.Standby.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Enable Server in Site:`t`t`t" $PA.Enabled.ToString()
				Line 3 "Server:`t`t`t`t`t" $PA.Server
				Line 3 "IP:`t`t`t`t`t" $PA.IP
				Line 3 "Alternative IPs:`t`t`t" $PA.AlternativeIPs
				Line 3 "Description:`t`t`t`t" $PA.Description
				Line 3 "Standby:`t`t`t`t" $PA.Standby.ToString()
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Enable Server in Site",($Script:htmlsb),$PA.Enabled.ToString(),$htmlwhite)
				$rowdata += @(,( "Server",($Script:htmlsb),$PA.Server,$htmlwhite))
				$rowdata += @(,( "IP",($Script:htmlsb),$PA.IP,$htmlwhite))
				$rowdata += @(,( "Alternative IPs",($Script:htmlsb),$PA.AlternativeIPs,$htmlwhite))
				$rowdata += @(,( "Description",($Script:htmlsb),$PA.Description,$htmlwhite))
				$rowdata += @(,( "Standby",($Script:htmlsb),$PA.Standby.ToString(),$htmlwhite))

				$msg = "Properties"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}
	
	#Enrollment Servers - not in PoSH
	
	#HALB - not in PoSH
	
	#Themes - not in PoSH
	
	#Certificates

	$Certs = Get-RASCertificate -Siteid $Site.Id -EA 0 4> $Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Certificates for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Certificates for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Certificates for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Certificates for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Certificates retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Certificates retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No Certificates retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Certificates retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Certificates"
		}
		If($Text)
		{
			Line 1 "Certificates"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Certificates"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Certificates"
		ForEach($Cert in $Certs)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Certificate $($Cert.Name)"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $Cert.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Status"; Value = $Cert.Status.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Expiration date"; Value = $Cert.ExpirationDate; }) > $Null
				$ScriptInformation.Add(@{Data = "Usage"; Value = $Cert.Usage; }) > $Null
				$ScriptInformation.Add(@{Data = "Common name"; Value = $Cert.CommonName; }) > $Null
				$ScriptInformation.Add(@{Data = "ID"; Value = $Cert.Id; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Certificate $($Cert.Name)"
				Line 3 "Name`t`t: " $Cert.Name
				Line 3 "Status`t`t: " $Cert.Status.ToString()
				Line 3 "Expiration date`t: " $Cert.ExpirationDate
				Line 3 "Usage`t`t: " $Cert.Usage
				Line 3 "Common name`t: " $Cert.CommonName
				Line 3 "ID`t`t: " $Cert.Id
				Line 0 ""
			}
			If($HTML)
			{
				$CertName = $Cert.Name.Replace("<","").Replace(">","")
				WriteHTMLLine 3 0 "Certificate $CertName"
				$rowdata = @()
				$columnHeaders = @("Name",($Script:htmlsb),$Cert.Name.Replace("<","").Replace(">",""),$htmlwhite)
				$rowdata += @(,("Status",($Script:htmlsb),$Cert.Status.ToString(),$htmlwhite))
				$rowdata += @(,("Expiration date",($Script:htmlsb),$Cert.ExpirationDate,$htmlwhite))
				$rowdata += @(,("Usage",($Script:htmlsb),$Cert.Usage,$htmlwhite))
				$rowdata += @(,("Common name",($Script:htmlsb),$Cert.CommonName,$htmlwhite))
				$rowdata += @(,("ID",($Script:htmlsb),$Cert.Id,$htmlwhite))

				$msg = ""
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Properties"
			}
			If($Text)
			{
				Line 2 "Properties"
			}
			If($HTML)
			{
				#WriteHTMLLine 4 0 "Properties"
			}
			
			Switch($Cert.KeySize)
			{
				"KeySize1024"		{$KeySize = "1024"; Break}
				"KeySize2048"		{$KeySize = "2048"; Break}
				"KeySize4096"		{$KeySize = "4096"; Break}
				"KeySizeUnknown"	{$KeySize = ""; Break}
				Default				{$KeySize = "Unable to determine certificate key size: $($Cert.KeySize)"; Break}
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $Cert.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $Cert.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Usage"; Value = $Cert.Usage; }) > $Null
				$ScriptInformation.Add(@{Data = "State"; Value = $Cert.Status.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Key size"; Value = $KeySize; }) > $Null
				$ScriptInformation.Add(@{Data = "Common name"; Value = $Cert.CommonName; }) > $Null
				$ScriptInformation.Add(@{Data = "Expiration date"; Value = $Cert.ExpirationDate; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Name:`t`t`t`t`t" $Cert.Name
				Line 3 "Description:`t`t`t`t" $Cert.Description
				Line 3 "Usage:`t`t`t`t`t" $Cert.Usage
				Line 3 "State:`t`t`t`t`t" $Cert.Status.ToString()
				Line 3 "Key size:`t`t`t`t" $KeySize
				Line 3 "Common name:`t`t`t`t" $Cert.CommonName
				Line 3 "Expiration date:`t`t`t" $Cert.ExpirationDate
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Name",($Script:htmlsb),$Cert.Name.Replace("<","").Replace(">",""),$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$Cert.Description,$htmlwhite))
				$rowdata += @(,("Usage",($Script:htmlsb),$Cert.Usage,$htmlwhite))
				$rowdata += @(,("State",($Script:htmlsb),$Cert.Status.ToString(),$htmlwhite))
				$rowdata += @(,("Key size",($Script:htmlsb),$KeySize,$htmlwhite))
				$rowdata += @(,("Common name",($Script:htmlsb),$Cert.CommonName,$htmlwhite))
				$rowdata += @(,("Expiration date",($Script:htmlsb),$Cert.ExpirationDate,$htmlwhite))

				$msg = "Properties"
				$columnWidths = @("200","275")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
	}

	#Settings
	
	#Auditing - not in PoSH
	#Global logging - not in PoSH

	$FarmSettings = Get-RASFarmSettings -ea 0 4>$Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Certificates for Site $($Site.Name)`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Certificates for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Certificates for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Certificates for Site $($Site.Name)"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Certificates retrieved for Site $($Site.Name).`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Certificates retrieved for Site $($Site.Name)"
		}
		If($Text)
		{
			Line 0 "No Certificates retrieved for Site $($Site.Name)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Certificates retrieved for Site $($Site.Name)"
		}
	}
	Else
	{
	
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Settings"
		}
		If($Text)
		{
			Line 1 "Settings"
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Settings"
		}

		Write-Verbose "$(Get-Date): `t`tOutput Settings"
	}
		
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "URL redirection"
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Do not redirect the following URLs"; Value = ""; }) > $Null
		$cnt = -1
		ForEach($Item in $FarmSettings.URLBlacklist)
		{
			$cnt++
			
			If($Cnt -eq 0)
			{
				$ScriptInformation.Add(@{Data = "Url"; Value = $Item; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
			}
		}

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		If($FarmSettings.ReplicateURLRedirection)
		{
			WriteWordLine 0 0 "Settings are replicated to all Sites"
		}
		Else
		{
			WriteWordLine 0 0 "Settings are not replicated to all Sites"
		}
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "URL redirection"
		Line 3 "Do not redirect the following URLs:"
		$cnt = -1
		ForEach($Item in $FarmSettings.URLBlacklist)
		{
			$cnt++
			
			If($Cnt -eq 0)
			{
				Line 8 "Url:" $Item
			}
			Else
			{
				Line 10 $Item
			}
		}
		If($FarmSettings.ReplicateURLRedirection)
		{
			Line 3 "Settings are replicated to all Sites"
		}
		Else
		{
			Line 3 "Settings are not replicated to all Sites"
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "URL redirection"
		$rowdata = @()
		$columnHeaders = @("Do not redirect the following URLs",($Script:htmlsb),"",$htmlwhite)
		$cnt = -1
		ForEach($Item in $FarmSettings.URLBlacklist)
		{
			$cnt++
			
			If($Cnt -eq 0)
			{
				$rowdata += @(,("URL",($Script:htmlsb),$Item,$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
			}
		}

		$msg = ""
		$columnWidths = @("200","275")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		If($FarmSettings.ReplicateURLRedirection)
		{
			WriteHTMLLine 0 0 "Settings are replicated to all Sites"
		}
		Else
		{
			WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
		}
		WriteHTMLLine 0 0 ""
	}
		
		
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Client settings"
		$ScriptInformation = New-Object System.Collections.ArrayList
		If($FarmSettings.SendHDIcons)
		{
			$ScriptInformation.Add(@{Data = "Published application icons"; Value = "Send high resolution icons (uses more network bandwidth)"; }) > $Null
		}
		Else
		{
			$ScriptInformation.Add(@{Data = "Published application icons"; Value = "Send standard resolution icons"; }) > $Null
		}
		$ScriptInformation.Add(@{Data = "Enable overlay icon"; Value = $FarmSettings.EnableOverlayIcons.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Show password expiration reminder"; Value = $FarmSettings.ShowPasswordExpiry.ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		If($FarmSettings.ReplicateSendHDIcons)
		{
			WriteWordLine 0 0 "Settings are replicated to all Sites"
		}
		Else
		{
			WriteWordLine 0 0 "Settings are not replicated to all Sites"
		}
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Client settings"
		If($FarmSettings.SendHDIcons)
		{
			Line 3 "Published application icons:`t`t" "Send high resolution icons (uses more network bandwidth)" ""
		}
		Else
		{
			Line 3 "Published application icons:`t`t" "Send standard resolution icons" ""
		}
		Line 3 "Enable overlay icon:`t`t`t" $FarmSettings.EnableOverlayIcons.ToString()
		Line 3 "Show password expiration reminder:`t" $FarmSettings.ShowPasswordExpiry.ToString()
		If($FarmSettings.ReplicateSendHDIcons)
		{
			Line 3 "Settings are replicated to all Sites"
		}
		Else
		{
			Line 3 "Settings are not replicated to all Sites"
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Client settings"
		$rowdata = @()
		If($FarmSettings.SendHDIcons)
		{
			$columnHeaders = @("Published application icons",($Script:htmlsb),"Send high resolution icons (uses more network bandwidth)",$htmlwhite)
		}
		Else
		{
			$columnHeaders = @("Published application icons",($Script:htmlsb),"Send standard resolution icons",$htmlwhite)
		}
		$rowdata += @(,("Enable overlay icon",($Script:htmlsb),$FarmSettings.EnableOverlayIcons.ToString(),$htmlwhite))
		$rowdata += @(,("Show password expiration reminder",($Script:htmlsb),$FarmSettings.ShowPasswordExpiry.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("200","275")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		If($FarmSettings.ReplicateSendHDIcons)
		{
			WriteHTMLLine 0 0 "Settings are replicated to all Sites"
		}
		Else
		{
			WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
		}
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region process load balancing
Function ProcessLoadBalancing
{
	Write-Verbose "$(Get-Date): Processing Load balancing"
	
	OutputLoadBalancingSectionPage
	
	Write-Verbose "$(Get-Date): `tProcessing Load balancing"
	
	$results = Get-RASLBSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Load balancing information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Load balancing information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Load balancing information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Load balancing information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No Load balancing information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Load balancing information was found"
		}
		If($Text)
		{
			Line 0 "No Load balancing information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Load balancing information was found"
		}
	}
	Else
	{
		OutputRASLBSettings $results
	}
}

Function OutputLoadBalancingSectionPage
{
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Load balancing"
	}
	If($Text)
	{
		Line 0 "Load balancing"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Load balancing"
	}
}

Function OutputRASLBSettings
{
	Param([object] $RASLBSettings)
	
	Write-Verbose "$(Get-Date): `t`tOutput Load balancing"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Load balancing"
	}
	If($Text)
	{
		Line 1 "Load balancing"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Load balancing"
	}
	
	Switch ($RASLBSettings.Method)
	{
		"ResourceBased"	{$RASLBSettingsMethod = "Resource Based"; Break}
		"RoundRobin"	{$RASLBSettingsMethod = "Round Robin"; Break}
		Default			{$RASLBSettingsMethod = "Unable to determine Load balancing method: $($RASLBSettings.Method)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Method"; Value = $RASLBSettingsMethod; }) > $Null
		$ScriptInformation.Add(@{Data = "Counters"; Value = ""; }) > $Null
		$ScriptInformation.Add(@{Data = "     User Sessions"; Value = $RASLBSettings.SessionsCounter; }) > $Null
		$ScriptInformation.Add(@{Data = "     Memory"; Value = $RASLBSettings.MemoryCounter; }) > $Null
		$ScriptInformation.Add(@{Data = "     CPU"; Value = $RASLBSettings.CPUCounter; }) > $Null
		$ScriptInformation.Add(@{Data = "Reconnect to disconnected sessions"; Value = $RASLBSettings.ReconnectDisconnect; }) > $Null
		$ScriptInformation.Add(@{Data = "Reconnect sessions using client's IP address only"; Value = $RASLBSettings.ReconnectUsingIPOnly; }) > $Null
		$ScriptInformation.Add(@{Data = "Limit each user to one session per desktop"; Value = $RASLBSettings.ReconnectUser; }) > $Null
		$ScriptInformation.Add(@{Data = "Disable Microsoft RD Connection Broker"; Value = $RASLBSettings.DisableRDSLB; }) > $Null
		$ScriptInformation.Add(@{Data = "Declare Agent dead if not responding for"; Value = "$($RASLBSettings.DeadTimeout) seconds"; }) > $Null
		$ScriptInformation.Add(@{Data = "Agent Refresh Time"; Value = "$($RASLBSettings.RefreshTimeout) seconds"; }) > $Null
		$ScriptInformation.Add(@{Data = "Enable CPU Load Balancer"; Value = $RASLBSettings.EnableCPULB.ToString(); }) > $Null
		$ScriptInformation.Add(@{Data = "Replicate settings"; Value = $RASLBSettings.Replicate; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 275;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Method`t`t`t : " $RASLBSettingsMethod
		Line 2 "Counters"
		Line 3 "User Sessions`t : " $RASLBSettings.SessionsCounter
		Line 3 "Memory`t`t : " $RASLBSettings.MemoryCounter
		Line 3 "CPU`t`t : " $RASLBSettings.CPUCounter
		Line 2 "Reconnect to disconnected sessions`t`t`t : " $RASLBSettings.ReconnectDisconnect
		Line 2 "Reconnect sessions using client's IP address only`t : " $RASLBSettings.ReconnectUsingIPOnly
		Line 2 "Limit each user to one session per desktop`t`t : " $RASLBSettings.ReconnectUser
		Line 2 "Disable Microsoft RD Connection Broker`t`t`t : " $RASLBSettings.DisableRDSLB
		Line 2 "Declare Agent dead if not responding for`t`t : " "$($RASLBSettings.DeadTimeout) seconds"
		Line 2 "Agent Refresh Time`t`t`t`t`t : " "$($RASLBSettings.RefreshTimeout) seconds"
		Line 2 "Enable CPU Load Balancer`t`t`t`t : " $RASLBSettings.EnableCPULB.ToString()
		Line 2 "Replicate settings`t`t`t`t`t : " $RASLBSettings.Replicate
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Method",($Script:htmlsb),$RASLBSettingsMethod,$htmlwhite)
		$rowdata += @(,("Counters",($Script:htmlsb),"",$htmlwhite))
		$rowdata += @(,("     User Sessions",($Script:htmlsb),$RASLBSettings.SessionsCounter.ToString(),$htmlwhite))
		$rowdata += @(,("     Memory",($Script:htmlsb),$RASLBSettings.MemoryCounter.ToString(),$htmlwhite))
		$rowdata += @(,("     CPU",($Script:htmlsb),$RASLBSettings.CPUCounter.ToString(),$htmlwhite))
		$rowdata += @(,("Reconnect to disconnected sessions",($Script:htmlsb),$RASLBSettings.ReconnectDisconnect.ToString(),$htmlwhite))
		$rowdata += @(,("Reconnect sessions using client's IP address only",($Script:htmlsb),$RASLBSettings.ReconnectUsingIPOnly.ToString(),$htmlwhite))
		$rowdata += @(,("Limit each user to one session per desktop",($Script:htmlsb),$RASLBSettings.ReconnectUser.ToString(),$htmlwhite))
		$rowdata += @(,("Disable Microsoft RD Connection Broker",($Script:htmlsb),$RASLBSettings.DisableRDSLB.ToString(),$htmlwhite))
		$rowdata += @(,("Declare Agent dead if not responding for",($Script:htmlsb),"$($RASLBSettings.DeadTimeout) seconds",$htmlwhite))
		$rowdata += @(,("Agent Refresh Time",($Script:htmlsb),"$($RASLBSettings.RefreshTimeout) seconds",$htmlwhite))
		$rowdata += @(,("Enable CPU Load Balancer",($Script:htmlsb),$RASLBSettings.EnableCPULB.ToString(),$htmlwhite))
		$rowdata += @(,("Replicate settings",($Script:htmlsb),$RASLBSettings.Replicate.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("300","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region process publishing
Function ListFolder($folder, $spaces) 
{
	#function provided 23-July-2020 by Ian Sant of Parallels
	#list published items in the order they appear in the console
	#modified 24-July-2020 by Webster to make it work with this script
	$prevId = 0

	Do
	{
		$item = $Script:AllItems | Where-Object {($_.ParentId -eq $folder) -and ($_.PreviousId -eq $prevId)}
		If ($Null -eq $item) {
			Return
		}

		#Write-Host $spaces $item.Name
		$Script:OrderedItems.Add($item) >$Null
		If ($item.Type -eq "Folder") {
			$newspaces = $spaces + "   "
			ListFolder $item.Id $newspaces
		}
		$prevId = $item.Id
	} While ($true) 
}

Function ProcessPublishing
{
	Write-Verbose "$(Get-Date): Processing Publishing"
	
	$Sites = Get-Site -EA 0 4> $Null

	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Sites for Publishing data`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Sites for Publishing data"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Sites for Publishing data"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Sites for Publishing data"
		}
	}
	ElseIf($? -and $Null -eq $Sites)
	{
		Write-Warning "
		`n`n
		`t`t
		No Sites were retrieved for Publishing data.`
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Sites were retrieved for Publishing data"
		}
		If($Text)
		{
			Line 0 "No Sites were retrieved for Publishing data"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Sites were retrieved for Publishing data"
		}
	}
	Else
	{
		ForEach($Site in $Sites)
		{
			OutputPublishingSectionPage $Site.Name
			
			Write-Verbose "$(Get-Date): `tProcessing Publishing for Site $($Site.Name)"
			
			$results = Get-PubItem -SiteId $Site.Id -EA 0 4>$Null
				
			If(!($?))
			{
				Write-Warning "
				`n`n
				`t`t
				Unable to retrieve Publishing information for Site $($Site.Name)
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Unable to retrieve Publishing information for Site $($Site.Name)"
				}
				If($Text)
				{
					Line 0 "Unable to retrieve Publishing information for Site $($Site.Name)"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Unable to retrieve Publishing information for Site $($Site.Name)"
				}
			}
			ElseIf($? -and $null -eq $results)
			{
				Write-Warning "
				`n`n
				`t`t
				No Publishing information was found for Site $($Site.Name)
				`n`n
				"
				If($MSWord -or $PDF)
				{
					WriteWordLine 1 0 "No Publishing information was found for Site $($Site.Name)"
				}
				If($Text)
				{
					Line 0 "No Publishing information was found for Site $($Site.Name)"
				}
				If($HTML)
				{
					WriteHTMLLine 1 0 "No Publishing information was found for Site $($Site.Name)"
				}
			}
			Else
			{
				$Script:OrderedItems = New-Object System.Collections.ArrayList
				$Script:AllItems = $results
				ListFolder 0 ""
				OutputPublishingSettings $Script:OrderedItems $Site.Id $Site.Name
			}
		}
	}
}

Function OutputPublishingSectionPage
{
	Param([string] $SiteName)
	
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Publishing for Site $($SiteName)"
	}
	If($Text)
	{
		Line 0 "Publishing for Site $($SiteName)"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Publishing for Site $($SiteName)"
	}
}

Function OutputPublishingSettings
{
	Param([object] $PubItems, [uint32] $SiteId, [string] $xSiteName)
	
	Write-Verbose "$(Get-Date): `t`tOutput Publishing for Site $xSiteName"
	
	<#
	Folder
	PCDesktop
	RDSDesktop
	VDIDesktop
	PCApp
	RDSApp
	VDIApp
	#>

	#Get the published items default settings
	Write-Verbose "$(Get-Date): `t`t`tRetrieve Publishing Default Site Settings for Site $xSiteName"
	$results = Get-PubDefaultSettings -SiteId $SiteId -EA 0 4>$Null
	
	If(!$?)
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Publishing Default Site Settings for Site $xSiteName, using built-in defaults
		`n`n
		"
		<#
		StartPath                     : RAS Remote Desktops & Applications\%Groups%
		CreateShortcutOnDesktop       : False
		CreateShortcutInStartFolder   : True
		CreateShortcutInStartUpFolder : False
		ReplicateShortcutSettings     : False
		ReplicateDisplaySettings      : False
		WaitForPrinters               : False
		StartMaximized                : True
		WaitForPrintersTimeout        : 20
		ColorDepth                    : ClientSpecified
		DisableSessionSharing         : False
		OneInstancePerUser            : False
		ConCurrentLicenses            : 0
		LicenseLimitNotify            : WarnUserAndNoStart
		ReplicateLicenseSettings      : False
		#>
		
		#Shortcuts tab
		$DefaultCreateShortcutOnDesktop       = "False"
		$DefaultCreateShortcutInStartFolder   = "True"
		$DefaultStartPath                     = "RAS Remote Desktops & Applications\%Groups%"
		$DefaultCreateShortcutInStartUpFolder = "False"
		$DefaultReplicateShortcutSettings     = "False"
		
		#License tab
		$DefaultDisableSessionSharing         = "False"
		$DefaultOneInstancePerUser            = "False"
		$DefaultConCurrentLicenses            = "Unlimited"
		$DefaultLicenseLimitNotify            = "Warn user and do not start"
		$DefaultReplicateLicenseSettings      = "False"

		#Display tab
		$DefaultWaitForPrinters               = "False"
		$DefaultWaitForPrintersTimeout        = "20"
		$DefaultColorDepth                    = "Client Specified"
		$DefaultStartMaximized                = "True"
		$DefaultReplicateDisplaySettings      = "False"

		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Publishing Default Site Settings for Site $xSiteName, using built-in defaults"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Publishing Default Site Settings for Site $xSiteName, using built-in defaults"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Publishing Default Site Settings for Site $xSiteName, using built-in defaults"
		}
	}
	Else
	{
		#Shortcuts tab
		$DefaultCreateShortcutOnDesktop       = $results.CreateShortcutOnDesktop.ToString()      
		$DefaultCreateShortcutInStartFolder   = $results.CreateShortcutInStartFolder.ToString()
		$DefaultStartPath                     = $results.StartPath                    
		$DefaultCreateShortcutInStartUpFolder = $results.CreateShortcutInStartUpFolder.ToString()
		$DefaultReplicateShortcutSettings     = $results.ReplicateShortcutSettings.ToString()

		#License tab
		$DefaultDisableSessionSharing         = $results.DisableSessionSharing.ToString()    
		$DefaultOneInstancePerUser            = $results.OneInstancePerUser.ToString()
		If($results.ConCurrentLicenses -eq 0)
		{
			$DefaultConCurrentLicenses        = "Unlimited"
		}
		Else
		{
			$DefaultConCurrentLicenses        = $results.ConCurrentLicenses.ToString()    
		}
		 
		Switch ($results.LicenseLimitNotify)
		{
			"WarnUserAndNoStart"		{$DefaultLicenseLimitNotify = "Warn user and do not start"; Break}
			"WarnUserAndStart"			{$DefaultLicenseLimitNotify = "Warn user and start"; Break}
			"NotifyAdminAndStart"		{$DefaultLicenseLimitNotify = "Notify administrator and start"; Break}
			"NotifyUserAdminAndStart"	{$DefaultLicenseLimitNotify = "Notify user, administrator and start"; Break}
			"NotifyUserAdminAndNoStart"	{$DefaultLicenseLimitNotify = "Notify user, administrator and do not start"; Break}
			Default						{$DefaultLicenseLimitNotify = "Unable to determine If limit is eceeded: $($results.LicenseLimitNotify)"; Break}
		}
		
		$DefaultReplicateLicenseSettings      = $results.ReplicateLicenseSettings.ToString()

		#Display tab
		$DefaultWaitForPrinters               = $results.WaitForPrinters.ToString()
		$DefaultWaitForPrintersTimeout        = $results.WaitForPrintersTimeout.ToString()
		Switch ($results.ColorDepth)
		{
			"Colors8Bit"		{$DefaultColorDepth = "256 Colors"; Break}
			"Colors15Bit"		{$DefaultColorDepth = "High Color (15 bit)"; Break}
			"Colors16Bit"		{$DefaultColorDepth = "High Color (16 bit)"; Break}
			"Colors24Bit"		{$DefaultColorDepth = "True Color (24 bit)"; Break}
			"Colors32Bit"		{$DefaultColorDepth = "Highest Quality (32 bit)"; Break}
			"ClientSpecified"	{$DefaultColorDepth = "Client Specified"; Break}
			Default				{$DefaultColorDepth = "Unable to determine Color Depth: $($PubItem.ColorDepth)"; Break}
		}
		$DefaultStartMaximized                = $results.StartMaximized.ToString()   
		$DefaultReplicateDisplaySettings      = $results.ReplicateDisplaySettings.ToString()
	}
	
	ForEach($PubItem in $PubItems)
	{
		Write-Verbose "$(Get-Date): `t`t`t`tOutput $($PubItem.Name)"

		If(ValidObject $PubItem WinType)
		{
			Switch ($PubItem.WinType)
			{
				"Normal"	{$WinType = "Normal Window"; Break}
				"Maximized"	{$WinType = "Maximized"; Break}
				"Minimized"	{$WinType = "Minimized"; Break}
				Default		{$WinType = "Unable to determine window Run type: $($PubItem.WinType)"; Break}
			}
		}

		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $PubItem.Name
		}
		If($Text)
		{
			Line 1 $PubItem.Name
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 $PubItem.Name
		}
	
		If($PubItem.Type -eq "Folder")
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
			}
			If($Text)
			{
				Line 2 "Information"
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Folder"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Folder"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Folder Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Use for administrative purposes"; Value = $PubItem.AdminOnly.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				OutputPubItemFilters $PubItem "MSWordPDF"
			}
			If($Text)
			{
				Line 3 "Folder:`t`t`t`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 Sites
				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "This published item will be available from the following Sites: " $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				
				Line 2 "Folder"
				Line 3 "Folder Name:`t`t`t`t`t`t" $PubItem.Name
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Use for administrative purposest:`t`t`t" $PubItem.AdminOnly.ToString()
				Line 0 ""
				
				OutputPubItemFilters $PubItem "Text"
			}
			If($HTML)
			{
				$rowdata = @()

				$columnHeaders = @("Folder",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Folder"
				$rowdata = @()

				$columnHeaders = @("Folder",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Use for administrative purposes",($Script:htmlsb),$PubItem.AdminOnly.ToString(),$htmlwhite))

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				OutputPubItemFilters $PubItem "HTML"
			}
		}
		ElseIf($PubItem.Type -eq "PCDesktop")
		{
			$DesktopSize = "Unable to determine"
			If($PubItem.UseAvailableArea)
			{
				$DesktopSize = "Use available area"
			}
			ElseIf($PubItem.Width -eq 0 -and $PubItem.Height -eq 0)
			{
				$DesktopSize = "Full Screen"
			}
			Else
			{
				$DesktopSize = "$($PubItem.Width.ToString())x$($PubItem.Height.ToString())"
			}
			
			If($PubItem.AllowMultiMonitor -eq "UseClientSettings")
			{
				$AllowMultiMonitor = "Use Client Settings"
			}
			Else
			{
				$AllowMultiMonitor = $PubItem.AllowMultiMonitor.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
				
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Remote PC Desktop"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Desktop Size"; Value = $DesktopSize; }) > $Null
				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($DefaultStartPath)'"; }) > $Null
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($PubItem.StartPath)'"; }) > $Null
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Remote PC Desktop"
				WriteWordLine 4 0 "Remote PC Desktop"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = $PubItem.StartOnLogon.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 4 0 "Properties"
				$ScriptInformation = New-Object System.Collections.ArrayList
				#$ScriptInformation.Add(@{Data = "Select Remote PC"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "Desktop Size"; Value = $DesktopSize; }) > $Null
				$ScriptInformation.Add(@{Data = "Multi-Monitor"; Value = $AllowMultiMonitor; }) > $Null
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				OutputPubItemFilters $PubItem "MSWordPDF"

				OutputPubItemShortcuts $PubItem "MSWordPDF" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($Text)
			{
				Line 2 "Information"
				Line 3 "Remote PC Desktop:`t`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Desktop Size:`t`t`t`t`t`t" $DesktopSize

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($DefaultStartPath)'"
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($PubItem.StartPath)'"
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 "Sites"
				Line 3 "This published item will be available from the following Sites"
				ForEach($Site in $PubItem.PublishToSite)
				{
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					Line 10 $SiteName
				}
				Line 0 ""

				Line 2 "Remote PC Desktop"
				Line 3 "Remote PC Desktop"
				Line 4 "Name:`t`t`t`t`t`t" $PubItem.Name
				Line 4 "Description:`t`t`t`t`t" $PubItem.Description
				Line 4 "Start automatically when user logs on:`t`t" $PubItem.StartOnLogon.ToString()
				Line 0 ""
				Line 3 "Properties"
				Line 4 "Desktop Size:`t`t`t`t`t" $DesktopSize
				Line 4 "Multi-Monitor:`t`t`t`t`t" $AllowMultiMonitor
				Line 0 ""

				OutputPubItemFilters $PubItem "Text"
				
				OutputPubItemShortcuts $PubItem "Text" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
				$rowdata = @()

				$columnHeaders = @("Remote PC Desktop",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Desktop Size",($Script:htmlsb),$DesktopSize,$htmlwhite))
				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($DefaultStartPath)'",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($PubItem.StartPath)'",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Sites"
				$rowdata = @()

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$columnHeaders = @("This published item will be available from the following Sites",($Script:htmlsb),$SiteName,$htmlwhite)
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Remote PC Desktop"
				#WriteHTMLLine 4 0 "Remote PC Desktop"
				$rowdata = @()
				$columnHeaders = @("Name",($Script:htmlsb),$PubItem.Name,$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),$PubItem.StartOnLogon.ToString(),$htmlwhite))

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 4 0 "Properties"
				$rowdata = @()
				$columnHeaders = @("Select Remote PC",($Script:htmlsb),"",$htmlwhite)
				$rowdata += @(,("Desktop Size",($Script:htmlsb),$DesktopSize,$htmlwhite))
				$rowdata += @(,("Multi-Monitor",($Script:htmlsb),$AllowMultiMonitor,$htmlwhite))
				
				$msg = "Remote PC Desktop"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				OutputPubItemFilters $PubItem "HTML"

				OutputPubItemShortcuts $PubItem "HTML" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
		}
		ElseIf($PubItem.Type -eq "RDSApp")
		{
			Switch ($PubItem.ConCurrentLicenses)
			{
				0 		{$ConCurrentLicenses = "Unlimited"; Break}
				Default	{$ConCurrentLicenses = $PubItem.ConCurrentLicenses.ToString(); Break} 
			}
			
			Switch ($PubItem.DisableSessionSharing)
			{
				$False	{$SessionSharing = "Enabled"; Break}
				$True	{$SessionSharing = "Disabled"; Break}
				Default	{$SessionSharing = "Unable to determine Session Sharing state: $($PubItem.DisableSessionSharing)"; Break}
			}
			
			Switch ($PubItem.PublishFrom)
			{
				"All"		{$PublishedFrom = "All Servers in Site"; Break}
				"Group"		{$PublishedFrom = "Server Groups:"; Break}
				"Server"	{$PublishedFrom = "Individual Servers:"; Break}
				Default		{$PublishedFrom = "Unable to determine Published From: $($PubItem.PublishFrom)"; Break}
			}
			
			Switch ($PubItem.LicenseLimitNotify)
			{
				"WarnUserAndNoStart"		{$LicenseLimitNotify ="Warn user and do not start"; Break}
				"WarnUserAndStart"			{$LicenseLimitNotify ="Warn user and start"; Break}
				"NotifyAdminAndStart"		{$LicenseLimitNotify ="Notify administrator and start"; Break}
				"NotifyUserAdminAndStart"	{$LicenseLimitNotify ="Notify user, administrator and start"; Break}
				"NotifyUserAdminAndNoStart"	{$LicenseLimitNotify ="Notify user, administrator and do not start"; Break}
				Default	{$LicenseLimitNotify ="Unable to determine If limit is exceeded: $($PubItem.LicenseLimitNotify)"; Break}
			}
			
			Switch ($PubItem.ColorDepth)
			{
				"Colors8Bit"		{$ColorDepth = "256 Colors"; Break}
				"Colors15Bit"		{$ColorDepth = "High Color (15 bit)"; Break}
				"Colors16Bit"		{$ColorDepth = "High Color (16 bit)"; Break}
				"Colors24Bit"		{$ColorDepth = "True Color (24 bit)"; Break}
				"Colors32Bit"		{$ColorDepth = "Highest Quality (32 bit)"; Break}
				"ClientSpecified"	{$ColorDepth = "Client Specified"; Break}
				Default				{$ColorDepth = "Unable to determine Color Depth: $($PubItem.ColorDepth)"; Break}
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Application"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Target"; Value = $PubItem.Target; }) > $Null
				$ScriptInformation.Add(@{Data = "Start In"; Value = $PubItem.StartIn; }) > $Null
				
				If($PubItem.StartOnLogon)
				{
					$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = ""; }) > $Null
				}

				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					$ScriptInformation.Add(@{Data = "Parameters"; Value = $PubItem.Parameters; }) > $Null
				}
				
				If($PubItem.EnableFileExtensions)
				{
					$ScriptInformation.Add(@{Data = "Associate the following file extensions"; Value = ""; }) > $Null
					ForEach($Item in $PubItem.FileExtensions)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				
				If($PubItem.InheritLicenseDefaultSettings)
				{
					If($DefaultOneInstancePerUser)
					{
						$ScriptInformation.Add(@{Data = "Allow users to start only one instance of this application"; Value = ""; }) > $Null
					}
					$ScriptInformation.Add(@{Data = "Concurrent licenses"; Value = $DefaultConCurrentLicenses; }) > $Null
					$ScriptInformation.Add(@{Data = "If limit is exceeded"; Value = $DefaultLicenseLimitNotify; }) > $Null
					$ScriptInformation.Add(@{Data = "Session Sharing"; Value = $DefaultDisableSessionSharing ; }) > $Null
				}
				Else
				{
					If($PubItem.OneInstancePerUser)
					{
						$ScriptInformation.Add(@{Data = "Allow users to start only one instance of this application"; Value = ""; }) > $Null
					}
					$ScriptInformation.Add(@{Data = "Concurrent licenses"; Value = $ConCurrentLicenses; }) > $Null
					$ScriptInformation.Add(@{Data = "If limit is exceeded"; Value = $LicenseLimitNotify; }) > $Null
					$ScriptInformation.Add(@{Data = "Session Sharing"; Value = $SessionSharing; }) > $Null
				}

				$ScriptInformation.Add(@{Data = "Settings for Site $xSiteName"; Value = ""; }) > $Null
				
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "Published from"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "Published from"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Published from"; Value = "All Servers in Site"; }) > $Null
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($DefaultStartPath)'"; }) > $Null
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($PubItem.StartPath)'"; }) > $Null
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Publish from"
				$ScriptInformation = New-Object System.Collections.ArrayList
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "$PublishedFrom"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "$PublishedFrom"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "$PublishedFrom"; Value = ""; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Application"
				WriteWordLine 4 0 "Application"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Run"; Value = $WinType; }) > $Null
				$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = $PubItem.StartOnLogon.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				WriteWordLine 4 0 "Server settings"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Server(s)"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "Target"; Value = $PubItem.Target; }) > $Null
				$ScriptInformation.Add(@{Data = "Start in"; Value = $PubItem.StartIn; }) > $Null
				$ScriptInformation.Add(@{Data = "Parameters"; Value = $PubItem.Parameters; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				OutputPubItemFilters $PubItem "MSWordPDF"
				
				OutputPubItemShortcuts $PubItem "MSWordPDF" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings

				WriteWordLine 3 0 "File extensions"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Associate File Extensions"; Value = $PubItem.EnableFileExtensions.ToString(); }) > $Null

				$cnt = -1
				ForEach($Item in $PubItem.FileExtensions)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Extension"; Value = $Item; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				If($PubItem.ReplicateFileExtensionSettings)
				{
					WriteWordLine 0 0 "Settings are replicated to all Sites"
				}
				Else
				{
					WriteWordLine 0 0 "Settings are not replicated to all Sites"
				}
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "License"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $PubItem.InheritLicenseDefaultSettings.ToString(); }) > $Null

				If($PubItem.InheritLicenseDefaultSettings)
				{
					$ScriptInformation.Add(@{Data = "Session Sharing"; Value = $DefaultDisableSessionSharing ; }) > $Null
					If($DefaultOneInstancePerUser)
					{
						$ScriptInformation.Add(@{Data = "Allow users to start only one instance of this application"; Value = "True"; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = "Allow users to start only one instance of this application"; Value = "False"; }) > $Null
					}
					$ScriptInformation.Add(@{Data = "Concurrent licenses"; Value = $DefaultConCurrentLicenses; }) > $Null
					$ScriptInformation.Add(@{Data = "If limit is exceeded"; Value = $DefaultLicenseLimitNotify; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Session Sharing"; Value = $SessionSharing; }) > $Null
					If($PubItem.OneInstancePerUser)
					{
						$ScriptInformation.Add(@{Data = "Allow users to start only one instance of this application"; Value = "True"; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = "Allow users to start only one instance of this application"; Value = "False"; }) > $Null
					}
					$ScriptInformation.Add(@{Data = "Concurrent licenses"; Value = $ConCurrentLicenses; }) > $Null
					$ScriptInformation.Add(@{Data = "If limit is exceeded"; Value = $LicenseLimitNotify; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				If($PubItem.InheritLicenseDefaultSettings)
				{
					If($DefaultReplicateLicenseSettings)
					{
						WriteWordLine 0 0 "Settings are replicated to all Sites"
					}
					Else
					{
						WriteWordLine 0 0 "Settings are not replicated to all Sites"
					}
				}
				Else
				{
					If($PubItem.ReplicateLicenseSettings)
					{
						WriteWordLine 0 0 "Settings are replicated to all Sites"
					}
					Else
					{
						WriteWordLine 0 0 "Settings are not replicated to all Sites"
					}
				}
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Display"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Inherit default settings"; Value = $PubItem.InheritDisplayDefaultSettings.ToString(); }) > $Null

				If($PubItem.InheritDisplayDefaultSettings)
				{
					$ScriptInformation.Add(@{Data = "Wait until all RAS Universal Printers are redirected before showing the application"; Value = $DefaultWaitForPrinters; }) > $Null
					$ScriptInformation.Add(@{Data = "Maximum time to wait is"; Value = "$($DefaultWaitForPrintersTimeout) seconds"; }) > $Null
					$ScriptInformation.Add(@{Data = "Color Depth"; Value = $DefaultColorDepth; }) > $Null
					$ScriptInformation.Add(@{Data = "Start the application as maximized when using mobile clients"; Value = $DefaultStartMaximized; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Wait until all RAS Universal Printers are redirected before showing the application"; Value = $PubItem.WaitForPrinters.ToString(); }) > $Null
					$ScriptInformation.Add(@{Data = "Maximum time to wait is"; Value = "$($PubItem.WaitForPrintersTimeout.ToString()) seconds"; }) > $Null
					$ScriptInformation.Add(@{Data = "Color Depth"; Value = $ColorDepth; }) > $Null
					$ScriptInformation.Add(@{Data = "Start the application as maximized when using mobile clients"; Value = $PubItem.StartMaximized.ToString(); }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null

				If($PubItem.InheritDisplayDefaultSettings)
				{
					If($DefaultReplicateDisplaySettings)
					{
						WriteWordLine 0 0 "Settings are replicated to all Sites"
					}
					Else
					{
						WriteWordLine 0 0 "Settings are not replicated to all Sites"
					}
				}
				Else
				{
					If($PubItem.ReplicateDisplaySettings)
					{
						WriteWordLine 0 0 "Settings are replicated to all Sites"
					}
					Else
					{
						WriteWordLine 0 0 "Settings are not replicated to all Sites"
					}
				}
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Information"
				Line 3 "Application:`t`t`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Target:`t`t`t`t`t`t`t" $PubItem.Target
				Line 3 "Start In:`t`t`t`t`t`t" $PubItem.StartIn
				
				If($PubItem.StartOnLogon)
				{
					Line 3 "Start automatically when user logs on" ""
				}
				
				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					Line 3 "Parameters:`t`t`t`t`t`t" $PubItem.Parameters
				}
				
				If($PubItem.EnableFileExtensions)
				{
					Line 3 "Associate the following file extensions"
					ForEach($Item in $PubItem.FileExtensions)
					{
						Line 10 $Item
					}
				}
				
				If($PubItem.InheritLicenseDefaultSettings)
				{
					If($DefaultOneInstancePerUser)
					{
						Line 3 "Allow users to start only one instance of this application" ""
					}
					Line 3 "Concurrent licenses: `t`t`t`t`t" $DefaultConCurrentLicenses
					Line 3 "If limit is exceeded: `t`t`t`t`t" $DefaultLicenseLimitNotify
					Line 3 "Session Sharing: `t`t`t`t`t" $DefaultDisableSessionSharing 
				}
				Else
				{
					If($PubItem.OneInstancePerUser)
					{
						Line 3 "Allow users to start only one instance of this application" ""
					}
					Line 3 "Concurrent licenses: `t`t`t`t`t" $ConCurrentLicenses
					Line 3 "If limit is exceeded: `t`t`t`t`t" $LicenseLimitNotify
					Line 3 "Session Sharing: `t`t`t`t`t" $SessionSharing
				}

				Line 3 "Settings for Site $xSiteName"
				
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						If($cnt -eq 0)
						{
							Line 3 "Published from:`t`t`t`t`t`t" ItemName
						}
						Else
						{
							Line 10 $ItemName
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						If($cnt -eq 0)
						{
							Line 3 "Published from:`t`t`t`t`t`t" ItemName
						}
						Else
						{
							Line 10 $ItemName
						}
					}
				}
				Else
				{
					Line 3 "Published from:`t`t`t`t`t`t" "All Servers in Site"
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($DefaultStartPath)'"
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($PubItem.StartPath)'"
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 "Sites"
				Line 3 "This published item will be available from the following Sites"
				ForEach($Site in $PubItem.PublishToSite)
				{
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					Line 10 $SiteName
				}
				Line 0 ""

				Line 2 "Publish from"
				Line 3 $PublishedFrom
				If($PubItem.PublishFrom -eq "Server")
				{
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						Line 6 $ItemName
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						Line 5 $ItemName
					}
				}
				Line 0 ""

				Line 2 "Application"
				Line 3 "Application"
				Line 4 "Name:`t`t`t`t`t`t" $PubItem.Name
				Line 4 "Description:`t`t`t`t`t" $PubItem.Description
				Line 4 "Run:`t`t`t`t`t`t" $WinType
				Line 4 "Start automatically when user logs on:`t`t" $PubItem.StartOnLogon.ToString()
				Line 0 ""
				
				Line 3 "Server settings"
				#Line 4 "Server(s):`t`t`t`t`t" ""
				Line 4 "Target:`t`t`t`t`t`t" $PubItem.Target
				Line 4 "Start in:`t`t`t`t`t" $PubItem.StartIn
				Line 4 "Parameters:`t`t`t`t`t" $PubItem.Parameters
				Line 0 ""
				
				OutputPubItemFilters $PubItem "Text"
				
				OutputPubItemShortcuts $PubItem "Text" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings

				Line 2 "File extensions"
				Line 3 "Associate File Extensions:`t`t`t`t" $PubItem.EnableFileExtensions.ToString()

				$cnt = -1
				ForEach($Item in $PubItem.FileExtensions)
				{
					$cnt++
					If($cnt -eq 0 )
					{
						Line 8 "Extension:`t" $Item
					}
					Else
					{
						Line 10 $Item
					}
				}
				If($PubItem.ReplicateFileExtensionSettings)
				{
					Line 3 "Settings are replicated to all Sites"
				}
				Else
				{
					Line 3 "Settings are not replicated to all Sites"
				}
				Line 0 ""

				Line 2 "Display"
				Line 3 "Inherit default settings:`t`t`t`t" $PubItem.InheritDisplayDefaultSettings.ToString()

				If($PubItem.InheritDisplayDefaultSettings)
				{
					Line 3 "Wait until all RAS Universal Printers are redirected "
					Line 3 "before showing the application:`t`t`t`t" $DefaultWaitForPrinters
					Line 3 "Maximum time to wait is:`t`t`t`t" "$($DefaultWaitForPrintersTimeout) seconds"
					Line 3 "Color Depth:`t`t`t`t`t`t" $DefaultColorDepth
					Line 3 "Start the application as maximized "
					Line 3 "when using mobile clients:`t`t`t`t" $DefaultStartMaximized
				}
				Else
				{
					Line 3 "Wait until all RAS Universal Printers are redirected "
					Line 3 "before showing the application:`t`t`t`t" $PubItem.WaitForPrinters.ToString()
					Line 3 "Maximum time to wait is:`t`t`t`t" "$($PubItem.WaitForPrintersTimeout.ToString()) seconds"
					Line 3 "Color Depth:`t`t`t`t`t`t" $ColorDepth
					Line 3 "Start the application as maximized "
					Line 3 "when using mobile clients:`t`t`t`t" $PubItem.StartMaximized.ToString()
				}

				If($PubItem.ReplicateDisplaySettings)
				{
					Line 3 "Settings are replicated to all Sites"
				}
				Else
				{
					Line 3 "Settings are not replicated to all Sites"
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
				$rowdata = @()

				$columnHeaders = @("Application",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Target",($Script:htmlsb),$PubItem.Target,$htmlwhite))
				$rowdata += @(,("Start In",($Script:htmlsb),$PubItem.StartIn,$htmlwhite))
				
				If($PubItem.StartOnLogon)
				{
					$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),"",$htmlwhite))
				}
				
				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					$rowdata += @(,("Parameters",($Script:htmlsb),$PubItem.Parameters,$htmlwhite))
				}
				
				If($PubItem.EnableFileExtensions)
				{
					$rowdata += @(,("Associate the following file extensions",($Script:htmlsb),"",$htmlwhite))
					ForEach($Item in $PubItem.FileExtensions)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				
				If($PubItem.InheritLicenseDefaultSettings)
				{
					If($DefaultOneInstancePerUser)
					{
						$rowdata += @(,("Allow users to start only one instance of this application",($Script:htmlsb),"",$htmlwhite))
					}
					$rowdata += @(,("Concurrent licenses",($Script:htmlsb),$DefaultConCurrentLicenses,$htmlwhite))
					$rowdata += @(,("If limit is exceeded",($Script:htmlsb),$DefaultLicenseLimitNotify,$htmlwhite))
					$rowdata += @(,("Session Sharing",($Script:htmlsb),$DefaultDisableSessionSharing,$htmlwhite))
				}
				Else
				{
					If($PubItem.OneInstancePerUser)
					{
						$rowdata += @(,("Allow users to start only one instance of this application",($Script:htmlsb),"",$htmlwhite))
					}
					$rowdata += @(,("Concurrent licenses",($Script:htmlsb),$ConCurrentLicenses,$htmlwhite))
					$rowdata += @(,("If limit is exceeded",($Script:htmlsb),$LicenseLimitNotify,$htmlwhite))
					$rowdata += @(,("Session Sharing",($Script:htmlsb),$SessionSharing,$htmlwhite))
				}

				$rowdata += @(,("Settings for Site $xSiteName",($Script:htmlsb),"",$htmlwhite))
				
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						If($cnt -eq 0)
						{
							$rowdata += @(,("Published from",($Script:htmlsb),$ItemName,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						If($cnt -eq 0)
						{
							$rowdata += @(,("Published from",($Script:htmlsb),$ItemName,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				Else
				{
					$rowdata += @(,("Published from",($Script:htmlsb),"All Servers in Site",$htmlwhite))
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($DefaultStartPath)'",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($PubItem.StartPath)'",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}
			
				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Sites"
				$rowdata = @()

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$columnHeaders = @("This published item will be available from the following Sites",($Script:htmlsb),$SiteName,$htmlwhite)
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Publish from"
				$rowdata = @()
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						
						If($cnt -eq 0)
						{
							$columnHeaders = @("$PublishedFrom",($Script:htmlsb),$ItemName,$htmlwhite)
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						
						If($cnt -eq 0)
						{
							$columnHeaders = @("$PublishedFrom",($Script:htmlsb),$ItemName,$htmlwhite)
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				Else
				{
					$columnHeaders = @("$PublishedFrom",($Script:htmlsb),"",$htmlwhite)
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Application"
				$rowdata = @()

				$columnHeaders = @("Name",($Script:htmlsb),$PubItem.Name,$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Run",($Script:htmlsb),$WinType,$htmlwhite))
				$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),$PubItem.StartOnLogon.ToString(),$htmlwhite))

				$msg = "Application"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
				
				$rowdata = @()
				$columnHeaders = @("Desktop size",($Script:htmlsb),$DesktopSize,$htmlwhite)
				$rowdata += @(,("Multi-Monitor",($Script:htmlsb),$AllowMultiMonitor,$htmlwhite))

				$msg = "Server settings"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				OutputPubItemFilters $PubItem "HTML"
				
				OutputPubItemShortcuts $PubItem "HTML" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings

				WriteHTMLLine 3 0 "File extensions"
				$rowdata = @()
				$columnHeaders = @("Associate File Extensions",($Script:htmlsb),$PubItem.EnableFileExtensions.ToString(),$htmlwhite)
				
				$cnt = -1
				ForEach($Item in $PubItem.FileExtensions)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Extension",($Script:htmlsb),$Item,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				If($PubItem.ReplicateFileExtensionSettings)
				{
					WriteHTMLLine 0 0 "Settings are replicated to all Sites"
				}
				Else
				{
					WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
				}
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Display"
				$rowdata = @()
				$columnHeaders = @("Inherit default settings",($Script:htmlsb),$PubItem.InheritDisplayDefaultSettings.ToString(),$htmlwhite)

				If($PubItem.InheritDisplayDefaultSettings)
				{
					$rowdata += @(,("Wait until all RAS Universal Printers are redirected before showing the application",($Script:htmlsb),$DefaultWaitForPrinters,$htmlwhite))
					$rowdata += @(,("Maximum time to wait is",($Script:htmlsb),"$($DefaultWaitForPrintersTimeout) seconds",$htmlwhite))
					$rowdata += @(,("Color Depth",($Script:htmlsb),$DefaultColorDepth,$htmlwhite))
					$rowdata += @(,("Start the application as maximized when using mobile clients",($Script:htmlsb),$DefaultStartMaximized,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,("Wait until all RAS Universal Printers are redirected before showing the application",($Script:htmlsb),$PubItem.WaitForPrinters.ToString(),$htmlwhite))
					$rowdata += @(,("Maximum time to wait is",($Script:htmlsb),"$($PubItem.WaitForPrintersTimeout.ToString()) seconds",$htmlwhite))
					$rowdata += @(,("Color Depth",($Script:htmlsb),$ColorDepth,$htmlwhite))
					$rowdata += @(,("Start the application as maximized when using mobile clients",($Script:htmlsb),$PubItem.StartMaximized.ToString(),$htmlwhite))
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				If($PubItem.ReplicateDisplaySettings)
				{
					WriteHTMLLine 0 0 "Settings are replicated to all Sites"
				}
				Else
				{
					WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
				}
				WriteHTMLLine 0 0 ""
			}
		}
		ElseIf($PubItem.Type -eq "PCApp")
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Remote PC Application"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				#$ScriptInformation.Add(@{Data = "Last modification by"; Value = ""; }) > $Null
				#$ScriptInformation.Add(@{Data = "Modified on"; Value = ""; }) > $Null
				#$ScriptInformation.Add(@{Data = "Created by"; Value = ""; }) > $Null
				#$ScriptInformation.Add(@{Data = "Created on"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "Target"; Value = $PubItem.Target; }) > $Null
				$ScriptInformation.Add(@{Data = "Start In"; Value = $PubItem.StartIn; }) > $Null
				
				If($PubItem.StartOnLogon)
				{
					$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = ""; }) > $Null
				}

				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					$ScriptInformation.Add(@{Data = "Parameters"; Value = $PubItem.Parameters; }) > $Null
				}
				
				$ScriptInformation.Add(@{Data = "Settings for Site $xSiteName"; Value = ""; }) > $Null
				
				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($DefaultStartPath)'"; }) > $Null
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($PubItem.StartPath)'"; }) > $Null
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Remote PC Application"
				WriteWordLine 4 0 "Application"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Run"; Value = $WinType; }) > $Null
				$ScriptInformation.Add(@{Data = "Target"; Value = $PubItem.Target; }) > $Null
				$ScriptInformation.Add(@{Data = "Start In"; Value = $PubItem.StartIn; }) > $Null
				$ScriptInformation.Add(@{Data = "Parameters"; Value = $PubItem.Parameters; }) > $Null
				$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = $PubItem.StartOnLogon.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				OutputPubItemFilters $PubItem "MSWordPDF"
				
				OutputPubItemShortcuts $PubItem "MSWordPDF" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($Text)
			{
				Line 2 "Information"
				Line 3 "Remote PC Application:`t`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Target:`t`t`t`t`t`t`t" $PubItem.Target
				Line 3 "Start In:`t`t`t`t`t`t" $PubItem.StartIn
				
				If($PubItem.StartOnLogon)
				{
					Line 3 "Start automatically when user logs on" ""
				}
				
				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					Line 3 "Parameters:`t`t`t`t`t`t" $PubItem.Parameters
				}
				
				Line 3 "Settings for Site $xSiteName"

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($DefaultStartPath)'"
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($PubItem.StartPath)'"
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 "Sites"
				Line 3 "This published item will be available from the following Sites"
				ForEach($Site in $PubItem.PublishToSite)
				{
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					Line 10 $SiteName
				}
				Line 0 ""

				Line 2 "Remote PC Application"
				Line 3 "Application"
				Line 4 "Name:`t`t`t`t`t`t" $PubItem.Name
				Line 4 "Description:`t`t`t`t`t" $PubItem.Description
				Line 4 "Run:`t`t`t`t`t`t" $WinType
				Line 4 "Target:`t`t`t`t`t`t" $PubItem.Target
				Line 4 "Start In:`t`t`t`t`t" $PubItem.StartIn
				Line 4 "Parameters:`t`t`t`t`t" $PubItem.Parameters
				Line 4 "Start automatically when user logs on:`t`t" $PubItem.StartOnLogon.ToString()
				Line 0 ""
				
				OutputPubItemFilters $PubItem "Text"
				
				OutputPubItemShortcuts $PubItem "Text" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
				$rowdata = @()

				$columnHeaders = @("Remote PC Application",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Target",($Script:htmlsb),$PubItem.Target,$htmlwhite))
				$rowdata += @(,("Start In",($Script:htmlsb),$PubItem.StartIn,$htmlwhite))
				
				If($PubItem.StartOnLogon)
				{
					$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),"",$htmlwhite))
				}
				
				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					$rowdata += @(,("Parameters",($Script:htmlsb),$PubItem.Parameters,$htmlwhite))
				}
				
				$rowdata += @(,("Settings for Site $xSiteName",($Script:htmlsb),"",$htmlwhite))
				
				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($DefaultStartPath)'",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($PubItem.StartPath)'",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}
			
				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Sites"
				$rowdata = @()

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$columnHeaders = @("This published item will be available from the following Sites",($Script:htmlsb),$SiteName,$htmlwhite)
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Remote PC Application"
				$rowdata = @()

				$columnHeaders = @("Name",($Script:htmlsb),$PubItem.Name,$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Run",($Script:htmlsb),$WinType,$htmlwhite))
				$rowdata += @(,("Target",($Script:htmlsb),$PubItem.Target,$htmlwhite))
				$rowdata += @(,("Start In",($Script:htmlsb),$PubItem.StartIn,$htmlwhite))
				$rowdata += @(,("Parameters",($Script:htmlsb),$PubItem.Parameters,$htmlwhite))
				$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),$PubItem.StartOnLogon.ToString(),$htmlwhite))

				$msg = "Application"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
				
				OutputPubItemFilters $PubItem "HTML"
				
				OutputPubItemShortcuts $PubItem "HTML" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
		}
		ElseIf($PubItem.Type -eq "VDIApp")
		{
			Switch($PubItem.ConnectTo)
			{
				"AnyGuest"				{$ConnectTo = "Any Guest VM"; Break}
				"SpecificRASTemplate"	{$ConnectTo = "Specific Template ($($PubItem.SelectedGuests[0].VDIGuestName))"; Break
										}
				Default					{$ConnectTo = "Unable to determine Connect To: $($PubItem.ConnectTo)"; Break}
			}
			
			$results = Get-VDIPool -Id $PubItem.VDIPoolId -EA 0 4>$Null
			
			If($? -and $Null -ne $results)
			{
				$FromPool = $results.Name
			}
			ElseIf($? -and $Null -eq $results)
			{
				$FromPool = "VDI Pool not found for Pool Id $($PubItem.VDIPoolId)"
			}
			Else
			{
				$FromPool = "Unable to retrieve VDI Pool for Pool Id $($PubItem.VDIPoolId)"
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Virtual Desktop Application"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Target"; Value = $PubItem.Target; }) > $Null
				$ScriptInformation.Add(@{Data = "Start In"; Value = $PubItem.StartIn; }) > $Null
				
				If($PubItem.StartOnLogon)
				{
					$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = ""; }) > $Null
				}

				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					$ScriptInformation.Add(@{Data = "Parameters"; Value = $PubItem.Parameters; }) > $Null
				}
				
				$ScriptInformation.Add(@{Data = "Connect To"; Value = $ConnectTo; }) > $Null
				$ScriptInformation.Add(@{Data = "From Pool"; Value = $FromPool; }) > $Null
				$ScriptInformation.Add(@{Data = "Settings for Site $xSiteName"; Value = ""; }) > $Null
				
				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($DefaultStartPath)'"; }) > $Null
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($PubItem.StartPath)'"; }) > $Null
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Virtual Desktop Application"
				WriteWordLine 4 0 "Application"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Run"; Value = $WinType; }) > $Null
				$ScriptInformation.Add(@{Data = "Target"; Value = $PubItem.Target; }) > $Null
				$ScriptInformation.Add(@{Data = "Start In"; Value = $PubItem.StartIn; }) > $Null
				$ScriptInformation.Add(@{Data = "Parameters"; Value = $PubItem.Parameters; }) > $Null
				$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = $PubItem.StartOnLogon.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				WriteWordLine 4 0 "Virtual Guest settings"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Connect to"; Value = $ConnectTo; }) > $Null
				$ScriptInformation.Add(@{Data = "from Pool"; Value = $FromPool; }) > $Null
				$ScriptInformation.Add(@{Data = "Persistent"; Value = $PubItem.Persistent.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				OutputPubItemFilters $PubItem "MSWordPDF"
				
				OutputPubItemShortcuts $PubItem "MSWordPDF" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($Text)
			{
				Line 2 "Information"
				Line 3 "Virtual Desktop Application:`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Target:`t`t`t`t`t`t`t" $PubItem.Target
				Line 3 "Start In:`t`t`t`t`t`t" $PubItem.StartIn
				
				If($PubItem.StartOnLogon)
				{
					Line 3 "Start automatically when user logs on" ""
				}
				
				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					Line 3 "Parameters:`t`t`t`t`t`t" $PubItem.Parameters
				}
				
				Line 3 "Connect To:`t`t`t`t`t`t" $ConnectTo
				Line 3 "From Pool:`t`t`t`t`t`t" $FromPool
				Line 3 "Settings for Site $xSiteName"

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($DefaultStartPath)'"
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($PubItem.StartPath)'"
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 "Sites"
				Line 3 "This published item will be available from the following Sites"
				ForEach($Site in $PubItem.PublishToSite)
				{
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					Line 10 $SiteName
				}
				Line 0 ""

				Line 2 "Virtual Guest settings"
				Line 3 "Application"
				Line 4 "Name:`t`t`t`t`t`t" $PubItem.Name
				Line 4 "Description:`t`t`t`t`t" $PubItem.Description
				Line 4 "Run:`t`t`t`t`t`t" $WinType
				Line 4 "Target:`t`t`t`t`t`t" $PubItem.Target
				Line 4 "Start In:`t`t`t`t`t" $PubItem.StartIn
				Line 4 "Parameters:`t`t`t`t`t" $PubItem.Parameters
				Line 4 "Start automatically when user logs on:`t`t" $PubItem.StartOnLogon.ToString()
				Line 0 ""
				
				Line 3 "Application"
				Line 4 "Connect to:`t`t`t`t`t" $ConnectTo
				Line 4 "from Pool:`t`t`t`t`t" $FromPool
				Line 4 "Persistent:`t`t`t`t`t" $PubItem.Persistent.ToString()
				Line 0 ""
				
				OutputPubItemFilters $PubItem "Text"
				
				OutputPubItemShortcuts $PubItem "Text" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
				$rowdata = @()

				$columnHeaders = @("Virtual Desktop Application",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Target",($Script:htmlsb),$PubItem.Target,$htmlwhite))
				$rowdata += @(,("Start In",($Script:htmlsb),$PubItem.StartIn,$htmlwhite))
				
				If($PubItem.StartOnLogon)
				{
					$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),"",$htmlwhite))
				}
				
				If(![String]::IsNullOrEmpty($PubItem.Parameters))
				{
					$rowdata += @(,("Parameters",($Script:htmlsb),$PubItem.Parameters,$htmlwhite))
				}
				
				$rowdata += @(,( "Connect To",($Script:htmlsb),$ConnectTo,$htmlwhite))
				$rowdata += @(,( "From Pool",($Script:htmlsb),$FromPool,$htmlwhite))
				$rowdata += @(,("Settings for Site $xSiteName",($Script:htmlsb),"",$htmlwhite))
				
				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($DefaultStartPath)'",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($PubItem.StartPath)'",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}
			
				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Sites"
				$rowdata = @()

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$columnHeaders = @("This published item will be available from the following Sites",($Script:htmlsb),$SiteName,$htmlwhite)
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Virtual Guest settings"
				$rowdata = @()

				$columnHeaders = @("Name",($Script:htmlsb),$PubItem.Name,$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Run",($Script:htmlsb),$WinType,$htmlwhite))
				$rowdata += @(,("Target",($Script:htmlsb),$PubItem.Target,$htmlwhite))
				$rowdata += @(,("Start In",($Script:htmlsb),$PubItem.StartIn,$htmlwhite))
				$rowdata += @(,("Parameters",($Script:htmlsb),$PubItem.Parameters,$htmlwhite))
				$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),$PubItem.StartOnLogon.ToString(),$htmlwhite))

				$msg = "Application"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
				
				$rowdata = @()

				$columnHeaders = @("Connect To",($Script:htmlsb),$ConnectTo,$htmlwhite)
				$rowdata += @(,( "from Pool",($Script:htmlsb),$FromPool,$htmlwhite))
				$rowdata += @(,("Persistent",($Script:htmlsb),$PubItem.Persistent.ToString(),$htmlwhite))

				$msg = "Virtual Guest settings"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
				
				OutputPubItemFilters $PubItem "HTML"
				
				OutputPubItemShortcuts $PubItem "HTML" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
		}
		ElseIf($PubItem.Type -eq "RDSDesktop")
		{
			$DesktopSize = "Unable to determine"
			If($PubItem.UseAvailableArea)
			{
				$DesktopSize = "Use available area"
			}
			ElseIf($PubItem.Width -eq 0 -and $PubItem.Height -eq 0)
			{
				$DesktopSize = "Full Screen"
			}
			Else
			{
				$DesktopSize = "$($PubItem.Width.ToString())x$($PubItem.Height.ToString())"
			}
			
			Switch ($PubItem.PublishFrom)
			{
				"All"		{$PublishedFrom = "All Servers in Site"; Break}
				"Group"		{$PublishedFrom = "Groups:"; Break}
				"Server"	{$PublishedFrom = "Individual Servers:"; Break}
				Default		{$PublishedFrom = "Unable to determine Published From: $($PubItem.PublishFrom)"; Break}
			}
			
			If($PubItem.AllowMultiMonitor -eq "UseClientSettings")
			{
				$AllowMultiMonitor = "Use Client Settings"
			}
			Else
			{
				$AllowMultiMonitor = $PubItem.AllowMultiMonitor.ToString()
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "RD Session Host Desktop"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Connect to administrative session"; Value = $PubItem.ConnectToConsole.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Desktop Size"; Value = $DesktopSize; }) > $Null
				
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "Published from"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "Published from"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "Published from"; Value = "All Servers in Site"; }) > $Null
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($DefaultStartPath)'"; }) > $Null
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($PubItem.StartPath)'"; }) > $Null
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Publish from"
				$ScriptInformation = New-Object System.Collections.ArrayList
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "$PublishedFrom"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "$PublishedFrom"; Value = $ItemName; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $ItemName; }) > $Null
						}
					}
				}
				Else
				{
					$ScriptInformation.Add(@{Data = "$PublishedFrom"; Value = ""; }) > $Null
					#$ScriptInformation.Add(@{Data = ""; Value = "All Servers in Site"; }) > $Null
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Desktop"
				WriteWordLine 4 0 "Desktop"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Connect to administrative session"; Value = $PubItem.ConnectToConsole.ToString(); }) > $Null
				$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = $PubItem.StartOnLogon.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				WriteWordLine 4 0 "Properties"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Desktop size"; Value = $DesktopSize; }) > $Null
				$ScriptInformation.Add(@{Data = "Multi-Monitor"; Value = $AllowMultiMonitor; }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
				
				OutputPubItemFilters $PubItem "MSWordPDF"
				
				OutputPubItemShortcuts $PubItem "MSWordPDF" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($Text)
			{
				Line 2 "Information"
				Line 3 "RD Session Host Desktop:`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Connect to administrative session:`t`t`t" $PubItem.ConnectToConsole.ToString()
				Line 3 "Desktop Size:`t`t`t`t`t`t" $DesktopSize
				
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						If($cnt -eq 0)
						{
							Line 3 "Published from:`t`t`t`t`t`t" ItemName
						}
						Else
						{
							Line 10 $ItemName
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						If($cnt -eq 0)
						{
							Line 3 "Published from:`t`t`t`t`t`t" ItemName
						}
						Else
						{
							Line 10 $ItemName
						}
					}
				}
				Else
				{
					Line 3 "Published from:`t`t`t`t`t`t" "All Servers in Site"
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($DefaultStartPath)'"
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($PubItem.StartPath)'"
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 "Sites"
				Line 3 "This published item will be available from the following Sites"
				ForEach($Site in $PubItem.PublishToSite)
				{
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					Line 10 $SiteName
				}
				Line 0 ""

				Line 2 "Publish from"
				Line 3 $PublishedFrom
				If($PubItem.PublishFrom -eq "Server")
				{
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						Line 6 $ItemName
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						Line 5 $ItemName
					}
				}
				Line 0 ""
				
				Line 2 "Desktop"
				Line 3 "Desktop"
				Line 4 "Name:`t`t`t`t`t`t" $PubItem.Name
				Line 4 "Description:`t`t`t`t`t" $PubItem.Description
				Line 4 "Connect to administrative session:`t`t" $PubItem.ConnectToConsole.ToString()
				Line 4 "Start automatically when user logs on:`t`t" $PubItem.StartOnLogon.ToString()
				Line 0 ""
				
				Line 3 "Properties"
				Line 4 "Desktop size:`t`t`t`t`t" $DesktopSize
				Line 4 "Multi-Monitor:`t`t`t`t`t" $AllowMultiMonitor
				Line 0 ""

				OutputPubItemFilters $PubItem "Text"
				
				OutputPubItemShortcuts $PubItem "Text" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
				$rowdata = @()
				$columnHeaders = @("RD Session Host Desktop",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Connect to administrative session",($Script:htmlsb),$PubItem.ConnectToConsole.ToString(),$htmlwhite))
				$rowdata += @(,("Desktop Size",($Script:htmlsb),$DesktopSize,$htmlwhite))
				
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						If($cnt -eq 0)
						{
							$rowdata += @(,("Published from",($Script:htmlsb),$ItemName,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						If($cnt -eq 0)
						{
							$rowdata += @(,("Published from",($Script:htmlsb),$ItemName,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				Else
				{
					$rowdata += @(,("Published from",($Script:htmlsb),"All Servers in Site",$htmlwhite))
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($DefaultStartPath)'",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($PubItem.StartPath)'",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}
			
				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Sites"
				$rowdata = @()

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$columnHeaders = @("This published item will be available from the following Sites",($Script:htmlsb),$SiteName,$htmlwhite)
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Publish from"
				$rowdata = @()
				If($PubItem.PublishFrom -eq "Server")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromServer)
					{
						$cnt++
						$ItemName = @(Get-RDS -Id $Item -EA 0 4>$Null).Server
						
						If($cnt -eq 0)
						{
							$columnHeaders = @("$PublishedFrom",($Script:htmlsb),$ItemName,$htmlwhite)
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				ElseIf($PubItem.PublishFrom -eq "Group")
				{
					$cnt = -1
					ForEach($Item in $PubItem.PublishFromGroup)
					{
						$cnt++
						$ItemName = @(Get-RDSGroup -Id $Item -EA 0 4>$Null).Name
						
						If($cnt -eq 0)
						{
							$columnHeaders = @("$PublishedFrom",($Script:htmlsb),$ItemName,$htmlwhite)
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$ItemName,$htmlwhite))
						}
					}
				}
				Else
				{
					$columnHeaders = @("$PublishedFrom",($Script:htmlsb),"",$htmlwhite)
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Desktop"
				$rowdata = @()

				$columnHeaders = @("Name",($Script:htmlsb),$PubItem.Name,$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Connect to administrative session",($Script:htmlsb),$PubItem.ConnectToConsole.ToString(),$htmlwhite))
				$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),$PubItem.StartOnLogon.ToString(),$htmlwhite))

				$msg = "Desktop"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
				
				$rowdata = @()
				$columnHeaders = @("Desktop size",($Script:htmlsb),$DesktopSize,$htmlwhite)
				$rowdata += @(,("Multi-Monitor",($Script:htmlsb),$AllowMultiMonitor,$htmlwhite))

				$msg = "Properties"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				OutputPubItemFilters $PubItem "HTML"
				
				OutputPubItemShortcuts $PubItem "HTML" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
		}
		ElseIf($PubItem.Type -eq "VDIDesktop")
		{
			Switch($PubItem.ConnectTo)
			{
				"AnyGuest"				{$ConnectTo = "Any Guest VM"; Break}
				"SpecificRASTemplate"	{$ConnectTo = "Specific Template ($($PubItem.SelectedGuests[0].VDIGuestName))"; Break
										}
				Default					{$ConnectTo = "Unable to determine Connect To: $($PubItem.ConnectTo)"; Break}
			}
			
			If($PubItem.AllowMultiMonitor -eq "UseClientSettings")
			{
				$AllowMultiMonitor = "Use Client Settings"
			}
			Else
			{
				$AllowMultiMonitor = $PubItem.AllowMultiMonitor.ToString()
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Information"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Virtual Desktop"; Value = "#$($PubItem.Id): $($PubItem.Name)"; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Settings for Site $xSiteName"; Value = ""; }) > $Null
				$ScriptInformation.Add(@{Data = "Connect to"; Value = $ConnectTo; }) > $Null
				
				If($PubItem.Persistent)
				{
					$ScriptInformation.Add(@{Data = "Published item is persistent"; Value = ""; }) > $Null
				}
				
				If($PubItem.UseAvailableArea)
				{
					$ScriptInformation.Add(@{Data = "Desktop Size"; Value = "Use available area"; }) > $Null
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -eq 0 -and $PubItem.Width -eq 0))
				{
					$ScriptInformation.Add(@{Data = "Desktop Size"; Value = "Full Screen"; }) > $Null
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -ne 0 -and $PubItem.Width -ne 0))
				{
					$ScriptInformation.Add(@{Data = "Desktop Size"; Value = "$($PubItem.Height)x$($PubItem.Width)"; }) > $Null
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($DefaultStartPath)'"; }) > $Null
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut on desktop"; Value = ""; }) > $Null
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Start Menu under "; Value = "'$($PubItem.StartPath)'"; }) > $Null
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = ""; }) > $Null
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = "Own Filters"; Value = ""; }) > $Null
				}

				If($PubItem.UserFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " User filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item.Account; }) > $Null
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " IP filtering is enabled"; Value = ""; }) > $Null

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$ScriptInformation.Add(@{Data = ""; Value = $item.From; }) > $Null
							}
							Else
							{
								$ScriptInformation.Add(@{Data = ""; Value = "$($item.From) - $($item.To)"; }) > $Null
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " MAC filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Gateway filtering is enabled"; Value = ""; }) > $Null
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Item; }) > $Null
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$ScriptInformation.Add(@{Data = " Client device operating system filtering is enabled"; Value = ""; }) > $Null
					
					If($PubItem.AllowedOSes.Android)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Android"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "ChromeApp"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "HTML5"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "iOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Linux"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "macOS"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "RAS Web Portal"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Windows"; }) > $Null
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$ScriptInformation.Add(@{Data = ""; Value = "Wyse"; }) > $Null
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "Available in Site(s)"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 3 0 "Sites"
				$ScriptInformation = New-Object System.Collections.ArrayList

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "This published item will be available from the following Sites"; Value = $SiteName; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SiteName; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""


				WriteWordLine 3 0 "Virtual Desktop"
				WriteWordLine 4 0 "Virtual desktop"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Name"; Value = $PubItem.Name; }) > $Null
				$ScriptInformation.Add(@{Data = "Description"; Value = $PubItem.Description; }) > $Null
				$ScriptInformation.Add(@{Data = "Start automatically when user logs on"; Value = $PubItem.StartOnLogon.ToString(); }) > $Null

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 4 0 "Properties"
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Connect to"; Value = $ConnectTo; }) > $Null

				If($PubItem.UseAvailableArea)
				{
					$ScriptInformation.Add(@{Data = "Desktop Size"; Value = "Use available area"; }) > $Null
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -eq 0 -and $PubItem.Width -eq 0))
				{
					$ScriptInformation.Add(@{Data = "Desktop Size"; Value = "Full Screen"; }) > $Null
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -ne 0 -and $PubItem.Width -ne 0))
				{
					$ScriptInformation.Add(@{Data = "Desktop Size"; Value = "$($PubItem.Height)x$($PubItem.Width)"; }) > $Null
				}

				$ScriptInformation.Add(@{Data = "Multi-Monitor"; Value = $AllowMultiMonitor; }) > $Null
				$ScriptInformation.Add(@{Data = "Persistent"; Value = $PubItem.Persistent.ToString(); }) > $Null
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				OutputPubItemFilters $PubItem "MSWordPDF"
				
				OutputPubItemShortcuts $PubItem "MSWordPDF" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($Text)
			{
				Line 2 "Information"
				Line 3 "Virtual Desktop:`t`t`t`t`t" "#$($PubItem.Id): $($PubItem.Name)"
				Line 3 "Description:`t`t`t`t`t`t" $PubItem.Description
				Line 3 "Settings for Site $xSiteName"
				Line 3 "Connect to:`t`t`t`t`t`t" $ConnectTo
				
				If($PubItem.Persistent)
				{
					Line 3 "Published item is persistent"
				}
				
				If($PubItem.UseAvailableArea)
				{
					Line 3 "Desktop Size:`t`t`t`t`t`t" "Use available area"
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -eq 0 -and $PubItem.Width -eq 0))
				{
					Line 3 "Desktop Size:`t`t`t`t`t`t" "Full Screen"
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -ne 0 -and $PubItem.Width -ne 0))
				{
					Line 3 "Desktop Size:`t`t`t`t`t`t" "$($PubItem.Height)x$($PubItem.Width)"
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($DefaultStartPath)'"
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						Line 3 "Create shortcut on desktop"
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						Line 3 "Create shortcut in Start Menu under " "'$($PubItem.StartPath)'"
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						Line 3 "Create shortcut in Auto Start Folder"
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					Line 3 "Own Filters"
				}

				If($PubItem.UserFilterEnabled)
				{
					Line 3 " User filtering is enabled"
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						Line 10 $Item.Account
					}
					Line 0 ""
				}
				If($PubItem.IPFilterEnabled)
				{
					Line 3 " IP filtering is enabled"
					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								Line 10 $item.From
							}
							Else
							{
								Line 10 "$($item.From) - $($item.To)"
							}
						}
					}
					Line 0 ""
				}
				If($PubItem.ClientFilterEnabled)
				{
					Line 3 " Client filtering is enabled"
					
					ForEach($item in $PubItem.AllowedClients)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.MACFilterEnabled)
				{
					Line 3 " MAC filtering is enabled"
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.GWFilterEnabled)
				{
					Line 3 " Gateway filtering is enabled"
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						Line 10 $Item
					}
					Line 0 ""
				}
				If($PubItem.OSFilterEnabled)
				{
					Line 3 " Client device operating system filtering is enabled"
					
					If($PubItem.AllowedOSes.Android)
					{
						Line 10 "Android"
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						Line 10 "ChromeApp"
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						Line 10 "HTML5"
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						Line 10 "iOS"
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						Line 10 "Linux"
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						Line 10 "macOS"
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						Line 10 "RAS Web Portal"
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						Line 10 "Windows"
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						Line 10 "Wyse"
					}
					Line 0 ""
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						Line 3 "Available in Site(s):`t`t`t`t`t" $SiteName
					}
					Else
					{
						Line 10 $SiteName
					}
				}
				Line 0 ""

				Line 2 "Sites"
				Line 3 "This published item will be available from the following Sites"
				ForEach($Site in $PubItem.PublishToSite)
				{
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					Line 10 $SiteName
				}
				Line 0 ""

				Line 2 "Virtual Desktop"
				Line 3 "Virtual desktop"
				Line 4 "Name:`t`t`t`t`t`t" $PubItem.Name
				Line 4 "Description:`t`t`t`t`t" $PubItem.Description
				Line 4 "Start automatically when user logs on:`t`t" $PubItem.StartOnLogon.ToString()
				Line 0 ""
				Line 3 "Properties"
				Line 4 "Connect to:`t`t`t`t`t" $ConnectTo
				If($PubItem.UseAvailableArea)
				{
					Line 4 "Desktop Size:`t`t`t`t`t" "Use available area"
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -eq 0 -and $PubItem.Width -eq 0))
				{
					Line 4 "Desktop Size:`t`t`t`t`t" "Full Screen"
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -ne 0 -and $PubItem.Width -ne 0))
				{
					Line 4 "Desktop Size:`t`t`t`t`t" "$($PubItem.Height)x$($PubItem.Width)"
				}
				Line 4 "Multi-Monitor:`t`t`t`t`t" $AllowMultiMonitor
				Line 4 "Persistent:`t`t`t`t`t" $PubItem.Persistent.ToString()
				Line 0 ""

				OutputPubItemFilters $PubItem "Text"
				
				OutputPubItemShortcuts $PubItem "Text" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Information"
				$rowdata = @()

				$columnHeaders = @("Virtual Desktop",($Script:htmlsb),"#$($PubItem.Id): $($PubItem.Name)",$htmlwhite)
				$rowdata += @(,("",($Script:htmlsb),,$htmlwhite))

				$rowdata += @(,( "Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,( "Settings for Site $xSiteName",($Script:htmlsb),"",$htmlwhite))
				$rowdata += @(,( "Connect to",($Script:htmlsb),$ConnectTo,$htmlwhite))
				
				If($PubItem.Persistent)
				{
					$rowdata += @(,( "Published item is persistent",($Script:htmlsb),"",$htmlwhite))
				}
				
				If($PubItem.UseAvailableArea)
				{
					$rowdata += @(,( "Desktop Size",($Script:htmlsb),"Use available area",$htmlwhite))
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -eq 0 -and $PubItem.Width -eq 0))
				{
					$rowdata += @(,( "Desktop Size",($Script:htmlsb),"Full Screen",$htmlwhite))
				}
				ElseIf($PubItem.UseAvailableArea -eq $False -and ($PubItem.Height -ne 0 -and $PubItem.Width -ne 0))
				{
					$rowdata += @(,( "Desktop Size",($Script:htmlsb),"$($PubItem.Height)x$($PubItem.Width)",$htmlwhite))
				}

				If($PubItem.InheritShortcutDefaultSettings)
				{
					If($DefaultCreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($DefaultStartPath)'",$htmlwhite))
					}
					If($DefaultCreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}
				Else
				{
					If($PubItem.CreateShortcutOnDesktop)
					{
						$rowdata += @(,("Create shortcut on desktop",($Script:htmlsb),"",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartFolder)
					{
						$rowdata += @(,("Create shortcut in Start Menu under ",($Script:htmlsb),"'$($PubItem.StartPath)'",$htmlwhite))
					}
					If($PubItem.CreateShortcutInStartUpFolder)
					{
						$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),"",$htmlwhite))
					}
				}

				If($PubItem.UserFilterEnabled -or 
				   $PubItem.ClientFilterEnabled -or 
				   $PubItem.IPFilterEnabled -or 
				   $PubItem.MACFilterEnabled -or 
				   $PubItem.GWFilterEnabled -or 
				   $PubItem.OSFilterEnabled)
				{
					$rowdata += @(,("Own Filters",($Script:htmlsb),"",$htmlwhite))
				}

				If($PubItem.UserFilterEnabled)
				{
					$rowdata += @(,(" User filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($Item in $PubItem.AllowedUsers)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item.Account,$htmlwhite))
					}
				}
				If($PubItem.IPFilterEnabled)
				{
					$rowdata += @(,(" IP filtering is enabled",($Script:htmlsb),"",$htmlwhite))

					If($PubItem.AllowedIP4s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP4s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}

					If($PubItem.AllowedIP6s.Count -gt 0)
					{
						ForEach($item in $PubItem.AllowedIP6s)
						{
							If($item.From -eq $item.To)
							{
								$rowdata += @(,("",($Script:htmlsb),$item.From,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("",($Script:htmlsb),"$($item.From) - $($item.To)",$htmlwhite))
							}
						}
					}
				}
				If($PubItem.ClientFilterEnabled)
				{
					$rowdata += @(,(" Client filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedClients)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.MACFilterEnabled)
				{
					$rowdata += @(,(" MAC filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedMACs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.GWFilterEnabled)
				{
					$rowdata += @(,(" Gateway filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					ForEach($item in $PubItem.AllowedGWs)
					{
						$rowdata += @(,("",($Script:htmlsb),$Item,$htmlwhite))
					}
				}
				If($PubItem.OSFilterEnabled)
				{
					$rowdata += @(,(" Client device operating system filtering is enabled",($Script:htmlsb),"",$htmlwhite))
					
					If($PubItem.AllowedOSes.Android)
					{
						$rowdata += @(,("",($Script:htmlsb),"Android",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Chrome)
					{
						$rowdata += @(,("",($Script:htmlsb),"ChromeApp",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.HTML5)
					{
						$rowdata += @(,("",($Script:htmlsb),"HTML5",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.iOS)
					{
						$rowdata += @(,("",($Script:htmlsb),"iOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Linux)
					{
						$rowdata += @(,("",($Script:htmlsb),"Linux",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Mac)
					{
						$rowdata += @(,("",($Script:htmlsb),"macOS",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.WebPortal)
					{
						$rowdata += @(,("",($Script:htmlsb),"RAS Web Portal",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Windows)
					{
						$rowdata += @(,("",($Script:htmlsb),"Windows",$htmlwhite))
					}
					
					If($PubItem.AllowedOSes.Wyse)
					{
						$rowdata += @(,("",($Script:htmlsb),"Wyse",$htmlwhite))
					}
				}

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("Available in Site(s)",($Script:htmlsb),$SiteName,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}
			
				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Sites"
				$rowdata = @()

				$cnt =-1
				ForEach($Site in $PubItem.PublishToSite)
				{
					$cnt++
					$SiteName = @(Get-Site -Id $Site -EA 0 4>$Null).Name
					
					If($cnt -eq 0)
					{
						$columnHeaders = @("This published item will be available from the following Sites",($Script:htmlsb),$SiteName,$htmlwhite)
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$SiteName,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 3 0 "Virtual Desktop"
				$rowdata = @()

				$columnHeaders = @("Name",($Script:htmlsb),$PubItem.Name,$htmlwhite)
				$rowdata += @(,("Description",($Script:htmlsb),$PubItem.Description,$htmlwhite))
				$rowdata += @(,("Start automatically when user logs on",($Script:htmlsb),$PubItem.StartOnLogon.ToString(),$htmlwhite))

				$msg = "Desktop"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
				
				$rowdata = @()
				$columnHeaders = @("Connect to",($Script:htmlsb),$ConnectTo,$htmlwhite)
				$rowdata += @(,("Desktop size",($Script:htmlsb),$DesktopSize,$htmlwhite))
				$rowdata += @(,("Multi-Monitor",($Script:htmlsb),$AllowMultiMonitor,$htmlwhite))
				$rowdata += @(,("Persistent",($Script:htmlsb),$PubItem.Persistent.ToString(),$htmlwhite))

				$msg = "Properties"
				$columnWidths = @("200","300")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				OutputPubItemFilters $PubItem "HTML"
				
				OutputPubItemShortcuts $PubItem "HTML" `
				$DefaultCreateShortcutOnDesktop `
				$DefaultCreateShortcutInStartFolder `
				$DefaultStartPath `
				$DefaultCreateShortcutInStartUpFolder `
				$DefaultReplicateShortcutSettings
			}
		}
	}
}

Function OutputPubItemFilters
{
	Param([object] $PubItem, [string] $OutputType)
	
	If($OutputType -eq "MSWordPDF")
	{
		WriteWordLine 3 0 "Filtering"
		If(!($PubItem.UserFilterEnabled))
		{
			WriteWordLine 0 0 "User filtering not enabled"
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "User filtering is enabled"
			WriteWordLine 0 0 "Allow the following Users:"

			$ScriptInformation = New-Object System.Collections.ArrayList
			$NameTable = @()
			
			ForEach($item in $PubItem.AllowedUsers)
			{
				$NameTable += @{
				User = $item.Account;
				Type = $item.Type;
				SID  = $item.Sid
				}
			}

			$Table = AddWordTable -Hashtable $NameTable `
			-Columns User,Type,SID `
			-Headers "User", "Type", "SID" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null

			If($PubItem.UserFilterReplicate)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteWordLine 0 0 ""
		}
		
		If(!($PubItem.ClientFilterEnabled))
		{
			WriteWordLine 0 0 "Client device name filtering not enabled"
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "Client device name filtering is enabled"
			WriteWordLine 0 0 "Allow the following Clients:"

			$ScriptInformation = New-Object System.Collections.ArrayList
			$NameTable = @()
			
			ForEach($item in $PubItem.AllowedClients)
			{
				$NameTable += @{
				Client = $item;
				}
			}

			$Table = AddWordTable -Hashtable $NameTable `
			-Columns Client `
			-Headers "Client" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null

			If($PubItem.ClientFilterReplicate)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteWordLine 0 0 ""
		}
		If(!($PubItem.OSFilterEnabled))
		{
			WriteWordLine 0 0 "Client device operating system filtering not enabled"
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "Client device operating system filtering is enabled"
			WriteWordLine 0 0 "Allow access to clients running on the following operating system:"
			$ScriptInformation = New-Object System.Collections.ArrayList
			If($PubItem.AllowedOSes.Android)
			{
				$ScriptInformation.Add(@{Data = "Android"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "Android"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.Chrome)
			{
				$ScriptInformation.Add(@{Data = "ChromeApp"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "ChromeApp"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.HTML5)
			{
				$ScriptInformation.Add(@{Data = "HTML5"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "HTML5"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.iOS)
			{
				$ScriptInformation.Add(@{Data = "iOS"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "iOS"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.Linux)
			{
				$ScriptInformation.Add(@{Data = "Linux"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "Linux"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.Mac)
			{
				$ScriptInformation.Add(@{Data = "macOS"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "macOS"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.WebPortal)
			{
				$ScriptInformation.Add(@{Data = "RAS Web Portal"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "RAS Web Portal"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.Windows)
			{
				$ScriptInformation.Add(@{Data = "Windows"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "Windows"; Value = "Disabled"; }) > $Null
			}
			
			If($PubItem.AllowedOSes.Wyse)
			{
				$ScriptInformation.Add(@{Data = "Wyse"; Value = "Enabled"; }) > $Null
			}
			Else
			{
				$ScriptInformation.Add(@{Data = "Wyse"; Value = "Disabled"; }) > $Null
			}
			
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null

			If($PubItem.OSFilterReplicate)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteWordLine 0 0 ""
		}
		
		If(!($PubItem.IPFilterEnabled))
		{
			WriteWordLine 0 0 "IP Address filtering not enabled"
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "IP Address filtering is enabled"
			WriteWordLine 0 0 "Allow the following IPs:"

			If($PubItem.AllowedIP4s.Count -gt 0)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$NameTable = @()
				
				ForEach($item in $PubItem.AllowedIP4s)
				{
					If($item.From -eq $item.To)
					{
						$NameTable += @{
						IPv4From = $item.From;
						IPv4To   = ""
						}
					}
					Else
					{
						$NameTable += @{
						IPv4From = $item.From;
						IPv4To   = $item.To
						}
					}
				}

				$Table = AddWordTable -Hashtable $NameTable `
				-Columns IPv4From,IPv4To `
				-Headers "IPv4 Address From", "IPv4 Address To" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				If($PubItem.AllowedIP6s.Count -gt 0)
				{
					WriteWordLine 0 0 ""
				}
			}
			
			If($PubItem.AllowedIP6s.Count -gt 0)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$NameTable = @()
				
				ForEach($item in $PubItem.AllowedIP6s)
				{
					If($item.From -eq $item.To)
					{
						$NameTable += @{
						IPv6From = $item.From;
						IPv6To   = ""
						}
					}
					Else
					{
						$NameTable += @{
						IPv6From = $item.From;
						IPv6To   = $item.To
						}
					}
				}

				$Table = AddWordTable -Hashtable $NameTable `
				-Columns IPv6From,IPv6To `
				-Headers "IPv6 Address From", "IPv6 Address To" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}

			If($PubItem.IPFilterReplicate)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteWordLine 0 0 ""
		}
		
		If(!($PubItem.MACFilterEnabled))
		{
			WriteWordLine 0 0 "MAC filtering not enabled"
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "MAC filtering is enabled"
			WriteWordLine 0 0 "Allow the following MACs:"

			$ScriptInformation = New-Object System.Collections.ArrayList
			$NameTable = @()
			
			ForEach($item in $PubItem.AllowedMACs)
			{
				$NameTable += @{
				MAC = $item;
				}
			}

			$Table = AddWordTable -Hashtable $NameTable `
			-Columns MAC `
			-Headers "MAC" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null

			If($PubItem.MACFilterReplicate)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteWordLine 0 0 ""
		}
		
		If(!($PubItem.GWFilterEnabled))
		{
			WriteWordLine 0 0 "Gateway filtering not enabled"
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "Gateway filtering is enabled"
			WriteWordLine 0 0 "Allow connections from the following Gateways:"

			$ScriptInformation = New-Object System.Collections.ArrayList
			$NameTable = @()
			
			ForEach($item in $PubItem.AllowedGWs)
			{
				$NameTable += @{
				GW = $item;
				}
			}

			$Table = AddWordTable -Hashtable $NameTable `
			-Columns GW `
			-Headers "Gateways" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null

			WriteWordLine 0 0 ""
		}
	}
	If($OutputType -eq "Text")
	{
		Line 2 "Filtering"
		If(!($PubItem.UserFilterEnabled))
		{
			Line 3 "User filtering not enabled"
			Line 0 ""
		}
		Else
		{
			Line 3 "User filtering is enabled"
			Line 3 "Allow the following Users:"
			Line 0 ""
			
			$maxLength = ($PubItem.AllowedUsers.Account | Measure-Object -Property length -Maximum).Maximum
			$NegativeMaxLength = $maxLength * -1
			Line 3 "User" -nonewline
			Line 0 (" " * ($maxLength - 3)) -nonewline
			LIne 0 "Type  SID"
			Line 3 ("=" * ($maxLength + 1 + 6 + 45)) # $maxLength, space, "Type" plus 2 spaces, length of SID
			ForEach($item in $PubItem.AllowedUsers)
			{
				Line 3 ("{0,$NegativeMaxLength} {1,-5} {2,-45}" -f $item.Account,$item.Type,$item.Sid)
			}
			Line 0 ""

			If($PubItem.UserFilterReplicate)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
		}
		
		If(!($PubItem.ClientFilterEnabled))
		{
			Line 3 "Client device name filtering not enabled"
			Line 0 ""
		}
		Else
		{
			Line 3 "Client device name filtering is enabled"
			Line 3 "Allow the following Clients:"
			Line 0 ""
			Line 3 "Client         "
			Line 3 "==============="
			ForEach($item in $PubItem.AllowedClients)
			{
				Line 3 ("{0,-15}" -f $item)
			}
			Line 0 ""

			If($PubItem.ClientFilterReplicate)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
			Line 0 ""
		}
		
		If(!($PubItem.OSFilterEnabled))
		{
			Line 3 "Client device operating system filtering not enabled"
			Line 0 ""
		}
		Else
		{
			Line 3 "Client device operating system filtering is enabled"
			Line 3 "Allow access to clients running on the following operating system:"
			Line 0 ""
			Line 3 "Operating system"
			Line 3 "================"
			If($PubItem.AllowedOSes.Android)
			{
				Line 3 "Android`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "Android`t`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.Chrome)
			{
				Line 3 "ChromeApp`t: " "Enabled"
			}
			Else
			{
				Line 3 "ChromeApp`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.HTML5)
			{
				Line 3 "HTML5`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "HTML5`t`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.iOS)
			{
				Line 3 "iOS`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "iOS`t`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.Linux)
			{
				Line 3 "Linux`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "Linux`t`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.Mac)
			{
				Line 3 "macOS`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "macOS`t`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.WebPortal)
			{
				Line 3 "RAS Web Portal`t: " "Enabled"
			}
			Else
			{
				Line 3 "RAS Web Portal`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.Windows)
			{
				Line 3 "Windows`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "Windows`t`t: " "Disabled"
			}
			
			If($PubItem.AllowedOSes.Wyse)
			{
				Line 3 "Wyse`t`t: " "Enabled"
			}
			Else
			{
				Line 3 "Wyse`t`t: " "Disabled"
			}
			Line 0 ""

			If($PubItem.OSFilterReplicate)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
			Line 0 ""
		}
		
		If(!($PubItem.IPFilterEnabled))
		{
			Line 3 "IP Address filtering not enabled"
			Line 0 ""
		}
		Else
		{
			Line 3 "IP Address filtering is enabled"
			Line 3 "Allow the following IPs:"
			Line 0 ""
			
			If($PubItem.AllowedIP4s.Count -gt 0)
			{
				Line 3 "IPv4 Address From  IPv4 Address To"
				Line 3 "=================================="
				#       255.255.255.255    255.255.255.255
				#       123456789012345SSSS123456789012345
				ForEach($item in $PubItem.AllowedIP4s)
				{
					If($item.From -eq $item.To)
					{
						Line 3 ("{0,-15}" -f $item.From)
					}
					Else
					{
						Line 3 ("{0,-15}    {1,-15}" -f $item.From, $Item.To)
					}
				}
				
				Line 0 ""
			}
			
			If($PubItem.AllowedIP6s.Count -gt 0)
			{
				$MaxFrom    = ($PubItem.AllowedIP6s.From | Measure-Object -Property length -maximum).Maximum
				$MaxTo      = ($PubItem.AllowedIP6s.To | Measure-Object -Property length -maximum).Maximum
				$NegMaxFrom = $MaxFrom * -1
				$NegMaxTo   = $MaxTo * -1
				
				$SpacesFrom = $MaxFrom - 17
				$SpacesTo   = $MaxTo - 15
				If($SpacesFrom -le 0)
				{
					$SpacesFrom = 17
				}
				If($SpacesTo -le 0)
				{
					$SpacesTo = 15
				}
				Line 3 "IPv6 Address From  " -nonewline
				Line 0 (" " * $SpacesFrom) -nonewline
				Line 0 "IPv6 Address To" -nonewline
				Line 0 (" " * $SpacesTo)
				Line 3 ("=" * (($MaxFrom + $MaxTo) + 2))

				ForEach($item in $PubItem.AllowedIP6s)
				{
					If($item.From -eq $item.To)
					{
						Line 3 ("{0,$NegMaxFrom}" -f $item.From)
					}
					Else
					{
						Line 3 ("{0,$NegMaxFrom}  {1,$NegMaxTo}" -f $item.From, $Item.To)
					}
				}
				
				Line 0 ""
			}

			If($PubItem.IPFilterReplicate)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
			Line 0 ""
		}
		
		If(!($PubItem.MACFilterEnabled))
		{
			Line 3 "MAC filtering not enabled"
			Line 0 ""
		}
		Else
		{
			Line 3 "MAC filtering is enabled"
			Line 3 "Allow the following MACs:"
			Line 0 ""
			Line 3 "MAC         "
			Line 3 "============"
			#       123456789012
			ForEach($item in $PubItem.AllowedMACs)
			{
				Line 3 ("{0,-12}" -f $item)
			}
			Line 0 ""

			If($PubItem.MACFilterReplicate)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
			Line 0 ""
		}
		
		If(!($PubItem.GWFilterEnabled))
		{
			Line 3 "Gateway filtering not enabled"
			Line 0 ""
		}
		Else
		{
			Line 3 "Gateway filtering is enabled"
			Line 3 "Allow connections from the following Gateways:"
			Line 0 ""
			Line 3 "Gateways       "
			Line 3 "==============="
			#       123456789012345
			ForEach($item in $PubItem.AllowedGWs)
			{
				Line 3 ("{0,-15}" -f $item)
			}
			Line 0 ""
		}
	}
	If($OutputType -eq "HTML")
	{
		WriteHTMLLine 3 0 "Filtering"
		If(!($PubItem.UserFilterEnabled))
		{
			WriteHTMLLine 0 0 "User filtering not enabled"
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			WriteHTMLLine 0 0 "User filtering is enabled"
			#WriteHTMLLine 0 0 "Allow the following Users:"

			$rowdata = @()
			
			ForEach($item in $PubItem.AllowedUsers)
			{
				$rowdata += @(,(
				$item.Account,$htmlwhite,
				$item.Type,$htmlwhite,
				$item.Sid,$htmlwhite))
			}

			$columnHeaders = @(
			'User',($Script:htmlsb),
			'Type',($Script:htmlsb),
			'SID',($Script:htmlsb))

			$msg = "Allow the following Users:"
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders

			If($PubItem.UserFilterReplicate)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteHTMLLine 0 0 ""
		}
		
		If(!($PubItem.ClientFilterEnabled))
		{
			WriteHTMLLine 0 0 "Client device name filtering not enabled"
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			WriteHTMLLine 0 0 "Client device name filtering is enabled"
			#WriteHTMLLine 0 0 "Allow the following Clients:"

			$rowdata = @()
			
			ForEach($item in $PubItem.AllowedClients)
			{
				$rowdata += @(,(
				$item,$htmlwhite))
			}

			$columnHeaders = @(
			'Client',($Script:htmlsb))

			$msg = "Allow the following Clients:"
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders 

			If($PubItem.ClientFilterReplicate)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteHTMLLine 0 0 ""
		}
		If(!($PubItem.OSFilterEnabled))
		{
			WriteHTMLLine 0 0 "Client device operating system filtering not enabled"
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			WriteHTMLLine 0 0 "Client device operating system filtering is enabled"
			#WriteHTMLLine 0 0 "Allow access to clients running on the following operating system:"
			$rowdata = @()

			If($PubItem.AllowedOSes.Android)
			{
				$columnHeaders = @("Android",($Script:htmlsb),"Enabled",$htmlwhite)
			}
			Else
			{
				$columnHeaders = @("Android",($Script:htmlsb),"Disabled",$htmlwhite)
			}
			
			If($PubItem.AllowedOSes.Chrome)
			{
				$rowdata += @(,("ChromeApp",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("ChromeApp",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.HTML5)
			{
				$rowdata += @(,("HTML5",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("HTML5",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.iOS)
			{
				$rowdata += @(,("iOS",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("iOS",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.Linux)
			{
				$rowdata += @(,("Linux",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("Linux",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.Mac)
			{
				$rowdata += @(,("macOS",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("macOS",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.WebPortal)
			{
				$rowdata += @(,("RAS Web Portal",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("RAS Web Portal",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.Windows)
			{
				$rowdata += @(,("Windows",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("Windows",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			If($PubItem.AllowedOSes.Wyse)
			{
				$rowdata += @(,("Wyse",($Script:htmlsb),"Enabled",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,("Wyse",($Script:htmlsb),"Disabled",$htmlwhite))
			}
			
			$msg = "Allow access to clients running on the following operating system:"
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders 

			If($PubItem.OSFilterReplicate)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteHTMLLine 0 0 ""
		}
		
		If(!($PubItem.IPFilterEnabled))
		{
			WriteHTMLLine 0 0 "IP Address filtering not enabled"
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			WriteHTMLLine 0 0 "IP Address filtering is enabled"
			#WriteHTMLLine 0 0 "Allow the following IPs:"

			If($PubItem.AllowedIP4s.Count -gt 0)
			{
				$rowdata = @()
				
				ForEach($item in $PubItem.AllowedIP4s)
				{
					If($item.From -eq $item.To)
					{
						$rowdata += @(,(
						$item.From,$htmlwhite,
						"",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,(
						$item.From,$htmlwhite,
						$item.To,$htmlwhite))
					}
				}

				$columnHeaders = @(
				'IPv4 Address From',($Script:htmlsb),
				'IPv4 Address To',($Script:htmlsb))

				$msg = "Allow the following IPs:"
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
				If($PubItem.AllowedIP6s.Count -gt 0)
				{
					WriteHTMLLine 0 0 ""
				}
			}

			If($PubItem.AllowedIP6s.Count -gt 0)
			{
				$rowdata = @()
				
				ForEach($item in $PubItem.AllowedIP6s)
				{
					If($item.From -eq $item.To)
					{
						$rowdata += @(,(
						$item.From,$htmlwhite,
						"",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,(
						$item.From,$htmlwhite,
						$item.To,$htmlwhite))
					}
				}

				$columnHeaders = @(
				'IPv6 Address From',($Script:htmlsb),
				'IPv6 Address To',($Script:htmlsb))

				If($PubItem.AllowedIP4s.Count -gt 0)
				{
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
				}
				Else
				{
					$msg = "Allow the following IPs:"
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
				}
			}

			If($PubItem.IPFilterReplicate)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteHTMLLine 0 0 ""
		}
		
		If(!($PubItem.MACFilterEnabled))
		{
			WriteHTMLLine 0 0 "MAC filtering not enabled"
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			WriteHTMLLine 0 0 "MAC filtering is enabled"
			WriteHTMLLine 0 0 "Allow the following MACs:"

			$rowdata = @()
			
			ForEach($item in $PubItem.AllowedMACs)
			{
				$rowdata += @(,(
				$item,$htmlwhite))
			}

			$columnHeaders = @(
			'MAC',($Script:htmlsb))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders 

			If($PubItem.MACFilterReplicate)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
			WriteHTMLLine 0 0 ""
		}
		
		If(!($PubItem.GWFilterEnabled))
		{
			WriteHTMLLine 0 0 "Gateway filtering not enabled"
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			WriteHTMLLine 0 0 "Gateway filtering is enabled"
			WriteHTMLLine 0 0 "Allow connections from the following Gateways:"

			$rowdata = @()
			
			ForEach($item in $PubItem.AllowedGWs)
			{
				$rowdata += @(,(
				$item,$htmlwhite))
			}

			$columnHeaders = @(
			'Gateways',($Script:htmlsb))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders 
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputPubItemShortCuts
{
	Param([object] $PubItem, 
	[string] $OutputType,
	[string] $DefaultCreateShortcutOnDesktop,
	[string] $DefaultCreateShortcutInStartFolder,
	[string] $DefaultStartPath,
	[string] $DefaultCreateShortcutInStartUpFolder,
	[string] $DefaultReplicateShortcutSettings
	)
	
	If($OutputType -eq "MSWordPDF")
	{
		WriteWordLine 3 0 "Shortcuts"

		$ScriptInformation = New-Object System.Collections.ArrayList
		If($PubItem.InheritShortcutDefaultSettings)
		{
			$ScriptInformation.Add(@{Data = "Create shortcut on Desktop"; Value = $DefaultCreateShortcutOnDesktop; }) > $Null
			$ScriptInformation.Add(@{Data = "Create shortcut in Start Folder"; Value = $DefaultCreateShortcutInStartFolder; }) > $Null
			If($DefaultCreateShortcutInStartFolder)
			{
				$ScriptInformation.Add(@{Data = ""; Value = $DefaultStartPath; }) > $Null
			}
			$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = $DefaultCreateShortcutInStartUpFolder; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			
			If($DefaultReplicateShortcutSettings)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
		}
		Else
		{
			$ScriptInformation.Add(@{Data = "Create shortcut on Desktop"; Value = $PubItem.CreateShortcutOnDesktop; }) > $Null
			$ScriptInformation.Add(@{Data = "Create shortcut in Start Folder"; Value = $PubItem.CreateShortcutInStartFolder; }) > $Null
			If($PubItem.CreateShortcutInStartFolder)
			{
				$ScriptInformation.Add(@{Data = ""; Value = $PubItem.StartPath; }) > $Null
			}
			$ScriptInformation.Add(@{Data = "Create shortcut in Auto Start Folder"; Value = $PubItem.CreateShortcutInStartUpFolder; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			
			If($PubItem.ReplicateShortcutSettings)
			{
				WriteWordLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteWordLine 0 0 "Settings are not replicated to all Sites"
			}
		}
		WriteWordLine 0 0 ""
	}
	If($OutputType -eq "Text")
	{
		Line 2 "Shortcuts"

		If($PubItem.InheritShortcutDefaultSettings)
		{
			Line 3 "Create shortcut on Desktop:`t`t`t`t" $DefaultCreateShortcutOnDesktop
			Line 3 "Create shortcut in Start Folder:`t`t`t" $DefaultCreateShortcutInStartFolder
			If($DefaultCreateShortcutInStartFolder)
			{
				Line 4 $DefaultStartPath
			}
			Line 3 "Create shortcut in Auto Start Folder:`t`t`t" $DefaultCreateShortcutInStartUpFolder
			
			If($DefaultReplicateShortcutSettings)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
		}
		Else
		{
			Line 3 "Create shortcut on Desktop:`t`t`t`t" $PubItem.CreateShortcutOnDesktop
			Line 3 "Create shortcut in Start Folder:`t`t`t" $PubItem.CreateShortcutInStartFolder
			If($PubItem.CreateShortcutInStartFolder)
			{
				Line 4 $PubItem.StartPath
			}
			Line 3 "Create shortcut in Auto Start Folder:`t`t`t" $PubItem.CreateShortcutInStartUpFolder
			
			If($PubItem.ReplicateShortcutSettings)
			{
				Line 3 "Settings are replicated to all Sites"
			}
			Else
			{
				Line 3 "Settings are not replicated to all Sites"
			}
		}
		Line 0 ""
	}
	If($OutputType -eq "HTML")
	{
		WriteHTMLLine 3 0 "Shortcuts"

		$rowdata = @()
		If($PubItem.InheritShortcutDefaultSettings)
		{
			$columnHeaders = @("Create shortcut on Desktop",($Script:htmlsb),$DefaultCreateShortcutOnDesktop.ToString(),$htmlwhite)
			$rowdata += @(,("Create shortcut in Start Folder",($Script:htmlsb),$DefaultCreateShortcutInStartFolder.ToString(),$htmlwhite))
			If($DefaultCreateShortcutInStartFolder)
			{
				$rowdata += @(,("",($Script:htmlsb),$DefaultStartPath,$htmlwhite))
			}
			$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),$DefaultCreateShortcutInStartUpFolder.ToString(),$htmlwhite))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			
			If($DefaultReplicateShortcutSettings)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
		}
		Else
		{
			$columnHeaders = @("Create shortcut on Desktop",($Script:htmlsb),$PubItem.CreateShortcutOnDesktop.ToString(),$htmlwhite)
			$rowdata += @(,("Create shortcut in Start Folder",($Script:htmlsb),$PubItem.CreateShortcutInStartFolder.ToString(),$htmlwhite))
			If($PubItem.CreateShortcutInStartFolder)
			{
				$rowdata += @(,("",($Script:htmlsb),$PubItem.StartPath,$htmlwhite))
			}
			$rowdata += @(,("Create shortcut in Auto Start Folder",($Script:htmlsb),$PubItem.CreateShortcutInStartUpFolder.ToString(),$htmlwhite))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			
			If($PubItem.ReplicateShortcutSettings)
			{
				WriteHTMLLine 0 0 "Settings are replicated to all Sites"
			}
			Else
			{
				WriteHTMLLine 0 0 "Settings are not replicated to all Sites"
			}
		}
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region process universal printing
Function ProcessUniversalPrinting
{
	Write-Verbose "$(Get-Date): Processing Universal Printing"
	
	OutputUniversalPrintingSectionPage
	
	Write-Verbose "$(Get-Date): `tProcessing Universal Printing"
	
	$results = Get-RASPrintingSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Universal Printing information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Universal Printing information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Universal Printing information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Universal Printing information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No Universal Printing information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Universal Printing information was found"
		}
		If($Text)
		{
			Line 0 "No Universal Printing information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Universal Printing information was found"
		}
	}
	Else
	{
		$RASPrinterSettings = $results
		
		$Printingobj = [PSCustomObject] @{
			PrinterNamePattern = $Results.PrinterNamePattern
			ReplicateSettings  = $Results.ReplicatePrinterPattern
			PrinterRetention   = $Results.PrinterRetention
		}
		
		$results = Get-RDS -EA 0 4>$Null
		
		If(!($?))
		{
			Write-Warning "
			`n`n
			`t`t
			Unable to retrieve RDS Printing information
			`n`n
			"
			$RDSobj = [PSCustomObject] @{
				Server = "Unable to retrieve RDS Printing information"
				Type = "Unable to retrieve RDS Printing information"
				PrintingState = "Unable to retrieve RDS Printing information"
			}
		}
		ElseIf($? -and $null -eq $results)
		{
			Write-Warning "
			`n`n
			`t`t
			No RDS Printing information was found
			`n`n
			"
			$RDSobj = [PSCustomObject] @{
				Server = "No RDS Printing information was found"
				Type = "No RDS Printing information was found"
				PrintingState = "No RDS Printing information was found"
			}
		}
		Else
		{
			$RDSobj = [PSCustomObject] @{
				Server = $Results.Server
				Type = "RD Session Hosts"
				PrintingState = $Results.EnablePrinting
			}
		}
		
		$results = Get-VDIHost -EA 0 4>$Null
		
		If(!($?))
		{
			Write-Warning "
			`n`n
			`t`t
			Unable to retrieve VDI Hosts Printing information"
			`n`n
			
			$VDIHostsobj = [PSCustomObject] @{
				Server = "Unable to retrieve VDI Hosts Printing information"
				Type = "Unable to retrieve VDI Hosts Printing information"
				PrintingState = "Unable to retrieve VDI Hosts Printing information"
			}
		}
		ElseIf($? -and $null -eq $results)
		{
			Write-Warning "
			`n`n
			`t`t
			No VDI Hosts Printing information was found
			`n`n
			"
			$VDIHostsobj = [PSCustomObject] @{
				Server = "No VDI Hosts Printing information was found"
				Type = "No VDI Hosts Printing information was found"
				PrintingState = "No VDI Hosts Printing information was found"
			}
		}
		Else
		{
			$VDIHostsobj = [PSCustomObject] @{
				Server = $Results.Server
				Type = "VDI Providers"
				PrintingState = $Results.EnablePrinting
			}
		}
		
		OutputUniversalPrintingSettings $Printingobj $RDSobj $VDIHostsobj
		
		OutputUniversalPrinterDriversSettings $RASPrinterSettings
		
		$results = Get-RASPrintingSettings -EA 0 4>$Null
		
		If(!($?))
		{
			Write-Warning "
			`n`n
			`t`t
			Unable to retrieve Fonts management information
			`n`n
			"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Unable to retrieve Fonts management information"
			}
			If($Text)
			{
				Line 0 "Unable to retrieve Fonts management information"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Unable to retrieve Fonts management information"
			}
		}
		ElseIf($? -and $null -eq $results)
		{
			Write-Warning "
			`n`n
			`t`t
			No Fonts management information was found
			`n`n
			"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "No Fonts management information was found"
			}
			If($Text)
			{
				Line 0 "No Fonts management information was found"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "No Fonts management information was found"
			}
		}
		Else
		{
			OutputUniversalPrinterFontsSettings $results
		}
	}
}

Function OutputUniversalPrintingSectionPage
{
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Universal printing"
	}
	If($Text)
	{
		Line 0 "Universal printing"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Universal printing"
	}
}

Function OutputUniversalPrintingSettings
{
	Param([object]$Printingobj, [object]$RDSobj, [object]$VDIHostsobj)
 
	Write-Verbose "$(Get-Date): `t`tOutput Universal printing"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Universal printing"
	}
	If($Text)
	{
		Line 1 "Universal printing"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Universal printing"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Options"
	}
	If($Text)
	{
		Line 2 "Options"
	}
	If($HTML)
	{
		#WriteHTMLLine 3 0 ""
	}

	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Printer rename pattern"; Value = $Printingobj.PrinterNamePattern; }) > $Null
		$ScriptInformation.Add(@{Data = "Printer retention"; Value = $Printingobj.PrinterRetention; }) > $Null
		$ScriptInformation.Add(@{Data = 'Replicate settings'; Value = $Printingobj.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 275;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Printer rename pattern`t: " $Printingobj.PrinterNamePattern
		Line 3 "Printer retention`t: " $Printingobj.PrinterRetention
		Line 3 "Replicate settings`t: " $Printingobj.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @(
			"Printer rename pattern",($Script:htmlsb),
			$Printingobj.PrinterNamePattern,$htmlwhite
		)
		$rowdata += @(,(
			'Printer retention',($Script:htmlsb),
			$Printingobj.PrinterRetention.ToString(),$htmlwhite)
		)
		$rowdata += @(,(
			'Replicate settings',($Script:htmlsb),
			$Printingobj.ReplicateSettings.ToString(),$htmlwhite)
		)

		$msg = "Options"
		$columnWidths = @("200","300")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Servers in Site"
	}
	If($Text)
	{
		Line 2 "Servers in Site"
	}
	If($HTML)
	{
		#WriteHTMLLine 3 0 ""
	}

	Switch ($RDSobj.PrintingState)
	{
		$True	{$RDSPrintingState = "Enabled"; Break}
		$False	{$RDSPrintingState = "Disabled"; Break}
		Default	{$RDSPrintingState = "Unable to determine RD Sessions Hosts Printing State: $($RDSobj.PrintingState)"; Break}
	}
	
	Switch ($VDIHostsobj.PrintingState)
	{
		$True	{$VDIHostsPrintingState = "Enabled"; Break}
		$False	{$VDIHostsPrintingState = "Disabled"; Break}
		Default	{$VDIHostsPrintingState = "Unable to determine VDI Hosts Printing State: $($VDIHostsobj.PrintingState)"; Break}
	}

	$RDSType = $RDSObj.Type
	$VDIType = $VDIHostsObj.Type
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ServersInSiteTable = @()
		
		ForEach($RDSServer in $RDSobj.Server)
		{
			$ServersInSiteTable += @{
				Server = $RDSServer
				Type   = $RDSType
				State  = $RDSPrintingState
			}
		}
		
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$ServersInSiteTable += @{
				Server = $VDIServer
				Type   = $VDIType
				State  = $VDIHostsPrintingState
			}
		}

		$Table = AddWordTable -Hashtable $ServersInSiteTable `
		-Columns Server, Type, State `
		-Headers "Server", "Type", "State" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 100;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Server                         Type                 State   "
		Line 3 "============================================================"
		#       123456789012345678901234567890S12345678901234567890S12345678

		ForEach($RDSServer in $RDSobj.Server)
		{
			Line 3 ( "{0,-30} {1,-20} {2,-8}" -f 
				$RDSServer, 
				$RDSType, 
				$RDSPrintingState
			)
		}
		
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			Line 3 ( "{0,-30} {1,-20} {2,-8}" -f 
				$VDIServer, 
				$VDIType, 
				$VDIHostsPrintingState
			)
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()

		ForEach($RDSServer in $RDSobj.Server)
		{
			$rowdata += @(,(
				$RDSServer,$htmlwhite,
				$RDSType,$htmlwhite,
				$RDSPrintingState,$htmlwhite)
			)
		}
		
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$rowdata += @(,(
				$VDIServer,$htmlwhite,
				$VDIType,$htmlwhite,
				$VDIHostsPrintingState,$htmlwhite)
			)
		}
		
		$columnHeaders = @(
			'Server',($Script:htmlsb),
			'Type',($Script:htmlsb),
			'State',($Script:htmlsb)
		)

		$msg = "Servers in Site"
		$columnWidths = @("200","100","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputUniversalPrinterDriversSettings 
{
	Param([object] $RASPrinterSettings)
 
	Write-Verbose "$(Get-Date): `t`tOutput Printer drivers"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Printer drivers"
	}
	If($Text)
	{
		Line 1 "Printer drivers"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Printer drivers"
	}

	Switch ($RASPrinterSettings.DriverAllowMode)
	{
		"AllowRedirUsingAnyDriver"
			{$RASPrinterSettingsDriverAllowMode = "Allow redirection of printers using any driver"; Break}
		"AllowRedirUsingSpecifiedDriver"
			{$RASPrinterSettingsDriverAllowMode = "Allow redirection of printers using one of the following drivers"; Break}
		"DoNotAllowRedirUsingSpecifiedDriver"
			{$RASPrinterSettingsDriverAllowMode = "Don't allow redirection of printers that use one of the following drivers"; Break}
		Default
			{$RASPrinterSettingsDriverAllowMode = "Unable to determine RAS Printer Setting Driver Allow Mode: $($RASPrinterSettings.DriverAllowMode)"; Break}
	}
	
	If($RASPrinterSettingsDriverAllowMode -eq "Allow redirection of printers using any driver")
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Mode"; Value = $RASPrinterSettingsDriverAllowMode; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 275;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Mode: " $RASPrinterSettingsDriverAllowMode
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @(
				"Mode",($Script:htmlsb),
				$RASPrinterSettingsDriverAllowMode,$htmlwhite
			)

			$msg = ""
			$columnWidths = @("100","300")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		$tmpArray = $RASPrinterSettings.PrinterDriversArray.Split(",")
		
		If($MSWord -or $PDF)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Mode"; Value = $RASPrinterSettingsDriverAllowMode; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 275;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			$ScriptInformation = New-Object System.Collections.ArrayList
			$DriverNameTable = @()
			
			ForEach($item in $tmpArray)
			{
				$DriverNameTable += @{DriverName = $item}
			}

			$Table = AddWordTable -Hashtable $DriverNameTable `
			-Columns DriverName `
			-Headers "Driver name" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			$maxLength = ($tmparray | Measure-Object -Property length -Maximum).Maximum
			Line 2 "Mode: " $RASPrinterSettingsDriverAllowMode
			Line 0 ""
			Line 2 "Driver name"
			Line 2 ("=" * $maxLength)
			
			ForEach($item in $tmpArray)
			{
				Line 2 $item
			}
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @(
				"Mode",($Script:htmlsb),
				$RASPrinterSettingsDriverAllowMode,$htmlwhite
			)

			$msg = ""
			$columnWidths = @("100","300")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
			
			$rowdata = @()

			ForEach($item in $tmpArray)
			{
				$rowdata += @(,($item,$htmlwhite))
			}
			
			$columnHeaders = @(
				'Driver name',($Script:htmlsb)
			)

			$msg = ""
			$columnWidths = @("100")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputUniversalPrinterFontsSettings 
{
	Param([object] $RASFontsSettings)

	Write-Verbose "$(Get-Date): `t`tOutput Fonts management"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Fonts management"
	}
	If($Text)
	{
		Line 1 "Fonts management"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Fonts management"
	}

	$tmpArray = $RASFontsSettings.ExcludedFontsArray.Split(",")
	$tmpArray = $tmpArray | Sort-Object
	$maxLength = ($tmparray | Measure-Object -Property length -Maximum).Maximum

	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Transfer Fonts"; Value = $RASFontsSettings.EmbedFonts; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		If($RASFontsSettings.EmbedFonts)
		{
			WriteWordLine 0 0 "Exclude the following Fonts from embedding:"
			$ScriptInformation = New-Object System.Collections.ArrayList
			$FontsNameTable = @()
			
			ForEach($item in $tmpArray)
			{
				$FontsNameTable += @{FontName = $item}
			}

			$Table = AddWordTable -Hashtable $FontsNameTable `
			-Columns FontName `
			-Headers "Font Name" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 125;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		Line 2 "Transfer Fonts: " $RASFontsSettings.EmbedFonts
		Line 0 ""
		If($RASFontsSettings.EmbedFonts)
		{
			Line 2 "Exclude the following Fonts from embedding:"
			Line 0 ""
			Line 2 "Font Name"
			Line 2 ("=" * $maxLength)
			
			ForEach($item in $tmpArray)
			{
				Line 2 $item
			}
			Line 0 ""
		}
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @(
			"Transfer Fonts",($Script:htmlsb),
			$RASFontsSettings.EmbedFonts.ToString(),$htmlwhite
		)

		$msg = ""
		$columnWidths = @("100","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
		
		If($RASFontsSettings.EmbedFonts)
		{
			$rowdata = @()

			ForEach($item in $tmpArray)
			{
				$rowdata += @(,($item,$htmlwhite))
			}
			
			$columnHeaders = @(
				'Font Name',($Script:htmlsb)
			)

			$msg = "Exclude the following Fonts from embedding:"
			$columnWidths = @("125")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}

	If($RASFontsSettings.AutoInstallFonts.Count -eq 0)
	{
		$tmpArray = @()
	}
	Else
	{
		$tmpArray = $RASFontsSettings.AutoInstallFonts.Split(",")
		$tmpArray = $tmpArray | Sort-Object
		$maxLength = ($tmparray | Measure-Object -Property length -Maximum).Maximum
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Auto install fonts:"
		
		If($tmpArray.Count -eq 0)
		{
			WriteWordLine 0 0 "There are no Auto install fonts"
		}
		Else
		{
			$FontsFileTable = @()
			
			ForEach($item in $tmpArray)
			{
				$FontsFileTable += @{FontFile = $item}
			}

			$Table = AddWordTable -Hashtable $FontsFileTable `
			-Columns FontFile `
			-Headers "Font File" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 125;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		WriteWordLine 0 0 ""

		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Replicate settings"; Value = $RASFontsSettings.ReplicatePrinterFont; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		If($tmpArray.Count -eq 0)
		{
			Line 2 "There are no Auto install fonts"
		}
		Else
		{
			Line 2 "Auto install fonts:"
			Line 0 ""
			Line 2 "Font File"
			Line 2 ("=" * $maxLength)
			
			ForEach($item in $tmpArray)
			{
				Line 2 $item
			}
		}
		Line 0 ""
		Line 2 "Replicate settings: " $RASFontsSettings.ReplicatePrinterFont
		Line 0 ""
	}
	If($HTML)
	{
		If($tmpArray.Count -eq 0)
		{
			WriteHTMLLine 0 0 "There are no Auto install fonts"
		}
		Else
		{
			$rowdata = @()

			ForEach($item in $tmpArray)
			{
				$rowdata += @(,($item,$htmlwhite))
			}
			
			$columnHeaders = @(
				'Font File',($Script:htmlsb)
			)

			$msg = "Auto install fonts:"
			$columnWidths = @("125")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		}
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @(
			"Replicate settings",($Script:htmlsb),
			$RASFontsSettings.ReplicatePrinterFont.ToString(),$htmlwhite
		)

		$msg = ""
		$columnWidths = @("125","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region process universal scanning
Function ProcessUniversalScanning
{
	Write-Verbose "$(Get-Date): Processing Universal scanning"
	
	OutputUniversalScanningSectionPage
	
	Write-Verbose "$(Get-Date): `tProcessing Universal Scanning"
	
	$results = Get-RASScanningSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Universal Scanning information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Universal Scanning information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Universal Scanning information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Universal Scanning information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No Universal Scanning information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Universal Scanning information was found"
		}
		If($Text)
		{
			Line 0 "No Universal Scanning information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Universal Scanning information was found"
		}
	}
	Else
	{
		$WIAobj = [PSCustomObject] @{
			WIANamePattern    = $Results.WIANamePattern
			ReplicateSettings = $Results.ReplicateWIAPattern
		}
		
		$TWAINobj = [PSCustomObject] @{
			TWAINNamePattern  = $Results.TWAINNamePattern
			ReplicateSettings = $Results.ReplicateTWAINPattern
		}
		
		$results = Get-RDS -EA 0 4>$Null
		
		If(!($?))
		{
			Write-Warning "
			`n`n
			`t`t
			Unable to retrieve RDS Scanning information
			`n`n
			"
			$RDSobj = [PSCustomObject] @{
				Server = "Unable to retrieve RDS Scanning information"
				Type = "Unable to retrieve RDS Scanning information"
				WIAState = "Unable to retrieve RDS Scanning information"
				TWAINState = "Unable to retrieve RDS Scanning information"
			}
		}
		ElseIf($? -and $null -eq $results)
		{
			Write-Warning "
			`n`n
			`t`t
			No RDS Scanning information was found
			`n`n
			"
			$RDSobj = [PSCustomObject] @{
				Server = "No RDS Scanning information was found"
				Type = "No RDS Scanning information was found"
				WIAState = "No RDS Scanning information was found"
				TWAINState = "No RDS Scanning information was found"
			}
		}
		Else
		{
			$RDSobj = [PSCustomObject] @{
				Server = $Results.Server
				Type = "RD Session Hosts"
				WIAState = $Results.EnableWIA
				TWAINState = $Results.EnableTWAIN
			}
		}
		
		$results = Get-VDIHost -EA 0 4>$Null
		
		If(!($?))
		{
			Write-Warning "
			`n`n
			`t`t
			Unable to retrieve VDI Hosts Scanning information
			`n`n
			"
			$VDIHostsobj = [PSCustomObject] @{
				Server = "Unable to retrieve VDI Hosts Scanning information"
				Type = "Unable to retrieve VDI Hosts Scanning information"
				WIAState = "Unable to retrieve VDI Hosts Scanning information"
				TWAINState = "Unable to retrieve VDI Hosts Scanning information"
			}
		}
		ElseIf($? -and $null -eq $results)
		{
			Write-Warning "
			`n`n
			`t`t
			No VDI Hosts Scanning information was found
			`n`n
			"
			$VDIHostsobj = [PSCustomObject] @{
				Server = "No VDI Hosts Scanning information was found"
				Type = "No VDI Hosts Scanning information was found"
				WIAState = "No VDI Hosts Scanning information was found"
				TWAINState = "No VDI Hosts Scanning information was found"
			}
		}
		Else
		{
			$VDIHostsobj = [PSCustomObject] @{
				Server = $Results.Server
				Type = "VDI Providers"
				WIAState = $Results.EnableWIA
				TWAINState = $Results.EnableTWAIN
			}
		}
		
		OutputUniversalScanningSettings $WIAobj $TWAINobj $RDSobj $VDIHostsobj
	}
}

Function OutputUniversalScanningSectionPage
{
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Universal Scanning"
	}
	If($Text)
	{
		Line 0 "Universal Scanning"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Universal Scanning"
	}
}

Function OutputUniversalScanningSettings
{
 Param([object]$WIAobj, [object]$TWAINobj, [object]$RDSobj, [object]$VDIHostsobj)
 
	Write-Verbose "$(Get-Date): `t`tOutput WIA"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "WIA"
	}
	If($Text)
	{
		Line 1 "WIA"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "WIA"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Scanner rename"
	}
	If($Text)
	{
		Line 2 "Scanner rename"
	}
	If($HTML)
	{
		#WriteHTMLLine 3 0 ""
	}

	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Pattern"; Value = $WIAObj.WIANamePattern; }) > $Null
		$ScriptInformation.Add(@{Data = 'Replicate settings'; Value = $WIAObj.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 275;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Pattern`t`t`t: " $WIAObj.WIANamePattern
		Line 3 "Replicate settings`t: " $WIAObj.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @(
			"Pattern",($Script:htmlsb),
			$WIAObj.WIANamePattern,$htmlwhite
		)
		$rowdata += @(,(
			'Replicate settings',($Script:htmlsb),
			$WIAobj.ReplicateSettings.ToString(),$htmlwhite)
		)

		$msg = "Scanner rename"
		$columnWidths = @("100","300")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Servers in Site"
	}
	If($Text)
	{
		Line 2 "Servers in Site"
	}
	If($HTML)
	{
		#WriteHTMLLine 3 0 ""
	}

	$RDSType = $RDSObj.Type
	$VDIType = $VDIHostsObj.Type
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ServersInSiteTable = @()
		
		$cnt = -1
		ForEach($RDSServer in $RDSobj.Server)
		{
			$cnt++
			$ServersInSiteTable += @{
				Server = $RDSServer
				Type   = $RDSType
				State  = $RDSobj.WIAState[$cnt].ToString()
			}
		}
		
		$cnt = -1
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$cnt++
			$ServersInSiteTable += @{
				Server = $VDIServer
				Type   = $VDIType
				State  = $VDIHostsobj.WIAState[$cnt].ToString()
			}
		}

		$Table = AddWordTable -Hashtable $ServersInSiteTable `
		-Columns Server, Type, State `
		-Headers "Server", "Type", "State" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 100;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Server                         Type                 State   "
		Line 3 "============================================================"
		#       123456789012345678901234567890S12345678901234567890S12345678
		$cnt = -1
		ForEach($RDSServer in $RDSobj.Server)
		{
			$cnt++
			Line 3 ( "{0,-30} {1,-20} {2,-8}" -f 
				$RDSServer, 
				$RDSType, 
				$RDSobj.WIAState[$cnt].ToString()
			)
		}
		
		$cnt = -1
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$cnt++
			Line 3 ( "{0,-30} {1,-20} {2,-8}" -f 
				$VDIServer, 
				$VDIType, 
				$VDIHostsobj.WIAState[$cnt].ToString()
			)
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()

		$cnt = -1
		ForEach($RDSServer in $RDSobj.Server)
		{
			$cnt++
			$rowdata += @(,(
				$RDSServer,$htmlwhite,
				$RDSType,$htmlwhite,
				$RDSobj.WIAState[$cnt].ToString(),$htmlwhite)
			)
		}
		
		$cnt = -1
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$cnt++
			$rowdata += @(,(
				$VDIServer,$htmlwhite,
				$VDIType,$htmlwhite,
				$VDIHostsobj.WIAState[$cnt].ToString(),$htmlwhite)
			)
		}

		$columnHeaders = @(
			'Server',($Script:htmlsb),
			'Type',($Script:htmlsb),
			'State',($Script:htmlsb)
		)

		$msg = "Servers in Site"
		$columnWidths = @("200","100","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

	Write-Verbose "$(Get-Date): `t`tOutput TWAIN"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "TWAIN"
	}
	If($Text)
	{
		Line 1 "TWAIN"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "TWAIN"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Scanner rename"
	}
	If($Text)
	{
		Line 2 "Scanner rename"
	}
	If($HTML)
	{
		#WriteHTMLLine 3 0 ""
	}

	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Pattern"; Value = $TWAINobj.TWAINNamePattern; }) > $Null
		$ScriptInformation.Add(@{Data = 'Replicate settings'; Value = $TWAINobj.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 275;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Pattern`t`t`t`t: " $TWAINobj.TWAINNamePattern
		Line 3 "Replicate settings`t: " $TWAINobj.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @(
			"Pattern",($Script:htmlsb),
			$TWAINobj.TWAINNamePattern,$htmlwhite
		)
		$rowdata += @(,(
			'Replicate settings',($Script:htmlsb),
			$TWAINobj.ReplicateSettings.ToString(),$htmlwhite)
		)

		$msg = "Scanner rename"
		$columnWidths = @("100","300")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Servers in Site"
	}
	If($Text)
	{
		Line 2 "Servers in Site"
	}
	If($HTML)
	{
		#WriteHTMLLine 3 0 ""
	}

	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ServersInSiteTable = @()
		
		$cnt = -1
		ForEach($RDSServer in $RDSobj.Server)
		{
			$cnt++
			$ServersInSiteTable += @{
				Server = $RDSServer
				Type   = $RDSType
				State  = $RDSobj.TwainState[$cnt].ToString()
			}
		}
		
		$cnt = -1
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$cnt++
			$ServersInSiteTable += @{
				Server = $VDIServer
				Type   = $VDIType
				State  = $VDIHostsobj.TwainState[$cnt].ToString()
			}
		}

		$Table = AddWordTable -Hashtable $ServersInSiteTable `
		-Columns Server, Type, State `
		-Headers "Server", "Type", "State" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 100;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Server                         Type                 State   "
		Line 3 "============================================================"
		#       123456789012345678901234567890S12345678901234567890S12345678
		$cnt = -1
		ForEach($RDSServer in $RDSobj.Server)
		{
			$cnt++
			Line 3 ( "{0,-30} {1,-20} {2,-8}" -f 
				$RDSServer, 
				$RDSType, 
				$RDSobj.TwainState[$cnt].ToString()
			)
		}
		
		$cnt = -1
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$cnt++
			Line 3 ( "{0,-30} {1,-20} {2,-8}" -f 
				$VDIServer, 
				$VDIType, 
				$VDIHostsobj.TwainState[$cnt].ToString()
			)
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()

		$cnt = -1
		ForEach($RDSServer in $RDSobj.Server)
		{
			$cnt++
			$rowdata += @(,(
				$RDSServer,$htmlwhite,
				$RDSType,$htmlwhite,
				$RDSobj.TwainState[$cnt].ToString(),$htmlwhite)
			)
		}
		
		$cnt = -1
		ForEach($VDIServer in $VDIHostsobj.Server)
		{
			$cnt++
			$rowdata += @(,(
				$VDIServer,$htmlwhite,
				$VDIType,$htmlwhite,
				$VDIHostsobj.TwainState[$cnt].ToString(),$htmlwhite)
			)
		}
		
		$columnHeaders = @(
			'Server',($Script:htmlsb),
			'Type',($Script:htmlsb),
			'State',($Script:htmlsb)
		)

		$msg = "Servers in Site"
		$columnWidths = @("200","100","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region process connection
Function ProcessConnection
{
	Write-Verbose "$(Get-Date): Processing Connection"
	
	OutputConnectionSectionPage
	
	Write-Verbose "$(Get-Date): `tProcessing Authentication"
	
	$results = Get-RASAuthSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve authentication information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve authentication information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve authentication information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve authentication information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No authentication information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No authentication information was found"
		}
		If($Text)
		{
			Line 0 "No authentication information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No authentication information was found"
		}
	}
	Else
	{
		OutputRASAuthSettings $results
	}

	Write-Verbose "$(Get-Date): `tProcessing Settings"
	
	$results = Get-RASSessionSetting -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve RAS Session settings information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve RAS Session settings information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve RAS Session settings information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve RAS Session settings information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No RAS Session settings information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No RAS Session settings information was found"
		}
		If($Text)
		{
			Line 0 "No RAS Session settings information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No RAS Session settings information was found"
		}
	}
	Else
	{
		OutputRASSessionSetting $results
	}

	Write-Verbose "$(Get-Date): `tProcessing Second level authentication"
	
	$results = Get-2FASetting -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Second level authentication information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Second level authentication information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Second level authentication information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Second level authentication information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No Second level authentication information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Second level authentication information was found"
		}
		If($Text)
		{
			Line 0 "No Second level authentication information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Second level authentication information was found"
		}
	}
	Else
	{
		Output2FASetting $results
	}

	Write-Verbose "$(Get-Date): `tProcessing Allowed devices"
	
	$results = Get-RASAllowedDevicesSetting -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Allowed devices information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Allowed devices information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Allowed devices information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Allowed devices information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No Allowed devices information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Allowed devices information was found"
		}
		If($Text)
		{
			Line 0 "No Allowed devices information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Allowed devices information was found"
		}
	}
	Else
	{
		OutputRASAllowedDevicesSetting $results
	}
}

Function OutputConnectionSectionPage
{
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Connection"
	}
	If($Text)
	{
		Line 0 "Connection"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Connection"
	}
}

Function OutputRASAuthSettings
{
	Param([object] $RASAuthSettings)
	
	Write-Verbose "$(Get-Date): `t`tOutput Authentication"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Authentication"
	}
	If($Text)
	{
		Line 1 "Authentication"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Authentication"
	}
	
	Switch ($RASAuthSettings.AuthType)
	{
		"UsernamePassword"				{$RASAuthSettingsAuthType = "Username/Password"; Break}
		"SmartCard"						{$RASAuthSettingsAuthType = "Smart Card"; Break}
		"UsernamePasswordOrSmartCard"	{$RASAuthSettingsAuthType = "Username/Password or Smart Card"; Break}
		Default							{$RASAuthSettingsAuthType = "Unable to determine AuthType: $($RASAuthSettings.AuthType)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Authentication Type"; Value = $RASAuthSettingsAuthType; }) > $Null
		If($RASAuthSettings.AllTrustedDomains)
		{
			$ScriptInformation.Add(@{Data = "All Trusted Domains"; Value = ""; }) > $Null
		}
		Else
		{
			$ScriptInformation.Add(@{Data = "Domain"; Value = $RASAuthSettings.Domain; }) > $Null
		}
		$ScriptInformation.Add(@{Data = "Use client domain if specified"; Value = $RASAuthSettings.UseClientDomain; }) > $Null
		$ScriptInformation.Add(@{Data = 'Force clients to use NetBIOS credentials'; Value = $RASAuthSettings.ForceNetBIOSCreds; }) > $Null
		$ScriptInformation.Add(@{Data = 'Replicate settings'; Value = $RASAuthSettings.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 275;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Authentication Type`t`t`t`t: " $RASAuthSettingsAuthType
		If($RASAuthSettings.AllTrustedDomains)
		{
			Line 2 "All Trusted Domains" ""
		}
		Else
		{
			Line 2 "Domain`t`t`t`t`t`t: " $RASAuthSettings.Domain
		}
		Line 2 "Use client domain if specified`t`t`t: " $RASAuthSettings.UseClientDomain
		Line 2 "Force clients to use NetBIOS credentials`t: " $RASAuthSettings.ForceNetBIOSCreds
		Line 2 "Replicate settings`t`t`t`t: " $RASAuthSettings.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Authentication Type",($Script:htmlsb),$RASAuthSettingsAuthType,$htmlwhite)
		If($RASAuthSettings.AllTrustedDomains)
		{
			$rowdata += @(,('All Trusted Domains',($Script:htmlsb),"",$htmlwhite))
		}
		Else
		{
			$rowdata += @(,('Domain',($Script:htmlsb),$RASAuthSettings.Domain,$htmlwhite))
		}
		$rowdata += @(,('Use client domain if specified',($Script:htmlsb),$RASAuthSettings.UseClientDomain.ToString(),$htmlwhite))
		$rowdata += @(,('Force clients to use NetBIOS credentials',($Script:htmlsb),$RASAuthSettings.ForceNetBIOSCreds.ToString(),$htmlwhite))
		$rowdata += @(,('Replicate settings',($Script:htmlsb),$RASAuthSettings.ReplicateSettings.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("300","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputRASSessionSetting
{
	Param([object] $RASSessionSettings)
	
	Write-Verbose "$(Get-Date): `t`tOutput Settings"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Settings"
	}
	If($Text)
	{
		Line 1 "Settings"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Settings"
	}
	
	Switch ($RASSessionSettings.RemoteIdleSessionTimeout)
	{
		60		{$RemoteIdleSessionTimeout = "1 minute"; Break}
		180		{$RemoteIdleSessionTimeout = "3 minutes"; Break}
		300		{$RemoteIdleSessionTimeout = "5 minutes"; Break}
		600		{$RemoteIdleSessionTimeout = "10 minutes"; Break}
		1800	{$RemoteIdleSessionTimeout = "30 minutes"; Break}
		3600	{$RemoteIdleSessionTimeout = "1 hour"; Break}
		5400	{$RemoteIdleSessionTimeout = "1 hour, 30 minutes"; Break}
		7200	{$RemoteIdleSessionTimeout = "2 hours"; Break}
		21600	{$RemoteIdleSessionTimeout = "6 hours"; Break}
		43200	{$RemoteIdleSessionTimeout = "12 hours"; Break}
		86400	{$RemoteIdleSessionTimeout = "1 day"; Break}
		Default	{$RemoteIdleSessionTimeout = "Unable to determine idle session timeout: $($RASSessionSettings.RemoteIdleSessionTimeout)"; Break}
	}
	
	Switch ($RASSessionSettings.LogoffIdleSessionTimeout)
	{
		0		{$LogoffIdleSessionTimeout = "Never"; Break}
		300		{$LogoffIdleSessionTimeout = "5 minutes"; Break}
		600		{$LogoffIdleSessionTimeout = "10 minutes"; Break}
		900		{$LogoffIdleSessionTimeout = "15 minutes"; Break}
		1200	{$LogoffIdleSessionTimeout = "20 minutes"; Break}
		1800	{$LogoffIdleSessionTimeout = "30 minutes"; Break}
		3600	{$LogoffIdleSessionTimeout = "1 hour"; Break}
		Default	{$LogoffIdleSessionTimeout = "Unable to determine idle session logoff: $($RASSessionSettings.LogoffIdleSessionTimeout)"; Break}
	}
	
	Switch ($RASSessionSettings.CachedSessionTimeout)
	{
		60		{$CachedSessionTimeout = "1 minute"; Break}
		180		{$CachedSessionTimeout = "3 minutes"; Break}
		300		{$CachedSessionTimeout = "5 minutes"; Break}
		600		{$CachedSessionTimeout = "10 minutes"; Break}
		1800	{$CachedSessionTimeout = "30 minutes"; Break}
		3600	{$CachedSessionTimeout = "1 hour"; Break}
		5400	{$CachedSessionTimeout = "1 hour, 30 minutes"; Break}
		7200	{$CachedSessionTimeout = "2 hours"; Break}
		21600	{$CachedSessionTimeout = "6 hours"; Break}
		43200	{$CachedSessionTimeout = "12 hours"; Break}
		86400	{$CachedSessionTimeout = "1 day"; Break}
		86400	{$CachedSessionTimeout = "30 days"; Break}
		Default	{$CachedSessionTimeout = "Unable to determine cached session timeout: $($RASSessionSettings.CachedSessionTimeout)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Declare remote session idle after"; Value = $RemoteIdleSessionTimeout; }) > $Null
		$ScriptInformation.Add(@{Data = "Automatic logoff RAS idle session after"; Value = $LogoffIdleSessionTimeout; }) > $Null
		$ScriptInformation.Add(@{Data = "Cached Session Timeout"; Value = $CachedSessionTimeout; }) > $Null
		$ScriptInformation.Add(@{Data = 'FIPS 140-2 encryption'; Value = $RASSessionSettings.FIPSMode; }) > $Null
		$ScriptInformation.Add(@{Data = 'Replicate settings'; Value = $RASSessionSettings.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 275;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Declare remote session idle after`t: " $RemoteIdleSessionTimeout
		Line 2 "Automatic logoff RAS idle session after`t: " $LogoffIdleSessionTimeout
		Line 2 "Cached Session Timeout`t`t`t: " $CachedSessionTimeout
		Line 2 "FIPS 140-2 encryption`t`t`t: " $RASSessionSettings.FIPSMode
		Line 2 "Replicate settings`t`t`t: " $RASSessionSettings.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Declare remote session idle after",($Script:htmlsb),$RemoteIdleSessionTimeout,$htmlwhite)
		$rowdata += @(,('Automatic logoff RAS idle session after',($Script:htmlsb),$LogoffIdleSessionTimeout,$htmlwhite))
		$rowdata += @(,('Cached Session Timeout',($Script:htmlsb),$CachedSessionTimeout,$htmlwhite))
		$rowdata += @(,('FIPS 140-2 encryption',($Script:htmlsb),$RASSessionSettings.FIPSMode.ToString(),$htmlwhite))
		$rowdata += @(,('Replicate settings',($Script:htmlsb),$RASSessionSettings.ReplicateSettings.ToString(),$htmlwhite))

		$msg = ""
		$columnWidths = @("300","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function Output2FASetting
{
	Param([object] $RAS2FASettings)
	
	Write-Verbose "$(Get-Date): `t`tOutput Second level authentication"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Multi-factor authentication"
	}
	If($Text)
	{
		Line 1 "Multi-factor authentication"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Multi-factor authentication"
	}

	Switch ($RAS2FASettings.Provider)
	{
		"None" 			{$RAS2FASettingsProvider = "None"; Break}
		"Deepnet" 		{$RAS2FASettingsProvider = "Deepnet"; Break}
		"SafeNet" 		{$RAS2FASettingsProvider = "SafeNet"; Break}
		"Radius" 		{$RAS2FASettingsProvider = "RADIUS"; Break}
		"AzureRadius"	{$RAS2FASettingsProvider = "Azure MFA server (RADIUS)"; Break}
		"DuoRadius" 	{$RAS2FASettingsProvider = "Duo (RADIUS)"; Break}
		"FortiRadius" 	{$RAS2FASettingsProvider = "FortiAuthenticator (RADIUS)"; Break}
		"TekRadius" 	{$RAS2FASettingsProvider = "TekRADIUS"; Break}
		"GAuthTOTP"		{$RAS2FASettingsProvider = "Google Authenticator"; Break}
		Default 		{$RAS2FASettingsProvider = "Unable to determine 2FA Provider: $($RAS2FASettings.Provider)"; Break}
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Provider settings"
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Provider"; Value = $RAS2FASettingsProvider; }) > $Null
		
		If($RAS2FASettingsProvider -ne "None")
		{
			$ScriptInformation.Add(@{Data = "Settings"; Value = ""; }) > $Null
			
			If($null -ne $RAS2FASettings.AzureRadiusSettings)
			{
				$ScriptInformation.Add(@{Data = "     Type Name"; Value = $RAS2FASettings.AzureRadiusSettings.TypeName; }) > $Null
				$ScriptInformation.Add(@{Data = "     Server"; Value = $RAS2FASettings.AzureRadiusSettings.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "     Port"; Value = $RAS2FASettings.AzureRadiusSettings.Port; }) > $Null
				$ScriptInformation.Add(@{Data = "     Timeout"; Value = $RAS2FASettings.AzureRadiusSettings.Timeout; }) > $Null
				$ScriptInformation.Add(@{Data = "     Retries"; Value = $RAS2FASettings.AzureRadiusSettings.Retries; }) > $Null
				$ScriptInformation.Add(@{Data = "     Password Encoding"; Value = $RAS2FASettings.AzureRadiusSettings.PasswordEncoding; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward username only to Radius Server"; Value = $RAS2FASettings.AzureRadiusSettings.UsernameOnly; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward the first password to Windows authentication provider"; Value = $RAS2FASettings.AzureRadiusSettings.ForwardFirstPwdToAD; }) > $Null
			}

			If($null -ne $RAS2FASettings.DuoRadiusSettings)
			{
				$ScriptInformation.Add(@{Data = "     Type Name"; Value = $RAS2FASettings.DuoRadiusSettings.TypeName; }) > $Null
				$ScriptInformation.Add(@{Data = "     Server"; Value = $RAS2FASettings.DuoRadiusSettings.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "     Port"; Value = $RAS2FASettings.DuoRadiusSettings.Port; }) > $Null
				$ScriptInformation.Add(@{Data = "     Timeout"; Value = $RAS2FASettings.DuoRadiusSettings.Timeout; }) > $Null
				$ScriptInformation.Add(@{Data = "     Retries"; Value = $RAS2FASettings.DuoRadiusSettings.Retries; }) > $Null
				$ScriptInformation.Add(@{Data = "     Password Encoding"; Value = $RAS2FASettings.DuoRadiusSettings.PasswordEncoding; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward username only to Radius Server"; Value = $RAS2FASettings.DuoRadiusSettings.UsernameOnly; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward the first password to Windows authentication provider"; Value = $RAS2FASettings.DuoRadiusSettings.ForwardFirstPwdToAD; }) > $Null
			}

			If($null -ne $RAS2FASettings.FortiRadiusSettings)
			{
				$ScriptInformation.Add(@{Data = "     Type Name"; Value = $RAS2FASettings.FortiRadiusSettings.TypeName; }) > $Null
				$ScriptInformation.Add(@{Data = "     Server"; Value = $RAS2FASettings.FortiRadiusSettings.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "     Port"; Value = $RAS2FASettings.FortiRadiusSettings.Port; }) > $Null
				$ScriptInformation.Add(@{Data = "     Timeout"; Value = $RAS2FASettings.FortiRadiusSettings.Timeout; }) > $Null
				$ScriptInformation.Add(@{Data = "     Retries"; Value = $RAS2FASettings.FortiRadiusSettings.Retries; }) > $Null
				$ScriptInformation.Add(@{Data = "     Password Encoding"; Value = $RAS2FASettings.FortiRadiusSettings.PasswordEncoding; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward username only to Radius Server"; Value = $RAS2FASettings.FortiRadiusSettings.UsernameOnly; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward the first password to Windows authentication provider"; Value = $RAS2FASettings.FortiRadiusSettings.ForwardFirstPwdToAD; }) > $Null
			}

			If($null -ne $RAS2FASettings.TekRadiusSettings)
			{
				$ScriptInformation.Add(@{Data = "     Type Name"; Value = $RAS2FASettings.TekRadiusSettings.TypeName; }) > $Null
				$ScriptInformation.Add(@{Data = "     Server"; Value = $RAS2FASettings.TekRadiusSettings.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "     Port"; Value = $RAS2FASettings.TekRadiusSettings.Port; }) > $Null
				$ScriptInformation.Add(@{Data = "     Timeout"; Value = $RAS2FASettings.TekRadiusSettings.Timeout; }) > $Null
				$ScriptInformation.Add(@{Data = "     Retries"; Value = $RAS2FASettings.TekRadiusSettings.Retries; }) > $Null
				$ScriptInformation.Add(@{Data = "     Password Encoding"; Value = $RAS2FASettings.TekRadiusSettings.PasswordEncoding; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward username only to Radius Server"; Value = $RAS2FASettings.TekRadiusSettings.UsernameOnly; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward the first password to Windows authentication provider"; Value = $RAS2FASettings.TekRadiusSettings.ForwardFirstPwdToAD; }) > $Null
			}

			If($null -ne $RAS2FASettings.RadiusSettings)
			{
				$ScriptInformation.Add(@{Data = "     Type Name"; Value = $RAS2FASettings.RadiusSettings.TypeName; }) > $Null
				$ScriptInformation.Add(@{Data = "     Server"; Value = $RAS2FASettings.RadiusSettings.Server; }) > $Null
				$ScriptInformation.Add(@{Data = "     Port"; Value = $RAS2FASettings.RadiusSettings.Port; }) > $Null
				$ScriptInformation.Add(@{Data = "     Timeout"; Value = $RAS2FASettings.RadiusSettings.Timeout; }) > $Null
				$ScriptInformation.Add(@{Data = "     Retries"; Value = $RAS2FASettings.RadiusSettings.Retries; }) > $Null
				$ScriptInformation.Add(@{Data = "     Password Encoding"; Value = $RAS2FASettings.RadiusSettings.PasswordEncoding; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward username only to Radius Server"; Value = $RAS2FASettings.RadiusSettings.UsernameOnly; }) > $Null
				$ScriptInformation.Add(@{Data = "     Forward the first password to Windows authentication provider"; Value = $RAS2FASettings.RadiusSettings.ForwardFirstPwdToAD; }) > $Null
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 275;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 3 0 "Exclusion"
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Client IP exclude list"; Value = $RAS2FASettings.ExcludeClientIPs; }) > $Null
			If($RAS2FASettings.ExcludeClientIPs)
			{
				If($RAS2FASettings.ExcludeClientIPList.Count -gt 0)
				{
					$cnt     = -1
					$MaxFrom = (($RAS2FASettings.ExcludeClientIPList.From | Measure-Object -Property length -maximum).Maximum * -1)
					$MaxTo   = (($RAS2FASettings.ExcludeClientIPList.To | Measure-Object -Property length -maximum).Maximum * -1)

					ForEach($Item in $RAS2FASettings.ExcludeClientIPList)
					{
						$cnt++
						$tmp = ("From: {0,$($MaxFrom)} To: {1,$($MaxTo)}" -f $Item.From, $Item.To)
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "     IPv4 Addresses"; Value = $tmp; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $tmp; }) > $Null
						}
					}
				}
				
				If($RAS2FASettings.ExcludeClientIPv6List.Count -gt 0)
				{
					$cnt     = -1
					$MaxFrom = (($RAS2FASettings.ExcludeClientIPv6List.From | Measure-Object -Property length -maximum).Maximum * -1)
					$MaxTo   = (($RAS2FASettings.ExcludeClientIPv6List.To | Measure-Object -Property length -maximum).Maximum * -1)

					ForEach($Item in $RAS2FASettings.ExcludeClientIPv6List)
					{
						$cnt++
						$tmp = ("From: {0,$($MaxFrom)} To: {1,$($MaxTo)}" -f $Item.From, $Item.To)
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "     IPv6 Addresses"; Value = $tmp; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $tmp; }) > $Null
						}
					}
				}
			}
			$ScriptInformation.Add(@{Data = "Client MAC exclude list"; Value = $RAS2FASettings.ExcludeClientMAC; }) > $Null
			If($RAS2FASettings.ExcludeClientMAC)
			{
				$cnt = -1
				ForEach($MAC in $RAS2FASettings.ExcludeClientMACList)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "     MAC Address"; Value = $MAC; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $MAC; }) > $Null
					}
				}
			}
			$ScriptInformation.Add(@{Data = "Connection to the following Gateway IPs"; Value = $RAS2FASettings.ExcludeClientGWIPs; }) > $Null
			If($RAS2FASettings.ExcludeClientGWIPs)
			{
				$cnt = -1
				ForEach($Server in $RAS2FASettings.ExcludeClientGWIPList)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "     Server Name"; Value = $Server; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $Server; }) > $Null
					}
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 275;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}

		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Replicate settings"; Value = $RAS2FASettings.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 275;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Provider settings"
		Line 3 "Provider: " $RAS2FASettingsProvider
		If($RAS2FASettingsProvider -ne "None")
		{
			Line 3 "Settings"
			
			If($null -ne $RAS2FASettings.AzureRadiusSettings)
			{
				Line 4 "Type Name`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.TypeName
				Line 4 "Server`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.Server
				Line 4 "Port`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.Port
				Line 4 "Timeout`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.Timeout
				Line 4 "Retries`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.Retries
				Line 4 "Password Encoding`t`t`t`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.PasswordEncoding
				Line 4 "Forward username only to Radius Server`t`t: " $RAS2FASettings.AzureRadiusSettings.UsernameOnly
				Line 4 "Forward the first password to "
				Line 4 "Windows authentication provider`t`t`t`t: " $RAS2FASettings.AzureRadiusSettings.ForwardFirstPwdToAD
				Line 0 ""
			}
			
			If($null -ne $RAS2FASettings.DuoRadiusSettings )
			{
				Line 4 "Type Name`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.TypeName
				Line 4 "Server`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.Server
				Line 4 "Port`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.Port
				Line 4 "Timeout`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.Timeout
				Line 4 "Retries`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.Retries
				Line 4 "Password Encoding`t`t`t`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.PasswordEncoding
				Line 4 "Forward username only to Radius Server`t`t: " $RAS2FASettings.DuoRadiusSettings.UsernameOnly
				Line 4 "Forward the first password to "
				Line 4 "Windows authentication provider`t`t`t`t: " $RAS2FASettings.DuoRadiusSettings.ForwardFirstPwdToAD
				Line 0 ""
			}
			
			If($null -ne $RAS2FASettings.FortiRadiusSettings)
			{
				Line 4 "Type Name`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.TypeName
				Line 4 "Server`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.Server
				Line 4 "Port`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.Port
				Line 4 "Timeout`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.Timeout
				Line 4 "Retries`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.Retries
				Line 4 "Password Encoding`t`t`t`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.PasswordEncoding
				Line 4 "Forward username only to Radius Server`t`t: " $RAS2FASettings.FortiRadiusSettings.UsernameOnly
				Line 4 "Forward the first password to "
				Line 4 "Windows authentication provider`t`t`t`t: " $RAS2FASettings.FortiRadiusSettings.ForwardFirstPwdToAD
				Line 0 ""
			}
			
			If($null -ne $RAS2FASettings.TekRadiusSettings)
			{
				Line 4 "Type Name`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.TypeName
				Line 4 "Server`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.Server
				Line 4 "Port`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.Port
				Line 4 "Timeout`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.Timeout
				Line 4 "Retries`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.Retries
				Line 4 "Password Encoding`t`t`t`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.PasswordEncoding
				Line 4 "Forward username only to Radius Server`t`t: " $RAS2FASettings.TekRadiusSettings.UsernameOnly
				Line 4 "Forward the first password to "
				Line 4 "Windows authentication provider`t`t`t`t: " $RAS2FASettings.TekRadiusSettings.ForwardFirstPwdToAD
				Line 0 ""
			}
			
			If($null -ne $RAS2FASettings.RadiusSettings )
			{
				Line 4 "Type Name`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.RadiusSettings.TypeName
				Line 4 "Server`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.RadiusSettings.Server
				Line 4 "Port`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.RadiusSettings.Port
				Line 4 "Timeout`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.RadiusSettings.Timeout
				Line 4 "Retries`t`t`t`t`t`t`t`t`t`t: " $RAS2FASettings.RadiusSettings.Retries
				Line 4 "Password Encoding`t`t`t`t`t`t`t: " $RAS2FASettings.RadiusSettings.PasswordEncoding
				Line 4 "Forward username only to Radius Server`t`t: " $RAS2FASettings.RadiusSettings.UsernameOnly
				Line 4 "Forward the first password to "
				Line 4 "Windows authentication provider`t`t`t`t: " $RAS2FASettings.RadiusSettings.ForwardFirstPwdToAD
				Line 0 ""
			}
		
			Line 2 "Exclusion"
			Line 3 "Client IP exclude list  : " $RAS2FASettings.ExcludeClientIPs
			If($RAS2FASettings.ExcludeClientIPs)
			{
				If($RAS2FASettings.ExcludeClientIPList.Count -gt 0)
				{
					$cnt     = -1
					$MaxFrom = (($RAS2FASettings.ExcludeClientIPList.From | Measure-Object -Property length -maximum).Maximum * -1)
					$MaxTo   = (($RAS2FASettings.ExcludeClientIPList.To | Measure-Object -Property length -maximum).Maximum * -1)

					ForEach($Item in $RAS2FASettings.ExcludeClientIPList)
					{
						$cnt++
						$tmp = ("From: {0,$($MaxFrom)} To: {1,$($MaxTo)}" -f $Item.From, $Item.To)
						
						If($cnt -eq 0)
						{
							Line 5 "IPv4 Addresses`t: " $tmp
						}
						Else
						{
							Line 9 "  " $tmp
						}
					}
				}
				
				If($RAS2FASettings.ExcludeClientIPv6List.Count -gt 0)
				{
					$cnt     = -1
					$MaxFrom = (($RAS2FASettings.ExcludeClientIPv6List.From | Measure-Object -Property length -maximum).Maximum * -1)
					$MaxTo   = (($RAS2FASettings.ExcludeClientIPv6List.To | Measure-Object -Property length -maximum).Maximum * -1)

					ForEach($Item in $RAS2FASettings.ExcludeClientIPv6List)
					{
						$cnt++
						$tmp = ("From: {0,$($MaxFrom)} To: {1,$($MaxTo)}" -f $Item.From, $Item.To)
						
						If($cnt -eq 0)
						{
							Line 5 "IPv6 Addresses`t: " $tmp
						}
						Else
						{
							Line 9 "  " $tmp
						}
					}
				}
			}
			Line 3 "Client MAC exclude list : " $RAS2FASettings.ExcludeClientMAC
			If($RAS2FASettings.ExcludeClientMAC)
			{
				$cnt = -1
				ForEach($MAC in $RAS2FASettings.ExcludeClientMACList)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						Line 5 "MAC Address`t`t: " $MAC
					}
					Else
					{
						Line 9 "  " $MAC
					}
				}
			}
			Line 3 "Connection to the following Gateway IPs: " $RAS2FASettings.ExcludeClientGWIPs
			If($RAS2FASettings.ExcludeClientGWIPs)
			{
				$cnt = -1
				ForEach($Server in $RAS2FASettings.ExcludeClientGWIPList)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						Line 5 "Server Name`t`t: " $Server
					}
					Else
					{
						Line 9 "  " $Server
					}
				}
			}
		}
		Line 0 ""
		Line 2 "Replicate settings: " $RAS2FASettings.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Provider",($Script:htmlsb),$RAS2FASettingsProvider,$htmlwhite)
		If($RAS2FASettingsProvider -ne "None")
		{
			If($null -ne $RAS2FASettings.AzureRadiusSettings)
			{
				$rowdata += @(,("     Type Name",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.TypeName,$htmlwhite))
				$rowdata += @(,("     Server",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.Server,$htmlwhite))
				$rowdata += @(,("     Port",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.Port,$htmlwhite))
				$rowdata += @(,("     Timeout",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.Timeout,$htmlwhite))
				$rowdata += @(,("     Retries",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.Retries,$htmlwhite))
				$rowdata += @(,("     Password Encoding",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.PasswordEncoding.ToString(),$htmlwhite))
				$rowdata += @(,("     Forward username only to Radius Server",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.UsernameOnly,$htmlwhite))
				$rowdata += @(,("     Forward the first password to Windows authentication provider",($Script:htmlsb),$RAS2FASettings.AzureRadiusSettings.ForwardFirstPwdToAD,$htmlwhite))
			}

			If($null -ne $RAS2FASettings.DuoRadiusSettings )
			{
				$rowdata += @(,("     Type Name",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.TypeName,$htmlwhite))
				$rowdata += @(,("     Server",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.Server,$htmlwhite))
				$rowdata += @(,("     Port",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.Port,$htmlwhite))
				$rowdata += @(,("     Timeout",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.Timeout,$htmlwhite))
				$rowdata += @(,("     Retries",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.Retries,$htmlwhite))
				$rowdata += @(,("     Password Encoding",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.PasswordEncoding.ToString(),$htmlwhite))
				$rowdata += @(,("     Forward username only to Radius Server",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.UsernameOnly,$htmlwhite))
				$rowdata += @(,("     Forward the first password to Windows authentication provider",($Script:htmlsb),$RAS2FASettings.DuoRadiusSettings.ForwardFirstPwdToAD,$htmlwhite))
			}

			If($null -ne $RAS2FASettings.FortiRadiusSettings)
			{
				$rowdata += @(,("     Type Name",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.TypeName,$htmlwhite))
				$rowdata += @(,("     Server",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.Server,$htmlwhite))
				$rowdata += @(,("     Port",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.Port,$htmlwhite))
				$rowdata += @(,("     Timeout",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.Timeout,$htmlwhite))
				$rowdata += @(,("     Retries",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.Retries,$htmlwhite))
				$rowdata += @(,("     Password Encoding",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.PasswordEncoding.ToString(),$htmlwhite))
				$rowdata += @(,("     Forward username only to Radius Server",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.UsernameOnly,$htmlwhite))
				$rowdata += @(,("     Forward the first password to Windows authentication provider",($Script:htmlsb),$RAS2FASettings.FortiRadiusSettings.ForwardFirstPwdToAD,$htmlwhite))
			}

			If($null -ne $RAS2FASettings.TekRadiusSettings)
			{
				$rowdata += @(,("     Type Name",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.TypeName,$htmlwhite))
				$rowdata += @(,("     Server",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.Server,$htmlwhite))
				$rowdata += @(,("     Port",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.Port,$htmlwhite))
				$rowdata += @(,("     Timeout",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.Timeout,$htmlwhite))
				$rowdata += @(,("     Retries",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.Retries,$htmlwhite))
				$rowdata += @(,("     Password Encoding",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.PasswordEncoding.ToString(),$htmlwhite))
				$rowdata += @(,("     Forward username only to Radius Server",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.UsernameOnly,$htmlwhite))
				$rowdata += @(,("     Forward the first password to Windows authentication provider",($Script:htmlsb),$RAS2FASettings.TekRadiusSettings.ForwardFirstPwdToAD,$htmlwhite))
			}

			If($null -ne $RAS2FASettings.RadiusSettings )
			{
				$rowdata += @(,("     Type Name",($Script:htmlsb),$RAS2FASettings.RadiusSettings.TypeName,$htmlwhite))
				$rowdata += @(,("     Server",($Script:htmlsb),$RAS2FASettings.RadiusSettings.Server,$htmlwhite))
				$rowdata += @(,("     Port",($Script:htmlsb),$RAS2FASettings.RadiusSettings.Port,$htmlwhite))
				$rowdata += @(,("     Timeout",($Script:htmlsb),$RAS2FASettings.RadiusSettings.Timeout,$htmlwhite))
				$rowdata += @(,("     Retries",($Script:htmlsb),$RAS2FASettings.RadiusSettings.Retries,$htmlwhite))
				$rowdata += @(,("     Password Encoding",($Script:htmlsb),$RAS2FASettings.RadiusSettings.PasswordEncoding.ToString(),$htmlwhite))
				$rowdata += @(,("     Forward username only to Radius Server",($Script:htmlsb),$RAS2FASettings.RadiusSettings.UsernameOnly,$htmlwhite))
				$rowdata += @(,("     Forward the first password to Windows authentication provider",($Script:htmlsb),$RAS2FASettings.RadiusSettings.ForwardFirstPwdToAD,$htmlwhite))
			}

			$msg = "Provider settings"
			$columnWidths = @("300","175")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""

			$rowdata = @()
			$columnHeaders = @("Client IP exclude list",($Script:htmlsb),$RAS2FASettings.ExcludeClientIPs.ToString(),$htmlwhite)
			If($RAS2FASettings.ExcludeClientIPs)
			{
				If($RAS2FASettings.ExcludeClientIPList.Count -gt 0)
				{
					$cnt     = -1
					$MaxFrom = (($RAS2FASettings.ExcludeClientIPList.From | Measure-Object -Property length -maximum).Maximum * -1)
					$MaxTo   = (($RAS2FASettings.ExcludeClientIPList.To | Measure-Object -Property length -maximum).Maximum * -1)

					ForEach($Item in $RAS2FASettings.ExcludeClientIPList)
					{
						$cnt++
						$tmp = ("From: {0,$($MaxFrom)} To: {1,$($MaxTo)}" -f $Item.From, $Item.To)
						
						If($cnt -eq 0)
						{
							$rowdata += @(,("     IPv4 Addresses",($Script:htmlsb),$tmp,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$tmp,$htmlwhite))
						}
					}
				}
				
				If($RAS2FASettings.ExcludeClientIPv6List.Count -gt 0)
				{
					$cnt     = -1
					$MaxFrom = (($RAS2FASettings.ExcludeClientIPv6List.From | Measure-Object -Property length -maximum).Maximum * -1)
					$MaxTo   = (($RAS2FASettings.ExcludeClientIPv6List.To | Measure-Object -Property length -maximum).Maximum * -1)

					ForEach($Item in $RAS2FASettings.ExcludeClientIPv6List)
					{
						$cnt++
						$tmp = ("From: {0,$($MaxFrom)} To: {1,$($MaxTo)}" -f $Item.From, $Item.To)
						
						If($cnt -eq 0)
						{
							$rowdata += @(,("     IPv6 Addresses",($Script:htmlsb),$tmp,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$tmp,$htmlwhite))
						}
					}
				}
			}
			$rowdata += @(,("Client MAC exclude list",($Script:htmlsb),$RAS2FASettings.ExcludeClientMAC.ToString(),$htmlwhite))
			If($RAS2FASettings.ExcludeClientMAC)
			{
				$cnt = -1
				ForEach($MAC in $RAS2FASettings.ExcludeClientMACList)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("     MAC Address",($Script:htmlsb),$MAC,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$MAC,$htmlwhite))
					}
				}
			}
			$rowdata += @(,("Connection to the following Gateway IPs",($Script:htmlsb),$RAS2FASettings.ExcludeClientGWIPs.ToString(),$htmlwhite))
			If($RAS2FASettings.ExcludeClientGWIPs)
			{
				$cnt = -1
				ForEach($Server in $RAS2FASettings.ExcludeClientGWIPList)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$rowdata += @(,("     Server Name",($Script:htmlsb),$Server,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("",($Script:htmlsb),$Server,$htmlwhite))
					}
				}
			}
		}
		$msg = "Exclusion"
		$columnWidths = @("300","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Replicate settings",($Script:htmlsb),$RAS2FASettings.ReplicateSettings.ToString(),$htmlwhite)

		$msg = ""
		$columnWidths = @("300","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputRASAllowedDevicesSetting
{
	Param([object] $RASAllowedDevices)
	
	Write-Verbose "$(Get-Date): `t`tOutput Allowed devices"
	
	Switch ($RASAllowedDevices.AllowClientMode)
	{
		"AllowAllClientsConnectToSystem"			
			{$RASAllowedDevicesAllowClientMode = "Allow all clients to connect to the system"; Break}
		"AllowSelectedClientsConnectToSystem"		
			{$RASAllowedDevicesAllowClientMode = "Allow only the selected clients to connect to the system"; Break}
		"AllowSelectedClientsListPublishedItems"	
			{$RASAllowedDevicesAllowClientMode = "Allow only the selected clients to list the published items"; Break}
		Default										
			{$RASAllowedDevicesAllowClientMode = "Unable to determine Allow Client Mode: $($RASAllowedDevices.AllowClientMode)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Allowed devices"
		WriteWordLine 4 0 $RASAllowedDevicesAllowClientMode
		$AllowedDevicesWordTable = @()
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientWindows
			ClientName         = "Windows client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildWindows
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientMAC
			ClientName         = "macOS client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildMAC
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientLinux
			ClientName         = "Linux client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildLinux
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientIOS
			ClientName         = "iOS client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildIOS
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientAndroid
			ClientName         = "Android client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildAndroid
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientHTML5
			ClientName         = "HTML5 client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildHTML5
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientChromeApp
			ClientName         = "Chrome client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildChromeApp
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientWebPortal
			ClientName         = "RAS Web Portal"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildWebPortal
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientWyse
			ClientName         = "Wyse client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildWyse
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientJava
			ClientName         = "Java client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildJava
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientWinPhone
			ClientName         = "Windows Phone client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildWinPhone
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClientBlackberry
			ClientName         = "Blackberry client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuildBlackberry
		}
		
		$AllowedDevicesWordTable += @{
			ClientStatus       = $RASAllowedDevices.AllowClient2XOS
			ClientName         = "2X OS client"
			ClientMinimumBuild = $RASAllowedDevices.MinBuild2XOS
		}

		$Table = AddWordTable -Hashtable $AllowedDevicesWordTable `
		-Columns ClientStatus, ClientName, ClientMinimumBuild `
		-Headers "Enabled", "Clients", "Minimum build" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 50;
		$Table.Columns.Item(2).Width = 125;
		$Table.Columns.Item(3).Width = 100;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Replicate settings"; Value = $RASAllowedDevices.ReplicateSettings; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Allowed devices"
		Line 2 $RASAllowedDevicesAllowClientMode
		Line 0 ""
		Line 2 "Enabled Clients              Minimum build"
		Line 2 "=========================================="
		#		1234567S12345678901234567890S1234
		#       False   WIndows Phone client 9999
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientWindows, "Windows client", $RASAllowedDevices.MinBuildWindows)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientMAC, "macOS client", $RASAllowedDevices.MinBuildMAC)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientLinux, "Linux client", $RASAllowedDevices.MinBuildLinux)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientIOS, "iOS client", $RASAllowedDevices.MinBuildIOS)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientAndroid, "Android client", $RASAllowedDevices.MinBuildAndroid)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientHTML5, "HTML5 client", $RASAllowedDevices.MinBuildHTML5)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientChromeApp, "Chrome client", $RASAllowedDevices.MinBuildChromeApp)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientWebPortal, "RAS Web Portal", $RASAllowedDevices.MinBuildWebPortal)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientWyse, "Wyse client", $RASAllowedDevices.MinBuildWyse)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientJava, "Java client", $RASAllowedDevices.MinBuildJava)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientWinPhone, "Windows Phone client", $RASAllowedDevices.MinBuildWinPhone)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClientBlackberry, "Blackberry client", $RASAllowedDevices.MinBuildBlackberry)
		Line 2 ( "{0,-7} {1,-20} {2,-4}" -f $RASAllowedDevices.AllowClient2XOS, "2X OS client", $RASAllowedDevices.MinBuild2XOS)
		Line 0 ""
		Line 2 "Replicate settings: " $RASAllowedDevices.ReplicateSettings
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Allowed devices"
		$rowdata = @()

		$rowdata += @(,($RASAllowedDevices.AllowClientWindows,$htmlwhite,"Windows client",$htmlwhite,$RASAllowedDevices.MinBuildWindows,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientMAC,$htmlwhite,"macOS client",$htmlwhite,$RASAllowedDevices.MinBuildMAC,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientLinux,$htmlwhite,"Linux client",$htmlwhite,$RASAllowedDevices.MinBuildLinux,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientIOS,$htmlwhite,"iOS client",$htmlwhite,$RASAllowedDevices.MinBuildIOS,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientAndroid,$htmlwhite,"Android client",$htmlwhite,$RASAllowedDevices.MinBuildAndroid,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientHTML5,$htmlwhite,"HTML5 client",$htmlwhite,$RASAllowedDevices.MinBuildHTML5,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientChromeApp,$htmlwhite,"Chrome client",$htmlwhite,$RASAllowedDevices.MinBuildChromeApp,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientWebPortal,$htmlwhite,"RAS Web Portal",$htmlwhite,$RASAllowedDevices.MinBuildWebPortal,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientWyse,$htmlwhite,"Wyse client",$htmlwhite,$RASAllowedDevices.MinBuildWyse,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientJava,$htmlwhite,"Java client",$htmlwhite,$RASAllowedDevices.MinBuildJava,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientWinPhone,$htmlwhite,"Windows Phone client",$htmlwhite,$RASAllowedDevices.MinBuildWinPhone,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClientBlackberry,$htmlwhite,"Blackberry client",$htmlwhite,$RASAllowedDevices.MinBuildBlackberry,$htmlwhite))
		$rowdata += @(,($RASAllowedDevices.AllowClient2XOS,$htmlwhite,"2X OS client",$htmlwhite,$RASAllowedDevices.MinBuild2XOS,$htmlwhite))
		
		$columnHeaders = @(
		'Enabled',($Script:htmlsb),
		'Clients',($Script:htmlsb),
		'Minimum build',($Script:htmlsb))

		$msg = $RASAllowedDevicesAllowClientMode
		$columnWidths = @("54","125","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Replicate settings",($Script:htmlsb),$RASAllowedDevices.ReplicateSettings.ToString(),$htmlwhite)

		$msg = ""
		$columnWidths = @("183","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region process administration
Function ProcessAdministration
{
	Write-Verbose "$(Get-Date): Processing Administration"
	
	OutputAdministrationSectionPage
	
	Write-Verbose "$(Get-Date): `tProcessing Accounts"
	
	$results = Get-RASAdminAccount -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve administration information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve administration information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve administration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve administration information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No administration information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No administration information was found"
		}
		If($Text)
		{
			Line 0 "No administration information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No administration information was found"
		}
	}
	Else
	{
		#make sure the results are sorted
		$results = $results | Sort-Object Name
		OutputRASAccounts $results
	}

	Write-Verbose "$(Get-Date): `tProcessing Features"
	
	$RASFeatures = Get-RASSystemSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve features information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve features information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve features information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve features information"
		}
	}
	ElseIf($? -and $null -eq $RASFeatures)
	{
		Write-Warning "
		`n`n
		`t`t
		No features information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No features information was found"
		}
		If($Text)
		{
			Line 0 "No features information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No features information was found"
		}
	}
	Else
	{
		OutputRASFeatures $RASFeatures
	}

	$RASTurbo = Get-RASTurboSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve Turbo.net information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve Turbo.net information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve Turbo.net information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve Turbo.net information"
		}
	}
	ElseIf($? -and $null -eq $RASTUrbo)
	{
		Write-Warning "
		`n`n
		`t`t
		No Turbo.net information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No Turbo.net information was found"
		}
		If($Text)
		{
			Line 0 "No Turbo.net information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No Turbo.net information was found"
		}
	}
	Else
	{
		OutputRASTurbo $RASTurbo
	}
	
	OutputRASSettings $RASFeatures
	
	$RASMailboxSettings = Get-RASMailboxSettings -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve mailbox settings information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve mailbox settings information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve mailbox settings information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve mailbox settings information"
		}
	}
	ElseIf($? -and $null -eq $RASMailboxSettings)
	{
		Write-Warning "
		`n`n
		`t`t
		No mailbox settings information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No mailbox settings information was found"
		}
		If($Text)
		{
			Line 0 "No mailbox settings information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No mailbox settings information was found"
		}
	}
	Else
	{
		OutputRASMailboxSettings $RASMailboxSettings
	}
	
	$RASNotificationHandlers = Get-RASNotification -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve notification handlers information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve notification handlers information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve notification handlers information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve notification handlers information"
		}
	}
	ElseIf($? -and $null -eq $RASMailboxSettings)
	{
		Write-Warning "
		`n`n
		`t`t
		No notification handlers information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No notification handlers information was found"
		}
		If($Text)
		{
			Line 0 "No notification handlers information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No notification handlers information was found"
		}
	}
	Else
	{
		OutputRASNotifications $RASNotificationHandlers
	}
	
	$RASNotificationScripts = Get-RASNotificationScript -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve notification scripts information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve notification scripts information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve notification scripts information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve notification scripts information"
		}
	}
	ElseIf($? -and $null -eq $RASMailboxSettings)
	{
		Write-Warning "
		`n`n
		`t`t
		No notification scripts information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No notification scripts information was found"
		}
		If($Text)
		{
			Line 0 "No notification scripts information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No notification scripts information was found"
		}
	}
	Else
	{
		OutputRASNotificationScripts $RASNotificationScripts
	}
}

Function OutputAdministrationSectionPage
{
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Administration"
	}
	If($Text)
	{
		Line 0 "Administration"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Administration"
	}
}

Function OutputRASAccounts
{
	Param([object] $RASAccounts)
	
	Write-Verbose "$(Get-Date): `t`tOutput Accounts"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Accounts"
	}
	If($Text)
	{
		Line 1 "Accounts"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Accounts"
	}
	
	ForEach($RASAccount in $RASAccounts)
	{
		Switch ($RASAccount.Permissions)
		{
			"CustomAdmin"	{$RASAccountPermissions = "Custom administration"; Break}
			"PowerAdmin"	{$RASAccountPermissions = "Power administration"; Break}
			"RootAdmin"		{$RASAccountPermissions = "Root administration"; Break}
			Default			{$RASAccountPermissions = "Unable to determine Permissions: $($RASAccount.Permissions)"; Break}
		}
		
		If($MSWord -or $PDF)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Group or user names"; Value = $RASAccount.Name; }) > $Null
			$ScriptInformation.Add(@{Data = "Type"; Value = $RASAccount.Type; }) > $Null
			$ScriptInformation.Add(@{Data = 'Permissions'; Value = $RASAccountPermissions; }) > $Null
			$ScriptInformation.Add(@{Data = 'Email'; Value = $RASAccount.Email; }) > $Null
			$ScriptInformation.Add(@{Data = 'Mobile'; Value = $RASAccount.Mobile; }) > $Null
			$ScriptInformation.Add(@{Data = 'Group'; Value = $RASAccount.GroupName; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 180;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Group or user names`t: " $RASAccount.Name
			Line 2 "Type`t`t`t: " $RASAccount.Type
			Line 2 "Permissions`t`t: " $RASAccountPermissions
			Line 2 "Email`t`t`t: " $RASAccount.Email
			Line 2 "Mobile`t`t`t: " $RASAccount.Mobile
			Line 2 "Group`t`t`t: " $RASAccount.GroupName
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Group or user names",($Script:htmlsb),$RASAccount.Name,$htmlwhite)
			$rowdata += @(,('Type',($Script:htmlsb),$RASAccount.Type,$htmlwhite))
			$rowdata += @(,('Permissions',($Script:htmlsb),$RASAccountPermissions,$htmlwhite))
			$rowdata += @(,('Email',($Script:htmlsb),$RASAccount.Email,$htmlwhite))
			$rowdata += @(,('Mobile',($Script:htmlsb),$RASAccount.Mobile,$htmlwhite))
			$rowdata += @(,('Group',($Script:htmlsb),$RASAccount.GroupName,$htmlwhite))

			$msg = ""
			$columnWidths = @("200","175")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputRASFeatures
{
	Param([object] $RASFeatures)
	
	Write-Verbose "$(Get-Date): `t`tOutput Features"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Features"
		WriteWordLine 3 0 "Helpdesk"
	}
	If($Text)
	{
		Line 1 "Features"
		Line 2 "Helpdesk"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Features"
		WriteHTMLLine 3 0 "Helpdesk"
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Enable Helpdesk functionality in Parallels Client"; Value = $RASFeatures.HelpDeskEnabled; }) > $Null
		$ScriptInformation.Add(@{Data = "Helpdesk email"; Value = $RASFeatures.HelpDeskEmail; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Enable Helpdesk functionality in Parallels Client`t: " $RASFeatures.HelpDeskEnabled
		Line 3 "Helpdesk email`t`t`t`t`t`t: " $RASFeatures.HelpDeskEmail
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Enable Helpdesk functionality in Parallels Client",($Script:htmlsb),$RASFeatures.HelpDeskEnabled.ToString(),$htmlwhite)
		$rowdata += @(,('Helpdesk email',($Script:htmlsb),$RASFeatures.HelpDeskEmail,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputRASTurbo
{
	Param([object] $RASTurbo)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Turbo.net configuration"
	}
	If($Text)
	{
		Line 1 "Turbo.net configuration"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Turbo.net configuration"
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Enable Turbo.net application publishing"; Value = $RASTurbo.Enabled.ToString(); }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Enable Turbo.net application publishing`t`t`t`t: " $RASTurbo.Enabled.ToString()
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Enable Turbo.net application publishing",($Script:htmlsb),$RASTurbo.Enabled.ToString(),$htmlwhite)

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputRASSettings
{
	Param([object] $RASFeatures)
	
	Write-Verbose "$(Get-Date): `t`tOutput Settings"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Settings"
		WriteWordLine 3 0 "Customer Experience Program"
	}
	If($Text)
	{
		Line 1 "Settings"
		Line 2 "Customer Experience Program"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Settings"
		WriteHTMLLine 3 0 "Customer Experience Program"
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Participate in the Customer Experience Program"; Value = $RASFeatures.CEPEnabled; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Participate in the Customer Experience Program`t`t: " $RASFeatures.CEPEnabled
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Participate in the Customer Experience Program",($Script:htmlsb),$RASFeatures.CEPEnabled.ToString(),$htmlwhite)

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "HTTP Proxy settings"
	}
	If($Text)
	{
		Line 2 "HTTP Proxy settings"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "HTTP Proxy settings"
	}
	
	If($MSWord -or $PDF)
	{
		If($RASFeatures.HttpProxyMode -eq "NoProxy")
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "No proxy server"; Value = ""; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 180;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		Else
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Manual HTTP proxy configuration"; Value = ""; }) > $Null
			$ScriptInformation.Add(@{Data = "Address"; Value = $RASFeatures.HttpProxyAddress; }) > $Null
			$ScriptInformation.Add(@{Data = "Port"; Value = $RASFeatures.HttpProxyPort; }) > $Null
			$ScriptInformation.Add(@{Data = "User name"; Value = $RASFeatures.HttpProxyUser; }) > $Null
			$ScriptInformation.Add(@{Data = "Password"; Value = $RASFeatures.HttpProxyPwd; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 180;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($Text)
	{
		If($RASFeatures.HttpProxyMode -eq "NoProxy")
		{
			Line 3 "No proxy server" 
			Line 0 ""
		}
		Else
		{
			Line 3 "Manual HTTP proxy configuration" 
			Line 3 "Address`t`t: " $RASFeatures.HttpProxyAddress
			Line 3 "Port`t`t: " $RASFeatures.HttpProxyPort
			Line 3 "User name`t: " $RASFeatures.HttpProxyUser
			Line 3 "Password`t: " $RASFeatures.HttpProxyPwd
			Line 0 ""
		}
	}
	If($HTML)
	{
		If($RASFeatures.HttpProxyMode -eq "NoProxy")
		{
			$rowdata = @()
			$columnHeaders = @("No proxy server",($Script:htmlsb),"",$htmlwhite)

			$msg = ""
			$columnWidths = @("200","175")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
		Else
		{
			$rowdata = @()
			$columnHeaders = @("Manual HTTP proxy configuration",($Script:htmlsb),"",$htmlwhite)
			$rowdata += @(,('Address',($Script:htmlsb),$RASFeatures.HttpProxyAddress,$htmlwhite))
			$rowdata += @(,('Port',($Script:htmlsb),$RASFeatures.HttpProxyPort.ToString(),$htmlwhite))
			$rowdata += @(,('User name',($Script:htmlsb),$RASFeatures.HttpProxyUser,$htmlwhite))
			$rowdata += @(,('Password',($Script:htmlsb),$RASFeatures.HttpProxyPwd,$htmlwhite))

			$msg = ""
			$columnWidths = @("200","175")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Miscellaneous"
	}
	If($Text)
	{
		Line 2 "Miscellaneous"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Miscellaneous"
	}
	
	Switch ($RASFeatures.ResetIdleSessionAfterMins)
	{
		0		{$RASFeaturesResetIdleSessionAfterMins = "Never"; Break}
		15		{$RASFeaturesResetIdleSessionAfterMins = "15 minutes"; Break}
		30		{$RASFeaturesResetIdleSessionAfterMins = "30 minutes"; Break}
		60		{$RASFeaturesResetIdleSessionAfterMins = "1 hour"; Break}
		180		{$RASFeaturesResetIdleSessionAfterMins = "3 hours"; Break}
		360		{$RASFeaturesResetIdleSessionAfterMins = "6 hours"; Break}
		720		{$RASFeaturesResetIdleSessionAfterMins = "12 hours"; Break}
		1440	{$RASFeaturesResetIdleSessionAfterMins = "1 day"; Break}
		Default	{$RASFeaturesResetIdleSessionAfterMins = "Unable to determine Console idle time: $($RASFeatures.ResetIdleSessionAfterMins)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Reset idle RAS Console session after"; Value = $RASFeaturesResetIdleSessionAfterMins; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Reset idle RAS Console session after`t`t`t: " $RASFeaturesResetIdleSessionAfterMins
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Reset idle RAS Console session after",($Script:htmlsb),$RASFeaturesResetIdleSessionAfterMins,$htmlwhite)

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputRASMailboxSettings
{
	Param([object] $RASMailboxSettings)
	
	Write-Verbose "$(Get-Date): `t`tOutput Mailbox"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Mailbox"
		WriteWordLine 3 0 "Mailbox configuration"
	}
	If($Text)
	{
		Line 1 "Mailbox"
		Line 2 "Mailbox configuration"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Mailbox"
		WriteHTMLLine 3 0 "Mailbox configuration"
	}
	
	Switch ($RASMailboxSettings.UseTLS)
	{
		"YesIfAvailable"	{$RASMailboxSettingsUseTLS = "Use TLS/SSL if available"; Break}
		"Yes"				{$RASMailboxSettingsUseTLS = "Use TLS/SSL"; Break}
		"No"				{$RASMailboxSettingsUseTLS = "Do not use"; Break}
		Default				{$RASMailboxSettingsUseTLS = "Unable to determine TLS/SSL setting: $($RASMailboxSettings.UseTLS)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Mail server"; Value = $RASMailboxSettings.SMTPServer; }) > $Null
		$ScriptInformation.Add(@{Data = "TLS/SSL"; Value = $RASMailboxSettingsUseTLS; }) > $Null
		$ScriptInformation.Add(@{Data = "SMTP server requires authentication"; Value = $RASMailboxSettings.RequireAuth; }) > $Null
		If($RASMailboxSettings.RequireAuth)
		{
			$ScriptInformation.Add(@{Data = "Username"; Value = $RASMailboxSettings.Username; }) > $Null
		}

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Mail server`t`t`t`t`t`t: " $RASMailboxSettings.SMTPServer
		Line 3 "TLS/SSL`t`t`t`t`t`t`t: " $RASMailboxSettingsUseTLS
		Line 3 "SMTP server requires authentication`t`t`t: " $RASMailboxSettings.RequireAuth
		If($RASMailboxSettings.RequireAuth)
		{
			Line 3 "Username`t`t`t`t`t`t: " $RASMailboxSettings.Username
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Mail server",($Script:htmlsb),$RASMailboxSettings.SMTPServer,$htmlwhite)
		$rowdata += @(,('TLS/SSL',($Script:htmlsb),$RASMailboxSettingsUseTLS,$htmlwhite))
		$rowdata += @(,('SMTP server requires authentication',($Script:htmlsb),$RASMailboxSettings.RequireAuth.ToString(),$htmlwhite))
		If($RASMailboxSettings.RequireAuth)
		{
			$rowdata += @(,('Username',($Script:htmlsb),$RASMailboxSettings.Username,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Sender information"
	}
	If($Text)
	{
		Line 2 "Sender information"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Sender information"
	}
	
	If($MSWord -or $PDF)
	{
		$ScriptInformation = New-Object System.Collections.ArrayList
		$ScriptInformation.Add(@{Data = "Email address"; Value = $RASMailboxSettings.SenderAddress; }) > $Null

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 3 "Email address`t`t`t`t`t`t: " $RASMailboxSettings.SenderAddress
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Email address",($Script:htmlsb),$RASMailboxSettings.SenderAddress,$htmlwhite)

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}

}

Function OutputRASNotifications
{
	Param([object] $RASNotificationHandlers)
	
	Write-Verbose "$(Get-Date): `t`tOutput Notifications"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Notifications"
		WriteWordLine 3 0 "Notification handlers"
	}
	If($Text)
	{
		Line 1 "Notifications"
		Line 2 "Notification handlers"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Notifications"
		WriteHTMLLine 3 0 "Notification handlers"
	}
	
	ForEach($RASNotificationHandler in $RASNotificationHandlers)
	{
		If($RASNotificationHandler.HasThreshold)
		{
			Switch ($RASNotificationHandler.Direction)
			{
				"LowersBelow"	{$RASNotificationHandlerDirection = "decreases below"; Break}
				"RisesAbove"	{$RASNotificationHandlerDirection = "rises above"; Break}
				Default			{$RASNotificationHandlerDirection = "Unable to determine notification direction: $($RASNotificationHandler.Direction)"; Break}
			}
		}
		Else
		{
			$RASNotificationHandlerDirection = ""
		}
		
		Switch ($RASNotificationHandler.Type)
		{
			"CPUEvent"
				{$RASNotificationText = "CPU utilization value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold) %"; Break}
			"MemoryEvent"
				{$RASNotificationText = "Memory utilization value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold) %"; Break}
			"ConnectedSessionEvent"
				{$RASNotificationText = "Number of connected RDSH sessions value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold)"; Break}
			"DisconnectSessionEvent"
				{$RASNotificationText = "Number of disconnected RDSH sessions value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold)"; Break}
			"RDSHConnectSessionEvent"
				{$RASNotificationText = "RDSH sessions utilization value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold) %"; Break}
			"RDSHDisconnectSessionEvent"
				{$RASNotificationText = "RDSH disconnected sessions utilization value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold)"; Break}
			"TunneledSess"
				{$RASNotificationText = "Number of gateway tunnelled sessions value $RASNotificationHandlerDirection $($RASNotificationHandler.Threshold)"; Break}
			"Agent"
				{$RASNotificationText = "RAS Agents events"; Break}
			"Licence"
				{$RASNotificationText = "Licensing events"; Break}
			"Authentication"
				{$RASNotificationText = "Authentication server events"; Break}
			"PubItem"
				{$RASNotificationText = "Published items events"; Break}
			"VDI"
				{$RASNotificationText = "VDI events"; Break}
			Default
				{$RASNotificationText = "Unable to determine event type: $($RASNotificationHandler.Type)"; Break}
		}
		
		$GracePeriod = ($RASNotificationHandler.GracePeriod / 60)
		$NotificationsInterval = ($RASNotificationHandler.Interval / 60)
		
		If($MSWord -or $PDF)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Events"; Value = $RASNotificationText; }) > $Null
			$ScriptInformation.Add(@{Data = 'Enabled'; Value = $RASNotificationHandler.Enabled.ToString(); }) > $Null
			If($RASNotificationHandler.Enabled)
			{
				$ScriptInformation.Add(@{Data = 'Send email to RAS administrators'; Value = $RASNotificationHandler.SendEmail; }) > $Null
				
				If($RASNotificationHandler.Recipients -eq "")
				{
					$ScriptInformation.Add(@{Data = 'Handler'; Value = ""; }) > $Null
				}
				Else
				{
					$tmpArray = $RASNotificationHandler.Recipients.Split(",").Split(";")
					
					$cnt = -1
					ForEach($item in $tmpArray)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = 'Handler'; Value = $item; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ''; Value = $item; }) > $Null
						}
					}
					$tmpArray = $Null
				}
				$ScriptInformation.Add(@{Data = 'Execute a notification script'; Value = $RASNotificationHandler.ExecuteScript; }) > $Null
				
				If($RASNotificationHandler.ExecuteScript)
				{
					$results = Get-RASNotificationScript -id $RASNotificationHandler.ScriptId -EA 0 4>$Null
					
					If(!($?))
					{
						$ScriptName = "Unable to retrieve script name"
					}
					Else
					{
						$ScriptName = $results.Name
					}
					
					$ScriptInformation.Add(@{Data = ''; Value = $ScriptName; }) > $Null
					$ScriptName = $Null
				}

				$ScriptInformation.Add(@{Data = 'Notification handler grace period'; Value = "$GracePeriod minutes"; }) > $Null
				If($RASNotificationHandler.EnableInterval)
				{
					$ScriptInformation.Add(@{Data = 'Notifications interval'; Value = "$NotificationsInterval minutes"; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = 'Send one notification and suspend further notifications until recovered'; Value = ""; }) > $Null
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 180;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 3 "Events`t`t`t`t`t`t`t`t: " $RASNotificationText
			Line 3 "Enabled`t`t`t`t`t`t`t`t: " $RASNotificationHandler.Enabled.ToString()
			If($RASNotificationHandler.Enabled)
			{
				Line 3 "Send email to RAS administrators`t: " $RASNotificationHandler.SendEmail
				
				If($RASNotificationHandler.Recipients -eq "")
				{
					Line 3 "Handler`t`t`t`t`t`t`t`t: " ""
				}
				Else
				{
					$tmpArray = $RASNotificationHandler.Recipients.Split(",").Split(";")
					
					$cnt = -1
					ForEach($item in $tmpArray)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							Line 3 "Handler`t`t`t`t`t`t`t`t: " $item
						}
						Else
						{
							Line 12 '  ' $item
						}
					}
					$tmpArray = $Null
				}
				Line 3 "Execute a notification script`t`t: " $RASNotificationHandler.ExecuteScript
				
				If($RASNotificationHandler.ExecuteScript)
				{
					$results = Get-RASNotificationScript -id $RASNotificationHandler.ScriptId -EA 0 4>$Null
					
					If(!($?))
					{
						$ScriptName = "Unable to retrieve script name"
					}
					Else
					{
						$ScriptName = $results.Name
					}
					
					Line 12 '  ' $ScriptName
					$ScriptName = $Null
				}

				Line 3 "Notification handler grace period`t: " "$GracePeriod minutes"
				If($RASNotificationHandler.EnableInterval)
				{
					Line 3 "Notifications interval`t`t`t`t: " "$NotificationsInterval minutes"
				}
				Else
				{
					Line 3 'Send one notification and suspend further notifications until recovered'
				}
			}
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Events",($Script:htmlsb),$RASNotificationText,$htmlwhite)
			$rowdata += @(,('Enabled',($Script:htmlsb),$RASNotificationHandler.Enabled.ToString(),$htmlwhite))
			If($RASNotificationHandler.Enabled)
			{
				$rowdata += @(,('Send email to RAS administrators',($Script:htmlsb),$RASNotificationHandler.SendEmail.ToString(),$htmlwhite))
				
				If($RASNotificationHandler.Recipients -eq "")
				{
					$rowdata += @(,('Handler',($Script:htmlsb),"",$htmlwhite))
				}
				Else
				{
					$tmpArray = $RASNotificationHandler.Recipients.Split(",").Split(";")
					
					$cnt = -1
					ForEach($item in $tmpArray)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							$rowdata += @(,('Handler',($Script:htmlsb),$item,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($Script:htmlsb),$item,$htmlwhite))
						}
					}
					$tmpArray = $Null
				}
				$rowdata += @(,('Execute a notification script',($Script:htmlsb),$RASNotificationHandler.ExecuteScript.ToString(),$htmlwhite))
				
				If($RASNotificationHandler.ExecuteScript)
				{
					$results = Get-RASNotificationScript -id $RASNotificationHandler.ScriptId -EA 0 4>$Null
					
					If(!($?))
					{
						$ScriptName = "Unable to retrieve script name"
					}
					Else
					{
						$ScriptName = $results.Name
					}
					
					$rowdata += @(,('',($Script:htmlsb),$ScriptName,$htmlwhite))
					$ScriptName = $Null
				}

				$rowdata += @(,( 'Notification handler grace period',($Script:htmlsb),"$GracePeriod minutes",$htmlwhite))
				If($RASNotificationHandler.EnableInterval)
				{
					$rowdata += @(,( 'Notifications interval',($Script:htmlsb),"$NotificationsInterval minutes",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,( 'Send one notification and suspend further notifications until recovered',($Script:htmlsb),"",$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200","175")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}

Function OutputRASNotificationScripts
{
	Param([object] $RASNotificationScripts)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Notification scripts"
	}
	If($Text)
	{
		Line 2 "Notification scripts"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Notification scripts"
	}
	
	ForEach($RASNotificationScript in $RASNotificationScripts)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation = New-Object System.Collections.ArrayList
			$ScriptInformation.Add(@{Data = "Script name"; Value = $RASNotificationScript.Name; }) > $Null
			$ScriptInformation.Add(@{Data = 'Command'; Value = $RASNotificationScript.Command; }) > $Null
			$ScriptInformation.Add(@{Data = 'Arguments'; Value = $RASNotificationScript.Arguments; }) > $Null
			$ScriptInformation.Add(@{Data = 'Initial directory'; Value = $RASNotificationScript.InitialDirectory; }) > $Null
			$ScriptInformation.Add(@{Data = 'User name'; Value = $RASNotificationScript.Username; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 180;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 3 "Script name`t`t`t: " $RASNotificationScript.Name
			Line 3 "Command`t`t`t`t: " $RASNotificationScript.Command
			Line 3 "Arguments`t`t`t: " $RASNotificationScript.Arguments
			Line 3 "Initial directory`t: " $RASNotificationScript.InitialDirectory
			Line 3 "User name`t`t`t: " $RASNotificationScript.Username
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Script name",($Script:htmlsb),$RASNotificationScript.Name,$htmlwhite)
			$rowdata += @(,('Command',($Script:htmlsb),$RASNotificationScript.Command,$htmlwhite))
			$rowdata += @(,('Arguments',($Script:htmlsb),$RASNotificationScript.Arguments,$htmlwhite))
			$rowdata += @(,('Initial directory',($Script:htmlsb),$RASNotificationScript.InitialDirectory,$htmlwhite))
			$rowdata += @(,('User name',($Script:htmlsb),$RASNotificationScript.Username,$htmlwhite))

			$msg = ""
			$columnWidths = @("200","175")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region process licensing
Function ProcessLicensing
{
	Write-Verbose "$(Get-Date): Processing Licensing"

	$results = Get-LicenseDetails -EA 0 4>$Null
	
	If(!($?))
	{
		Write-Warning "
		`n`n
		`t`t
		Unable to retrieve licensing information
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Unable to retrieve licensing information"
		}
		If($Text)
		{
			Line 0 "Unable to retrieve licensing information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Unable to retrieve licensing information"
		}
	}
	ElseIf($? -and $null -eq $results)
	{
		Write-Warning "
		`n`n
		`t`t
		No licensing information was found
		`n`n
		"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No licensing information was found"
		}
		If($Text)
		{
			Line 0 "No licensing information was found"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No licensing information was found"
		}
	}
	Else
	{
		OutputRASLicense $results
	}
}

Function OutputRASLicense
{
	Param([object] $RASLicense)
	
	Write-Verbose "$(Get-Date): `tOutput RAS License"

	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "License Details"
		$ScriptInformation = New-Object System.Collections.ArrayList
	}
	If($Text)
	{
		Line 0 "License Details"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "License Details"
		$rowdata = @()
	}

	If($MSWord -or $PDF)
	{
		$ScriptInformation.Add(@{Data = "License Type"; Value = $RASLicense.LicenseType; }) > $Null
		If(ValidObject $RASLicense LicenseKey)
		{
			$ScriptInformation.Add(@{Data = "License Key"; Value = $RASLicense.LicenseKey; }) > $Null
		}
		$ScriptInformation.Add(@{Data = ""; Value = ""; }) > $Null
		If(ValidObject $RASLicense SupportExpireDate)
		{
			$ScriptInformation.Add(@{Data = "Support Expiration Date"; Value = $RASLicense.SupportExpireDate; }) > $Null
		}
		If(ValidObject $RASLicense ExpiryDate)
		{
			$ScriptInformation.Add(@{Data = "Expiration Date"; Value = $RASLicense.ExpiryDate; }) > $Null
		}
		If(ValidObject $RASLicense LicenseFirstActive)
		{
			$ScriptInformation.Add(@{Data = "First Activation"; Value = $RASLicense.LicenseFirstActive; }) > $Null
		}
		$ScriptInformation.Add(@{Data = ""; Value = ""; }) > $Null
		$ScriptInformation.Add(@{Data = 'Maximum allowed concurrent users'; Value = $RASLicense.InstalledUsers; }) > $Null
		$ScriptInformation.Add(@{Data = 'Peak Users'; Value = $RASLicense.UsersPeak; }) > $Null
		$ScriptInformation.Add(@{Data = 'Concurrent Users'; Value = $RASLicense.UsersLicenseInfo; }) > $Null
		$ScriptInformation.Add(@{Data = ""; Value = ""; }) > $Null
		If(ValidObject $RASLicense PAUserEmail)
		{
			$ScriptInformation.Add(@{Data = 'Parallels Account user email'; Value = $RASLicense.PAUserEmail; }) > $Null
		}
		If(ValidObject $RASLicense PAUserName)
		{
			$ScriptInformation.Add(@{Data = 'Parallels Account user name'; Value = $RASLicense.PAUserName; }) > $Null
		}
		If(ValidObject $RASLicense PACompanyName)
		{
			$ScriptInformation.Add(@{Data = 'Parallels Account company'; Value = $RASLicense.PACompanyName; }) > $Null
		}

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "License Type`t`t`t: " $RASLicense.LicenseType
		If(ValidObject $RASLicense LicenseKey)
		{
			Line 1 "License Key`t`t`t: " $RASLicense.LicenseKey
		}
		Line 0 ""
		If(ValidObject $RASLicense SupportExpireDate)
		{
			Line 1 "Support Expiration Date`t`t: " $RASLicense.SupportExpireDate
		}
		If(ValidObject $RASLicense ExpiryDate)
		{
			Line 1 "Expiration Date`t`t`t: " $RASLicense.ExpiryDate
		}
		If(ValidObject $RASLicense LicenseFirstActive)
		{
			Line 1 "First Activation`t`t: " $RASLicense.LicenseFirstActive
		}
		Line 0 ""
		Line 1 "Maximum allowed concurrent users: " $RASLicense.InstalledUsers
		Line 1 "Peak Users`t`t`t: " $RASLicense.UsersPeak
		Line 1 "Concurrent Users`t`t: " $RASLicense.UsersLicenseInfo
		Line 0 ""
		If(ValidObject $RASLicense PAUserEmail)
		{
			Line 1 "Parallels Account user email`t: " $RASLicense.PAUserEmail
		}
		If(ValidObject $RASLicense PAUserName)
		{
			Line 1 "Parallels Account user name`t: " $RASLicense.PAUserName
		}
		If(ValidObject $RASLicense PACompanyName)
		{
			Line 1 "Parallels Account company`t: " $RASLicense.PACompanyName
		}
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @("License Type",($Script:htmlsb),$RASLicense.LicenseType,$htmlwhite) #
		If(ValidObject $RASLicense LicenseKey)
		{
			$rowdata += @(,("License Key",($Script:htmlsb),$RASLicense.LicenseKey,$htmlwhite))
		}
		$rowdata += @(,("",($Script:htmlsb),"",$htmlwhite))
		If(ValidObject $RASLicense SupportExpireDate)
		{
			$rowdata += @(,("Support Expiration Date",($Script:htmlsb),$RASLicense.SupportExpireDate,$htmlwhite))
		}
		$rowdata += @(,('Expiration Date',($Script:htmlsb),$RASLicense.ExpiryDate,$htmlwhite)) #
		If(ValidObject $RASLicense ExpiryDate)
		{
			$rowdata += @(,("Expiration Date",($Script:htmlsb),$RASLicense.ExpiryDate,$htmlwhite))
		}
		If(ValidObject $RASLicense LicenseFirstActive)
		{
			$rowdata += @(,("First Activation",($Script:htmlsb),$RASLicense.LicenseFirstActive,$htmlwhite))
		}
		$rowdata += @(,("",($Script:htmlsb),"",$htmlwhite))
		$rowdata += @(,('Maximum allowed concurrent users',($Script:htmlsb),$RASLicense.InstalledUsers,$htmlwhite)) #
		$rowdata += @(,('Peak Users',($Script:htmlsb),$RASLicense.UsersPeak,$htmlwhite)) #
		$rowdata += @(,('Concurrent Users',($Script:htmlsb),$RASLicense.UsersLicenseInfo,$htmlwhite)) #
		$rowdata += @(,("",($Script:htmlsb),"",$htmlwhite))
		If(ValidObject $RASLicense PAUserEmail)
		{
			$rowdata += @(,('Parallels Account user email',($Script:htmlsb),$RASLicense.PAUserEmail,$htmlwhite))
		}
		If(ValidObject $RASLicense PAUserName)
		{
			$rowdata += @(,('Parallels Account user name',($Script:htmlsb),$RASLicense.PAUserName,$htmlwhite))
		}
		If(ValidObject $RASLicense PACompanyName)
		{
			$rowdata += @(,('Parallels Account company',($Script:htmlsb),$RASLicense.PACompanyName,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("200","175")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

SetFilenames "Parallels_RAS"

ProcessFarm

ProcessLoadBalancing

ProcessPublishing

ProcessUniversalPrinting

ProcessUniversalScanning

ProcessConnection

ProcessAdministration

ProcessLicensing

#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

If(($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
{
	$AbstractTitle = "Parallels RAS Inventory Report"
	$SubjectTitle = "Parallels RAS Inventory Report"
	UpdateDocumentProperties $AbstractTitle $SubjectTitle
}

ProcessDocumentOutput "Regular"

ProcessScriptEnd
#endregion
