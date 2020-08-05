# Parallels RAS - RAS Inventory by Carl Webster

## Description 
Creates a complete inventory of Parallels Remote Application Server (RAS) using Microsoft PowerShell, Word, plain text, or HTML.

The script requires at least PowerShell version 3 but runs best in version 5.

Word is NOT needed to run the script. This script will output in Text and HTML.
The default output format is HTML.

Creates an output file named RASInventory.<fileextension>.

You do NOT have to run this script on a server running RAS. This script was developed and run from a Windows 10 VM.

Word and PDF document includes a Cover Page, Table of Contents and Footer.
Includes support for the following language versions of Microsoft Word:
* Catalan
* Chinese
* Danish
* Dutch
* English
* Finnish
* French
* German
* Norwegian
* Portuguese
* Spanish
* Swedish

## Documentaion  

Check out the [RAS_Inventory_V1_ReadMe.rtf](https://github.com/Parallels/RAS-PowerShell/raw/RAS-Inventory-CWebster/RAS_Inventory_V1_ReadMe.rtf) for detailed documentation about this script.
		
## Examples 
```
    -------------------------- EXAMPLE 1 --------------------------
 
    PS C:\PSScript >.\RAS_Inventory_V1.ps1
 
    Outputs, by default to HTML.
    Prompts for credentials for the LocalHost RAS Server.
 
 
    -------------------------- EXAMPLE 2 --------------------------
 
    PS C:\>PS C:\PSScript .\RAS_Inventory_V1.ps1 -MSWord -CompanyName "Carl Webster
    Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ComputerName RAS01
 
    Will use:
        Carl Webster Consulting for the Company Name.
        Mod for the Cover Page format.
        Carl Webster for the User Name.
        RAS server named RAS01 for the ComputerName.
 
    Outputs to Microsoft Word.
    Prompts for credentials for the RAS Server RAS01.
 
 
    -------------------------- EXAMPLE 3 --------------------------
 
    PS C:\>PS C:\PSScript .\RAS_Inventory_V1.ps1 -PDF -CN "Carl Webster Consulting" -CP
    "Mod" -UN "Carl Webster"
 
    Will use:
        Carl Webster Consulting for the Company Name (alias CN).
        Mod for the Cover Page format (alias CP).
        Carl Webster for the User Name (alias UN).
 
    Outputs to PDF.
    Prompts for credentials for the LocalHost RAS Server.
 
 
    -------------------------- EXAMPLE 4 --------------------------
 
    PS C:\>PS C:\PSScript .\RAS_Inventory_V1.ps1 -CompanyName "Sherlock Holmes
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
 
 
    -------------------------- EXAMPLE 5 --------------------------
 
    PS C:\>PS C:\PSScript .\RAS_Inventory_V1.ps1 -CompanyName "Sherlock Holmes
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
 
 
    -------------------------- EXAMPLE 6 --------------------------
 
    PS C:\PSScript >.\RAS_Inventory_V1.ps1
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
 
 
    -------------------------- EXAMPLE 7 --------------------------
 
    PS C:\PSScript >.\RAS_Inventory_V1.ps1
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
 
 
    -------------------------- EXAMPLE 8 --------------------------
 
    PS C:\PSScript >.\RAS_Inventory_V1.ps1
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
 
 
    -------------------------- EXAMPLE 9 --------------------------
 
    PS C:\PSScript >.\RAS_Inventory_V1.ps1
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
 

    -------------------------- EXAMPLE 10 --------------------------
 
    PS C:\PSScript >.\RAS_Inventory_V1.ps1
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
``` 

## License 

These scripts are under [GNU General Public License v2.0](LICENSE).
