

This PowerShell script requires PowerShell 7.

This repository contains two parts.

1) An export script, to export data from Parallels Secure Workspace (PSW).

2) A compatibility check.

3) An import script, to import data into Remote Application Server (RAS).
This should be executed from a machine where the RAS PowerShell module has been installed.


# Pre-requisites

* These scripts were written for PowerShell 7. ( https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell-on-windows?view=powershell-7.6 )
* Start by deploying a RAS single farm, with a single site. Don't configure anything.
* Execute as a Microsoft Windows administrator that has administrative permissions on both the Microsoft Remote Application Server, as well as the connection brokers and session hosts.  
* Ensure communication from the machine on which the scripts are executed, is possible with the various Microsoft Windows Server Remote Desktop Session Hosts and Connection Brokers.


# In scope


* Simple migrations.



# Out of scope

* Reverse proxied web applications.
  Consider Parallels Browser Isolation as an alternative option.

* Pre-authentication (SAML and OpenID) is out of scope.  
  For SAML, some info would be missing; and the redirect URIs in RAS are different.  
  OIDC is not yet supported. Workarounds are possible, e.g. by federating identity providers.

* Single Sign-On. 
  RAS implements this in a different way.

* Branding.
  The image size is different.

* Categories.
  An application can be present in multiple categories in PSW. 
  In RAS, it's always part of one folder.

* Session recording.  
  If interested in this feature, please file a feature request for RAS.


# Not yet supported

At this time, no automatic migration is in place for:

* Feature restrictions (e.g. clipboard).
* File extensions.
