

This PowerShell script requires PowerShell 7.

This repository contains two parts.

1) An export script, to export data from Parallels Secure Workspace (PSW).

2) A compatibility check.

3) An import script, to import data into Remote Application Server (RAS).
This should be executed from a machine where the RAS PowerShell module has been installed.


# Prerequisites

* These scripts were written for PowerShell 7. ( https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell-on-windows?view=powershell-7.6 )
* Start by deploying a RAS single farm, with a single site. Don't configure anything.
* Ensure that the Parallels Remote Application Server PowerShell module has been installed on the host where you are executing these scripts.
* Execute as a Microsoft Windows administrator that has administrative permissions on both the Microsoft Remote Application Server, as well as the connection brokers and session hosts.  
* Ensure communication from the machine on which the scripts are executed, is possible with the various Microsoft Windows Server Remote Desktop Session Hosts and Connection Brokers.


# In scope


This script intends to support migrating one or more Parallels Secure Workspace domains that point to the same Microsoft Windows infrastructure. 

It should assist in migrating:

* PSW Workspace domains to RAS themes (with user filtering partially covered: group assignments).
* PSW Application Servers to RAS remote PCs.
* PSW Applications (Desktop and RemoteApp) to published items in RAS; also including the user/group assignments.

For automatically managed Let's Encrypt certificates, the import script asks you to accept
the Let's Encrypt terms. Only acceptance creates these certificates. Rejecting the terms
or entering any other response displays a warning and skips certificate migration; it does
not undo or skip the other migration changes. Configure certificates manually if you reject
the terms.


## Export behavior

List requests use `limit=1000` and follow the API's `next` links until there are no
more pages. All records are saved in the existing `results` wrapper for the
compatibility check and import scripts. Non-paginated responses retain their
original shape. Exported JSON uses a serialization depth of 10.

API token input is masked and CSRF tokens are not printed by default. To enable
visible tokens, set `$ShowSensitiveInformation = $true` at the beginning of
`1_export_from_Parallels_Secure_Workspace.ps1`. A warning then requires typing
`ACCEPT` before continuing. Password input remains masked.

HTTP API URLs remain supported, and credential authentication still logs out
other sessions.

Run the export regression tests with PowerShell 7:

```powershell
.\tests\export.tests.ps1
```

The tests mock API requests and file writes; they do not connect to a workspace
or overwrite exported data.


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
