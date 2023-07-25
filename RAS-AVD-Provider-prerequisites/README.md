# Auto configure Azure Virtual Desktop Provider requirements in Azure for Parallels RAS

## Description

Parallels RAS provides integration with Azure Virtual Desktop (AVD).  This repostitory contains a script that configures all prerequisites in Azure in an automated way. You can [download the script here](./pras-create-avd-prereq_v1.0.ps1)

> Step 1: [Running the script](./1.runscript.md)

> Step 2: [Deploying the Parallels RAS Azure Virtual Desktop Provider](./2.deployprovider.md)

> Step 3: [Optional: confirm configured permissions and resources](./3.confirmpermissions.md)

## Summary

Below is a summary of the actions that the script performs.

-	Create the App Registration
-	Create a new Client Secret
- Create an Azure Key Vault and securely store the Client Secret in it
-	Set the Graph API permissions (user.read.all & group.read.all) for the App Registration
-	Set User Access Administrator permissions on subscription level for the App Registration
-	Add VM Reader permission on subscription to support standalone host pools & custom images
-	Create a resource group (to host AVD Control Plane components & session hosts later)
-	Add contributor permissions on the resource group for the App Registration
-	Add contributor permissions on the vNet for the App Registration (to add AVD hosts to the vNet later)
-	Output all values the admin needs to complete the AVD Provider wizard & deploy AVD & Session Hosts

## More info

This guides explains how to use a script to automate the configuration of all requirements that are needed to create an Azure Virtual Desktop provider in in Parallels Remote Application Server (RAS). For more information on managing Azure Virtual Desktop using Parallels RAS visit our [Azure Virtual Desktop Landing Page](https://www.parallels.com/products/ras/capabilities/avd/)

## License 

The scripts are MIT-licensed, so you are free to use it in your commercial setting.
