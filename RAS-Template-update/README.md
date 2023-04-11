# Schedule Parallels RAS template updates with PowerShell

![alt text](https://github.com/Parallels/RAS-PowerShell/blob/master/RAS-Template-update/pslogo.png)

## Description 
Administrators of Parallels RAS occasionally need to update templates to address system change requirements. In [this article](https://www.parallels.com/blogs/ras/parallels-ras-template-updates/), we’re going to show just how easy it is to use PowerShell to update RAS templates and use sample template scripts. 


## Example of how a script can work 
The main steps for using a script to update your template are:

1. Update your template by placing it into maintenance mode.
2. Exit maintenance mode when your updates are complete (don’t recreate it on exit).
3. Run the script at a scheduled time to update the guest VMs automatically.

The template script follows these steps:

1. Finds any user sessions in the RDSH or VDI pool and logs them off.
2. Disables the RDSH Group. (RDSH servers are unregistered and removed from the RDSH Group. In case of VDI, VMs are shut down.)
3. Enters maintenance mode.
4. Exits maintenance mode with updated guest VMs switch.
5. Recreates Guest VMs.
6. Enables RDSH Group.
7. Parallels RAS autoscaling adds RDSH servers back into the group and registers them in the Parallels RAS Console.

## Prerequisites

* Parallels RAS Farm.
* RDSH or VDI pools with template.
* Parallels RAS PowerShell module installed.

## Direct links
* Get the [RDSH Template script](https://github.com/Parallels/RAS-PowerShell/blob/master/RAS-Template-update/schedule-template-redeploy.ps1)
* Get the [VDI Template script](https://github.com/Parallels/RAS-PowerShell/blob/master/RAS-Template-update/schedule-vdi-template-redeploy.ps1)

## Step by step instructions
Follow [this blog post](https://www.parallels.com/blogs/ras/parallels-ras-template-updates/) for detailed instructions and guidance. This code is provided as a community effort without support, use at your own risk.

## License 

These scripts are under [GNU General Public License v2.0](LICENSE).
