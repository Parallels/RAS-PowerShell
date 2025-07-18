# Custom Package creator for Parallels Client for MacOS 

## Summary

This tool allows administrators to easily create a custom Parallels MacOS client install package for mass deployment using an exported .2xc file from a template MacOS client configuration downloaded from this location [Parallels Client for MacOS](https://www.parallels.com/products/ras/download/links/).
It allows administrators to preconfigure common client settings during the deployment which is useful when deploying clients at scale. 

## Download and Extract the tool
* Download the file [Ras_With_Pkg_2XC_Search.zip](https://github.com/Parallels/RAS-PowerShell/tree/master/Tools) and extract it's contents
* It will extract to a folder called RasClientDeploy with a number of sub-folders and two bash scripts within. One called 'createpackage' and the other called 'postinstall'


## Tool contents
* The 'createpackage' script takes a standard Parallels Client for MacOS install and the custom .2xc file you export (see below) and creates a custom .pkg file to be used in mass deployment tools such as Jamf
* The 'postinstall' script is incorporated into the final custom package and runs after the Parallels Client for MacOS has been deployed and is what delivers the customisations the administrator has saved in the .2xc file

## Preparation
#### Creating a 2xc file
* install the Parallels Client for MacOS (the current tool has been tested up to Client version 20.2.1 but should also work with subsequent versions, though full testing is advised prior to mass deployment)
* Configure the client settings and connections that are needed using this client install as a template configuration.
* Navigate to File - Export Settings to export a .2xc file with those settings configured.
* If editing is needed (for example to remove the username used in testing) you can open this .2xc file with a text editor and customize
* In order to be used by this tool the .2xc file you save needs to be located in the following folder: RasClientDeploy > Scripts > Contents. The tool looks in this folder for a file ending .2xc. 
> IMPORTANT make sure there is only one .2xc file in this folder as the script will use the first .2xc file it finds


#### Download the notarized Parallels MacOS client
* Download the latest [Parallels Client for MacOS](https://www.parallels.com/products/ras/download/links/) - or otherwise use the version you wish to deploy
* In order to be used by this tool the .pkg file you download above needs to be saved in the following folder: RasClientDeploy > Scripts > Contents. The tool looks in this folder for a file ending .pkg. 
> IMPORTANT make sure there is only one .pkg file in this folder as the script will use the first .pkg file it finds

## Creating the Custom Install

* Once both the .2xc file and the .pkg install are located within the folder: RasClientDeploy > Scripts > Contents - you are ready to create the custom package
* Open a Terminal window at the root RasClientDeploy folder.
* Run the 'createpackage' script by entering the following command ./createpackage
* A file called ParallelsClient.pkg is created within the root folder of the tool
* You are now able to mass deploy your custom Parallels MacOS Client package using tools such as Jamf by adding this file ParallelsClient.pkg as a package in such tools 


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
