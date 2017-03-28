# VBScript Wrappers for SCCM Administrators and Application Repackagers

This ones for the System Center Configuration Manager people out there. A number of years ago, a colleague (Glenn Turner) and I wrote these scripts and put them at <https://vbscriptwrapper.codeplex.com> to streamline the import of our existing applications into SCCM. We wanted to ensure continuity with every application we created in SCCM, maintaining a standardise installation command line.

All of the scripts can run from anywhere. The script has logic built-in to determine its location, allowing the folder to be moved around without the need to re-code to have the installations still work. It is designed to look in the script folder for  the relevant MSI or EXE to install.

# Basic Setup
## Obtain the files
1. Click ***[Download Zip](https://github.com/jeffreyhunt/vbscript-wrapper/archive/master.zip)*** or on the right sidebar of this Github page.

1. Extract the zip file into a folder

## Preparing the environment
1. Create the source file directory and structure to store the installation files

	  > I tend to use the following structure for my Definitive Media Library (DML):  
	  > `\\<server>\<share>\<vendor>\<application-name>\<application-version> <application-architecture>` 
	  >  
	  > So, for Microsoft Visio Pro 2010 x86, I would extract the zip file to a folder named:  
	  > `\\server01\share01\Microsoft\Visio Pro\2010 x86`  
  
1. Navigate to the extracted files and copy the two vbs files into the source file directory:  
	* If you are installing an `MSI`, copy files in the **MSI** folder
	* If you are installing an `EXE`, copy files in the **EXE** folder
	
1. Copy the installation files into the same folder as the two vbs files

1. Rename the vbs files to represent the application you are installing:  

	> i.e. Rename _Install-**ApplicationName**.vbs_ to _Install-**MicrosoftVisioPro**.vbs_ if you are installing Microsoft Visio Pro  
	> i.e. Rename _Uninstall-**ApplicationName**.vbs_ to _Uninstall-**MicrosoftVisioPro**.vbs_ if you are uninstalling Microsoft Visio Pro  
	>
	> **note:** the version is omitted to reduce the amount of editing required when copying the directory and application/package inside SCCM for a new version
	
## Configuring the MSI scripts
### Installation script
1. Open the installation vbs, e.g. `Install-MicrosoftVisioPro.vbs`

1. change the following variables at the top of the file to match up with the new application:
	* "ApplicationShortName", e.g. `MicrosoftVisioPro`
	* "MSIVersion", e.g. `2010_x86`
	* "MSIName", e.g. `MicrosoftVisioPro`
	
### Uninstallation script
1. Open the uninstall vbs, e.g. `Uninstall-MicrosoftVisioPro.vbs`

1. change the following variables at the top of the file to match up with the new application:
	```vbs
	strApplicationShortName = "ApplicationShortName" 'application name without spaces, e.g. MicrosoftVisioPro
	strCurrentVersion = "MSIVersion" 'application version, e.g. 2010_x86
	strUninstallGUID = "GUID" 'insert application GUID, WITHOUT the curly brackets, e.g. 90150000-003B-0000-0000-0000000FF1CE
	```

## Configuring the EXE scripts
### Installation script
1. Open the installation vbs, e.g. `Install-MicrosoftVisioPro.vbs`

1. change the following variables at the top of the file to match up with the new application:
	```vbs
	strApplicationShortName = "ApplicationShortName" 'application name without spaces, e.g. MicrosoftVisioPro
	strCurrentVersion = "EXEVersion" 'application version, e.g. 2010_x86
	strInstallEXE = "EXEName" 'name of EXE file WITHOUT the .exe extension, e.g. MicrosoftVisioPro
	strCommandLineOptions = "Command-line options" 'command-line options WITHOUT the leading space, e.g. /S
	```
	
### Uninstallation script
1. Open the uninstall vbs, e.g. `Uninstall-MicrosoftVisioPro.vbs`

1. change the following variables at the top of the file to match up with the new application:
	```vbs
	strApplicationShortName = "ApplicationShortName" 'application name without spaces, e.g. MicrosoftVisioPro
	strCurrentVersion = "EXEVersion" 'application version, e.g. 2010_x86
	strInstallEXE = "EXEName" 'name of EXE file WITHOUT the .exe extension, e.g. MicrosoftVisioPro
	strCommandLineOptions = "Command-line options" 'command-line options WITHOUT the leading space, e.g. /S
	```
## Important notes

* Ensure that if you use an MST, it has the same name as the MSI. This is because the _ApplicationShortName_ reference in the VBS file is used against the MSI & MST file names.
	> e.g. if you have an MSI named MicrosoftVisioPro.msi, name your MST MicrosoftVisioPro.mst

* In order to install with MST and apply MSPs, you need to comment out line 32, and uncomment line 36.  
	> This requires a little more editing to have the MSP filenames (future edits???)

* We have commented in the scripts enough that it should be explanatory.

# Feedback

If you need any extra help, let me know. There could be some fine-tuning needed on some of the scripts (hey, we all make mistakes), but they should work out-of-the-box.

# Issue Reporting

Please raise any issues in [the vbscript-wrapper issues page](https://github.com/jeffreyhunt/vbscript-wrapper/issues).  
To raise a new issue, click here: <https://github.com/jeffreyhunt/vbscript-wrapper/issues/new>  
To see open issues, click here: <https://github.com/jeffreyhunt/vbscript-wrapper/issues?q=is:open>  

# Contributing

1. Fork it!
1. Create your feature branch: `git checkout -b my-new-feature`
1. Commit your changes: `git commit -am 'Add some feature'`
1. Push to the branch: `git push origin my-new-feature`
1. Submit a pull request :D

# Versioning

We use [SemVer](http://semver.org/) for versioning.

# Authors

* **Jeffrey Hunt**
* **Glenn Turner**

See also the list of [contributors](https://github.com/jeffreyhunt/vbscript-wrapper/contributors) who participated in this project.

# License

This project is licensed under the MIT License - see the [license.md](license.md) file for details
