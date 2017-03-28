# VBScript Wrappers for SCCM Administrators and Application Repackagers

This ones for the System Center Configuration Manager people out there. A number of years ago, a colleague (Glenn Turner) and I wrote these scripts and put them at <https://vbscriptwrapper.codeplex.com> to streamline the import of our existing applications into SCCM. We wanted to ensure continuity with every application we created in SCCM, maintaining a standardise installation command line.

All of the scripts can run from anywhere. The script has logic built-in to determine its location, allowing the folder to be moved around without the need to re-code to have the installations still work. It is designed to look in the script folder for  the relevant MSI or EXE to install.

# Basic Setup
1. Click ***[Download Zip](https://github.com/jeffreyhunt/vbscript-wrapper/archive/master.zip)*** on the right sidebar of this Github page.
1. Extract the zip file into a folder

	  > I tend to use the following structure for my Definitive Media Library (DML):  
	  > `\\<server>\<share>\<vendor>\<application-name>\<application-version>\<application-architecture>` 
	  >  
	  > So, for Microsoft Visio Pro 2010 x86, I would extract the zip file to a folder named:  
	  > `\\server01\share01\Microsoft\Visio Pro\2010\x86`  
  
1. Copy the installation files into the same folder as the Install-ApplicationShortName.vbs file
1. Edit the vbs file name to represent the application you are installing:  

	> i.e. Rename _Install-**ApplicationName**.vbs_ to _Install-**MicrosoftVisioPro**.vbs_ if you are installing Microsoft Visio Pro  
	>
	> **note:** the version is omitted to reduce the amount of editing required when copying the directory and application/package inside SCCM for a new version
  
1. Edit the vbs file, changing the following to match up with the new application:
	* "ApplicationShortName"
	* "MSIVersion"
	* "MSIName"

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

# License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
