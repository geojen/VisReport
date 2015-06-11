# VisReport
Visual Report Designer

This project contains the source code for a programming project from April 2008. It is uploaded here solely as an archive of past work.

This is a Windows desktop application (tested primarily on Windows XP) written in Visual Basic 6. It was designed to be a tool for fishery population dynamics modellers using one of 7 NOAA Fisheries Toolbox software products. The purpose of this tool was to compare multiple model runs using a "Consumer Reports" style or red-yellow-green "stop light" style graphic display.

This software tool is still available for download and use at the NOAA Fisheries Toolbox web site. In 2008 I wrote an overview description, which is available here: http://nft.nefsc.noaa.gov/VisRpt.html

BRIEF DESCRIPTON OF DEV DIRECTORY

-- Visual Basic Source and Project Files --

VReport.bas - Main VB source code for the application.

Utils.bas - VB code for working with a grid editing tool I developed

*.frm - contains the VB code for each of the application's forms. "frmMain.frx" is the root-level form in this multiple document interface application.

*.frx - stores binary information for each VB form

VisualReport.vbp - Visual Studio project file; stores project-level information.

ScreenCapture.bas - VB code from Microsoft to enable Screen Capture

-- Project Utilities and Misc --

CUTIL.dll - Utilities common to all Toolbox models

ScanAgepro.exe - reads in data from Toolbox model "AGEPRO"
SanAim.exe - reads in data from Toolbox model "AIM"
scanasap.exe - reads in data from Toolbox model "ASAP"
ScanAspic.exe - reads in data from Toolbox model "ASPIC"
ScanCSA.exe - reads in data from Toolbox model "CSA"
ScanSampLenWt.exe - reads in data from Toolbox model "SAGA"
ScanVPA.exe - reads in data from Toolbox model "VPA"

VRSymbols/* - images used for the application's symbol palette

-- User Files --

VisualReport.cfg - default configuration file for the Visual Report Designer application. 

VRHelp.chm - User documentation for the Visual Report Designer application.

ReadMe.txt - User readme file.

-- Package and Deployment Files --

visual_report_setup.iss - Inno Setup Script to create installation package

VisualReport.exe - The compiled executable from the VB source

warning.rtf - custom notes to display to the user during installation

Package_SystemFiles/* - Windows system files necessary for installation

Package_ExampleFiles/* - Example to include with installation