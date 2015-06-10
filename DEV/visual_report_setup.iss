; Visual Report Designer Version 1.6.1 (4/2/08)

[Setup]
; Name of program: Do not include the version number; use AppVerName instead
; Note: this name is also used on the first page of the installer, as the main title:
;  "Welcome to the XXX Setup Wizard"
AppName=Visual Report Designer
; Version name: should be the same (or similar to) the value of AppName,
; but it should also include the program's version number
; Note: this title is also used on the first page of the installer, as the complete title:
;  "This will install XXX on your computer."
AppVerName=Visual Report Designer Version 1.6.1
; Admin priveleges required to install some of the files
PrivilegesRequired=admin
; Default directory
DefaultDirName=C:\NFT\VisualReport
; Start menu group name
DefaultGroupName=NFT
; Adds a warnings and notes page to the installer that appears at the beginning of the install
InfoBeforeFile=warning.rtf
; Replaces default intall images with the specified images
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardSmallImageFile=compiler:WizModernSmallImage-IS.bmp
; Stuff that will be displayed in the Add/Remove Programs "support info" section
AppPublisher=NOAA, National Marine Fisheries Service
AppComments=Contact times: M-F, 8:00 AM to 5:00 PM (Eastern Time). When e-mailing a support problem please attach the problem input files.
AppContact=NFToolbox.support@noaa.gov or 508-495-2024 (Phone)
AppPublisherURL=http://www.nmfs.noaa.gov/
AppSupportURL=http://nft.nefsc.noaa.gov/
AppVersion=1.6.1
UninstallDisplayIcon={app}\VisualReport.exe
UninstallDisplayName=Visual Report Designer Version 1.6.1
; Stuff that will be displayed in the properties of the setup.exe file
VersionInfoVersion=1.6.1
VersionInfoDescription=Visual Report Designer Version 1.6.1 Installer
; AppId is stored inside uninstall log files (unins???.dat), and is checked by subsequent
; installations to determine whether it may append to a particular existing uninstall log.
AppID=VisualReport

[Dirs]
; Specifies what permissions to grant in the installation directory's ACL (access control list).
;   Specifically, it grants "Modify" permission, which allows everybody in the "Everyone group"
;   to read, execute, create, modify, and delete files in the directory and its subdirectories.
Name: "{app}"; Permissions: everyone-modify

[Files]
; [Bootstrap Files]
; Use OnlyBelowVersion to prevent installation of these 6 files on Windows Vista
Source: Package_SystemFiles\ASYCFILT.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile promptifolder
Source: Package_SystemFiles\COMCAT.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver promptifolder
Source: Package_SystemFiles\MSVBVM60.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver promptifolder
Source: Package_SystemFiles\OLEAUT32.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver promptifolder
Source: Package_SystemFiles\OLEPRO32.DLL; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver promptifolder
Source: Package_SystemFiles\STDOLE2.TLB; DestDir: {sys}; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib promptifolder
; GDI+ is needed for ChartFX. GDI+ comes standard with Windows XP, so only need to
; install on 2000 and earlier. The first number in OnlyBelowVersion represents
; the minimum Windows version (i.e., Win 95, Win 98 or Win ME) that the command will NOT be processed,
; the second number represents the minimum Windows NT version (i.e., Win NT, Win 2000 or Win XP) that
; the command will NOT be processed. "0" means there is no upper version limit.
Source: Package_SystemFiles\gdiplus.dll; DestDir: {sys}; OnlyBelowVersion: 0, 5.01; Flags: promptifolder sharedfile

; Visual Report Designer Application Files
; make sure VisualReport.exe is first so that users can be clued in
; if they are trying to install an older version.
; Remember to save version number in VB project!
Source: VisualReport.exe; DestDir: {app}; Flags: ignoreversion
; Automatically overwrite other exe's and files.

; These exe's import data from NFT executables
Source: ScanAgepro.exe; DestDir: {app}
Source: ScanAIM.exe; DestDir: {app}
Source: ScanAsap.exe; DestDir: {app}
Source: ScanAspic.exe; DestDir: {app}
Source: ScanCSA.exe; DestDir: {app}
Source: ScanSampLenWt.exe; DestDir: {app}
Source: ScanVPA.exe; DestDir: {app}
; This file is needed for recent versions of SAGA
Source: species.txt; DestDir: {app}

; Contains GUI common NFT utilities
Source: CUTIL.dll; DestDir: {app}

; Symbols used in Visual Report Designer
Source: VRsymbols\green.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\greenminus.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\greenplus.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\minus.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\plus.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q1a.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q1b.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q1c.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q1d.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q1e.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q2a.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q2b.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q2c.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q2d.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q2e.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q3a.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q3b.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q3c.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q3d.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\q3e.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\red.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\redminus.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\redplus.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\white.bmp; DestDir: {app}\VRsymbols
Source: VRsymbols\yellow.bmp; DestDir: {app}\VRsymbols

; Documentation
Source: VRHELP.chm; DestDir: {app}
Source: ReadMe.txt; DestDir: {app}

; Sample Files
Source: Package_ExampleFiles\example.log; DestDir: {app}\Example
Source: Package_ExampleFiles\example.csv; DestDir: {app}\Example
Source: Package_ExampleFiles\example.txt; DestDir: {app}\Example
Source: Package_ExampleFiles\example_rpt1.html; DestDir: {app}\Example
Source: Package_ExampleFiles\VRsymbols\q1a.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\q1b.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\q1c.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\q1d.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\q1e.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\green.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\red.bmp; DestDir: {app}\Example\VRsymbols
Source: Package_ExampleFiles\VRsymbols\yellow.bmp; DestDir: {app}\Example\VRsymbols

; Necessary DLLs and OCXs
Source: Package_SystemFiles\ChartFX.ClientServer.Core.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: Package_SystemFiles\COMDLG32.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; Note that due to a recent Windows security fix, the four files associated with Windows Help can only be
;   modified by Windows Update. They are included here for users with older computers.
;  (hh.exe, hhctrl.ocx, itircl.dll, and itss.dll)
Source: Package_SystemFiles\hh.exe; DestDir: {sys}; Flags: promptifolder sharedfile
Source: Package_SystemFiles\hhctrl.ocx; DestDir: {sys}; Flags: regserver sharedfile
;Source: Package_SystemFiles\hhctrl.ocx; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: Package_SystemFiles\itircl.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: Package_SystemFiles\itss.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: Package_SystemFiles\MSFLXGRD.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: Package_SystemFiles\MSCOMCTL.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: Package_SystemFiles\TABCTL32.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile

[Icons]
; Installs Programs menu shortcut.
Name: {group}\Visual Report Designer 1.6.1; Filename: {app}\VisualReport.exe; WorkingDir: {app}
; Installs desktop shortcut. See Tasks below.
Name: {commondesktop}\Visual Report Designer 1.6.1; Filename: {app}\VisualReport.exe; WorkingDir: {app}; Tasks: desktopicon
; installs an internet shortcut to the Toolbox web site
Name: "{app}\NOAA Fisheries Toolbox Web Site"; Filename: "http://nft.nefsc.noaa.gov/"

[Run]
; This section lists stuff you can have users launch immediately after installation.
Filename: {app}\ReadMe.txt; Description: View the README File; Flags: postinstall shellexec skipifsilent unchecked
Filename: {app}\VRHELP.chm; Description: View the Help File; Flags: postinstall shellexec skipifsilent unchecked
Filename: {app}\VisualReport.exe; Description: Launch Visual Report Designer; Flags: postinstall nowait skipifsilent unchecked

[Tasks]
; This section inserts a page that performs tasks users select with a checkbox.
Name: desktopicon; Description: Create a desktop icon; Flags: unchecked

