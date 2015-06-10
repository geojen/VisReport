NOAA Fisheries Toolbox
Visual Report Designer Version 1.6.1 (4/2/08)

1. ABOUT VISUAL REPORT DESIGNER
2. INSTALLATION
3. GETTING STARTED
4. GETTING HELP
5. VERSION CHANGES


=== 1. ABOUT VISUAL REPORT DESIGNER =====

Visual Report Designer allows you to easily import the input and output files from several NOAA Fisheries Toolbox models and display the entire set or a subset of the collected data in a report resembling a stoplight chart. You can also import data stored in text files, spreadsheets or other applications.

Visual Report Designer supports the following Toolbox models: AGEPRO, AIM, ASAP, ASPIC, CSA, and VPA / ADAPT, as well as SAGA length-weight files.

You can construct reports using several varieties of color-coded symbols. In addition to binning the data into the three categories commonly used in stoplight charts, you have the option to bin the data into two categories or into quintiles.

Many improvements and new features have been added beginning with version 1.5:
- The process of creating reports has been streamlined.
- Data can be added to a report or edited within the report in batch mode, reducing the time necessary to design or modify reports.
- All reports as well as data that have been imported into Visual Report Designer are saved automatically.
- Year ranges for the imported data as well as each report are handled automatically in the background, eliminating the need for you to set the year ranges explicitly.
- More symbol sets are available.
- Reports are more fully customizable and include integrated copy and save functions.
- Reports are saved as HTML files allowing for greater portability and can be edited outside of Visual Report Designer.
- Entering custom or user-supplied data into the User Added Data grid is easier and more intuitive.

In version 1.6, several new features have been added:
- A dispersion statistic can be printed next to each line of data in the report.
- Missing data in the Data Collection Grid are allowed.
- You can select a missing data indicator (e.g., "N/A", -9999, etc.). Any datum that exactly matches this value will be excluded from the binning calculations.
- Backups of your project are created automatically and you can restore a backed up project from the saved copies.

In version 1.6, the set of symbols used in quintiles have been improved for better clarity.


=== 2. INSTALLATION =====================

Visual Report Designer can be installed on Windows 95, Windows 98, Windows 2000, and Windows XP.

If you have a previous version installed on your computer, it is recommended that you uninstall the old version before proceeding.

To begin the installation process, double-click on the file "setup.exe". Follow the prompts in the installation dialogue. Visual Report Designer is installed at C:\NFT\VisualReport by default.

During the installation, you may encounter the message: "The existing file is newer than the one Setup is trying to install. It is recommended that you keep the existing file." This simply indicates that you have a more recent copy of a helper file needed to run the program. At the prompt, "Do you want to keep the existing file?" click "Yes" to keep the file that is on your computer.  Clicking "No" will overwrite the file with the one supplied by the installation package. It is generally recommended that you keep the file that is newer.

If you install Visual Report Designer on a remote or network drive, be aware that a recent Windows security update prevents the help file (VRHELP.chm) from being displayed correctly if opened from a remote computer. This has the effect of disabling the context sensitive help feature (see section 4. Getting Help, below, for more information on context sensitive help). However, you can still view the help file manually if you copy the file (VRHELP.chm) to your computer's hard drive and then double-click on the file to open it.

Please note that Visual Report Designer’s internal Report Viewer requires Microsoft Internet Explorer to be installed on your computer. (Most computers using Microsoft Windows come with Microsoft Internet Explorer already installed.) If so desired you may use another application other than the Report Viewer to view the reports you create, but some useful features will not be available. See the Help File for more information.


=== 3. GETTING STARTED ==================

You can start Visual Report Designer by going to the Start Menu, selecting Programs, then selecting NFT, and finally selecting Visual Report Designer.

A sample project ("example.log") is provided in the installation directory.

In addition to pop-up or free-floating forms, Visual Report Designer displays information in "windows" or panels which hide behind each other in the main display. When using Visual Report Designer, if multiple windows are open, simply go to the Window Menu (located in the top row of the Visual Report Designer application) and select the desired item to bring it to the front.


=== 4. GETTING HELP =====================

Context sensitive help when using Visual Report Designer is available by pressing the F1 key on your keyboard. Pressing F1 brings up the help pages, and in many cases goes directly to the topic relating to what is currently displayed on your screen.

You can also access the help pages from the Help menu of the Visual Report Designer application. Or you can manually open the file by navigating to the directory where Visual Report Designer is installed (typically C:\NFT\VisualReport) and double-clicking on the help file "VRHELP.chm".

For answers to Frequently Asked Questions about the NOAA Fisheries Toolbox, please visit the NOAA Fisheries Toolbox website, http://nft.nefsc.noaa.gov/, and select the Frequently Asked Questions topic.

For user support, please contact:

Alan Seaver
NOAA Fisheries
Northeast Fisheries Science Center
166 Water Street
Woods Hole, MA  02543  USA

Phone: 508-495-2024
Fax: 508-495-2393
email: NFToolbox.support@noaa.gov

Contact times are Monday through Friday,  8:00 AM to 5:00 PM (Eastern Time).

When e-mailing a problem: 
Please include the project files and which version you are using.


=== 5. VERSION CHANGES ==================

Visual Report Designer 1.6.1 (April 2008)
Replaces Visual Report Designer 1.6 (February 2008).
- Visual Report Designer now includes the ability to import AIM Version 2.0 files and ASAP Version 2.0.x files.
- Fixed a bug in the AIM import utility that prevented data files with a blank model description from being imported.

Visual Report Designer 1.6 (February 2008)

New Features:
- You have the option to print a dispersion statistic next to each line of data in the report.
- Missing data in the Data Collection Grid are allowed.
- You can select a missing data indicator (e.g., "N/A", -9999, etc.). Any datum that exactly matches this value will be excluded from the binning calculations.
- You can globally replace values in the Data Collection Grid that represent your missing value indicator with a different missing value indicator.
- ASPIC files containing missing data can be imported with a user-selected missing value indicator, or as missing values are displayed in ASPIC (any negative number for input data; zeros in output data).
- Better handling of Data Type and Item descriptors when importing ASPIC data.
- New Backup capabilities: Backups of your project are created automatically when you open an existing project and before you close the program. You can create a backup manually at any time. Up to 20 backups can be created before recycling the old backups. You can restore a backed up project from the saved copies.
- The set of symbols used in quintiles have been improved for better clarity.
- Minor enhancements to the graphic interface.
Bug Fixes:
- Fixed a bug that incorrectly changed the project's start and end years when adding user-supplied data which falls outside the existing data's year range.
- Fixed a bug that did not allow Toolbox models to be imported using the Add Data button on the Data Collection form.
- Fixed a bug that incorrectly re-loaded data into the Data Collection Grid for an existing project in cases where commas were present in any of the descriptive fields.
- Fixed a bug that didn't automatically write out the Report Layout file when the report's Page Preferences are changed, and so modifications to the page preferences may not have been saved right away.
- Fixed a bug that erroneously excluded the Starting Biomass from imported ASPIC files.

Visual Report Designer 1.5.1 (April 2007)

- Primarily a bug fix version.
- Fixed a bug that disallowed spaces or periods in the file name or directory name.
- Fixed a bug that began the data collection years at zero for new cases.

Visual Report Designer 1.5 (December 2006)

Many improvements and new features were added to version 1.5:
- The process of creating reports was streamlined.
- Data can be added to a report or edited within the report in batch mode, reducing the time necessary to design or modify reports.
- All reports as well as data that have been imported into Visual Report Designer can be saved automatically.
- Year ranges for the imported data as well as each report are handled automatically in the background, eliminating the need for you to set the year ranges explicitly.
- More symbol sets are available.
- Reports are more fully customizable and include integrated copy and save functions.
- Reports are saved as HTML files allowing for greater portability and can be edited outside of Visual Report Designer.
- Entering custom or user-supplied data into the User Added Data grid is easier and more intuitive.


=========================================
