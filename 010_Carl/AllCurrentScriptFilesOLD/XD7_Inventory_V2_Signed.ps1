﻿#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#
.SYNOPSIS
	Creates an inventory of a Citrix XenDesktop 7.8+ Site.
.DESCRIPTION
	Creates an inventory of a Citrix XenDesktop 7.8+ Site using Microsoft PowerShell, Word,
	plain text or HTML.
	
	This Script requires at least PowerShell version 3 but runs best in version 5.

	Word is NOT needed to run the script. This script will output in Text and HTML.
	
	You do NOT have to run this script on a Controller. This script was developed and run 
	from a Windows 10 VM.
	
	You can run this script remotely using the –AdminAddress (AA) parameter.
	
	This script supports versions of XenApp/XenDesktop starting with 7.8.
	
	By default, only gives summary information for:
		Machine Catalogs
		AppDisks
		Delivery Groups
		Applications
		Application Groups
		Policies
		Logging
		Administrators
		Hosting
		StoreFront
		App-V Publishing
		AppDNA
		Zones

	The Summary information is what is shown in the top half of Citrix Studio for:
		Machine Catalogs
		AppDisks
		Delivery Groups
		Applications
		Policies
		Logging
		Administrators
		Hosting
		StoreFront

	Using the MachineCatalogs parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.
	
	Using the DeliveryGroups parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.

	Using both the MachineCatalogs and DeliveryGroups parameters can cause the report to 
	take an extremely long time to complete and generate an exceptionally long report.

	Creates an output file named after the XenDesktop 7.8+ Site.
	
	Word and PDF Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
		
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)

	Valid input is:
			Alphabet (Word 2010. Works)
			Annual (Word 2010. Doesn't work well for this report)
			Austere (Word 2010. Works)
			Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
			works in 2010 but Subtitle/Subject & Author fields need to be moved 
			after title box is moved up)
			Banded (Word 2013/2016. Works)
			Conservative (Word 2010. Works)
			Contrast (Word 2010. Works)
			Cubicles (Word 2010. Works)
			Exposure (Word 2010. Works if you like looking sideways)
			Facet (Word 2013/2016. Works)
			Filigree (Word 2013/2016. Works)
			Grid (Word 2010/2013/2016. Works in 2010)
			Integral (Word 2013/2016. Works)
			Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
			manually resized or font changed to 8 point)
			Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
			manually resized or font changed to 8 point)
			Mod (Word 2010. Works)
			Motion (Word 2010/2013/2016. Works if top date is manually changed to 
			36 point)
			Newsprint (Word 2010. Works but date is not populated)
			Perspective (Word 2010. Works)
			Pinstripes (Word 2010. Works)
			Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
			resized or font changed to 14 point)
			Retrospect (Word 2013/2016. Works)
			Semaphore (Word 2013/2016. Works)
			Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
			2010)
			Slice (Dark) (Word 2013/2016. Doesn't work)
			Slice (Light) (Word 2013/2016. Doesn't work)
			Stacks (Word 2010. Works)
			Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
			Transcend (Word 2010. Works)
			ViewMaster (Word 2013/2016. Works)
			Whisp (Word 2013/2016. Works)

	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER AdminAddress
	Specifies the address of a XenDesktop controller the PowerShell snapins will connect 
	to. 
	This can be provided as a host name or an IP address. 
	This parameter defaults to LocalHost.
	This parameter has an alias of AA.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
.PARAMETER MachineCatalogs
	Gives detailed information for all machines in all Machine Catalogs.
	
	Using the MachineCatalogs parameter can cause the report to take a very long 
	time to complete and can generate an extremely long report.
	
	Using both the MachineCatalogs and DeliveryGroups parameters can cause the 
	report to take an extremely long time to complete and generate an exceptionally 
	long report.
	
	This parameter is disabled by default.
	This parameter has an alias of MC.
.PARAMETER AppDisks
	Gives detailed information for all AppDisks.
	
	This parameter is disabled by default.
	This parameter has an alias of AD.
.PARAMETER DeliveryGroups
	Gives detailed information on all desktops in all Desktop (Delivery) Groups.
	
	Using the DeliveryGroups parameter can cause the report to take a very long 
	time to complete and can generate an extremely long report.
	
	Using both the MachineCatalogs and DeliveryGroups parameters can cause the 
	report to take an extremely long time to complete and generate an exceptionally 
	long report.
	
	This parameter is disabled by default.
	This parameter has an alias of DG.
.PARAMETER DeliveryGroupsUtilization
	Gives a chart with the delivery group utilization for the last 7 days 
	depending on the information in the database.
	
	This option is only available when the report is generated in Word and requires 
	Microsoft Excel to be locally installed.
	
	Using the DeliveryGroupsUtilization parameter causes the report to take a longer 
	time to complete and generates a longer report.
	
	This parameter is disabled by default.
	This parameter has an alias of DGU.
.PARAMETER Applications
	Gives detailed information for all applications.
	This parameter is disabled by default.
	This parameter has an alias of Apps.
.PARAMETER Policies
	Give detailed information for both Site and Citrix AD based Policies.
	
	Using the Policies parameter can cause the report to take a very long time 
	to complete and can generate an extremely long report.
	
	There are three related parameters: Policies, NoPolicies and NoADPolicies.
	
	Policies and NoPolicies are mutually exclusive and priority is given to NoPolicies.
	
	This parameter is disabled by default.
	This parameter has an alias of Pol.
.PARAMETER NoPolicies
	Excludes all Site and Citrix AD-based policy information from the output document.
	
	Using the NoPolicies parameter will cause the Policies parameter to be set to False.
	
	This parameter is disabled by default.
	This parameter has an alias of NP.
.PARAMETER NoADPolicies
	Excludes all Citrix AD-based policy information from the output document.
	Includes only Site policies created in Studio.
	
	This switch is useful in large AD environments, where there may be thousands
	of policies, to keep SYSVOL from being searched.
	
	This parameter is disabled by default.
	This parameter has an alias of NoAD.
.PARAMETER Logging
	Give the Configuration Logging report with, by default, details for the previous 
	seven days.
	This parameter is disabled by default.
	This parameter has an alias of Log.
.PARAMETER Administrators
	Give detailed information for Administrator Scopes and Roles.
	This parameter is disabled by default.
	This parameter has an alias of Admins.
.PARAMETER Hosting
	Give detailed information for Hosts, Host Connections, and Resources.
	This parameter is disabled by default.
	This parameter has an alias of Host.
.PARAMETER StartDate
	Start date for the Configuration Logging report.
	
	Format for date only is MM/DD/YYYY.
	
	Format to include a specific time range is "MM/DD/YYYY HH:MM:SS" in 24-hour format.
	The double quotes are needed.
	
	The default is today's date minus seven days.
	This parameter has an alias of SD.
.PARAMETER EndDate
	End date for the Configuration Logging report.
	
	Format for date only is MM/DD/YYYY.
	
	Format to include a specific time range is "MM/DD/YYYY HH:MM:SS" in 24-hour format.
	The double quotes are needed.
	
	The default is today's date.
	This parameter has an alias of ED.
.PARAMETER StoreFront
	Give detailed information for StoreFront.
	This parameter is disabled by default.
	This parameter has an alias of SF.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2017 at 6PM is 2017-06-01_1800.
	Output filename will be ReportName_2017-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER Hardware
	Use WMI to gather hardware information on Computer System, Disks, Processor and 
	Network Interface Cards

	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin 
	or Local Administrator).

	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.

	This parameter is disabled by default.
	This parameter has an alias of HW.
.PARAMETER Section
	Processes a specific section of the report.
	Valid options are:
		Admins (Administrators)
		AppDisks
		AppDNA
		Apps (Applications and Application Group Details)
		AppV
		Catalogs (Machine Catalogs)
		Config (Configuration)
		Controllers
		Groups (Delivery Groups)
		Hosting
		Licensing
		Logging
		Policies
		StoreFront
		Zones
		All
	This parameter defaults to All sections.
	
	Notes:
	Using Logging will force the Logging switch to True.
	Using Policies will force the Policies switch to True.
	If Policies is selected and the NoPolicies switch is used, the script will terminate.
	
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	The default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -AdminAddress DDC01
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	DDC01 for the AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -TEXT

	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -HTML

	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -MachineCatalogs
	
	Creates a report with full details for all machines in all Machine Catalogs.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -DeliveryGroups
	
	Creates a report with full details for all desktops in all Desktop (Delivery) Groups.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -DeliveryGroupsUtilization
	
	Creates a report with utilization details for all Desktop (Delivery) Groups.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -DeliveryGroups -MachineCatalogs
	
	Creates a report with full details for all machines in all Machine Catalogs and 
	all desktops in all Delivery Groups.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Applications
	
	Creates a report with full details for all applications.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Policies
	
	Creates a report with full details for Policies.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -NoPolicies
	
	Creates a report with no Policy information.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -NoADPolicies
	
	Creates a report with no Citrix AD based Policy information.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Policies -NoADPolicies
	
	Creates a report with full details on Site policies created in Studio but 
	no Citrix AD based Policy information.
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Administrators
	
	Creates a report with full details on Administrator Scopes and Roles.
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Logging -StartDate 01/01/2017 -EndDate 01/31/2017
	
	Creates a report with Configuration Logging details for the dates 01/01/2017 through 
	01/31/2017.
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Logging -StartDate "06/01/2017 10:00:00" -EndDate "06/01/2017 14:00:00"
	
	Creates a report with Configuration Logging details for the time range 
	06/01/2017 10:00:00AM through 06/01/2017 02:00:00PM.
	
	Narrowing the report down to seconds does not work. Seconds must be either 00 or 59.
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Hosting
	
	Creates a report with full details for Hosts, Host Connections, and Resources.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -StoreFront
	
	Creates a report with full details for StoreFront.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -MachineCatalogs -DeliveryGroups -Applications -Policies -Hosting -StoreFront
	
	Creates a report with full details for all:
		Machines in all Machine Catalogs
		Desktops in all Delivery Groups
		Applications
		Policies
		Hosts, Host Connections, and Resources
		StoreFront
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -MC -DG -Apps -Policies -Hosting
	
	Creates a report with full details for all:
		Machines in all Machine Catalogs
		Desktops in all Delivery Groups
		Applications
		Policies
		Hosts, Host Connections, and Resources
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\XD7_Inventory_V2.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster" -AdminAddress DDC01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		Controller named DDC01 for the AdminAddress.
.EXAMPLE
	PS C:\PSScript .\XD7_Inventory_V2.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -AddDateTime
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2017 at 6PM is 2017-06-01_1800.
	Output filename will be XD7SiteName_2017-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -PDF -AddDateTime
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2017 at 6PM is 2017-06-01_1800.
	Output filename will be XD7SiteName_2017-06-01_1800.pdf
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Hardware
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -SmtpServer mail.domain.tld -From XDAdmin@domain.tld -To ITGroup@domain.tld
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, sending to ITGroup@domain.tld.
	Script will use the default SMTP port 25 and will not use SSL.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script will use the email server smtp.office365.com on port 587 using SSL, sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Section Policies
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Processes only the Policies section of the report.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Section Groups -DG
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Processes only the Delivery Groups section of the report with Delivery Group details.
.EXAMPLE
	PS C:\PSScript > .\XD7_Inventory_V2.ps1 -Section Groups
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Processes only the Delivery Groups section of the report with no Delivery Group details.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word, PDF
	plain text or HTML document.
.NOTES
	NAME: XD7_Inventory_V2.ps1
	VERSION: 2.04
	AUTHOR: Carl Webster
	LASTEDIT: March 6, 2017
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[ValidateNotNullOrEmpty()]
	[Alias("AA")]
	[string]$AdminAddress="LocalHost",

	[parameter(Mandatory=$False)] 
	[Alias("MC")]
	[Switch]$MachineCatalogs=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("AD")]
	[Switch]$AppDisks=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("DG")]
	[Switch]$DeliveryGroups=$False,	

	[parameter(Mandatory=$False)] 
	[Alias("DGU")]
	[Switch]$DeliveryGroupsUtilization=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Apps")]
	[Switch]$Applications=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Pol")]
	[Switch]$Policies=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("NP")]
	[Switch]$NoPolicies=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("NoAD")]
	[Switch]$NoADPolicies=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Log")]
	[Switch]$Logging=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Admins")]
	[Switch]$Administrators=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("Host")]
	[Switch]$Hosting=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("SF")]
	[Switch]$StoreFront=$False,	
	
	[parameter(Mandatory=$False)] 
	[Alias("SD")]
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-7)),

	[parameter(Mandatory=$False)] 
	[Alias("ED")]
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("HW")]
	[Switch]$Hardware=$False,

	[parameter(Mandatory=$False)] 
	[string]$Section="All",
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From="",

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False
	
	)
#endregion

#region script change log	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on October 20, 2013
#started updating for version 7.8+ on April 17, 2016

# This script is based on the 1.20 script

#Version 2.04 6-Mar-2017
#	Fixed wording of more policy names that changed from 7.13 prerelease to RTM
#		URL Redirection -> Bidirectional Content Redirection
#		Allow URL Redirection -> Allow Bidirectional Content Redirection
#		Allow Client URLs -> Allowed URLs to be redirected to Client
#		Allow VDA URLs -> Allowed URLs to be redirected to VDA
#		UPM - Enable Default Exclusion List - directories -> Enable Default Exclusion List - directories
#		UPM - !ctx_localappdata!\Microsoft\Application Shortcuts -> UPM - !ctx_localappdata!\Microsoft\Windows\Application Shortcuts
#		UPM - !ctx_localappdata!\Microsoft\Burn -> UPM - !ctx_localappdata!\Microsoft\Windows\Burn

#Version 2.03 3-Mar-2017
#	Fixed bug reported by P. Ewing in Functions ConfigLogPreferences and OutputDatastores
#	Fixed wording of policy setting (thanks to Esther Barthel):
#		"Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\CD Burning" to 
#		"Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\CD Burning"

#Version 2.02 released 1-Mar-2017
#	Added Application Group details
#	Added Application Group "SingleAppPerSession" property
#	Updated help text

#Version 2.01 released 28-Feb-2017
#	Added -Dev and -ScriptInfo parameters
#	Fixed several undefined variables
#	Updated help text

# Version 2.0 released on 21-Feb-2017
#	Added "Launch in user's home zone" to Delivery Group details
#	Added AppDisks
#	Added AppDNA and the ability to process just the AppDNA section
#	Added Break statements to most of the Switch statements
#	Added Chinese language support
#	Added Configuration Logging Preferences
#		Show correct database size, not the wrong size reported in Studio
#	Added Description to Machine Catalog details
#	Added Desktop Entitlement settings to Delivery Groups that are configured to deliver desktops
#	Added new policies
#	Added RemotePC OU and Subfolder properties to RemotePC Machine Catalog details
#	Added "Restrict launches to machines with tag" to Random/Desktops only and RemotePC Delivery Groups
#	Added Summary Report page
#	Added support for VDA versions 7.8 and 7.9 (which includes 7.11/7.12/7.13)
#	Added support for XenApp/XenDesktop 7.8, 7.9, 7.11, 7.12, 7.13
#	Added to machine catalog information for RemotePC, "No. of Machines" and "Allocated Machines"
#	Brought core functions up-to-date with the other scripts
#	Fix numerous typos
#	Fixed formatting issues with HTML headings output
#	Fixed French wording for Table of Contents 2 (Thanks to David Rouquier)
#	Fixed the “No. of machines” for Machine Catalogs so it is now accurate
#	Fixed the Machine Catalog details to match what is shown in Studio
#	For Machine Catalog details, for PVS provisioned catalogs, add the PVS Server address
#	For Persistent machines with changes stored on the local disk, added the “VM copy mode”
#	For Personal vDisk catalogs, added PvD size and drive letter
#	For Random catalog types (SingleSession and MultiSession), added "Temporary memory cache size (MB)" and "Temporary disk cache size (GB)"
#	Removed unnecessary blank lines in policy value output
#	Removed snapin citrix.common.commands as it is removed in 7.13 and no cmdlets are used from that snapin
#	Updated error message for missing snapins to state requires a 7.8 or later Controller
#	Updated help text
#	Updated Machine/Desktop details to match what is shown in Studio
#	Updated the Delivery Group details section with the changes to how "Delivering" is determined for XenApp Delivery Groups
#	Updated the Delivery Group "Restart Schedule" wording to match the changes in Studio
#	Updated the Delivery Group "Restart Schedule" to include the PowerShell only setting of "Restrict to tag" for 7.12 and later
#	Updated the Delivery Group section to match all the changes made in Studio
#	Updated version checking
#		Now display running version in error messages
#	Updated version checking registry access to allow 32-bit PowerShell access to 64-bit registry
#
#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($Null -eq $MSWord)
{
	$MSWord = $False
}
If($Null -eq $PDF)
{
	$PDF = $False
}
If($Null -eq $Text)
{
	$Text = $False
}
If($Null -eq $HTML)
{
	$HTML = $False
}
If($Null -eq $MachineCatalogs)
{
	$MachineCatalogs = $False
}
If($Null -eq $AppDisks)
{
	$AppDisks = $False
}
If($Null -eq $DeliveryGroups)
{
	$DeliveryGroups = $False
}
If($Null -eq $DeliveryGroupsUtilization)
{
	$DeliveryGroupsUtilization = $False
}
If($Null -eq $Applications)
{
	$Applications = $False
}
If($Null -eq $Policies)
{
	$Policies = $False
}
If($Null -eq $NoPolicies)
{
	$NoPolicies = $False
}
If($Null -eq $NoADPolicies)
{
	$NoADPolicies = $False
}
If($Null -eq $Logging)
{
	$Logging = $False
}
If($Null -eq $Administrators)
{
	$Administrators = $False
}
If($Null -eq $Hosting)
{
	$Hosting = $False
}
If($Null -eq $StoreFront)
{
	$StoreFront = $False
}
If($Null -eq $StartDate)
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-7))
}
If($Null -eq $EndDate)
{
	$EndDate = ((Get-Date -displayhint date))
}
If($Null -eq $AddDateTime)
{
	$AddDateTime = $False
}
If($Null -eq $Hardware)
{
	$Hardware = $False
}
If($Null -eq $AdminAddress)
{
	$AdminAddress = "LocalHost"
}
If($Null -eq $Section)
{
	$Section = "All"
}
If($Null -eq $Folder)
{
	$Folder = ""
}
If($Null -eq $SmtpServer)
{
	$SmtpServer = ""
}
If($Null -eq $SmtpPort)
{
	$SmtpPort = 25
}
If($Null -eq $UseSSL)
{
	$UseSSL = $False
}
If($Null -eq $From)
{
	$From = ""
}
If($Null -eq $To)
{
	$To = ""
}
If($Null -eq $Dev)
{
	$Dev = $False
}
If($Null -eq $ScriptInfo)
{
	$ScriptInfo = $False
}

If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
}
If(!(Test-Path Variable:MachineCatalogs))
{
	$MachineCatalogs = $False
}
If(!(Test-Path Variable:AppDisks))
{
	$AppDisks = $False
}
If(!(Test-Path Variable:DeliveryGroups))
{
	$DeliveryGroups = $False
}
If(!(Test-Path Variable:DeliveryGroupsUtilization))
{
	$DeliveryGroupsUtilization = $False
}
If(!(Test-Path Variable:Applications))
{
	$Applications = $False
}
If(!(Test-Path Variable:Policies))
{
	$Policies = $False
}
If(!(Test-Path Variable:NoPolicies))
{
	$NoPolicies = $False
}
If(!(Test-Path Variable:NoADPolicies))
{
	$NoADPolicies = $False
}
If(!(Test-Path Variable:Logging))
{
	$Logging = $False
}
If(!(Test-Path Variable:Administrators))
{
	$Administrators = $False
}
If(!(Test-Path Variable:Hosting))
{
	$Hosting = $False
}
If(!(Test-Path Variable:StoreFront))
{
	$StoreFront = $False
}
If(!(Test-Path Variable:StartDate))
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-7))
}
If(!(Test-Path Variable:EndDate))
{
	$EndDate = ((Get-Date -displayhint date))
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Hardware))
{
	$Hardware = $False
}
If(!(Test-Path Variable:AdminAddress))
{
	$AdminAddress = "LocalHost"
}
If(!(Test-Path Variable:Section))
{
	$Section = "All"
}
If(!(Test-Path Variable:Folder))
{
	$Folder = ""
}
If(!(Test-Path Variable:SmtpServer))
{
	$SmtpServer = ""
}
If(!(Test-Path Variable:SmtpPort))
{
	$SmtpPort = 25
}
If(!(Test-Path Variable:UseSSL))
{
	$UseSSL = $False
}
If(!(Test-Path Variable:From))
{
	$From = ""
}
If(!(Test-Path Variable:To))
{
	$To = ""
}
If(!(Test-Path Variable:Dev))
{
	$Dev = $False
}
If(!(Test-Path Variable:ScriptInfo))
{
	$ScriptInfo = $False
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($pwd.Path)\XAXDV2InventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Null -eq $Text)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($Null -eq $HTML)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($NoPolicies)
{
	$Policies = $False
}

If($NoPolicies -and $Section -eq "Policies")
{
	#conflict
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "`n`tYou specified conflicting parameters.`n`n`tYou specified the $($Section) section but also selected NoPolicies.`n`n`tPlease change one of these options and rerun the script.`n`n
	Script cannot continue."
	Exit
}

$ValidSection = $False
Switch ($Section)
{
	"Admins"		{$ValidSection = $True; Break}
	"AppDisks"		{$ValidSection = $True; Break}
	"AppDNA"		{$ValidSection = $True; Break}
	"Apps"			{$ValidSection = $True; Break}
	"AppV"			{$ValidSection = $True; Break}
	"Catalogs"		{$ValidSection = $True; Break}
	"Config"		{$ValidSection = $True; Break}
	"Controllers"	{$ValidSection = $True; Break}
	"Groups"		{$ValidSection = $True; Break}
	"Hosting"		{$ValidSection = $True; Break}
	"Licensing"		{$ValidSection = $True; Break}
	"Logging"		{$ValidSection = $True; $Logging = $True; Break}	#force $logging true if the config logging section is specified
	"Policies"		{$ValidSection = $True; $Policies = $True; Break} #force $policies true if the policies section is specified
	"StoreFront"	{$ValidSection = $True; Break}
	"Zones"			{$ValidSection = $True; Break}
	"All"			{$ValidSection = $True; Break}
}

If($ValidSection -eq $False)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "`n`tThe Section parameter specified, $($Section), is an invalid Section option.`n`tValid options are:
	
	`t`tAdmins
	`t`tAppDisks
	`t`tAppDNA
	`t`tApps
	`t`tAppV
	`t`tCatalogs
	`t`tConfig
	`t`tControllers
	`t`tGroups
	`t`tHosting
	`t`tLicensing
	`t`tLogging
	`t`tPolicies
	`t`tStoreFront
	`t`tZones
	`t`tAll
	
	`tScript cannot continue."
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}
#endregion

#region initialize variables for Word, HTML, and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155
	[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
	
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}
Else
{
	$Script:CoName = ""
}

If($HTML)
{
    Set htmlredmask         -Option AllScope -Value "#FF0000" 4>$Null
    Set htmlcyanmask        -Option AllScope -Value "#00FFFF" 4>$Null
    Set htmlbluemask        -Option AllScope -Value "#0000FF" 4>$Null
    Set htmldarkbluemask    -Option AllScope -Value "#0000A0" 4>$Null
    Set htmllightbluemask   -Option AllScope -Value "#ADD8E6" 4>$Null
    Set htmlpurplemask      -Option AllScope -Value "#800080" 4>$Null
    Set htmlyellowmask      -Option AllScope -Value "#FFFF00" 4>$Null
    Set htmllimemask        -Option AllScope -Value "#00FF00" 4>$Null
    Set htmlmagentamask     -Option AllScope -Value "#FF00FF" 4>$Null
    Set htmlwhitemask       -Option AllScope -Value "#FFFFFF" 4>$Null
    Set htmlsilvermask      -Option AllScope -Value "#C0C0C0" 4>$Null
    Set htmlgraymask        -Option AllScope -Value "#808080" 4>$Null
    Set htmlblackmask       -Option AllScope -Value "#000000" 4>$Null
    Set htmlorangemask      -Option AllScope -Value "#FFA500" 4>$Null
    Set htmlmaroonmask      -Option AllScope -Value "#800000" 4>$Null
    Set htmlgreenmask       -Option AllScope -Value "#008000" 4>$Null
    Set htmlolivemask       -Option AllScope -Value "#808000" 4>$Null

    Set htmlbold        -Option AllScope -Value 1 4>$Null
    Set htmlitalics     -Option AllScope -Value 2 4>$Null
    Set htmlred         -Option AllScope -Value 4 4>$Null
    Set htmlcyan        -Option AllScope -Value 8 4>$Null
    Set htmlblue        -Option AllScope -Value 16 4>$Null
    Set htmldarkblue    -Option AllScope -Value 32 4>$Null
    Set htmllightblue   -Option AllScope -Value 64 4>$Null
    Set htmlpurple      -Option AllScope -Value 128 4>$Null
    Set htmlyellow      -Option AllScope -Value 256 4>$Null
    Set htmllime        -Option AllScope -Value 512 4>$Null
    Set htmlmagenta     -Option AllScope -Value 1024 4>$Null
    Set htmlwhite       -Option AllScope -Value 2048 4>$Null
    Set htmlsilver      -Option AllScope -Value 4096 4>$Null
    Set htmlgray        -Option AllScope -Value 8192 4>$Null
    Set htmlolive       -Option AllScope -Value 16384 4>$Null
    Set htmlorange      -Option AllScope -Value 32768 4>$Null
    Set htmlmaroon      -Option AllScope -Value 65536 4>$Null
    Set htmlgreen       -Option AllScope -Value 131072 4>$Null
    Set htmlblack       -Option AllScope -Value 262144 4>$Null
}

If($TEXT)
{
	$global:output = ""
}
#endregion

#region code for -hardware switch
Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteHTMLLine 4 0 "General Computer"
	}
	
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Drive(s)"
	}

	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Processor(s)"
	}

	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Nics -eq $Null ) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					OutputNicItem $Nic $ThisNic
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 " "
	}
}

Function OutputComputerItem
{
	Param([object]$Item)
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemInformation = @()
		$ItemInformation += @{Data = "Manufacturer"; Value = $Item.manufacturer; }
		$ItemInformation += @{Data = "Model"; Value = $Item.model; }
		$ItemInformation += @{Data = "Domain"; Value = $Item.domain; }
		$ItemInformation += @{Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }
		$ItemInformation += @{Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }
		$ItemInformation += @{Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 2 "Model`t`t`t`t: " $Item.model
		Line 2 "Domain`t`t`t`t: " $Item.domain
		Line 2 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlbold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlbold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"; Break}
		1	{$xDriveType = "No Root Directory"; Break}
		2	{$xDriveType = "Removable Disk"; Break}
		3	{$xDriveType = "Local Disk"; Break}
		4	{$xDriveType = "Network Drive"; Break}
		5	{$xDriveType = "Compact Disc"; Break}
		6	{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $DriveInformation = @()
		$DriveInformation += @{Data = "Caption"; Value = $Drive.caption; }
		$DriveInformation += @{Data = "Size"; Value = "$($drive.drivesize) GB"; }
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += @{Data = "File System"; Value = $Drive.filesystem; }
		}
		$DriveInformation += @{Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += @{Data = "Volume Name"; Value = $Drive.volumename; }
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation += @{Data = "Volume is Dirty"; Value = $xVolumeDirty; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		$DriveInformation += @{Data = "Drive Type"; Value = $xDriveType; }
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlbold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlbold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlbold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlbold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlbold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlbold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlbold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlbold),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break}
		2	{$xAvailability = "Unknown"; Break}
		3	{$xAvailability = "Running or Full Power"; Break}
		4	{$xAvailability = "Warning"; Break}
		5	{$xAvailability = "In Test"; Break}
		6	{$xAvailability = "Not Applicable"; Break}
		7	{$xAvailability = "Power Off"; Break}
		8	{$xAvailability = "Off Line"; Break}
		9	{$xAvailability = "Off Duty"; Break}
		10	{$xAvailability = "Degraded"; Break}
		11	{$xAvailability = "Not Installed"; Break}
		12	{$xAvailability = "Install Error"; Break}
		13	{$xAvailability = "Power Save - Unknown"; Break}
		14	{$xAvailability = "Power Save - Low Power Mode"; Break}
		15	{$xAvailability = "Power Save - Standby"; Break}
		16	{$xAvailability = "Power Cycle"; Break}
		17	{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ProcessorInformation = @()
		$ProcessorInformation += @{Data = "Name"; Value = $Processor.name; }
		$ProcessorInformation += @{Data = "Description"; Value = $Processor.description; }
		$ProcessorInformation += @{Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += @{Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += @{Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += @{Data = "Number of Cores"; Value = $Processor.numberofcores; }
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += @{Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }
		}
		$ProcessorInformation += @{Data = "Availability"; Value = $xAvailability; }
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t`t: " $processor.name
		Line 2 "Description`t`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t`t: " $xAvailability
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlbold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlbold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	
	$powerMgmt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi | where {$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

	If($? -and $Null -ne $powerMgmt)
	{
		If($powerMgmt.Enable -eq $True)
		{
			$PowerSaving = "Enabled"
		}
		Else
		{
			$PowerSaving = "Disabled"
		}
	}
	Else
	{
        $PowerSaving = "N/A"
	}
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break}
		2	{$xAvailability = "Unknown"; Break}
		3	{$xAvailability = "Running or Full Power"; Break}
		4	{$xAvailability = "Warning"; Break}
		5	{$xAvailability = "In Test"; Break}
		6	{$xAvailability = "Not Applicable"; Break}
		7	{$xAvailability = "Power Off"; Break}
		8	{$xAvailability = "Off Line"; Break}
		9	{$xAvailability = "Off Duty"; Break}
		10	{$xAvailability = "Degraded"; Break}
		11	{$xAvailability = "Not Installed"; Break}
		12	{$xAvailability = "Install Error"; Break}
		13	{$xAvailability = "Power Save - Unknown"; Break}
		14	{$xAvailability = "Power Save - Low Power Mode"; Break}
		15	{$xAvailability = "Power Save - Standby"; Break}
		16	{$xAvailability = "Power Cycle"; Break}
		17	{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $NicInformation = @()
		$NicInformation += @{Data = "Name"; Value = $ThisNic.Name; }
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation += @{Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		If(validObject $Nic Manufacturer)
		{
			$NicInformation += @{Data = "Manufacturer"; Value = $Nic.manufacturer; }
		}
		$NicInformation += @{Data = "Availability"; Value = $xAvailability; }
		$NicInformation += @{Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }
		$NicInformation += @{Data = "Physical Address"; Value = $Nic.macaddress; }
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation += @{Data = "IP Address"; Value = $xIPAddress[0]; }
			$NicInformation += @{Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{Data = "Subnet Mask"; Value = $xIPSubnet[0]; }
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{Data = "IP Address"; Value = $tmp; }
					$NicInformation += @{Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }
				}
			}
		}
		Else
		{
			$NicInformation += @{Data = "IP Address"; Value = $xIPAddress; }
			$NicInformation += @{Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{Data = "Subnet Mask"; Value = $xIPSubnet; }
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += @{Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }
			$NicInformation += @{Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }
			$NicInformation += @{Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }
			$NicInformation += @{Data = "DHCP Server"; Value = $Nic.dhcpserver; }
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += @{Data = "DNS Domain"; Value = $Nic.dnsdomain; }
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation += @{Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation += @{Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }
		$NicInformation += @{Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += @{Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += @{Data = "Primary Server"; Value = $Nic.winsprimaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += @{Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += @{Data = "Scope ID"; Value = $Nic.winsscopeid; }
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		If(validObject $Nic Manufacturer)
		{
			Line 2 "Manufacturer`t`t: " $Nic.manufacturer
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Allow computer to turn "
		Line 2 "off device to save power: " $PowerSaving
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t: " $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t: " $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlbold),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlbold),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($htmlsilver -bor $htmlbold),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlbold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlbold),$Nic.Defaultipgateway[0],$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlbold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlbold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlbold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlbold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlbold),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlbold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlbold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlbold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlbold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlbold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlbold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlbold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlbold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlbold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region GetComputerServices
Function GetComputerServices 
{
	Param([string]$RemoteComputerName)
	
	#Get Computer services info
	Write-Verbose "$(Get-Date): `t`tProcessing Computer services information"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "Services"
	}
	ElseIf($Text)
	{
		Line 0 "Services"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Services"
	}

	Try
	{
		#Iain Brighton optimization 5-Jun-2014
		#Replaced with a single call to retrieve services via WMI. The repeated
		## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
		## If we need to retrieve the StartUp type might as well just use WMI.
		$Services = Get-WMIObject Win32_Service -ComputerName $RemoteComputerName | Sort DisplayName
	}
	
	Catch
	{
		$Services = $Null
	}
	
	If($? -and $Null -ne $Services)
	{
		If($Services -is [array])
		{
			[int]$NumServices = $Services.count
		}
		Else
		{
			[int]$NumServices = 1
		}
		Write-Verbose "$(Get-Date): `t`t$($NumServices) Services found"

		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Services ($NumServices Services found)"

			[System.Collections.Hashtable[]] $ServicesWordTable = @();
			## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
			[System.Collections.Hashtable[]] $HighlightedCells = @();
			## Seed the $Services row index from the second row
			[int] $CurrentServiceIndex = 2;
		}
		ElseIf($Text)
		{
			Line 0 "Services ($NumServices Services found)"
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "Services ($NumServices Services found)"
			$rowdata = @()
		}

		ForEach($Service in $Services) 
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{

				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ 
				DisplayName = $Service.DisplayName; 
				Status = $Service.State; 
				StartMode = $Service.StartMode
				}

				## Add the hash to the array
				$ServicesWordTable += $WordTableRowHash;

				## Store "to highlight" cell references
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
				}
				$CurrentServiceIndex++;
			}
			ElseIf($Text)
			{
				Line 0 "Display Name`t: " $Service.DisplayName
				Line 0 "Status`t`t: " $Service.State
				Line 0 "Start Mode`t: " $Service.StartMode
				Line 0 ""
			}
			ElseIf($HTML)
			{
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					$HighlightedCells = $htmlred
				}
				Else
				{
					$HighlightedCells = $htmlwhite
				} 
				$rowdata += @(,($Service.DisplayName,$htmlwhite,
								$Service.State,$HighlightedCells,
								$Service.StartMode,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			$Table = AddWordTable -Hashtable $ServicesWordTable `
			-Columns DisplayName, Status, StartMode `
			-Headers "Display Name", "Status", "Startup Type" `
			-AutoFit $wdAutoFitContent;

			## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
			## IB - Set the required highlighted cells
			SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

			#indent the entire table 1 tab stop
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			#nothing to do
		}
		ElseIf($HTML)
		{
			$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "No services were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
			WriteWordLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
			WriteWordLine 0 1 "script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Services were retrieved"
			Line 1 "If this is a trusted Forest, you may need to rerun the"
			Line 1 "script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $htmlbold
			WriteHTMLLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $htmlbold
			WriteHTMLLine 0 1 "script with Domain Admin credentials from the trusted Forest." "" $Null 0 $htmlbold
		}
	}
	Else
	{
		Write-Warning "Services retrieval was successful but no services were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Services retrieval was successful but no services were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $htmlbold
		}
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
#			'fr-'	{ 'Sommaire Automatique 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:Title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $Script:CoName"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}
#endregion

#region Word, text, and HTML line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlbold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlbold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not give you
	a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$options=$htmlblack)


	#Build output style
	[string]$output = ""
	[string]$HTMLStyle1 = ""
	[string]$HTMLStyle2 = ""
	
	If([String]::IsNullOrEmpty($Name))	
	{
		#$HTMLBody = "<p></p>"
		$HTMLBody = ""
	}
	Else
	{
		$color = CheckHTMLColor $options

		#build # of tabs

		While($tabs -gt 0)
		{ 
			$output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
		}

		$HTMLFontName = $fontName		

		$HTMLBody = ""

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "<i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "<b>"
		} 

		#output the rest of the parameters.
		$output += $name + $value

		Switch ($style)
		{
			1 {$HTMLStyle1 = "<h1>"; Break}
			2 {$HTMLStyle1 = "<h2>"; Break}
			3 {$HTMLStyle1 = "<h3>"; Break}
			4 {$HTMLStyle1 = "<h4>"; Break}
			Default {$HTMLStyle1 = ""; Break}
		}

		Switch ($style)
		{
			1 {$HTMLStyle2 = "</h1>"; Break}
			2 {$HTMLStyle2 = "</h2>"; Break}
			3 {$HTMLStyle2 = "</h3>"; Break}
			4 {$HTMLStyle2 = "</h4>"; Break}
			Default {$HTMLStyle2 = ""; Break}
		}

		#added by webster 12-oct-2016
		#if a heading, don't add the <br>
		If($HTMLStyle1 -eq "")
		{
			$HTMLBody += "<br><font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		Else
		{
			$HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		
		$HTMLBody += $HTMLStyle1 + $output

		$HTMLBody += $HTMLStyle2 +  "</font>"

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "</i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "</b>"
		} 
	}
	
	#added by webster 12-oct-2016
	#if a heading, don't add the <br />
	#If($HTMLStyle1 -eq "")
	#{
	#	$HTMLBody += "<br />"
	#}

	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$colCount=0,
	[int]$rowCount=0,
	[object[]]$rowInfo=@(),
	[object[]]$fixedInfo=@())

	For($rowidx = $RowIndex;$rowidx -le $rowCount;$rowidx++)
	{
		$rd = @($rowInfo[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $colCount; $columnindex+=2)
		{
			$fontitalics = $False
			$fontbold = $false
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			If($fixedInfo.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedInfo[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $rd[$columnIndex])
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border='0')
.PARAMETER noHeadCols
	This parameter should be used when generating tables without column headers
	Set this parameter equal to the number of columns in the table
.PARAMETER rowArray
	This parameter contains the row data array for the table
.PARAMETER columnArray
	This parameter contains column header data for the table
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $columnWidths = @("100px","110px","120px","130px","140px")

.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' parameter is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param([string]$tableheader,
	[string]$tablewidth="auto",
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[switch]$noBorder=$false,
	[int]$noHeadCols=1,
	[object[]]$rowArray=@(),
	[object[]]$fixedWidth=@(),
	[object[]]$columnArray=@())

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If($Null -ne $rowArray)
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If($noBorder)
	{
		$htmlbody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$htmlbody += "<table border='1' width='" + $tablewidth + "'>"
	}

	If(!($columnArray.Length -eq 0))
	{
		$htmlbody += "<tr>"

		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $columnArray[$columnIndex+1]
			If($fixedWidth.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $columnArray[$columnIndex])
			{
				If($columnArray[$columnIndex] -eq " " -or $columnArray[$columnIndex].length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					$found = $false
					For($i=0;$i -lt $columnArray[$columnIndex].length;$i+=2)
					{
						If($columnArray[$columnIndex][$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($columnArray[$columnIndex][$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $columnArray[$columnIndex]
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	$rowindex = 2
	If($Null -ne $rowArray)
	{
		AddHTMLTable $fontName $fontSize -colCount $numCols -rowCount $NumRows -rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = @()
		$htmlbody = "</table>"
	}
	Else
	{
		$HTMLBody += "</table>"
	}	
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:Filename1 -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -eq $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$True, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $True; }
				If($Italic) { $Cell.Range.Font.Italic = $True; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$True, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$True, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function CheckExcelPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Excel.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tFor the Delivery Groups Utilization option, this script directly outputs to Microsoft Excel, `n`t`tplease install Microsoft Excel or do not use the DeliveryGroupsUtilization (DGU) switch`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if excel is running in our session
	[bool]$excelrunning = ((Get-Process 'Excel' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($excelrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Excel before running this report.`n`n"
		Exit
	}
}

Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	
	[bool]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If(!$ModuleFound) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0 4>$Null
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Add-PSSnapin -Name $snapin -EA 0 *>$Null
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {Write-Warning "($_)"}
		Return $False
	}
	Else
	{
		Return $True
	}
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $Script:Filename1 and $Script:Filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		If($DeliveryGroupsUtilization)
		{
			CheckExcelPreReq
		}

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
		ShowScriptOptions
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
		SetupHTML
		ShowScriptOptions
	}
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Add DateTime    : $($AddDateTime)"
	Write-Verbose "$(Get-Date): AdminAddress    : $($AdminAddress)"
	Write-Verbose "$(Get-Date): Administrators  : $($Administrators)"
	Write-Verbose "$(Get-Date): Applications    : $($Applications)"
	Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
	Write-Verbose "$(Get-Date): DeliveryGroups  : $($DeliveryGroups)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile  : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): DGUtilization   : $($DeliveryGroupsUtilization)"
	Write-Verbose "$(Get-Date): Filename1       : $($Script:filename1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $($Script:Filename2)"
	}
	Write-Verbose "$(Get-Date): Folder          : $($Folder)"
	Write-Verbose "$(Get-Date): From            : $($From)"
	Write-Verbose "$(Get-Date): Hosting         : $($Hosting)"
	Write-Verbose "$(Get-Date): HW Inventory    : $($Hardware)"
	Write-Verbose "$(Get-Date): Logging         : $($Logging)"
	If($Logging)
	{
		Write-Verbose "$(Get-Date):    Start Date   : $($StartDate)"
		Write-Verbose "$(Get-Date):    End Date     : $($EndDate)"
	}
	Write-Verbose "$(Get-Date): MachineCatalogs : $($MachineCatalogs)"
	Write-Verbose "$(Get-Date): NoADPolicies    : $($NoADPolicies)"
	Write-Verbose "$(Get-Date): NoPolicies      : $($NoPolicies)"
	Write-Verbose "$(Get-Date): Policies        : $($Policies)"
	Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
	Write-Verbose "$(Get-Date): Save As HTML    : $($HTML)"
	Write-Verbose "$(Get-Date): Save As TEXT    : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Section         : $($Section)"
	Write-Verbose "$(Get-Date): Site Name       : $($XDSiteName)"
	Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
	Write-Verbose "$(Get-Date): StoreFront      : $($StoreFront)"
	Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
	Write-Verbose "$(Get-Date): To              : $($To)"
	Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
	Write-Verbose "$(Get-Date): User Name       : $($UserName)"
	Write-Verbose "$(Get-Date): XA/XD Version   : $($Script:XDSiteVersion)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version  : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start  : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function OutputWarning
{
	Param([string] $txt)
	Write-Warning $txt
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 $txt
		WriteWordLIne 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 1 $txt
		WriteHTMLLine 0 0 " "
	}
}

Function OutputAdminsForDetails
{
	Param([object] $Admins)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Administrators"
		## Create an array of hashtables to store our admins
		[System.Collections.Hashtable[]] $AdminsWordTable = @();
		## Seed the row index from the second row
	}
	ElseIf($Text)
	{
		Line 0 "Administrators"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Administrators"
		$rowdata = @()
	}
	
	ForEach($Admin in $Admins)
	{
		$Tmp = ""
		If($Admin.Enabled)
		{
			$Tmp = "Enabled"
		}
		Else
		{
			$Tmp = "Disabled"
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			AdminName = $Admin.Name;
			Role = $Admin.Rights[0].RoleName;
			Status = $Tmp;
			}
			$AdminsWordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "Administrator Name`t: " $Admin.Name
			Line 1 "Role`t`t`t: " $Admin.Rights[0].RoleName
			Line 1 "Status`t`t`t: " $tmp
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Admin.Name,$htmlwhite,
			$Admin.Rights[0].RoleName,$htmlwhite,
			$tmp,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $AdminsWordTable `
		-Columns AdminName, Role, Status `
		-Headers "Administrator Name", "Role", "Status" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 225;
		$Table.Columns.Item(2).Width = 200;
		$Table.Columns.Item(3).Width = 60;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Administrator Name',($htmlsilver -bor $htmlbold),
		'Role',($htmlsilver -bor $htmlbold),
		'Status',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("225","200","60")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "485"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"
	
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()

	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

#region getadmins function from Citrix
Function GetAdmins
{
	Param([string]$xType="", [string]$xName="")
	
	Switch ($xType)
	{
		"AppDisk" {
			$scopes = $Null
			$permissions = Get-AdminPermission @XDParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "AppDisk" } | `
			Select-Object -ExpandProperty Id
			$roles = Get-AdminRole @XDParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id
			#this is an unscoped object type as $admins is done differently than the others
			$Admins = Get-AdminAdministrator @XDParams2 | `
			Where-Object {$_.Rights | Where-Object {$roles -contains $_.RoleId}}
		}
		"ApplicationGroup" {
			$scopes = $Null
			$permissions = Get-AdminPermission @XDParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "ApplicationGroup" } | `
			Select-Object -ExpandProperty Id
			$roles = Get-AdminRole @XDParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id
			#this is an unscoped object type as $admins is done differently than the others
			$Admins = Get-AdminAdministrator @XDParams2 | `
			Where-Object {$_.Rights | Where-Object {$roles -contains $_.RoleId}}
		}
		"Catalog" {
			$scopes = (Get-BrokerCatalog -Name $xName @XDParams2).Scopes | `
			Select-Object -ExpandProperty ScopeId
			$permissions = Get-AdminPermission @XDParams2 | Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Catalog" } | `
			Select-Object -ExpandProperty Id
			$roles = Get-AdminRole @XDParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id
			$Admins = Get-AdminAdministrator @XDParams2 | `
			Where-Object {$_.Rights | Where-Object {($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
		}
		"DesktopGroup" {
			$scopes = (Get-BrokerDesktopGroup -Name $xName @XDParams2).Scopes | `
			Select-Object -ExpandProperty ScopeId
			$permissions = Get-AdminPermission @XDParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "DesktopGroup" } | `
			Select-Object -ExpandProperty Id
			$roles = Get-AdminRole @XDParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id
			$Admins = Get-AdminAdministrator @XDParams2 | `
			Where-Object {$_.Rights | Where-Object {($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
		}
		"Host" {
			$scopes = (Get-hypscopedobject -ObjectName $xName @XDParams2).ScopeId
			$permissions = Get-AdminPermission @XDParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Connection" -or `
			$_.MetadataMap["Citrix_ObjectType"] -eq "Host"} | `
			Select-Object -ExpandProperty Id		
			$roles = Get-AdminRole @XDParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id
			$Admins = Get-AdminAdministrator @XDParams2 | `
			Where-Object {$_.Rights | Where-Object {($_.ScopeId -eq [guid]::Empty -or `
			$scopes -contains $_.ScopeId) -and $roles -contains $_.RoleId}}
		}
		"Storefront" {
			$scopes = $Null
			$permissions = Get-AdminPermission @XDParams2 | `
			Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Storefront" } | `
			Select-Object -ExpandProperty Id
			$roles = Get-AdminRole @XDParams2 | `
			Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | `
			Select-Object -ExpandProperty Id
			#this is an unscoped object type as $admins is done differently than the others
			$Admins = Get-AdminAdministrator @XDParams2 | `
			Where-Object {$_.Rights | Where-Object {$roles -contains $_.RoleId}}
		}
	}
	
	# $scopes = (Get-BrokerCatalog -Name "XenApp 75" -adminaddress xd75 ).Scopes | Select-Object -ExpandProperty ScopeId

	# First, get all the permissions which are relevant to this object type
	# Change "Catalog" here as appropriate for the object type you're interested in
	# $permissions = Get-AdminPermission @XDParams2 | Where-Object { $_.MetadataMap["Citrix_ObjectType"] -eq "Catalog" } | Select-Object -ExpandProperty Id

	# Now, get all the roles which include at least one of those permissions
	# $roles = Get-AdminRole @XDParams2 | Where-Object {$_.Permissions | Where-Object { $permissions -contains $_ }} | Select-Object -ExpandProperty Id

	# Finally, get all administrators which have a scope/role pair which matches
	#$Admins = Get-AdminAdministrator @XDParams2 | Where-Object {
	#	$_.Rights | Where-Object {
	#		# [guid]::Empty is the GUID for the All scope
	#		# Remove the next line if you're dealing with an unscoped object type
	#		($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and
	#		$roles -contains $_.RoleId
	#	}
	#}
	#$Admins = Get-AdminAdministrator @XDParams2 | Where-Object {$_.Rights | Where-Object {($_.ScopeId -eq [guid]::Empty -or $scopes -contains $_.ScopeId) -and	$roles -contains $_.RoleId}}

	$Admins = $Admins | Sort Name
	Return ,$Admins
}
#endregion

#region Machine Catalog functions
Function ProcessMachineCatalogs
{
	Write-Verbose "$(Get-Date): Retrieving Machine Catalogs"

	$Global:TotalServerOSCatalogs = 0
	$Global:TotalDesktopOSCatalogs = 0
	$Global:TotalRemotePCCatalogs = 0

	$AllMachineCatalogs = Get-BrokerCatalog @XDParams2 -SortBy Name 

	If($? -and $Null -ne $AllMachineCatalogs)
	{
		OutputMachines $AllMachineCatalogs
	}
	ElseIf($? -and ($Null -eq $AllMachineCatalogs))
	{
		$txt = "There are no Machines"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Machines"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputMachines
{
	Param([object]$Catalogs)
	
	Write-Verbose "$(Get-Date): `tProcessing Machine Catalogs"
	
	$txt = "Machine Catalogs"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	#add 16-jun-2015, summary table of catalogs to match what is shown in Studio
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	ForEach($Catalog in $Catalogs)
	{
		$xCatalogType = ""
		$xAllocationType = ""
		$xPersistType = ""
		$xProvisioningType = ""
		
		If($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Desktop OS"
			$Global:TotalDesktopOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Server OS"
			$Global:TotalServerOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Desktop OS (Virtual)"
			$Global:TotalDesktopOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Server OS (Virtual)"
			$Global:TotalServerOSCatalogs++
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $True)
		{
			$xCatalogType = "Desktop OS (Remote PC Access)"
			$Global:TotalRemotePCCatalogs++
		}
		
		Switch ($Catalog.AllocationType)
		{
			"Static"	{$xAllocationType = "Static"; Break}
			"Permanent"	{$xAllocationType = "Static"; Break}
			"Random"	{$xAllocationType = "Random"; Break}
			Default		{$xAllocationType = "Allocation type could not be determined: $($Catalog.AllocationType)"; Break}
		}
		Switch ($Catalog.PersistUserChanges)
		{
			"OnLocal" {$xPersistType = "On local disk"; Break}
			"Discard" {$xPersistType = "Discard"; Break}
			"OnPvd"   {$xPersistType = "On personal vDisk"; Break}
			Default   {$xPersistType = "User data could not be determined: $($Catalog.PersistUserChanges)"; Break}
		}
		Switch ($Catalog.ProvisioningType)
		{
			"Manual" {$xProvisioningType = "Manual"; Break}
			"PVS"    {$xProvisioningType = "Provisioning Services"; Break}
			"MCS"    {$xProvisioningType = "Machine creation services"; Break}
			Default  {$xProvisioningType = "Provisioning method could not be determined: $($Catalog.ProvisioningType)"; Break}
		}

		$Machines = Get-BrokerMachine @XDParams2 -CatalogName $Catalog.Name -SortBy DNSName 
		If($? -and ($Null -ne $Machines))
		{
		
			If($Machines -is [array])
			{
				$NumberOfMachines = $Machines.Count
			}
			Else
			{
				$NumberOfMachines = 1
			}
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			MachineCatalogName = $Catalog.Name; 
			MachineType = $xCatalogType; 
			NoOfMachines = $NumberOfMachines;
			AllocatedMachines = $Catalog.UsedCount; 
			AllocationType = $xAllocationType;
			UserData = $xPersistType;
			ProvisioningMethod = $xProvisioningType;
			}
			$WordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 0 "Machine Catalog`t`t: " $Catalog.Name
			Line 0 "Machine type`t`t: " $xCatalogType
			Line 0 "No. of machines`t`t: " $NumberOfMachines
			Line 0 "Allocated machines`t: " $Catalog.UsedCount
			Line 0 "Allocation type`t`t: " $xAllocationType
			Line 0 "User data`t`t: " $xPersistType
			Line 0 "Provisioning method`t: " $xProvisioningType
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Catalog.Name,$htmlwhite,
			$xCatalogType,$htmlwhite,
			$NumberOfMachines,$htmlwhite,
			$Catalog.UsedCount,$htmlwhite,
			$xAllocationType,$htmlwhite,
			$xPersistType,$htmlwhite,
			$xProvisioningType,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  MachineCatalogName, MachineType, NoOfMachines, AllocatedMachines, AllocationType, UserData, ProvisioningMethod `
		-Headers  "Machine Catalog", "Machine type", "No. of machines", "Allocated machines", "Allocation Type", "User data", "Provisioning method" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 105;
		$Table.Columns.Item(2).Width = 100;
		$Table.Columns.Item(3).Width = 75;
		$Table.Columns.Item(4).Width = 50;
		$Table.Columns.Item(5).Width = 55;
		$Table.Columns.Item(6).Width = 50;
		$Table.Columns.Item(7).Width = 65;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Machine Catalog',($htmlsilver -bor $htmlbold),
		'Machine type',($htmlsilver -bor $htmlbold),
		'No. of machines',($htmlsilver -bor $htmlbold),
		'Allocated machines',($htmlsilver -bor $htmlbold),
		'Allocation Type',($htmlsilver -bor $htmlbold),
		'User data',($htmlsilver -bor $htmlbold),
		'Provisioning method',($htmlsilver -bor $htmlbold)
		)

		$columnWidths = @("105","100","75","50","55","50","65")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
	
	ForEach($Catalog in $Catalogs)
	{
		Write-Verbose "$(Get-Date): `t`tAdding Catalog $($Catalog.Name)"
		$xCatalogType = ""
		$xAllocationType = ""
		$xPersistType = ""
		$xProvisioningType = ""
		$xVDAVersion = ""
		
		If($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Desktop OS"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Server OS"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "SingleSession")
		{
			$xCatalogType = "Desktop OS (Virtual)"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $False -and $Catalog.IsRemotePC -eq $False -and $Catalog.SessionSupport -eq "MultiSession")
		{
			$xCatalogType = "Server OS (Virtual)"
		}
		ElseIf($Catalog.MachinesArePhysical -eq $True -and $Catalog.IsRemotePC -eq $True)
		{
			$xCatalogType = "Desktop OS (Remote PC Access)"
		}

		Switch ($Catalog.AllocationType)
		{
			"Static"	{$xAllocationType = "Permanent"; Break}
			"Permanent"	{$xAllocationType = "Permanent"; Break}
			"Random"	{$xAllocationType = "Random"; Break}
			Default		{$xAllocationType = "Allocation type could not be determined: $($Catalog.AllocationType)"; Break}
		}
		Switch ($Catalog.PersistUserChanges)
		{
			"OnLocal" {$xPersistType = "On local disk"; Break}
			"Discard" {$xPersistType = "Discard"; Break}
			"OnPvd"   {$xPersistType = "On personal vDisk"; Break}
			Default   {$xPersistType = "User data could not be determined: $($Catalog.PersistUserChanges)"; Break}
		}
		Switch ($Catalog.ProvisioningType)
		{
			"Manual" {$xProvisioningType = "Manual"; Break}
			"PVS"    {$xProvisioningType = "Provisioning Services"; Break}
			"MCS"    {$xProvisioningType = "Machine creation services"; Break}
			Default  {$xProvisioningType = "Provisioning method could not be determined: $($Catalog.ProvisioningType)"; Break}
		}
		Switch ($Catalog.MinimumFunctionalLevel)
		{
			"L5" 	{$xVDAVersion = "5.6 FP1 (Windows XP and Windows Vista)"; Break}
			"L7"	{$xVDAVersion = "7.0 (or newer)"; Break}
			"L7_6"	{$xVDAVersion = "7.6 (or newer)"; Break}
			"L7_7"	{$xVDAVersion = "7.7 (or newer)"; Break}
			"L7_8"	{$xVDAVersion = "7.8 (or newer)"; Break}
			"L7_9"	{$xVDAVersion = "7.9 (or newer)"; Break}
			Default {"Unable to determine VDA version: $($Catalog.MinimumFunctionalLevel)"; Break}
		}

		If($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
		{
			$Results = Get-BrokerRemotePCAccount @XDParams2 -CatalogUid $Catalog.Uid
			
			If($? -and $Null -ne $Results)
			{
				$RemotePCOU = $Results.OU
				$RemotePCSubOU = $Results.AllowSubfolderMatches.ToString()
			}
			ElseIf($? -and $Null -eq $Results)
			{
				$RemotePCOU = "No RemotePC OU configured"
				$RemotePCSubOU = ""
			}
			Else
			{
				$RemotePCOU = "Unable to retrieve"
				$RemotePCSubOU = ""
			}
		}
		
		$MachineData = $Null
		$Machines = Get-BrokerMachine @XDParams2 -CatalogName $Catalog.Name -SortBy DNSName 
		If($? -and ($Null -ne $Machines))
		{
		
			If($Machines -is [array])
			{
				$NumberOfMachines = $Machines.Count
			}
			Else
			{
				$NumberOfMachines = 1
			}
			
			$MachineData = Get-ProvScheme -ProvisioningSchemeUid $Catalog.ProvisioningSchemeId @XDParams1
			If($? -and $Null -ne $MachineData)
			{
				$tmp1 = $MachineData.MasterImageVM.Split("\")
				$tmp2 = $tmp1[$tmp1.count -1]
				$tmp3 = $tmp2.Split(".")
				$xDiskImage = $tmp3[0]

				#28-sep-2016 add the VM the catalog is based on
				$MasterVM = ""
				ForEach($Item in $tmp1)
				{
					If($Item.EndsWith(".vm"))
					{
						$MasterVM = $Item
					}
				}
				
				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and (($xAllocationType -eq "Random") -or ($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					$TempDiskCacheSize = $MachineData.WriteBackCacheDiskSize
					$TempMemoryCacheSize = $MachineData.WriteBackCacheMemorySize
				}
				
				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and ($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and ((Get-ConfigEnabledFeature @XDParams1) -contains "DedicatedFullDiskClone"))
				{
					If($MachineData.UseFullDiskCloneProvisioning -eq $True)
					{
						$VMCopyMode = "Full Copy"
					}
					Else
					{
						$VMCopyMode = "Fast Clone"
					}
				}
				
				If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
				{
					$xPvDDriveLetter = $MachineData.PersonalVDiskDriveLetter
					$xPvDSize = $MachineData.PersonalVDiskDriveSize
				}

			}
			Else
			{
				$xDiskImage = "Unable to retrieve details"
			}
		}
		Else
		{
			Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
		}
		
		If($MSWord -or $PDF)
		{

			$Selection.InsertNewPage()
			WriteWordLine 2 0 "Machine Catalog: $($Catalog.Name)"
			[System.Collections.Hashtable[]] $CatalogInformation = @()
			
			If($Catalog.ProvisioningType -eq "MCS")
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "No. of machines"; Value = $NumberOfMachines; }
				$CatalogInformation += @{Data = "Allocated machines"; Value = $Catalog.UsedCount; }
				$CatalogInformation += @{Data = "Allocation type"; Value = $xAllocationType; }
				$CatalogInformation += @{Data = "User data"; Value = $xPersistType; }
				$CatalogInformation += @{Data = "Provisioning method"; Value = $xProvisioningType; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				$CatalogInformation += @{Data = "Resources"; Value = $MachineData.HostingUnitName; }
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }
				}

				If($Null -ne $MachineData)
				{
					$CatalogInformation += @{Data = "Master VM"; Value = $MasterVM; }
					$CatalogInformation += @{Data = "Disk Image"; Value = $xDiskImage; }
					$CatalogInformation += @{Data = "Virtual CPUs"; Value = $MachineData.CpuCount; }
					$CatalogInformation += @{Data = "Memory"; Value = "$($MachineData.MemoryMB) MB"; }
					$CatalogInformation += @{Data = "Hard disk"; Value = "$($MachineData.DiskSize) GB"; }
					If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
					{
						$CatalogInformation += @{Data = "Personal vDisk size"; Value = "$($xPvDSize) GB"; }
						$CatalogInformation += @{Data = "Personal vDisk drive letter"; Value = $xPvDDriveLetter; }
					}
				}
				ElseIf($Null -eq $MachineData)
				{
					$CatalogInformation += @{Data = "Master VM"; Value = $MasterVM; }
					$CatalogInformation += @{Data = "Disk Image"; Value = $xDiskImage; }
					$CatalogInformation += @{Data = "Virtual CPUs"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Memory"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Hard disk"; Value = "Unable to retrieve details"; }
					If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
					{
						$CatalogInformation += @{Data = "Personal vDisk size"; Value = "Unable to retrieve details"; }
						$CatalogInformation += @{Data = "Personal vDisk drive letter"; Value = "Unable to retrieve details"; }
					}
				}
			
				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and (($xAllocationType -eq "Random") -or ($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					$CatalogInformation += @{Data = "Temporary memory cache size"; Value = "$($MachineData.WriteBackCacheMemorySize) MB"; }
					$CatalogInformation += @{Data = "Temporary disk cache size"; Value = "$($MachineData.WriteBackCacheDiskSize) GB"; }
				}

				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and ($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and ((Get-ConfigEnabledFeature @XDParams1) -contains "DedicatedFullDiskClone"))
				{
					$CatalogInformation += @{Data = "VM copy mode"; Value = $VMCopyMode; }
				}
				
				If($Null -ne $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
					$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "PVS")
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "Provisioning method"; Value = $xProvisioningType; }
				$CatalogInformation += @{Data = "PVS address"; Value = $Catalog.PvsAddress; }
				$CatalogInformation += @{Data = "Allocation type"; Value = $xAllocationType; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				$CatalogInformation += @{Data = "Resources"; Value = $MachineData.HostingUnitName; }
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }
				}
			
				If($Null -ne $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
					$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
			{
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Organizational Units"; Value = $RemotePCOU; }
				$CatalogInformation += @{Data = "Allow subfolder matches"; Value = $RemotePCSubOU; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "No. of machines"; Value = $NumberOfMachines; }
				$CatalogInformation += @{Data = "Allocated machines"; Value = $Catalog.UsedCount; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }
				}

				If($Null -ne $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
					$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -ne $True)
			{
				#(added 29-feb-2016 thanks to Michael Foster)
				$CatalogInformation += @{Data = "Description"; Value = $Catalog.Description; }
				$CatalogInformation += @{Data = "Machine type"; Value = $xCatalogType; }
				$CatalogInformation += @{Data = "No. of machines"; Value = $NumberOfMachines; }
				$CatalogInformation += @{Data = "Allocated machines"; Value = $Catalog.UsedCount; }
				$CatalogInformation += @{Data = "Allocation type"; Value = $xAllocationType; }
				$CatalogInformation += @{Data = "User data"; Value = $xPersistType; }
				$CatalogInformation += @{Data = "Provisioning method"; Value = $xProvisioningType; }
				$CatalogInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$CatalogInformation += @{Data = "Zone"; Value = $Catalog.ZoneName; }
				}
				
				If($Null -ne $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = $Machines[0].AgentVersion; }
					$CatalogInformation += @{Data = "Operating System"; Value = $Machines[0].OSType; }
				}
				ElseIf($Null -eq $Machines)
				{
					$CatalogInformation += @{Data = "Installed VDA version"; Value = "Unable to retrieve details"; }
					$CatalogInformation += @{Data = "Operating System"; Value = "Unable to retrieve details"; }
					Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
				}
			}
			
			$Table = AddWordTable -Hashtable $CatalogInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 225;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 "Machine Catalog: $($Catalog.Name)"
			If($Catalog.ProvisioningType -eq "MCS")
			{
				Line 1 "Description`t`t`t: " $Catalog.Description
				Line 1 "Machine type`t`t`t: " $xCatalogType
				Line 1 "No. of machines`t`t`t: " $NumberOfMachines
				Line 1 "Allocated machines`t`t: " $Catalog.UsedCount
				Line 1 "Allocation type`t`t`t: " $xAllocationType
				Line 1 "User data`t`t`t: " $xPersistType
				Line 1 "Provisioning method`t`t: " $xProvisioningType
				Line 1 "Set to VDA version`t`t: " $xVDAVersion
				Line 1 "Resources`t`t`t: " $MachineData.HostingUnitName
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					Line 1 "Zone`t`t`t`t: " $Catalog.ZoneName
				}
				
				If($Null -ne $MachineData)
				{
					Line 1 "Master VM`t`t`t: " $MasterVM
					Line 1 "Disk Image`t`t`t: " $xDiskImage
					Line 1 "Virtual CPUs`t`t`t: " $MachineData.CpuCount
					Line 1 "Memory`t`t`t`t: " "$($MachineData.MemoryMB) MB"
					Line 1 "Hard disk`t`t`t: " "$($MachineData.DiskSize) GB"
					If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
					{
						Line 1 "Personal vDisk size`t`t: " "$($xPvDSize) GB"
						Line 1 "Personal vDisk drive letter`t: " $xPvDDriveLetter
					}
				}
				ElseIf($Null -eq $MachineData)
				{
					Line 1 "Master VM`t`t`t: " $MasterVM
					Line 1 "Disk Image`t`t`t: " $xDiskImage
					Line 1 "Virtual CPUs`t`t`t: " "Unable to retrieve details"
					Line 1 "Memory`t`t`t`t: " "Unable to retrieve details"
					Line 1 "Hard disk`t`t`t: " "Unable to retrieve details"
					If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
					{
						Line 1 "Personal vDisk size`t`t: " "Unable to retrieve details"
						Line 1 "Personal vDisk drive letter`t: " "Unable to retrieve details"
					}
				}
				
				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and (($xAllocationType -eq "Random") -or ($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					Line 1 "Temporary memory cache size`t: " "$($MachineData.WriteBackCacheMemorySize) MB"
					Line 1 "Temporary disk cache size`t: " "$($MachineData.WriteBackCacheDiskSize) GB"
				}

				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and ($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and ((Get-ConfigEnabledFeature @XDParams1) -contains "DedicatedFullDiskClone"))
				{
					Line 1 "VM copy mode`t`t`t: " $VMCopyMode
				}
				
				If($Null -ne $Machines)
				{
					Line 1 "Installed VDA version`t`t: " $Machines[0].AgentVersion
					Line 1 "Operating System`t`t: " $Machines[0].OSType
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t: " "Unable to retrieve details"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "PVS")
			{
				Line 1 "Description`t`t`t: " $Catalog.Description
				Line 1 "Machine type`t`t`t: " $xCatalogType
				Line 1 "Provisioning method`t`t: " $xProvisioningType
				Line 1 "PVS address`t`t`t: " $Catalog.PvsAddress
				Line 1 "Allocation type`t`t`t: " $xAllocationType
				Line 1 "Set to VDA version`t`t: " $xVDAVersion
				Line 1 "Resources`t`t`t: " $MachineData.HostingUnitName
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					Line 1 "Zone`t`t`t`t: " $Catalog.ZoneName
				}
				
				If($Null -ne $Machines)
				{
					Line 1 "Installed VDA version`t`t: " $Machines[0].AgentVersion
					Line 1 "Operating System`t`t: " $Machines[0].OSType
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t: " "Unable to retrieve details"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
			{
				Line 1 "Description`t`t`t: " $Catalog.Description
				Line 1 "Organizational Units`t`t: " $RemotePCOU
				Line 1 "Allow subfolder matches`t`t: " $RemotePCSubOU
				Line 1 "Machine type`t`t`t: " $xCatalogType
				Line 1 "No. of machines`t`t`t: "$NumberOfMachines
				Line 1 "Allocated machines`t`t: " $Catalog.UsedCount
				Line 1 "Set to VDA version`t`t: " $xVDAVersion
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					Line 1 "Zone`t`t`t`t: " $Catalog.ZoneName
				}
				
				If($Null -ne $Machines)
				{
					Line 1 "Installed VDA version`t`t: " $Machines[0].AgentVersion
					Line 1 "Operating System`t`t: " $Machines[0].OSType
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t: " "Unable to retrieve details"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -ne $True)
			{
				#(added 29-feb-2016 thanks to Michael Foster)
				Line 1 "Description`t`t`t: " $Catalog.Description
				Line 1 "Machine type`t`t`t: " $xCatalogType
				Line 1 "No. of machines`t`t`t: "$NumberOfMachines
				Line 1 "Allocated machines`t`t: " $Catalog.UsedCount
				Line 1 "Allocation type`t`t`t: " $xAllocationType
				Line 1 "User data`t`t`t: " $xPersistType
				Line 1 "Provisioning method`t`t: " $xProvisioningType
				Line 1 "Set to VDA version`t`t: " $xVDAVersion
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					Line 1 "Zone`t`t`t`t: " $Catalog.ZoneName
				}
				
				If($Null -ne $Machines)
				{
					Line 1 "Installed VDA version`t`t: " $Machines[0].AgentVersion
					Line 1 "Operating System`t`t: " $Machines[0].OSType
				}
				ElseIf($Null -eq $Machines)
				{
					Line 1 "Installed VDA version`t`t: " "Unable to retrieve details"
					Line 1 "Operating System`t`t: " "Unable to retrieve details"
					Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
				}
			}
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 "Machine Catalog: $($Catalog.Name)"
			$rowdata = @()
			$columnHeaders = @("Machine type",($htmlsilver -bor $htmlbold),$xCatalogType,$htmlwhite)
			If($Catalog.ProvisioningType -eq "MCS")
			{
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Machine Type',($htmlsilver -bor $htmlbold),$xCatalogType,$htmlwhite))
				$rowdata += @(,('No. of machines',($htmlsilver -bor $htmlbold),$NumberOfMachines,$htmlwhite))
				$rowdata += @(,('Allocated machines',($htmlsilver -bor $htmlbold),$Catalog.UsedCount,$htmlwhite))
				$rowdata += @(,('Allocation type',($htmlsilver -bor $htmlbold),$xAllocationType,$htmlwhite))
				$rowdata += @(,('User data',($htmlsilver -bor $htmlbold),$xPersistType,$htmlwhite))
				$rowdata += @(,('Provisioning method',($htmlsilver -bor $htmlbold),$xProvisioningType,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($htmlsilver -bor $htmlbold),$xVDAVersion,$htmlwhite))
				$rowdata += @(,('Resources',($htmlsilver -bor $htmlbold),$MachineData.HostingUnitName,$htmlwhite))
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$Catalog.ZoneName,$htmlwhite))
				}
				
				If($Null -ne $MachineData)
				{
					$rowdata += @(,('Master VM',($htmlsilver -bor $htmlbold),$MasterVM,$htmlwhite))
					$rowdata += @(,('Disk Image',($htmlsilver -bor $htmlbold),$xDiskImage,$htmlwhite))
					$rowdata += @(,('Virtual CPUs',($htmlsilver -bor $htmlbold),$MachineData.CpuCount,$htmlwhite))
					$rowdata += @(,('Memory',($htmlsilver -bor $htmlbold),"$($MachineData.MemoryMB) MB",$htmlwhite))
					$rowdata += @(,('Hard disk',($htmlsilver -bor $htmlbold),"$($MachineData.DiskSize) GB",$htmlwhite))
					If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
					{
						$rowdata += @(,('Personal vDisk size',($htmlsilver -bor $htmlbold),"$($xPvDSize) GB",$htmlwhite))
						$rowdata += @(,('Personal vDisk drive letter',($htmlsilver -bor $htmlbold),$xPvDDriveLetter,$htmlwhite))
					}
				}
				ElseIf($Null -eq $MachineData)
				{
					$rowdata += @(,('Master VM',($htmlsilver -bor $htmlbold),$MasterVM,$htmlwhite))
					$rowdata += @(,('Disk Image',($htmlsilver -bor $htmlbold),$xDiskImage,$htmlwhite))
					$rowdata += @(,('Virtual CPUs',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Memory',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Hard disk',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					If($xAllocationType -eq "Permanent" -and $xPersistType -eq "On personal vDisk" )
					{
						$rowdata += @(,('Personal vDisk size',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
						$rowdata += @(,('Personal vDisk drive letter',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					}
				}
				
				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and (($xAllocationType -eq "Random") -or ($xAllocationType -eq "Permanent" -and $xPersistType -eq "Discard" )))
				{
					$rowdata += @(,('Temporary memory cache size',($htmlsilver -bor $htmlbold),"$($MachineData.WriteBackCacheMemorySize) MB",$htmlwhite))
					$rowdata += @(,('Temporary disk cache size',($htmlsilver -bor $htmlbold), "$($MachineData.WriteBackCacheDiskSize) GB",$htmlwhite))
				}

				If($Catalog.MinimumFunctionalLevel -eq "L7_9" -and ($xAllocationType -eq "Permanent" -and $xPersistType -eq "On local disk" ) -and ((Get-ConfigEnabledFeature @XDParams1) -contains "DedicatedFullDiskClone"))
				{
					$rowdata += @(,('VM copy mode',($htmlsilver -bor $htmlbold),$VMCopyMode,$htmlwhite))
				}
				
				If($Null -ne $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),$Machines[0].AgentVersion,$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$Machines[0].OSType,$htmlwhite))
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "PVS")
			{
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Machine Type',($htmlsilver -bor $htmlbold),$xCatalogType,$htmlwhite))
				$rowdata += @(,('Provisioning method',($htmlsilver -bor $htmlbold),$xProvisioningType,$htmlwhite))
				$rowdata += @(,('PVS address',($htmlsilver -bor $htmlbold),$Catalog.PvsAddress,$htmlwhite))
				$rowdata += @(,('Allocation type',($htmlsilver -bor $htmlbold),$xAllocationType,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($htmlsilver -bor $htmlbold),$xVDAVersion,$htmlwhite))
				$rowdata += @(,('Resources',($htmlsilver -bor $htmlbold),$MachineData.HostingUnitName,$htmlwhite))
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$Catalog.ZoneName,$htmlwhite))
				}
				
				If($Null -ne $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),$Machines[0].AgentVersion,$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$Machines[0].OSType,$htmlwhite))
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -eq $True)
			{
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Organizational Units',($htmlsilver -bor $htmlbold),$RemotePCOU,$htmlwhite))
				$rowdata += @(,('Allow subfolder matches',($htmlsilver -bor $htmlbold),$RemotePCSubOU,$htmlwhite))
				$rowdata += @(,('Machine Type',($htmlsilver -bor $htmlbold),$xCatalogType,$htmlwhite))
				$rowdata += @(,('No. of machines',($htmlsilver -bor $htmlbold),$NumberOfMachines,$htmlwhite))
				$rowdata += @(,('Allocated machines',($htmlsilver -bor $htmlbold),$Catalog.UsedCount,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($htmlsilver -bor $htmlbold),$xVDAVersion,$htmlwhite))
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$Catalog.ZoneName,$htmlwhite))
				}
				
				If($Null -ne $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),$Machines[0].AgentVersion,$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$Machines[0].OSType,$htmlwhite))
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
				}
			}
			ElseIf($Catalog.ProvisioningType -eq "Manual" -and $Catalog.IsRemotePC -ne $True)
			{
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Catalog.Description,$htmlwhite))
				$rowdata += @(,('Machine Type',($htmlsilver -bor $htmlbold),$xCatalogType,$htmlwhite))
				$rowdata += @(,('No. of machines',($htmlsilver -bor $htmlbold),$NumberOfMachines,$htmlwhite))
				$rowdata += @(,('Allocated machines',($htmlsilver -bor $htmlbold),$Catalog.UsedCount,$htmlwhite))
				$rowdata += @(,('Allocation type',($htmlsilver -bor $htmlbold),$xAllocationType,$htmlwhite))
				$rowdata += @(,('User data',($htmlsilver -bor $htmlbold),$xPersistType,$htmlwhite))
				$rowdata += @(,('Provisioning method',($htmlsilver -bor $htmlbold),$xProvisioningType,$htmlwhite))
				$rowdata += @(,('Set to VDA version',($htmlsilver -bor $htmlbold),$xVDAVersion,$htmlwhite))
				If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
				{
					$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$Catalog.ZoneName,$htmlwhite))
				}
				
				If($Null -ne $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),$Machines[0].AgentVersion,$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$Machines[0].OSType,$htmlwhite))
				}
				ElseIf($Null -eq $Machines)
				{
					$rowdata += @(,('Installed VDA version',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),"Unable to retrieve details",$htmlwhite))
					Write-Warning "Unable to retrieve details for Machine Catalog $($Catalog.Name)"
				}
			}
			
			$msg = ""
			$columnWidths = @("225px","200px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "425"
			WriteHTMLLine 0 0 " "
		}
			
		#scopes
		$Scopes = (Get-BrokerCatalog -Name $Catalog.Name @XDParams2).Scopes
		
		If($? -and ($Null -eq $Scopes))
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Scopes"
				[System.Collections.Hashtable[]] $ScopesWordTable = @();

				$WordTableRowHash = @{ 
				Scope = "All";
				}

				$ScopesWordTable += $WordTableRowHash;

				$Table = AddWordTable -Hashtable $ScopesWordTable `
				-Columns Scope `
				-Headers  "Scopes" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 225;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Scopes"
				Line 2 "All"
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 4 0 "Scopes"
				$rowdata = @()
				$rowdata += @(,("All",$htmlwhite))

				$columnHeaders = @(
				'Scopes',($htmlsilver -bor $htmlbold))

				$msg = ""
				$columnWidths = @("225")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
				WriteHTMLLine 0 0 " "
			}
		}
		ElseIf($? -and ($Null -ne $Scopes))
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Scopes"
				[System.Collections.Hashtable[]] $ScopesWordTable = @();

				$WordTableRowHash = @{ 
				Scope = "All";
				}

				$ScopesWordTable += $WordTableRowHash;

				$CurrentServiceIndex++;
				
				ForEach($Scope in $Scopes)
				{
					$WordTableRowHash = @{ 
					Scope = $Scope.ScopeName;
					}

					$ScopesWordTable += $WordTableRowHash;
				}
				$Table = AddWordTable -Hashtable $ScopesWordTable `
				-Columns Scope `
				-Headers  "Scopes" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 225;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Scopes"
				Line 2 "All"

				ForEach($Scope in $Scopes)
				{
					Line 2 $Scope.ScopeName;
				}
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 4 0 "Scopes"
				$rowdata = @()
				$rowdata += @(,("All",$htmlwhite))

				ForEach($Scope in $Scopes)
				{
					$rowdata += @(,($Scope.ScopeName,$htmlwhite))
				}
				$columnHeaders = @(
				'Scopes',($htmlsilver -bor $htmlbold))

				$msg = ""
				$columnWidths = @("225")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			$txt = "Unable to retrieve Scopes for Machine Catalog $($Catalog.Name)"
			OutputWarning $txt
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 " "
			}
		}
		
		If($MachineCatalogs)
		{
			If($Null -ne $Machines)
			{
				Write-Verbose "$(Get-Date): `t`tProcessing Machines in $($Catalog.Name)"
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Machines"
					[System.Collections.Hashtable[]] $MachinesWordTable = @();
				}
				ElseIf($Text)
				{
					Line 1 "Machines"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 4 0 "Machines"
					$rowdata = @()
				}
				
				ForEach($Machine in $Machines)
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ MachineName = $Machine.MachineName;}
						$MachinesWordTable += $WordTableRowHash;
					}
					ElseIf($Text)
					{
						Line 2 $Machine.MachineName
					}
					ElseIf($HTML)
					{
						$rowdata += @(,($Machine.MachineName,$htmlwhite))
					}
				}
				
				If($MSWord -or $PDF)
				{
					$Table = AddWordTable -Hashtable $MachinesWordTable `
					-Columns MachineName `
					-Headers "Machine Names" `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 225;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				ElseIf($Text)
				{
					Line 0 ""
				}
				ElseIf($HTML)
				{
					$columnHeaders = @(
					'Machine Names',($htmlsilver -bor $htmlbold))

					$msg = ""
					$columnWidths = @("225")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
					WriteHTMLLine 0 0 " "
				}
				
				Write-Verbose "$(Get-Date): `t`tProcessing administrators for Machines in $($Catalog.Name)"
				$Admins = GetAdmins "Catalog" $Catalog.Name
				
				If($? -and ($Null -ne $Admins))
				{
					OutputAdminsForDetails $Admins
				}
				ElseIf($? -and ($Null -eq $Admins))
				{
					$txt = "There are no administrators for Machines in $($Catalog.Name)"
					OutputWarning $txt
				}
				Else
				{
					$txt = "Unable to retrieve administrators for Machines in $($Catalog.Name)"
					OutputWarning $txt
				}
				
				ForEach($Machine in $Machines)
				{
					OutputMachineDetails $Machine
				}
			}
		}
	}
}
#endregion

#region AppDisks
Function ProcessAppDisks
{
	Write-Verbose "$(Get-Date): Retrieving AppDisks"
	$Global:TotalAppDisks = 0

	$AllAppDisks = Get-AppLibAppDisk @XDParams2 -SortBy AppDiskName 

	If($? -and ($Null -ne $AllAppDisks))
	{
		Write-Verbose "$(Get-Date): `tProcessing AppDisks"
		
		OutputAppDiskTable $AllAppDisks
		
		ForEach($AppDisk in $AllAppDisks)
		{
			$Global:TotalAppDisks++
			
			If($AppDisks)
			{
				OutputAppDisk $AppDisk
			}
		}
	}
	ElseIf($? -and ($Null -eq $AllAppDisks))
	{
		$txt = "There are no AppDisks"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve AppDisks"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputAppDiskTable
{
	Param([object] $AllAppDisks)

	$txt = "AppDisks"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}
		
	ForEach($AppDisk in $AllAppDisks)
	{
		Write-Verbose "$(Get-Date): `t`tAdding AppDisk $($AppDisk.AppDiskName)"
		
		$xAppDiskState = ""
		Switch($AppDisk.State)
		{
			"Cloning"		{$xAppDiskState = "Cloning"; Break}
			"Creating"		{$xAppDiskState = "Creating"; Break}
			"Deleting"		{$xAppDiskState = "Deleting"; Break}
			"Error"			{$xAppDiskState = "Error"; Break}
			"Importing"		{$xAppDiskState = "Importing"; Break}
			"None"			{$xAppDiskState = "None"; Break}
			"Populating"	{$xAppDiskState = "Ready to install applications"; Break}
			"Ready"			{$xAppDiskState = "Ready"; Break}
			"Sealing"		{$xAppDiskState = "Sealing"; Break}
			"Unknown"		{$xAppDiskState = "Unknown"; Break}
			"Unsupported"	{$xAppDiskState = "Unsupported"; Break}
			Default			{"Unable to determine AppDisk State: $($AppDisk.State)"; Break}
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			AppDiskName = $AppDisk.AppDiskName; 
			AppDiskState = $xAppDiskState
			}
			$WordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "AppDisk Name`t: " $AppDisk.AppDiskName
			Line 1 "State`t`t: " $xAppDiskState
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$AppDisk.AppDiskName,$htmlwhite,
			$xAppDiskState,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  AppDiskName, AppDiskState `
		-Headers  "Name", "State" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'State',($htmlsilver -bor $htmlbold)
		)

		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputAppDisk
{
	Param([object] $AppDisk)
	
	If($MSWORD -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 3 0 "AppDisk: " $AppDisk.AppDiskName
	}
	ElseIf($Text)
	{
		Line 0 "AppDisk: " $AppDisk.AppDiskName
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHtmlLine 3 0 "AppDisk: " $AppDisk.AppDiskName
	}
	
	OutputAppDiskDetails $AppDisk
	OutputAppDiskApplications $AppDisk
	OutputAppDiskDeliveryGroups $AppDisk
	OutputAppDiskAdmins $AppDisk
}

Function OutputAppDiskDetails
{
	Param([object] $AppDisk)

	Write-Verbose "$(Get-Date): `t`tProcessing details"
	$xCompatibility = ""
	Switch ($AppDisk.Compatibility)
	{
		"Compatible"		{$xCompatibility = "Compatible"; Break}
		"None"				{$xCompatibility = "None"; Break}
		"ProblemsDetected"	{$xCompatibility = "Problems detected"; Break}
		"Unknown"			{$xCompatibility = "Unknown"; Break}
		"Unsupported"		{$xCompatibility = "Unknown (AppDNA required)"; Break}
		Default				{$xCompatibility = "Unable to determine Compatibility: $($AppDisk.Compatibility)"; Break}
	}
	
	#get preparation machine details
	$PrepMachine = Get-BrokerMachine @XDParams2 -SID $AppDisk.ReservedMachineSid
	
	If($? -and $Null -ne $PrepMachine)
	{
		$PrepMachineName = $PrepMachine.MachineName
		$PrepMachinePowerState = $PrepMachine.PowerState
		$PrepMachineRegistrationState = $PrepMachine.RegistrationState
	}
	ElseIf($? -and ($Null -eq $PrepMachine))
	{
		$txt = "There is no Preparation machine"
		OutputWarning $txt
		$PrepMachineName = "Unknown"
		$PrepMachinePowerState = "Unknown"
		$PrepMachineRegistrationState = "Unknown"
	}
	Else
	{
		$txt = "Unable to retrieve Preparation machine"
		OutputWarning $txt
		$PrepMachineName = "Unknown"
		$PrepMachinePowerState = "Unknown"
		$PrepMachineRegistrationState = "Unknown"
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 4 0 "Details"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Name"; Value = $AppDisk.AppDiskName; }
		$ScriptInformation += @{Data = "Description"; Value = $AppDisk.Description; }
		$ScriptInformation += @{Data = "Resources"; Value = $AppDisk.HostingUnitName[0]; }
		$cnt = -1
		ForEach($tmp in $AppDisk.HostingUnitName)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		$ScriptInformation += @{Data = "Compatibility"; Value = $xCompatibility; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		
		WriteWordLine 4 0 "Preparation machine"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Machine name"; Value = $PrepMachineName; }
		$ScriptInformation += @{Data = "Power state"; Value = $PrepMachinePowerState; }
		$ScriptInformation += @{Data = "Registration state"; Value = $PrepMachineRegistrationState; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		
	}
	ElseIf($Text)
	{
		Line 0 "Details"
		Line 1 "Name`t`t: " $AppDisk.AppDiskName
		Line 1 "Description`t: " $AppDisk.Description
		Line 1 "Resources`t: " $AppDisk.HostingUnitName[0]
		$cnt = -1
		ForEach($tmp in $AppDisk.HostingUnitName)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 3 "" $tmp
			}
		}
		Line 1 "Compatibility`t: " $xCompatibility
		Line 0 ""
		
		Line 0 "Preparation machine"
		Line 1 "Machine name`t`t: " $PrepMachineName
		Line 1 "Power state`t`t: " $PrepMachinePowerState
		Line 1 "Registration state`t: " $PrepMachineRegistrationState
		Line 0 ""
		
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Details"
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$AppDisk.AppDiskName,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$AppDisk.Description,$htmlwhite))
		$rowdata += @(,('Resources',($htmlsilver -bor $htmlbold),$AppDisk.HostingUnitName[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $AppDisk.HostingUnitName)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Resources',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Compatibility',($htmlsilver -bor $htmlbold),$xCompatibility,$htmlwhite))
		$msg = ""
		$columnWidths = @("100px","250px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		
		WriteHTMLLine 4 0 "Preparation machine"
		$columnHeaders = @("Machine name",($htmlsilver -bor $htmlbold),$PrepMachineName,$htmlwhite)
		$rowdata += @(,('Power state',($htmlsilver -bor $htmlbold),$PrepMachinePowerState,$htmlwhite))
		$rowdata += @(,('Registration state',($htmlsilver -bor $htmlbold),$PrepMachineRegistrationState,$htmlwhite))
		$msg = ""
		$columnWidths = @("100px","250px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
	}
}

Function OutputAppDiskApplications
{
	Param([object] $AppDisk)
	
	Write-Verbose "$(Get-Date): `t`tProcessing applications"
	$Apps = $AppDisk.Packages
	
	If($Null -ne $Apps)
	{
		$Apps = $Apps | Sort Name
		
		OutputApplicationsOnStartMenu $Apps
		OutputInstalledPackages $Apps
	}
}

Function OutputApplicationsOnStartMenu
{
	Param([object] $Apps)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Applications on Start Menu"
		[System.Collections.Hashtable[]] $AppWordTable = @();
		ForEach($App in $Apps)
		{
			$WordTableRowHash = @{ 
			AppName = $App.Name;
			}

			$AppWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $AppWordTable `
		-Columns AppName `
		-Headers "Name" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Applications on Start Menu"
		Line 0 ""
		$cnt = -1
		ForEach($App in $Apps)
		{
			$cnt++
			If($cnt -eq 0)
			{
				Line 1 "Name:   " $App.Name
			}
			Else
			{
				Line 2 "" $App.Name
			}
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Applications on Start Menu"
		$rowdata = @()
		ForEach($App in $Apps)
		{
			$rowdata += @(,(
			$App.Name,$htmlwhite))
		}

		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputInstalledPackages
{
	Param([object] $Apps)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Installed packages"
		[System.Collections.Hashtable[]] $AppWordTable = @();
		ForEach($App in $Apps)
		{
			$WordTableRowHash = @{ 
			AppName = $App.Name;
			AppPublisher = $App.Manufacturer;
			AppVersion = $App.Version;
			}

			$AppWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $AppWordTable `
		-Columns AppName,AppPublisher,AppVersion `
		-Headers "Name", "Publisher", "Version" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Installed packages"
		Line 0 ""
		ForEach($App in $Apps)
		{
			Line 1 "Name`t`t: " $App.Name
			Line 1 "Publisher`t: " $App.Manufacturer
			Line 1 "Version`t`t: " $App.Version
			Line 0 ""
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Installed packages"
		$rowdata = @()
		ForEach($App in $Apps)
		{
			$rowdata += @(,($App.Name,$htmlwhite,
							$App.Manufacturer,$htmlwhite,
							$App.Version,$htmlwhite))
		}

		$columnHeaders = @('Name',($htmlsilver -bor $htmlbold),
							'Publisher',($htmlsilver -bor $htmlbold),
							'Version',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputAppDiskDeliveryGroups
{
	Param([object] $AppDisk)
	
	Write-Verbose "$(Get-Date): `t`tProcessing delivery groups"
	
	$DGs = Get-BrokerDesktopGroup @XDParams2 -AppDisk $AppDisk.AppDiskUid
	
	If($? -and $Null -ne $DGs)
	{
		[array]$DeliveryGroups = $DGs | Sort Name
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 "Delivery Groups"
			[System.Collections.Hashtable[]] $DGWordTable = @();
			ForEach($Group in $DeliveryGroups)
			{
				$WordTableRowHash = @{ 
				DGName = $Group.Name;
				}

				$DGWordTable += $WordTableRowHash;
			}
			$Table = AddWordTable -Hashtable $DGWordTable `
			-Columns DGName `
			-Headers "Name" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($Text)
		{
			Line 0 "Delivery Groups"
			Line 0 ""
			$cnt = -1
			ForEach($Group in $DeliveryGroups)
			{
				$cnt++
				If($cnt -eq 0)
				{
					Line 1 "Name:   " $Group.Name
				}
				Else
				{
					Line 2 "" $Group.Name
				}
			}
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 4 0 "Delivery Groups"
			$rowdata = @()
			ForEach($Group in $DeliveryGroups)
			{
				$rowdata += @(,(
				$Group.Name,$htmlwhite))
			}

			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf($? -and ($Null -eq $DGs))
	{
		$txt = "There are no Delivery Groups for $($AppDisk.AppDiskName)"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Delivery Groups for $($AppDisk.AppDiskName)"
		OutputWarning $txt
	}
}

Function OutputAppDiskAdmins
{
	Param([object] $AppDisk)
	
	Write-Verbose "$(Get-Date): `t`tProcessing administrators"
	$Admins = GetAdmins "AppDisk"
	
	If($? -and ($Null -ne $Admins))
	{
		OutputAdminsForDetails $Admins
	}
	ElseIf($? -and ($Null -eq $Admins))
	{
		$txt = "There are no administrators for $($AppDisk.AppDiskName)"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve administrators for $($AppDisk.AppDiskName)"
		OutputWarning $txt
	}
}
#endregion

#region function to output machine/desktop details
Function OutputMachineDetails
{
	Param([object] $Machine)
	
	Write-Verbose "$(Get-Date): `t`tOutput Machine $($Machine.HostedMachineName)"
	
	$xAssociatedUserFullNames = @()
	ForEach($Value in $Machine.AssociatedUserFullNames)
	{
		$xAssociatedUserFullNames += "$($Value)"
	}
		
	$xAssociatedUserNames = @()
	ForEach($Value in $Machine.AssociatedUserNames)
	{
		$xAssociatedUserNames += "$($Value)"
	}
	
	$xAssociatedUserUPNs = @()
	ForEach($Value in $Machine.AssociatedUserUPNs)
	{
		$xAssociatedUserUPNs += "$($Value)"
	}

	$xDesktopConditions = @()
	ForEach($Value in $Machine.DesktopConditions)
	{
		$xDesktopConditions += "$($Value)"
	}
	
	If($xDesktopConditions.Count -eq 0)
	{
		$xDesktopConditions += "-"
	}

	$xAllocationType = ""
	If($Machine.AllocationType -eq "Static")
	{
		$xAllocationType = "Private"
	}
	Else
	{
		$xAllocationType = $Machine.AllocationType
	}

	$xInMaintenanceMode = ""
	If($Machine.InMaintenanceMode)
	{
		$xInMaintenanceMode = "On"
	}
	Else
	{
		$xInMaintenanceMode ="Off"
	}

	$xWindowsConnectionSetting = ""
	If($Machine.SessionSupport -eq "MultiSession")
	{
		Switch ($Machine.WindowsConnectionSetting)
		{
			"LogonEnabled"			{$xWindowsConnectionSetting = "Logon Enabled"}
			"Draining"				{$xWindowsConnectionSetting = "Draining"}
			"DrainingUntilRestart"	{$xWindowsConnectionSetting = "Draining until restart"}
			"LogonDisabled"			{$xWindowsConnectionSetting = "Logons Disabled"}
			Default					{$xWindowsConnectionSetting = "Unable to determine WindowsConnectionSetting: $($Machine.WindowsConnectionSetting)"; Break}
		}
	}

	$xIsPhysical = ""
	If($Machine.IsPhysical)
	{
		$xIsPhysical = "Physical"
	}
	Else
	{
		$xIsPhysical ="Virtual"
	}

	$xPvdStage = "-"
	If($xAllocationType -eq "Private")
	{
		If($Machine.PvdStage -eq "None")
		{
			$xPvdStage = "Ready"
		}
		Else
		{
			$xPvdStage = $Machine.PvdStage
		}
	}

	$xSummaryState = ""
	If($Machine.SummaryState -eq "InUse")
	{
		$xSummaryState = "In Use"
	}
	Else
	{
		$xSummaryState = $Machine.SummaryState
	}

	$xTags = @()
	ForEach($Value in $Machine.Tags)
	{
		$xTags += "$($Value)"
	}
	
	If($xTags.Count -eq 0)
	{
		$xTags += "-"
	}

	$xApplicationsInUse = @()
	ForEach($value in $Machine.ApplicationsInUse)
	{
		$xApplicationsInUse += "$($value)"
	}
	
	If($xApplicationsInUse.Count -eq 0)
	{
		$xApplicationsInUse += "-"
	}

	$xPublishedApplications = @()
	ForEach($value in $Machine.PublishedApplications)
	{
		$xPublishedApplications += "$($value)"
	}
	
	If($xPublishedApplications.Count -eq 0)
	{
		$xPublishedApplications += "-"
	}

	$xSessionSecureIcaActive = ""
	If($Machine.SessionSecureIcaActive)
	{
		$xSessionSecureIcaActive = "Yes"
	}
	Else
	{
		$xSessionSecureIcaActive = "-"
	}

	$xLastDeregistrationReason = ""
	Switch ($Machine.LastDeregistrationReason)
	{
		$Null									{$xLastDeregistrationReason = ""; Break}
		"AgentAddressResolutionFailed"			{$xLastDeregistrationReason = "Agent Address Resolution Failed"; Break}
		"AgentNotContactable"					{$xLastDeregistrationReason = "Agent Not Contactable"; Break}
		"AgentRejectedSettingsUpdate"			{$xLastDeregistrationReason = "Agent Rejected Settings Update"; Break}
		"AgentRequested"						{$xLastDeregistrationReason = "Agent Requested"; Break}
		"AgentShutdown"							{$xLastDeregistrationReason = "Agent Shutdown"; Break}
		"AgentSuspended"						{$xLastDeregistrationReason = "Agent Suspended"; Break}
		"AgentWrongActiveDirectoryOU"			{$xLastDeregistrationReason = "Agent Wrong Active Directory OU"; Break}
		"BrokerRegistrationLimitReached"		{$xLastDeregistrationReason = "Broker Registration Limit Reached"; Break}
		"ContactLost"							{$xLastDeregistrationReason = "Contact Lost"; Break}
		"DesktopRemoved"						{$xLastDeregistrationReason = "Desktop Removed"; Break}
		"DesktopRestart"						{$xLastDeregistrationReason = "Desktop Restart"; Break}
		"EmptyRegistrationRequest"				{$xLastDeregistrationReason = "Empty Registration Request"; Break}
		"FunctionalLevelTooLowForCatalog"		{$xLastDeregistrationReason = "Functional Level Too Low For Catalog"; Break}
		"FunctionalLevelTooLowForDesktopGroup"	{$xLastDeregistrationReason = "Functional Level Too Low For Desktop Group"; Break}
		"IncompatibleVersion"					{$xLastDeregistrationReason = "Incompatible Version"; Break}
		"InconsistentRegistrationCapabilities"	{$xLastDeregistrationReason = "Inconsistent Registration Capabilities"; Break}
		"InvalidRegistrationRequest"			{$xLastDeregistrationReason = "Invalid Registration Request"; Break}
		"MissingAgentVersion"					{$xLastDeregistrationReason = "Missing Agent Version"; Break}
		"MissingRegistrationCapabilities"		{$xLastDeregistrationReason = "Missing Registration Capabilities"; Break}
		"NotLicensedForFeature"					{$xLastDeregistrationReason = "Not Licensed For Feature"; Break}
		"PowerOff"								{$xLastDeregistrationReason = "Power Off"; Break}
		"SendSettingsFailure"					{$xLastDeregistrationReason = "Send Settings Failure"; Break}
		"SessionAuditFailure"					{$xLastDeregistrationReason = "Session Audit Failure"; Break}
		"SessionPrepareFailure"					{$xLastDeregistrationReason = "Session Prepare Failure"; Break}
		"SettingsCreationFailure"				{$xLastDeregistrationReason = "Settings Creation Failure"; Break}
		"SingleMultiSessionMismatch"			{$xLastDeregistrationReason = "Single Multi Session Mismatch"; Break}
		"UnknownError"							{$xLastDeregistrationReason = "Unknown Error"; Break}
		"UnsupportedCredentialSecurityVersion"	{$xLastDeregistrationReason = "Unsupported Credential Security Version"; Break} 
		Default {$xLastDeregistrationReason = "Unable to determine LastDeregistrationReason: $($Machine.LastDeregistrationReason)"; Break}
	}

	$xPersistUserChanges = ""
	Switch ($Machine.PersistUserChanges)
	{
		"OnLocal"	{$xPersistUserChanges = "On Local"; Break}
		"Discard"	{$xPersistUserChanges = "Discard"; Break}
		"OnPvD"		{$xPersistUserChanges = "On personal vDisk"; Break}
		Default		{$xPersistUserChanges = "Unable to determine the value of PersistUserChanges: $($Machine.PersistUserChanges)"; Break}
	}

	$xWillShutdownAfterUse = ""
	If($Machine.WillShutdownAfterUse)
	{
		$xWillShutdownAfterUse = "Yes"
	}
	Else
	{
		$xWillShutdownAfterUse = "No"
	}

	$xSessionSmartAccessTags = @()
	ForEach($value in $Machine.SessionSmartAccessTags)
	{
		$xSessionSmartAccessTags += "$($value)"
	}
	
	If($xSessionSmartAccessTags.Count -eq 0)
	{
		$xSessionSmartAccessTags += "-"
	}
	
	Switch($Machine.FaultState)
	{
		"None"			{$xMachineFaultState = "None"}
		"FailedToStart"	{$xMachineFaultState = "Failed to start"}
		"StuckOnBoot"	{$xMachineFaultState = "Stuck on boot"}
		"Unregistered"	{$xMachineFaultState = "Unregistered"}
		"MaxCapacity"	{$xMachineFaultState = "Maximum capacity"}
		Default			{$xMachineFaultState = "Unable to determine the value of FaultState: $($Machine.FaultState)"; Break}
	}

	If($Null -eq $Machine.SessionLaunchedViaHostName)
	{
		$xSessionLaunchedViaHostName = "-"
	}
	Else
	{
		$xSessionLaunchedViaHostName = $Machine.SessionLaunchedViaHostName
	}
	
	If($Null -eq $Machine.SessionLaunchedViaIP)
	{
		$xSessionLaunchedViaIP = "-"
	}
	Else
	{
		$xSessionLaunchedViaIP = $Machine.SessionLaunchedViaIP
	}
	
	If($Null -eq $Machine.SessionClientAddress)
	{
		$xSessionClientAddress = "-"
	}
	Else
	{
		$xSessionClientAddress = $Machine.SessionClientAddress
	}
	
	If($Null -eq $Machine.SessionClientName)
	{
		$xSessionClientName = "-"
	}
	Else
	{
		$xSessionClientName = $Machine.SessionClientName
	}
	
	If($Null -eq $Machine.SessionClientVersion)
	{
		$xSessionClientVersion = "-"
	}
	Else
	{
		$xSessionClientVersion = $Machine.SessionClientVersion
	}
	
	If($Null -eq $Machine.SessionConnectedViaHostName)
	{
		$xSessionConnectedViaHostName = "-"
	}
	Else
	{
		$xSessionConnectedViaHostName = $Machine.SessionConnectedViaHostName
	}
	
	If($Null -eq $Machine.SessionConnectedViaIP)
	{
		$xSessionConnectedViaIP = "-"
	}
	Else
	{
		$xSessionConnectedViaIP = $Machine.SessionConnectedViaIP
	}
	
	If($Null -eq $Machine.SessionProtocol)
	{
		$xSessionProtocol = "-"
	}
	Else
	{
		$xSessionProtocol = $Machine.SessionProtocol
	}
	
	If($Null -eq $Machine.SessionStateChangeTime)
	{
		$xSessionStateChangeTime = "-"
	}
	Else
	{
		$xSessionStateChangeTime = $Machine.SessionStateChangeTime
	}
	
	If($Null -eq $Machine.SessionState)
	{
		$xSessionState = "-"
	}
	Else
	{
		$xSessionState = $Machine.SessionState
	}
	
	If($Null -eq $Machine.SessionUserName)
	{
		$xSessionUserName = "-"
	}
	Else
	{
		$xSessionUserName = $Machine.SessionUserName
	}
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 3 0 $Machine.DNSName
		If($Machine.SessionSupport -eq "MultiSession")
		{
			WriteWordLine 4 0 "Machine"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Name"; Value = $Machine.DNSName; }
			$ScriptInformation += @{Data = "Machine Catalog"; Value = $Machine.CatalogName; }
			$ScriptInformation += @{Data = "Delivery Group"; Value = $Machine.DesktopGroupName; }
			$ScriptInformation += @{Data = "User Display Name"; Value = $xAssociatedUserFullNames[0]; }
			$cnt = -1
			ForEach($tmp in $xAssociatedUserFullNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "User"; Value = $xAssociatedUserNames[0]; }
			$cnt = -1
			ForEach($tmp in $xAssociatedUserNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "UPN"; Value = $xAssociatedUserUPNs[0]; }
			$cnt = -1
			ForEach($tmp in $xAssociatedUserUPNs)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Desktop Conditions"; Value = $xDesktopConditions[0]; }
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Allocation Type"; Value = $xAllocationType; }
			$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xInMaintenanceMode; }
			$ScriptInformation += @{Data = "Windows Connection Setting"; Value = $xWindowsConnectionSetting; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned; }
			$ScriptInformation += @{Data = "Is Physical"; Value = $xIsPhysical; }
			$ScriptInformation += @{Data = "Provisioning Type"; Value = $Machine.ProvisioningType; }
			$ScriptInformation += @{Data = "PvD State"; Value = $xPvdStage; }
			$ScriptInformation += @{Data = "Scheduled Reboot"; Value = $Machine.ScheduledReboot; }
			If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
			{
				$ScriptInformation += @{Data = "Zone"; Value = $Machine.ZoneName; }
			}
			$ScriptInformation += @{Data = "Summary State"; Value = $xSummaryState; }
			$ScriptInformation += @{Data = "Tags"; Value = $xTags[0]; }
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Load Index"; Value = $Machine.LoadIndex; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Machine Details"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Agent Version"; Value = $Machine.AgentVersion; }
			$ScriptInformation += @{Data = "IP Address"; Value = $Machine.IPAddress; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Applications"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			
			$ScriptInformation += @{Data = "Applications In Use"; Value = $xApplicationsInUse[0]; }
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Published Applications"; Value = $xPublishedApplications[0]; }
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Registration"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Broker"; Value = $Machine.ControllerDNSName; }
			$ScriptInformation += @{Data = "Last registration failure"; Value = $xLastDeregistrationReason; }
			$ScriptInformation += @{Data = "Last registration failure time"; Value = $Machine.LastDeregistrationTime; }
			$ScriptInformation += @{Data = "Registration State"; Value = $Machine.RegistrationState; }
			$ScriptInformation += @{Data = "Fault State"; Value = $xMachineFaultState; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		
			WriteWordLine 4 0 "Hosting"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "VM"; Value = $Machine.HostedMachineName; }
			$ScriptInformation += @{Data = "Hosting Server Name"; Value = $Machine.HostingServerName; }
			$ScriptInformation += @{Data = "Connection"; Value = $Machine.HypervisorConnectionName ; }
			$ScriptInformation += @{Data = "Pending Update"; Value = $Machine.ImageOutOfDate; }
			$ScriptInformation += @{Data = "Persist User Changes"; Value = $xPersistUserChanges; }
			$ScriptInformation += @{Data = "Power Action Pending"; Value = $Machine.PowerActionPending; }
			$ScriptInformation += @{Data = "Power State"; Value = $Machine.PowerState; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Connection"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Last Connection Time"; Value = $Machine.LastConnectionTime.ToString() ; }
			$ScriptInformation += @{Data = "Last Connection User"; Value = $Machine.LastConnectionUser; }
			$ScriptInformation += @{Data = "Secure ICA Active"; Value = $xSessionSecureIcaActive ; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Session Details"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Launched Via"; Value = $xSessionLaunchedViaHostName; }
			$ScriptInformation += @{Data = "Launched Via (IP)"; Value = $xSessionLaunchedViaIP; }
			$ScriptInformation += @{Data = "SmartAccess Filters"; Value = $xSessionSmartAccessTags[0]; }
			$cnt = -1
			ForEach($tmp in $xSessionSmartAccessTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Session"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Session Count"; Value = $Machine.SessionCount.ToString(); }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($Machine.SessionSupport -eq "SingleSession")
		{
			WriteWordLine 4 0 "Machine"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Name"; Value = $Machine.DNSName; }
			$ScriptInformation += @{Data = "Machine Catalog"; Value = $Machine.CatalogName; }
			$ScriptInformation += @{Data = "Delivery Group"; Value = $Machine.DesktopGroupName; }
			$ScriptInformation += @{Data = "User Display Name"; Value = $xAssociatedUserFullNames[0]; }
			$cnt = -1
			ForEach($tmp in $xAssociatedUserFullNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "User"; Value = $xAssociatedUserNames[0]; }
			$cnt = -1
			ForEach($tmp in $xAssociatedUserNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "UPN"; Value = $xAssociatedUserUPNs[0]; }
			$cnt = -1
			ForEach($tmp in $xAssociatedUserUPNs)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Desktop Conditions"; Value = $xDesktopConditions[0]; }
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Allocation Type"; Value = $xAllocationType; }
			$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xInMaintenanceMode; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned; }
			$ScriptInformation += @{Data = "Is Physical"; Value = $xIsPhysical; }
			$ScriptInformation += @{Data = "Provisioning Type"; Value = $Machine.ProvisioningType; }
			$ScriptInformation += @{Data = "PvD State"; Value = $xPvdStage; }
			If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
			{
				$ScriptInformation += @{Data = "Zone"; Value = $Machine.ZoneName; }
			}
			$ScriptInformation += @{Data = "Scheduled Reboot"; Value = $Machine.ScheduledReboot; }
			$ScriptInformation += @{Data = "Summary State"; Value = $xSummaryState; }
			$ScriptInformation += @{Data = "Tags"; Value = $xTags[0]; }
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Machine Details"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Agent Version"; Value = $Machine.AgentVersion; }
			$ScriptInformation += @{Data = "IP Address"; Value = $Machine.IPAddress; }
			$ScriptInformation += @{Data = "Is Assigned"; Value = $Machine.IsAssigned; }
			$ScriptInformation += @{Data = "OS Type"; Value = $Machine.OSType; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Applications"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			
			$ScriptInformation += @{Data = "Applications In Use"; Value = $xApplicationsInUse[0]; }
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			$ScriptInformation += @{Data = "Published Applications"; Value = $xPublishedApplications[0]; }
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Connection"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Client (IP)"; Value = $xSessionClientAddress; }
			$ScriptInformation += @{Data = "Client"; Value = $xSessionClientName; }
			$ScriptInformation += @{Data = "Plug-in Version"; Value = $xSessionClientVersion; }
			$ScriptInformation += @{Data = "Connected Via"; Value = $xSessionConnectedViaHostName; }
			$ScriptInformation += @{Data = "Connected Via (IP)"; Value = $xSessionConnectedViaIP; }
			$ScriptInformation += @{Data = "Last Connection Time"; Value = $Machine.LastConnectionTime.ToString() ; }
			$ScriptInformation += @{Data = "Last Connection User"; Value = $Machine.LastConnectionUser; }
			$ScriptInformation += @{Data = "Connection Type"; Value = $xSessionProtocol; }
			$ScriptInformation += @{Data = "Secure ICA Active"; Value = $xSessionSecureIcaActive ; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Registration"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Broker"; Value = $Machine.ControllerDNSName; }
			$ScriptInformation += @{Data = "Last registration failure"; Value = $xLastDeregistrationReason; }
			$ScriptInformation += @{Data = "Last registration failure time"; Value = $Machine.LastDeregistrationTime; }
			$ScriptInformation += @{Data = "Registration State"; Value = $Machine.RegistrationState; }
			$ScriptInformation += @{Data = "Fault State"; Value = $xMachineFaultState; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Hosting"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "VM"; Value = $Machine.HostedMachineName; }
			$ScriptInformation += @{Data = "Hosting Server Name"; Value = $Machine.HostingServerName; }
			$ScriptInformation += @{Data = "Connection"; Value = $Machine.HypervisorConnectionName ; }
			$ScriptInformation += @{Data = "Pending Update"; Value = $Machine.ImageOutOfDate; }
			$ScriptInformation += @{Data = "Persist User Changes"; Value = $xPersistUserChanges; }
			$ScriptInformation += @{Data = "Power Action Pending"; Value = $Machine.PowerActionPending; }
			$ScriptInformation += @{Data = "Power State"; Value = $Machine.PowerState; }
			$ScriptInformation += @{Data = "Will Shutdown After Use"; Value = $xWillShutdownAfterUse; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Session Details"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Launched Via"; Value = $xSessionLaunchedViaHostName; }
			$ScriptInformation += @{Data = "Launched Via (IP)"; Value = $xSessionLaunchedViaIP; }
			$ScriptInformation += @{Data = "Session Change Time"; Value = $xSessionStateChangeTime; }
			$ScriptInformation += @{Data = "SmartAccess Filters"; Value = $xSessionSmartAccessTags[0]; }
			$cnt = -1
			ForEach($tmp in $xSessionSmartAccessTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 4 0 "Session"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Session State"; Value = $xSessionState; }
			$ScriptInformation += @{Data = "Current User"; Value = $xSessionUserName; }
			$ScriptInformation += @{Data = "Start Time"; Value = $xSessionStateChangeTime; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		If($Machine.SessionSupport -eq "MultiSession")
		{
			Line 1 "Machine"
			Line 2 "Name`t`t`t`t: " $Machine.DNSName
			Line 2 "Machine Catalog`t`t`t: " $Machine.CatalogName
			Line 2 "Delivery Group`t`t`t: " $Machine.DesktopGroupName
			Line 2 "User Display Name`t`t: " $xAssociatedUserFullNames[0]
			$cnt = -1
			ForEach($tmp in $xAssociatedUserFullNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "User`t`t`t`t: " $xAssociatedUserNames[0]
			$cnt = -1
			ForEach($tmp in $xAssociatedUserNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "UPN`t`t`t`t: " $xAssociatedUserUPNs[0]
			$cnt = -1
			ForEach($tmp in $xAssociatedUserUPNs)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Desktop Conditions`t`t: " $xDesktopConditions[0]
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Allocation Type`t`t`t: " $xAllocationType
			Line 2 "Maintenance Mode`t`t: " $xInMaintenanceMode
			Line 2 "Windows Connection Setting`t: " $xWindowsConnectionSetting
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned
			Line 2 "Is Physical`t`t`t: " $xIsPhysical
			Line 2 "Provisioning Type`t`t: " $Machine.ProvisioningType
			Line 2 "PvD State`t`t`t: " $xPvdStage
			Line 2 "Scheduled Reboot`t`t: " $Machine.ScheduledReboot
			If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
			{
				Line 2 "Zone`t`t`t`t: " $Machine.ZoneName
			}
			Line 2 "Summary State`t`t`t: " $xSummaryState
			Line 2 "Tags`t`t`t`t: " $xTags[0]
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Load Index`t`t`t: " $Machine.LoadIndex
			Line 0 ""

			Line 1 "Machine Details"
			Line 2 "Agent Version`t`t`t: " $Machine.AgentVersion
			Line 2 "IP Address`t`t`t: " $Machine.IPAddress
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned
			Line 0 ""
			
			Line 1 "Applications"
			Line 2 "Applications In Use`t`t: " $xApplicationsInUse[0]
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Published Applications`t`t: " $xPublishedApplications[0]
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 0 ""
			
			Line 1 "Registration"
			Line 2 "Broker`t`t`t`t: " $Machine.ControllerDNSName
			Line 2 "Last registration failure`t: " $xLastDeregistrationReason
			Line 2 "Last registration failure time`t: " $Machine.LastDeregistrationTime
			Line 2 "Registration State`t`t: " $Machine.RegistrationState
			Line 2 "Fault State`t`t`t: " $xMachineFaultState
			Line 0 ""
			
			Line 1 "Hosting"
			Line 2 "VM`t`t`t`t: " $Machine.HostedMachineName
			Line 2 "Hosting Server Name`t`t: " $Machine.HostingServerName
			Line 2 "Connection`t`t`t: " $Machine.HypervisorConnectionName 
			Line 2 "Pending Update`t`t`t: " $Machine.ImageOutOfDate
			Line 2 "Persist User Changes`t`t: " $xPersistUserChanges
			Line 2 "Power Action Pending`t`t: " $Machine.PowerActionPending
			Line 2 "Power State`t`t`t: " $Machine.PowerState
			Line 0 ""
			
			Line 1 "Connection"
			Line 2 "Last Connection Time`t`t: " $Machine.LastConnectionTime.ToString() 
			Line 2 "Last Connection User`t`t: " $Machine.LastConnectionUser
			Line 2 "Secure ICA Active`t`t: " $xSessionSecureIcaActive 
			Line 0 ""
			
			Line 1 "Session Details"
			Line 2 "Launched Via`t`t`t: " $xSessionLaunchedViaHostName
			Line 2 "Launched Via (IP)`t`t: " $xSessionLaunchedViaIP
			Line 2 "SmartAccess Filters`t`t: " $xSessionSmartAccessTags[0]
			$cnt = -1
			ForEach($tmp in $xSessionSmartAccessTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
			Line 0 ""
			
			Line 1 "Session"
			Line 2 "Session Count`t`t`t: " $Machine.SessionCount.ToString()
			Line 0 ""
		}
		ElseIf($Machine.SessionSupport -eq "SingleSession")
		{
			Line 1 "Machine"
			Line 2 "Name`t`t`t`t: " $Machine.DNSName
			Line 2 "Machine Catalog`t`t`t: " $Machine.CatalogName
			Line 2 "Delivery Group`t`t`t: " $Machine.DesktopGroupName
			Line 2 "User Display Name`t`t: " $xAssociatedUserFullNames[0]
			$cnt = -1
			ForEach($tmp in $xAssociatedUserFullNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "User`t`t`t`t: " $xAssociatedUserNames[0]
			$cnt = -1
			ForEach($tmp in $xAssociatedUserNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "UPN`t`t`t`t: " $xAssociatedUserUPNs[0]
			$cnt = -1
			ForEach($tmp in $xAssociatedUserUPNs)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Desktop Conditions`t`t: " $xDesktopConditions[0]
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Allocation Type`t`t`t: " $xAllocationType
			Line 2 "Maintenance Mode`t`t: " $xInMaintenanceMode
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned
			Line 2 "Is Physical`t`t`t: " $xIsPhysical
			Line 2 "Provisioning Type`t`t: " $Machine.ProvisioningType
			Line 2 "PvD State`t`t`t: " $xPvdStage
			If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
			{
				Line 2 "Zone`t`t`t`t: " $Machine.ZoneName
			}
			Line 2 "Scheduled Reboot`t`t: " $Machine.ScheduledReboot
			Line 2 "Summary State`t`t`t: " $xSummaryState
			Line 2 "Tags`t`t`t`t: " $xTags[0]
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 0 ""

			Line 1 "Machine Details"
			Line 2 "Agent Version`t`t`t: " $Machine.AgentVersion
			Line 2 "IP Address`t`t`t: " $Machine.IPAddress
			Line 2 "Is Assigned`t`t`t: " $Machine.IsAssigned
			Line 2 "OS Type`t`t`t`t: " $Machine.OSType
			Line 0 ""
			
			Line 1 "Applications"
			Line 2 "Applications In Use`t`t: " $xApplicationsInUse[0]
			$cnt = -1
			ForEach($tmp in $xApplicationsInUse)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 2 "Published Applications`t`t: " $xPublishedApplications[0]
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 6 "  " $tmp
				}
			}
			Line 0 ""
			
			Line 1 "Connection"
			Line 2 "Client (IP)`t`t`t: " $xSessionClientAddress
			Line 2 "Client`t`t`t`t: " $xSessionClientName
			Line 2 "Plug-in Version`t`t`t: " $xSessionClientVersion
			Line 2 "Connected Via`t`t`t: " $xSessionConnectedViaHostName
			Line 2 "Connect Via (IP)`t`t: " $xSessionConnectedViaIP
			Line 2 "Last Connection Time`t`t: " $Machine.LastConnectionTime.ToString() 
			Line 2 "Last Connection User`t`t: " $Machine.LastConnectionUser
			Line 2 "Connection Type`t`t`t: " $xSessionProtocol
			Line 2 "Secure ICA Active`t`t: " $xSessionSecureIcaActive 
			Line 0 ""
			
			Line 1 "Registration"
			Line 2 "Broker`t`t`t`t: " $Machine.ControllerDNSName
			Line 2 "Last registration failure`t: " $xLastDeregistrationReason
			Line 2 "Last registration failure time`t: " $Machine.LastDeregistrationTime
			Line 2 "Registration State`t`t: " $Machine.RegistrationState
			Line 2 "Fault State`t`t`t: " $xMachineFaultState
			Line 0 ""
			
			Line 1 "Hosting"
			Line 2 "VM`t`t`t`t: " $Machine.HostedMachineName
			Line 2 "Hosting Server Name`t`t: " $Machine.HostingServerName
			Line 2 "Connection`t`t`t: " $Machine.HypervisorConnectionName 
			Line 2 "Pending Update`t`t`t: " $Machine.ImageOutOfDate
			Line 2 "Persist User Changes`t`t: " $xPersistUserChanges
			Line 2 "Power Action Pending`t`t: " $Machine.PowerActionPending
			Line 2 "Power State`t`t`t: " $Machine.PowerState
			Line 2 "Will Shutdown After Use`t`t: " $xWillShutdownAfterUse
			Line 0 ""
			
			Line 1 "Session Details"
			Line 2 "Launched Via`t`t`t: " $xSessionLaunchedViaHostName
			Line 2 "Launched Via (IP)`t`t: " $xSessionLaunchedViaIP
			Line 2 "Session Change Time`t`t: " $xSessionStateChangeTime
			Line 2 "SmartAccess Filters`t`t: " $xSessionSmartAccessTags[0]
			$cnt = -1
			ForEach($tmp in $xSessionSmartAccessTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
			Line 0 ""
			
			Line 1 "Session"
			Line 2 "Session State`t`t`t: " $xSessionState
			Line 2 "Current User`t`t`t: " $xSessionUserName
			Line 2 "Start Time`t`t`t: " $xSessionStateChangeTime
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		If($Machine.SessionSupport -eq "MultiSession")
		{
			WriteHTMLLine 4 0 "Machine"
			$rowdata = @()

			$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Machine.DNSName,$htmlwhite)
			$rowdata += @(,('Machine Catalog',($htmlsilver -bor $htmlbold),$Machine.CatalogName,$htmlwhite))
			$rowdata += @(,('Delivery Group',($htmlsilver -bor $htmlbold),$Machine.DesktopGroupName,$htmlwhite))
			$rowdata += @(,('User Display Name',($htmlsilver -bor $htmlbold),$xAssociatedUserFullNames[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xAssociatedUserFullNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('User',($htmlsilver -bor $htmlbold),$xAssociatedUserNames[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xAssociatedUserNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('UPN',($htmlsilver -bor $htmlbold),$xAssociatedUserUPNs[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xAssociatedUserUPNs)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Desktop Conditions',($htmlsilver -bor $htmlbold),$xDesktopConditions[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Allocation Type',($htmlsilver -bor $htmlbold),$xAllocationType,$htmlwhite))
			$rowdata += @(,('Maintenance Mode',($htmlsilver -bor $htmlbold),$xInMaintenanceMode,$htmlwhite))
			$rowdata += @(,('Windows Connection Setting',($htmlsilver -bor $htmlbold),$xWindowsConnectionSetting,$htmlwhite))
			$rowdata += @(,('Is Assigned',($htmlsilver -bor $htmlbold),$Machine.IsAssigned,$htmlwhite))
			$rowdata += @(,('Is Physical',($htmlsilver -bor $htmlbold),$xIsPhysical,$htmlwhite))
			$rowdata += @(,('Provisioning Type',($htmlsilver -bor $htmlbold),$Machine.ProvisioningType,$htmlwhite))
			$rowdata += @(,('PvD State',($htmlsilver -bor $htmlbold),$xPvdStage,$htmlwhite))
			$rowdata += @(,('Scheduled Reboot',($htmlsilver -bor $htmlbold),$Machine.ScheduledReboot,$htmlwhite))
			If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
			{
				$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$Machine.ZoneName,$htmlwhite))
			}
			$rowdata += @(,('Summary State',($htmlsilver -bor $htmlbold),$xSummaryState,$htmlwhite))
			$rowdata += @(,('Tags',($htmlsilver -bor $htmlbold),$xTags[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Load Index',($htmlsilver -bor $htmlbold),$Machine.LoadIndex,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Machine Details"
			$rowdata = @()
			$columnHeaders = @("Agent Version",($htmlsilver -bor $htmlbold),$Machine.AgentVersion,$htmlwhite)
			$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$Machine.IPAddress,$htmlwhite))
			$rowdata += @(,('Is Assigned',($htmlsilver -bor $htmlbold),$Machine.IsAssigned,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Applications"
			$rowdata = @()
			$columnHeaders = @("Applications In Use",($htmlsilver -bor $htmlbold),$xApplicationsInUse[0],$htmlwhite)
			$cnt = -1
			ForEach($tmp in $xApplicationsInUSe)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Published Applications',($htmlsilver -bor $htmlbold),$xPublishedApplications[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Registration"
			$rowdata = @()
			$columnHeaders = @("Broker",($htmlsilver -bor $htmlbold),$Machine.ControllerDNSName,$htmlwhite)
			$rowdata += @(,('Last registration failure',($htmlsilver -bor $htmlbold),$xLastDeregistrationReason,$htmlwhite))
			$rowdata += @(,('Last registration failure time',($htmlsilver -bor $htmlbold),$Machine.LastDeregistrationTime,$htmlwhite))
			$rowdata += @(,('Registration State',($htmlsilver -bor $htmlbold),$Machine.RegistrationState,$htmlwhite))
			$rowdata += @(,('Fault State',($htmlsilver -bor $htmlbold),$xMachineFaultState,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Hosting"
			$rowdata = @()
			$columnHeaders = @("VM",($htmlsilver -bor $htmlbold),$Machine.HostedMachineName,$htmlwhite)
			$rowdata += @(,('Hosting Server Name',($htmlsilver -bor $htmlbold),$Machine.HostingServerName,$htmlwhite))
			$rowdata += @(,('Connection',($htmlsilver -bor $htmlbold),$Machine.HypervisorConnectionName,$htmlwhite))
			$rowdata += @(,('Pending Update',($htmlsilver -bor $htmlbold),$Machine.ImageOutOfDate,$htmlwhite))
			$rowdata += @(,('Persist User Changes',($htmlsilver -bor $htmlbold),$xPersistUserChanges,$htmlwhite))
			$rowdata += @(,('Power Action Pending',($htmlsilver -bor $htmlbold),$Machine.PowerActionPending,$htmlwhite))
			$rowdata += @(,('Power State',($htmlsilver -bor $htmlbold),$Machine.PowerState,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Connection"
			$rowdata = @()
			$columnHeaders = @("Last Connection Time",($htmlsilver -bor $htmlbold),$Machine.LastConnectionTime.ToString(),$htmlwhite)
			$rowdata += @(,('Last Connection User',($htmlsilver -bor $htmlbold),$Machine.LastConnectionUser,$htmlwhite))
			$rowdata += @(,('Secure ICA Active',($htmlsilver -bor $htmlbold),$xSessionSecureIcaActive,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Session Details"
			$rowdata = @()
			$columnHeaders = @("Launched Via",($htmlsilver -bor $htmlbold),$xSessionLaunchedViaHostName,$htmlwhite)
			$rowdata += @(,('Launched Via (IP)',($htmlsilver -bor $htmlbold),$xSessionLaunchedViaIP,$htmlwhite))
			$rowdata += @(,('SmartAccess Filters',($htmlsilver -bor $htmlbold),$xSessionSmartAccessTags[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xSessionSmartAccessTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Session"
			$rowdata = @()
			$columnHeaders = @("Session Count",($htmlsilver -bor $htmlbold),$Machine.SessionCount.ToString(),$htmlwhite)

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "
		}
		ElseIf($Machine.SessionSupport -eq "SingleSession")
		{
			WriteHTMLLine 4 0 "Machine"
			$rowdata = @()

			$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Machine.DNSName,$htmlwhite)
			$rowdata += @(,('Machine Catalog',($htmlsilver -bor $htmlbold),$Machine.CatalogName,$htmlwhite))
			$rowdata += @(,('Delivery Group',($htmlsilver -bor $htmlbold),$Machine.DesktopGroupName,$htmlwhite))
			$rowdata += @(,('User Display Name',($htmlsilver -bor $htmlbold),$xAssociatedUserFullNames[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xAssociatedUserFullNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('User',($htmlsilver -bor $htmlbold),$xAssociatedUserNames[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xAssociatedUserNames)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('UPN',($htmlsilver -bor $htmlbold),$xAssociatedUserUPNs[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xAssociatedUserUPNs)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Desktop Conditions',($htmlsilver -bor $htmlbold),$xDesktopConditions[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xDesktopConditions)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Allocation Type',($htmlsilver -bor $htmlbold),$xAllocationType,$htmlwhite))
			$rowdata += @(,('Maintenance Mode',($htmlsilver -bor $htmlbold),$xInMaintenanceMode,$htmlwhite))
			$rowdata += @(,('Is Assigned',($htmlsilver -bor $htmlbold),$Machine.IsAssigned,$htmlwhite))
			$rowdata += @(,('Is Physical',($htmlsilver -bor $htmlbold),$xIsPhysical,$htmlwhite))
			$rowdata += @(,('Provisioning Type',($htmlsilver -bor $htmlbold),$Machine.ProvisioningType,$htmlwhite))
			$rowdata += @(,('PvD State',($htmlsilver -bor $htmlbold),$xPvdStage,$htmlwhite))
			If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
			{
				$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$Machine.ZoneName,$htmlwhite))
			}
			$rowdata += @(,('Scheduled Reboot',($htmlsilver -bor $htmlbold),$Machine.ScheduledReboot,$htmlwhite))
			$rowdata += @(,('Summary State',($htmlsilver -bor $htmlbold),$xSummaryState,$htmlwhite))
			$rowdata += @(,('Tags',($htmlsilver -bor $htmlbold),$xTags[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Machine Details"
			$rowdata = @()
			$columnHeaders = @("Agent Version",($htmlsilver -bor $htmlbold),$Machine.AgentVersion,$htmlwhite)
			$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$Machine.IPAddress,$htmlwhite))
			$rowdata += @(,('Is Assigned',($htmlsilver -bor $htmlbold),$Machine.IsAssigned,$htmlwhite))
			$rowdata += @(,('OS Type',($htmlsilver -bor $htmlbold),$Machine.OSType,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Applications"
			$rowdata = @()
			$columnHeaders = @("Applications In Use",($htmlsilver -bor $htmlbold),$xApplicationsInUse[0],$htmlwhite)
			$cnt = -1
			ForEach($tmp in $xApplicationsInUSe)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			$rowdata += @(,('Published Applications',($htmlsilver -bor $htmlbold),$xPublishedApplications[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xPublishedApplications)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Connection"
			$rowdata = @()
			$columnHeaders = @("Client (IP)",($htmlsilver -bor $htmlbold),$xSessionClientAddress,$htmlwhite)
			$rowdata += @(,('Client',($htmlsilver -bor $htmlbold),$xSessionClientName,$htmlwhite))
			$rowdata += @(,('Plug-in Version',($htmlsilver -bor $htmlbold),$xSessionClientVersion,$htmlwhite))
			$rowdata += @(,('Connected Via',($htmlsilver -bor $htmlbold),$xSessionConnectedViaHostName,$htmlwhite))
			$rowdata += @(,('Connect Via (IP)',($htmlsilver -bor $htmlbold),$xSessionConnectedViaIP,$htmlwhite))
			$rowdata += @(,('Last Connection Time',($htmlsilver -bor $htmlbold),$Machine.LastConnectionTime.ToString(),$htmlwhite))
			$rowdata += @(,('Last Connection User',($htmlsilver -bor $htmlbold),$Machine.LastConnectionUser,$htmlwhite))
			$rowdata += @(,('Connection Type',($htmlsilver -bor $htmlbold),$xSessionProtocol,$htmlwhite))
			$rowdata += @(,('Secure ICA Active',($htmlsilver -bor $htmlbold),$xSessionSecureIcaActive,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Registration"
			$rowdata = @()
			$columnHeaders = @("Broker",($htmlsilver -bor $htmlbold),$Machine.ControllerDNSName,$htmlwhite)
			$rowdata += @(,('Last registration failure',($htmlsilver -bor $htmlbold),$xLastDeregistrationReason,$htmlwhite))
			$rowdata += @(,('Last registration failure time',($htmlsilver -bor $htmlbold),$Machine.LastDeregistrationTime,$htmlwhite))
			$rowdata += @(,('Registration State',($htmlsilver -bor $htmlbold),$Machine.RegistrationState,$htmlwhite))
			$rowdata += @(,('Fault State',($htmlsilver -bor $htmlbold),$xMachineFaultState,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Hosting"
			$rowdata = @()
			$columnHeaders = @("VM",($htmlsilver -bor $htmlbold),$Machine.HostedMachineName,$htmlwhite)
			$rowdata += @(,('Hosting Server Name',($htmlsilver -bor $htmlbold),$Machine.HostingServerName,$htmlwhite))
			$rowdata += @(,('Connection',($htmlsilver -bor $htmlbold),$Machine.HypervisorConnectionName,$htmlwhite))
			$rowdata += @(,('Pending Update',($htmlsilver -bor $htmlbold),$Machine.ImageOutOfDate,$htmlwhite))
			$rowdata += @(,('Persist User Changes',($htmlsilver -bor $htmlbold),$xPersistUserChanges,$htmlwhite))
			$rowdata += @(,('Power Action Pending',($htmlsilver -bor $htmlbold),$Machine.PowerActionPending,$htmlwhite))
			$rowdata += @(,('Power State',($htmlsilver -bor $htmlbold),$Machine.PowerState,$htmlwhite))
			$rowdata += @(,('Will Shutdown After Use',($htmlsilver -bor $htmlbold),$xWillShutdownAfterUse,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Session Details"
			$rowdata = @()
			$columnHeaders = @("Launched Via",($htmlsilver -bor $htmlbold),$xSessionLaunchedViaHostName,$htmlwhite)
			$rowdata += @(,('Launched Via (IP)',($htmlsilver -bor $htmlbold),$xSessionLaunchedViaIP,$htmlwhite))
			$rowdata += @(,('Session Change Time',($htmlsilver -bor $htmlbold),$xSessionStateChangeTime,$htmlwhite))
			$rowdata += @(,('SmartAccess Filters',($htmlsilver -bor $htmlbold),$xSessionSmartAccessTags[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xSessionSmartAccessTags)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "

			WriteHTMLLine 4 0 "Session"
			$rowdata = @()
			$columnHeaders = @("Session State",($htmlsilver -bor $htmlbold),$xSessionState,$htmlwhite)
			$rowdata += @(,('Current User',($htmlsilver -bor $htmlbold),$xSessionUserName,$htmlwhite))
			$rowdata += @(,('Start Time',($htmlsilver -bor $htmlbold),$xSessionStateChangeTime,$htmlwhite))

			$msg = ""
			$columnWidths = @("200px","250px")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "
		}
	}
}
#endregion

#region Delivery Group functions
Function ProcessDeliveryGroups
{
	Write-Verbose "$(Get-Date): Retrieving Delivery Groups"
	
	$Global:TotalApplicationGroups = 0
	$Global:TotalDesktopGroups = 0
	$Global:TotalAppsAndDesktopGroups = 0

	$AllDeliveryGroups = Get-BrokerDesktopGroup @XDParams2 -SortBy Name 

	If($? -and ($Null -ne $AllDeliveryGroups))
	{
		Write-Verbose "$(Get-Date): `tProcessing Delivery Groups"
		
		#add 16-jun-2015, summary table of delivery groups to match what is shown in Studio
		OutputDeliveryGroupTable $AllDeliveryGroups
		
		ForEach($Group in $AllDeliveryGroups)
		{
			OutputDeliveryGroup $Group
		}
	}
	ElseIf($? -and ($Null -eq $AllDeliveryGroups))
	{
		$txt = "There are no Delivery Groups"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Delivery Groups"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputDeliveryGroupTable 
{
	Param([object] $AllDeliveryGroups)
	
	$txt = "Delivery Groups"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}
	
	ForEach($Group in $AllDeliveryGroups)
	{
		$xSingleSession = ""
		$xState = ""
		$xDeliveryType = ""
		$xGroupName = ""
		[int]$NumApps = 0
		[int]$NumAppGroups = 0
		[int]$NumDesktops = 0
		
		If($Group.SessionSupport -eq "SingleSession")
		{
			$xSingleSession = "Desktop OS"
		}
		Else
		{
			$xSingleSession = "Server OS"
		}
		
		If($Group.InMaintenanceMode)
		{
			$xState = "(Maint) "
		}
		
		$xGroupName = "$($xState)$($Group.Name)"
		
		If($Group.DesktopKind -eq "Private")
		{
			$xSingleSession += " (Static machine assignment)"
		}
		
		$NumApps      = (@(Get-BrokerApplication @XDParams2 -DesktopGroupUid $Group.Uid)).Count
		$NumAppGroups = (@(Get-BrokerApplicationGroup @XDParams2 -DesktopGroupUid $Group.Uid)).Count
		$NumDesktops  = (@(Get-BrokerEntitlementPolicyRule @XDParams2 -DesktopGroupUid $Group.Uid)).Count

		If($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -eq 0)
		{
			$xDeliveryType = "Applications"
			$Global:TotalApplicationGroups++
		}
		ElseIf($NumApps -eq 0 -and $NumAppGroups -eq 0 -and $NumDesktops -gt 0)
		{
			$xDeliveryType = "Desktops"
			$Global:TotalDesktopGroups++
		}
		ElseIf($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -gt 0)
		{
			$xDeliveryType = "Applications and Desktops"
			$Global:TotalAppsAndDesktopGroups++
		}
		Else
		{
			$xDeliveryType = "Delivery type could not be determined: Apps($NumApps) AppGroups($NumAppGroups) Desktops($NumDesktops)"
		}

		If($Group.DeliveryType -eq "DesktopsOnly" -and $Group.DesktopKind -eq "Private")
		{
			$xDeliveryType = "Desktops"
		}
		
		[int]$xAppDisks = $Group.AppDisks.Count
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			DeliveryGroupName = $xGroupName; 
			DeliveryType = $xDeliveryType
			NoOfMachines = $Group.TotalDesktops; 
			SessionsInUse = $Group.Sessions; 
			AppDisks = $xAppDisks
			MachineType = $xSingleSession; 
			Unregistered = $Group.DesktopsUnregistered; 
			Disconnected = $Group.DesktopsDisconnected; 
			}
			$WordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "Delivery Group`t`t: " $xGroupName
			Line 1 "Delivering`t`t: " $xDeliveryType
			Line 1 "No. of machines`t`t: " $Group.TotalDesktops
			Line 1 "Sessions in use`t`t: " $Group.Sessions
			Line 1 "AppDisks`t`t: " $xAppDisks
			Line 1 "Machine type`t`t: " $xSingleSession
			Line 1 "Unregistered`t`t: " $Group.DesktopsUnregistered
			Line 1 "Disconnected`t`t: " $Group.DesktopsDisconnected
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$xGroupName,$htmlwhite,
			$xDeliveryType,$htmlwhite,
			$Group.TotalDesktops,$htmlwhite,
			$Group.Sessions,$htmlwhite,
			$xAppDisks,$htmlwhite,
			$xSingleSession,$htmlwhite,
			$Group.DesktopsUnregistered,$htmlwhite,
			$Group.DesktopsDisconnected,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  DeliveryGroupName, DeliveryType, NoOfMachines, SessionsInUse, AppDisks, MachineType, Unregistered, Disconnected `
		-Headers  "Delivery Group", "Delivering", "No. of machines", "Sessions in use", "AppDisks", "Machine type", "Unregistered", "Disconnected" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 8
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 110;
		$Table.Columns.Item(2).Width = 95;
		$Table.Columns.Item(3).Width = 43;
		$Table.Columns.Item(4).Width = 40;
		$Table.Columns.Item(5).Width = 42;
		$Table.Columns.Item(6).Width = 58;
		$Table.Columns.Item(7).Width = 55;
		$Table.Columns.Item(8).Width = 57;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Delivery Group',($htmlsilver -bor $htmlbold),
		'Delivering',($htmlsilver -bor $htmlbold),
		'No. of machines',($htmlsilver -bor $htmlbold),
		'Sessions in use',($htmlsilver -bor $htmlbold),
		'AppDisks',($htmlsilver -bor $htmlbold),
		'Machine type',($htmlsilver -bor $htmlbold),
		'Unregistered',($htmlsilver -bor $htmlbold),
		'Disconnected',($htmlsilver -bor $htmlbold)
		)

		$columnWidths = @("135","130","50","45","50","65","60","65")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "600"
		WriteHTMLLine 0 0 " "
	}
	
}
Function OutputDeliveryGroup
{
	Param([object] $Group)
	
	Write-Verbose "$(Get-Date): `t`tAdding Delivery Group $($Group.Name)"
	$xSingleSession = ""
	$xState = ""
	If($Group.SessionSupport -eq "SingleSession")
	{
		$xSingleSession = "Desktop OS"
	}
	Else
	{
		$xSingleSession = "Server OS"
	}
	If($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $True)
	{
		$xState = "Maintenance Mode"
	}
	ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $True)
	{
		$xState = "Maintenance Mode"
	}
	ElseIf($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $False)
	{
		$xState = "Enabled"
	}
	ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $False)
	{
		$xState = "Disabled"
	}

	If($MSWord -or$PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 2 0 "Delivery Group: " $Group.Name
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Machine type"; Value = $xSingleSession; }
		$ScriptInformation += @{Data = "Number of machines"; Value = $Group.TotalDesktops; }
		$ScriptInformation += @{Data = "Sessions in use"; Value = $Group.Sessions; }
		$ScriptInformation += @{Data = "Number of applications"; Value = $Group.TotalApplications; }
		$ScriptInformation += @{Data = "State"; Value = $xState; }
		$ScriptInformation += @{Data = "Unregistered"; Value = $Group.DesktopsUnregistered; }
		$ScriptInformation += @{Data = "Disconnected"; Value = $Group.DesktopsDisconnected; }
		$ScriptInformation += @{Data = "Available"; Value = $Group.DesktopsAvailable; }
		$ScriptInformation += @{Data = "In Use"; Value = $Group.DesktopsInUse; }
		$ScriptInformation += @{Data = "Never Registered"; Value = $Group.DesktopsNeverRegistered; }
		$ScriptInformation += @{Data = "Preparing"; Value = $Group.DesktopsPreparing; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Delivery Group: " $Group.Name
		Line 1 "Machine type`t`t: " $xSingleSession
		Line 1 "No. of machines`t`t: " $Group.TotalDesktops
		Line 1 "Sessions in use`t`t: " $Group.Sessions
		Line 1 "No. of applications`t: " $Group.TotalApplications
		Line 1 "State`t`t`t: " $xState
		Line 1 "Unregistered`t`t: " $Group.DesktopsUnregistered
		Line 1 "Disconnected`t`t: " $Group.DesktopsDisconnected
		Line 1 "Available`t`t: " $Group.DesktopsAvailable
		Line 1 "In Use`t`t`t: " $Group.DesktopsInUse
		Line 1 "Never Registered`t: " $Group.DesktopsNeverRegistered
		Line 1 "Preparing`t`t: " $Group.DesktopsPreparing
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 2 0 "Delivery Group: " $Group.Name
		$columnHeaders = @("Machine type",($htmlsilver -bor $htmlbold),$xSingleSession,$htmlwhite)
		$rowdata += @(,('No. of machines',($htmlsilver -bor $htmlbold),$Group.TotalDesktops,$htmlwhite))
		$rowdata += @(,('Sessions in use',($htmlsilver -bor $htmlbold),$Group.Sessions,$htmlwhite))
		$rowdata += @(,('No. of applications',($htmlsilver -bor $htmlbold),$Group.TotalApplications,$htmlwhite))
		$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$xState,$htmlwhite))
		$rowdata += @(,('Unregistered',($htmlsilver -bor $htmlbold),$Group.DesktopsUnregistered,$htmlwhite))
		$rowdata += @(,('Disconnected',($htmlsilver -bor $htmlbold),$Group.DesktopsDisconnected,$htmlwhite))
		$rowdata += @(,('Available',($htmlsilver -bor $htmlbold),$Group.DesktopsAvailable,$htmlwhite))
		$rowdata += @(,('In Use',($htmlsilver -bor $htmlbold),$Group.DesktopsInUse,$htmlwhite))
		$rowdata += @(,('Never Registered',($htmlsilver -bor $htmlbold),$Group.DesktopsNeverRegistered,$htmlwhite))
		$rowdata += @(,('Preparing',($htmlsilver -bor $htmlbold),$Group.DesktopsPreparing,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}
	
	If($DeliveryGroups)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing details"
		$txt = "Delivery Group Details: "
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $txt $Group.Name
		}
		ElseIf($text)
		{
			Line 0 $txt $Group.Name
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 $txt $Group.Name
		}
		OutputDeliveryGroupDetails $Group
		
		Write-Verbose "$(Get-Date): `t`tProcessing applications"
		OutputDeliveryGroupApplicationDetails $Group

		#retrieve machines in delivery group
		$Machines = Get-BrokerMachine -DesktopGroupName $Group.name @XDParams2 -SortBy DNSName
		If($? -and $Null -ne $Machines)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Desktops"
			}
			ElseIf($text)
			{
				Line 0 "Desktops"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 4 0 "Desktops"
			}
			ForEach($Machine in $Machines)
			{
				OutputMachineDetails $Machine
			}
		}
		ElseIf($? -and $Null -eq $Machines)
		{
			$txt = "There are no Machines for Delivery Group $($Group.name)"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Machines for Delivery Group $($Group.name)"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date): `t`tProcessing machine catalogs"
		OutputDeliveryGroupCatalogs $Group

		Write-Verbose "$(Get-Date): `t`tProcessing AppDisks"
		OutputDeliveryGroupAppDisks $Group

		If($DeliveryGroupsUtilization)
		{
			Write-Verbose "$(Get-Date): `t`t`tCreating Delivery Group Utilization report"
			OutputDeliveryGroupUtilization $Group
		}

		Write-Verbose "$(Get-Date): `t`tProcessing Tags"
		OutputDeliveryGroupTags $Group

		Write-Verbose "$(Get-Date): `t`tProcessing Application Groups"
		OutputDeliveryGroupApplicationGroups $Group

		Write-Verbose "$(Get-Date): `t`tProcessing administrators"
		$Admins = GetAdmins "DesktopGroup" $Group.Name
		
		If($? -and ($Null -ne $Admins))
		{
			OutputAdminsForDetails $Admins
		}
		ElseIf($? -and ($Null -eq $Admins))
		{
			$txt = "There are no administrators for $($Group.Name)"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve administrators for $($Group.Name)"
			OutputWarning $txt
		}
		Write-Verbose "$(Get-Date): "
	}
	
	If($DeliveryGroupsUtilization)
	{
		Write-Verbose "$(Get-Date): `t`t`tCreating Delivery Group Utilization report"
		OutputDeliveryGroupUtilization $Group
	}
}

Function OutputDeliveryGroupDetails 
{
	Param([object] $Group)

	$xDGType = "Delivery Group Type cannot be determined: $($Group.DeliveryType) $($Group.DesktopKind)"
	$xVDAVersion = ""
	$xDeliveryType = ""
	$xColorDepth = ""
	$xShutdownDesktopsAfterUse = "No"
	$xTurnOnAddedMachine = "No"
	$DGIncludedUsers = @()
	$DGExcludedUsers = @()
	$DGScopes = @()
	$DGSFServers = @()
	$xSessionPrelaunch = "Off"
	[int]$xSessionPrelaunchAvgLoad = 0
	[int]$xSessionPrelaunchAnyLoad = 0
	$xSessionLinger = "Off"
	[int]$xSessionLingerAvgLoad = 0
	[int]$xSessionLingerAnyLoad = 0
	$xEndPrelaunchSession = ""
	$xEndLinger = ""
	$PwrMgmt1 = $False
	$PwrMgmt2 = $False
	$PwrMgmt3 = $False
	$xUsersHomeZone = "No"
	
	If($Group.DeliveryType -eq "AppsOnly" -and $Group.DesktopKind -eq "Shared")
	{
		$xDGType = "Random Applications"
	}
	ElseIf($Group.DeliveryType -eq "DesktopsOnly" -and $Group.DesktopKind -eq "Shared")
	{
		$xDGType = "Random Desktops"
	}
	ElseIf($Group.DeliveryType -eq "DesktopsOnly" -and $Group.DesktopKind -eq "Private")
	{
		$xDGType = "Static Desktops"
		$xDeliveryType = "Desktops"
	}
	ElseIf($Group.DeliveryType -eq "DesktopsAndApps" -and $Group.DesktopKind -eq "Shared")
	{
		$xDGType = "Random Desktops and applications"
	}
	
	$NumApps      = (@(Get-BrokerApplication @XDParams2 -DesktopGroupUid $Group.Uid)).Count
	$NumAppGroups = (@(Get-BrokerApplicationGroup @XDParams2 -DesktopGroupUid $Group.Uid)).Count
	$NumDesktops  = (@(Get-BrokerEntitlementPolicyRule @XDParams2 -DesktopGroupUid $Group.Uid)).Count

	If($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -eq 0)
	{
		$xDeliveryType = "Applications"
	}
	ElseIf($NumApps -eq 0 -and $NumAppGroups -eq 0 -and $NumDesktops -gt 0)
	{
		$xDeliveryType = "Desktops"
	}
	ElseIf($NumApps -gt 0 -or $NumAppGroups -gt 0 -and $NumDesktops -gt 0)
	{
		$xDeliveryType = "Applications and Desktops"
	}
	Else
	{
		$xDeliveryType = "Delivery type could not be determined: Apps($NumApps) AppGroups($NumAppGroups) Desktops($NumDesktops)"
	}
	
	If((Get-BrokerServiceAddedCapability @XDParams1) -contains "RestrictToTag")
	{
		$CanRestrictToTags = $True
	}
	Else
	{
		$CanRestrictToTags = $False
	}

	$DesktopSettings = $Null
	If($xDeliveryType -ne "Applications")
	{
		$DesktopSettingsIncludedUsers = @()
		$DesktopSettingsExcludedUsers = @()
		$DesktopSettings = Get-BrokerEntitlementPolicyRule @XDParams2 -DesktopGroupUid $Group.Uid
		
		If($? -and $Null -ne $DesktopSettings)
		{
			If($CanRestrictToTags -eq $True)
			{
				$RestrictedToTag = $DesktopSettings.RestrictToTag
			}
			Else
			{
				$RestrictedToTag = "-"
			}
			
			If($DesktopSettings.IncludedUserFilterEnabled -eq $True)
			{
				ForEach($User in $DesktopSettings.IncludedUsers)
				{
					$DesktopSettingsIncludedUsers += $User.Name
				}
				
				[array]$DesktopSettingsIncludedUsers = $DesktopSettingsIncludedUsers | Sort -unique
			}
			ElseIf($DesktopSettings.IncludedUserFilterEnabled -eq $False)
			{
				$DesktopSettingsIncludedUsers += "Allow everyone with access to this Delivery Group to use a desktop"
			}
			
			If($DesktopSettings.ExcludedUserFilterEnabled -eq $True)
			{
				ForEach($User in $DesktopSettings.ExcludedUsers)
				{
					$DesktopSettingsExcludedUsers += $User.Name
				}

				[array]$DesktopSettingsExcludedUsers = $DesktopSettingsExcludedUsers | Sort -unique
			}
		}
		ElseIf($? -and $Null -eq $DesktopSettings)
		{
			$RestrictedToTag = "-"
		}
		Else
		{
			$DesktopSettings = $Null
			$RestrictedToTag = "-"
		}
	}
	
	Switch ($Group.MinimumFunctionalLevel)
	{
		"L5" 	{$xVDAVersion = "5.6 FP1 (Windows XP and Windows Vista)"; Break}
		"L7"	{$xVDAVersion = "7.0 (or newer)"; Break}
		"L7_6"	{$xVDAVersion = "7.6 (or newer)"; Break}
		"L7_7"	{$xVDAVersion = "7.7 (or newer)"; Break}
		"L7_8"	{$xVDAVersion = "7.8 (or newer)"; Break}
		"L7_9"	{$xVDAVersion = "7.9 (or newer)"; Break}
		Default {"Unable to determine VDA version: $($Group.MinimumFunctionalLevel)"; Break}
	}
	
	If($Group.ColorDepth -eq "FourBit")
	{
		$xColorDepth = "4bit - 16 colors"
	}
	ElseIf($Group.ColorDepth -eq "EightBit")
	{
		$xColorDepth = "8bit - 256 colors"
	}
	ElseIf($Group.ColorDepth -eq "SixteenBit")
	{
		$xColorDepth = "16bit - High color"
	}
	ElseIf($Group.ColorDepth -eq "TwentyFourBit")
	{
		$xColorDepth = "24bit - True color"
	}
	
	If($Group.ShutdownDesktopsAfterUse)
	{
		$xShutdownDesktopsAfterUse = "Yes"
	}
	
	If($Group.TurnOnAddedMachine)
	{
		$xTurnOnAddedMachine = "Yes"
	}

	ForEach($Scope in $Group.Scopes)
	{
		$DGScopes += $Scope
	}
	$DGScopes += "All"
	
	ForEach($Server in $Group.MachineConfigurationNames)
	{
		$SFTmp = Get-BrokerMachineConfiguration -Name $Server
		If($? -and $Null -ne $SFTmp)
		{
			$SFByteArray = $SFTmp.Policy
			$SFServer = Get-SFStoreFrontAddress -ByteArray $SFByteArray 4>$Null
			If($? -and $Null -ne $SFServer)
			{
				$DGSFServers += $SFServer.Url
			}
		}
	}
	
	If($DGSFServers.Count -eq 0)
	{
		$DGSFServers += "-"
	}
	
	$Results = Get-BrokerSessionPreLaunch -DesktopGroupUid $Group.Uid @XDParams1
	If($? -and $Null -ne $Results)
	{
		If($Results.Enabled -and $Results.AssociatedUserFullNames.Count -eq 0)
		{
			$xSessionPrelaunch = "Prelaunch for any user"
		}
		ElseIf($Results.Enabled -and $Results.AssociatedUserFullNames.Count -gt 0)
		{
			$xSessionPrelaunch = "Prelaunch for specific users"
		}
		
		If($Results.MaxAverageLoadThreshold -gt 0)
		{
			$xSessionPrelaunchAvgLoad = ($Results.MaxAverageLoadThreshold/100)
		}
		If($Results.MaxLoadPerMachineThreshold -gt 0)
		{
			$xSessionPrelaunchAnyLoad = ($Results.MaxLoadPerMachineThreshold/100)
		}
		$Mins = $Results.MaxTimeBeforeTerminate.Minutes
		$Hours = $Results.MaxTimeBeforeTerminate.Hours
		$Days = $Results.MaxTimeBeforeTerminate.Days
		If($Mins -gt 0)
		{
			$xEndPrelaunchSession = "$($Mins) Minutes"
		}
		If($Hours -gt 0)
		{
			$xEndPrelaunchSession = "$($Hours) Hours"
		}
		ElseIf($Days -gt 0)
		{
			$xEndPrelaunchSession = "$($Days) Days"
		}
	}
	
	$Results = Get-BrokerSessionLinger -DesktopGroupUid $Group.Uid @XDParams1
	If($? -and $Null -ne $Results)
	{
		$xSessionLinger = "Keep session active"
		If($Results.MaxAverageLoadThreshold -gt 0)
		{
			$xSessionLingerAvgLoad = ($Results.MaxAverageLoadThreshold/100)
		}
		If($Results.MaxLoadPerMachineThreshold -gt 0)
		{
			$xSessionLingerAnyLoad = ($Results.MaxLoadPerMachineThreshold/100)
		}
		$Mins = $Results.MaxTimeBeforeTerminate.Minutes
		$Hours = $Results.MaxTimeBeforeTerminate.Hours
		$Days = $Results.MaxTimeBeforeTerminate.Days
		If($Mins -gt 0)
		{
			$xEndPrelaunchSession = "$($Mins) Minutes"
		}
		If($Hours -gt 0)
		{
			$xEndPrelaunchSession = "$($Hours) Hours"
		}
		ElseIf($Days -gt 0)
		{
			$xEndPrelaunchSession = "$($Days) Days"
		}
	}

	If($Group.ZonePreferences -Contains "UserHomeOnly")
	{
		$xUsersHomeZone = "Yes, if configured"
	}
	
	#get a desktop in an associated delivery group to get the catalog
	$Desktop = Get-BrokerDesktop @XDParams1 -DesktopGroupUid $Group.Uid -Property CatalogName
	
	If($? -and $Null -ne $Desktop)
	{
		$Catalog = Get-BrokerCatalog @XDParams1 -Name $Desktop[0].CatalogName
		
		If($? -and $Null -ne $Catalog)
		{
			If($Catalog.AllocationType -eq "Static" -and $Catalog.PersistUserChanges -eq "Discard" -and $Group.DesktopKind -eq "Private" -and $Group.SessionSupport -eq "SingleSession")
			{
				$PwrMgmt1 = $True
				$PwrMgmt2 = $False
				$PwrMgmt3 = $False
			}
			ElseIf($Catalog.AllocationType -eq "Static" -and $Catalog.PersistUserChanges -ne "Discard" -and $Group.DesktopKind -eq "Private" -and $Group.SessionSupport -eq "SingleSession")
			{
				$PwrMgmt1 = $False
				$PwrMgmt2 = $True
				$PwrMgmt3 = $False
			}
			ElseIf($Catalog.AllocationType -eq "Random" -and $Catalog.PersistUserChanges -eq "Discard" -and $Group.DesktopKind -eq "Shared" -and $Group.SessionSupport -eq "SingleSession")
			{
				$PwrMgmt1 = $False
				$PwrMgmt2 = $False
				$PwrMgmt3 = $True
			}
		}
	}

	If($PwrMgmt2 -or $PwrMgmt3)
	{
		$PwrMgmts = Get-BrokerPowerTimeScheme @XDParams1 -DesktopGroupUid $Group.Uid 
	}
	
	$xOffPeakBufferSizePercent = $Group.OffPeakBufferSizePercent
	$xOffPeakDisconnectTimeout = $Group.OffPeakDisconnectTimeout
	$xOffPeakExtendedDisconnectTimeout = $Group.OffPeakExtendedDisconnectTimeout
	$xOffPeakLogOffTimeout = $Group.OffPeakLogOffTimeout
	$xPeakBufferSizePercent = $Group.PeakBufferSizePercent
	$xPeakDisconnectTimeout = $Group.PeakDisconnectTimeout
	$xPeakExtendedDisconnectTimeout = $Group.PeakExtendedDisconnectTimeout
	$xPeakLogOffTimeout = $Group.PeakLogOffTimeout

	$xOffPeakDisconnectAction = ""
	$xOffPeakExtendedDisconnectAction = ""
	$xOffPeakLogOffAction = ""
	$xPeakDisconnectAction = ""
	$xPeakExtendedDisconnectAction = ""
	$xPeakLogOffAction = ""

	Switch ($Group.OffPeakDisconnectAction)
	{
		"Nothing"	{ $xOffPeakDisconnectAction = "No action"; Break}
		"Suspend"	{ $xOffPeakDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xOffPeakDisconnectAction = "Shut down"; Break}
		Default		{ $xOffPeakDisconnectAction = "Unable to determine the OffPeakDisconnectAction action: $($Group.OffPeakDisconnectAction)"; Break}
	}
	
	Switch ($Group.OffPeakExtendedDisconnectAction)
	{
		"Nothing"	{ $xOffPeakExtendedDisconnectAction = "No action"; Break}
		"Suspend"	{ $xOffPeakExtendedDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xOffPeakExtendedDisconnectAction = "Shut down"; Break}
		Default		{ $xOffPeakExtendedDisconnectAction = "Unable to determine the OffPeakExtendedDisconnectAction action: $($Group.OffPeakExtendedDisconnectAction)"; Break}
	}
	
	Switch ($Group.OffPeakLogOffAction)
	{
		"Nothing"	{ $xOffPeakLogOffAction = "No action"; Break}
		"Suspend"	{ $xOffPeakLogOffAction = "Suspend"; Break}
		"Shutdown"	{ $xOffPeakLogOffAction = "Shut down"; Break}
		Default		{ $xOffPeakLogOffAction = "Unable to determine $xOffPeakLogOffAction action: $($Group.OffPeakLogOffAction)"; Break}
	}
	
	Switch ($Group.PeakDisconnectAction)
	{
		"Nothing"	{ $xPeakDisconnectAction = "No action"; Break}
		"Suspend"	{ $xPeakDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xPeakDisconnectAction = "Shut down"; Break}
		Default		{ $xPeakDisconnectAction = "Unable to determine $xPeakDisconnectAction action: $($Group.PeakDisconnectAction)"; Break}
	}
	
	Switch ($Group.PeakExtendedDisconnectAction)
	{
		"Nothing"	{ $xPeakExtendedDisconnectAction = "No action"; Break}
		"Suspend"	{ $xPeakExtendedDisconnectAction = "Suspend"; Break}
		"Shutdown"	{ $xPeakExtendedDisconnectAction = "Shut down"; Break}
		Default		{ $xPeakExtendedDisconnectAction = "Unable to determine $xPeakExtendedDisconnectAction action: $($Group.PeakExtendedDisconnectAction)"; Break}
	}
	
	Switch ($Group.PeakLogOffAction)
	{
		"Nothing"	{ $xPeakLogOffAction = "No action"; Break}
		"Suspend"	{ $xPeakLogOffAction = "Suspend"; Break}
		"Shutdown"	{ $xPeakLogOffAction = "Shut down"; Break}
		Default		{ $xPeakLogOffAction = "Unable to determine $xPeakLogOffAction action: $($Group.PeakLogOffAction)"; Break}
	}

	$xEnabled = "Disabled"
	If($Group.Enabled)
	{
		$xEnabled = "Enabled"
	}

	$xSecureICA = "Disabled"
	If($Group.SecureICARequired)
	{
		$xSecureICA = "Enabled"
	}
	
	#added 17-jun-2015
	$xAutoPowerOnForAssigned = "Disabled"
	$xAutoPowerOnForAssignedDuringPeak = "Disabled"
	
	If($Group.AutomaticPowerOnForAssigned)
	{
		$xAutoPowerOnForAssigned = "Enabled"
	}
	If($Group.AutomaticPowerOnForAssignedDuringPeak)
	{
		$xAutoPowerOnForAssignedDuringPeak = "Enabled"
	}

	$SFAnonymousUsers = $False
	$Results = Get-BrokerAccessPolicyRule -DesktopGroupUid $Group.Uid @XDParams1
	
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			
			If($Result.AllowedUsers -eq "Any" -or $Result.AllowedUsers -eq "FilteredOrAnonymous" -or $Result.AllowedUsers -eq "AnonymousOnly")
			{
				$SFAnonymousUsers = $True
			}
			
			If($Result.IncludedUserFilterEnabled -and $Result.AllowedUsers -eq "Filtered")
			{
				ForEach($User in $Result.IncludedUsers)
				{
					$DGIncludedUsers += $User.Name
				}
			}
			ElseIf($Result.IncludedUserFilterEnabled -and ($Result.AllowedUsers -eq "AnyAuthenticated" -or $Result.AllowedUsers -eq "Any"))
			{
				$DGIncludedUsers += "Allow any authenticated users to use this Delivery Group"
			}
			
			If($Result.ExcludedUserFilterEnabled)
			{
				ForEach($User in $Result.ExcludedUsers)
				{
					$DGExcludedUsers += $User.Name
				}
			}
			
			If($Result.Name -like '*_AG')
			{
				If($Result.AllowedConnections -eq "ViaAG" -and $Result.IncludedSmartAccessFilterEnabled -eq $False -and $Result.Enabled -eq $True)
				{
					$xAllConnections = "Enabled"
					$xNSConnection = "Disabled"
					$xAGFilters = @()
					If($HTML)
					{
						$xAGFilters += "N/A"
					}
					Else
					{
						$xAGFilters += "<N/A>"
					}
				}
				ElseIf($Result.AllowedConnections -eq "ViaAG" -and $Result.IncludedSmartAccessFilterEnabled -eq $True -and $Result.Enabled -eq $True)
				{
					$xAllConnections = "Enabled"
					$xNSConnection = "Enabled"
					$xAGFilters = @()
					ForEach($AccessCondition in $Result.IncludedSmartAccessTags)
					{
						$xAGFilters += $AccessCondition
					}
					If($xAGFilters.Count -eq 0)
					{
						If($HTML)
						{
							$xAGFilters += "none"
						}
						Else
						{
							$xAGFilters += "<none>"
						}
					}
				}
				ElseIf($Result.AllowedConnections -eq "ViaAG" -and $Result.IncludedSmartAccessFilterEnabled -eq $False -and $Result.Enabled -eq $False)
				{
					$xAllConnections = "Disabled"
					$xNSConnection = "Disabled"
					$xAGFilters = @()
					If($HTML)
					{
						$xAGFilters += "N/A"
					}
					Else
					{
						$xAGFilters += "<N/A>"
					}
				}
			}
		}
		
		[array]$DGIncludedUsers = $DGIncludedUsers | Sort -unique
		[array]$DGExcludedUsers = $DGExcludedUsers | Sort -unique
	}
	
	#desktops per user for singlesession OS
	If($Group.SessionSupport -eq "SingleSession")
	{
		If($xDGType -eq "Static Desktops")
		{
			#static desktops have a maxdesktops count stored as a property
			$xMaxDesktops = 0
			$MaxDesktops = Get-BrokerAssignmentPolicyRule @XDParams1 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $MaxDesktops)
			{
				$xMaxDesktops = $MaxDesktops.MaxDesktops
			}
		}
		ElseIf($xDGType -like "*Random*")
		{
			#random desktops are a count of the number of entitlement policy rules
			$xMaxDesktops = 0
			$MaxDesktops = Get-BrokerEntitlementPolicyRule @XDParams1 -DesktopGroupUid $Group.Uid
			
			If($? -and $Null -ne $MaxDesktops)
			{
				$xMaxDesktops = $MaxDesktops.Count
			}
		}
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Details: " $Group.Name
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Description"; Value = $Group.Description; }
		If(![String]::IsNullOrEmpty($Group.PublishedName))
		{
			$ScriptInformation += @{Data = "Display Name"; Value = $Group.PublishedName; }
		}
		$ScriptInformation += @{Data = "Type"; Value = $xDGType; }
		$ScriptInformation += @{Data = "Set to VDA version"; Value = $xVDAVersion; }
		If($Group.SessionSupport -eq "SingleSession" -and ($xDGType -eq "Static Desktops" -or $xDGType -like "*Random*"))
		{
			$ScriptInformation += @{Data = "Desktops per user"; Value = $xMaxDesktops; }
		}
		$ScriptInformation += @{Data = "Time zone"; Value = $Group.TimeZone; }
		$ScriptInformation += @{Data = "Enable Delivery Group"; Value = $xEnabled; }
		$ScriptInformation += @{Data = "Enable Secure ICA"; Value = $xSecureICA; }
		$ScriptInformation += @{Data = "Color Depth"; Value = $xColorDepth; }
		$ScriptInformation += @{Data = "Shutdown Desktops After Use"; Value = $xShutdownDesktopsAfterUse; }
		$ScriptInformation += @{Data = "Turn On Added Machine"; Value = $xTurnOnAddedMachine; }
		$ScriptInformation += @{Data = "Included Users"; Value = $DGIncludedUsers[0]; }
		$cnt = -1
		ForEach($tmp in $DGIncludedUsers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($DGExcludedUsers.Count -gt 0)
		{
			$ScriptInformation += @{Data = "Excluded Users"; Value = $DGExcludedUsers[0]; }
			$cnt = -1
			ForEach($tmp in $DGExcludedUsers)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
		}

		If($Group.SessionSupport -eq "MultiSession")
		{
			$ScriptInformation += @{Data = 'Give access to unathenticated (anonymous) users'; Value = $SFAnonymousUsers; }
		}

		If($xDeliveryType -ne "Applications" -and $Null -ne $DesktopSettings)
		{
			$ScriptInformation += @{Data = "Desktop Entitlement"; Value = ""; }
			$ScriptInformation += @{Data = "     Display name"; Value = $DesktopSettings.PublishedName; }
			$ScriptInformation += @{Data = "     Description"; Value = $DesktopSettings.Description; }
			If($CanRestrictToTags)
			{
				$ScriptInformation += @{Data = "     Restrict launches to machines with tag"; Value = $RestrictedToTag; }
			}
			$ScriptInformation += @{Data = "     Included Users"; Value = $DesktopSettingsIncludedUsers[0]; }
			$cnt = -1
			ForEach($tmp in $DesktopSettingsIncludedUsers)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{Data = ""; Value = $tmp; }
				}
			}
			
			If($DesktopSettingsExcludedUsers.Count -gt 0)
			{
				$ScriptInformation += @{Data = "     Excluded Users"; Value = $DesktopSettingsExcludedUsers[0]; }
				$cnt = -1
				ForEach($tmp in $DGExcludedUsers)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
			}
			$ScriptInformation += @{Data = "     Enable desktop"; Value = $DesktopSettings.Enabled; }
		}
		
		$ScriptInformation += @{Data = "Scopes"; Value = $DGScopes[0]; }
		$cnt = -1
		ForEach($tmp in $DGScopes)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		$ScriptInformation += @{Data = "StoreFronts"; Value = $DGSFServers[0]; }
		$cnt = -1
		ForEach($tmp in $DGSFServers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession" -and $xDeliveryType -like '*App*')
		{
			$ScriptInformation += @{Data = "Session prelaunch"; Value = $xSessionPrelaunch; }
			If($xSessionPrelaunch -ne "Off")
			{
				$ScriptInformation += @{Data = "Prelaunched session will end in"; Value = $xEndPrelaunchSession; }
				
				If($xSessionPrelaunchAvgLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When avg load on all machines exceeds (%)"; Value = $xSessionPrelaunchAvgLoad; }
				}
				If($xSessionPrelaunchAnyLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When load on any machines exceeds (%)"; Value = $xSessionPrelaunchAnyLoad; }
				}
			}
			$ScriptInformation += @{Data = "Session lingering"; Value = $xSessionLinger; }
			If($xSessionLinger -ne "Off")
			{
				$ScriptInformation += @{Data = "Keep sessions active until after"; Value = $xEndPrelaunchSession; }
				
				If($xSessionLingerAvgLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When avg load on all machines exceeds (%)"; Value = $xSessionPrelaunchAvgLoad; }
				}
				If($xSessionLingerAnyLoad -gt 0)
				{
					$ScriptInformation += @{Data = "When load on any machines exceeds (%)"; Value = $xSessionPrelaunchAnyLoad; }
				}
			}
		}

		$ScriptInformation += @{Data = "Launch in user's home zone"; Value = $xUsersHomeZOne; }
		
		If($Group.SessionSupport -eq "MultiSession")
		{
			If((Get-BrokerServiceAddedCapability @XDParams1) -contains "MultipleRebootSchedulesPerGroup")
			{
				$ChkTags = $True
				$RestartSchedule = Get-BrokerRebootScheduleV2 @XDParams1 -DesktopGroupUid $Group.Uid
			}
			Else
			{
				$ChkTags = $False
				$RestartSchedule = Get-BrokerRebootSchedule @XDParams1 -DesktopGroupUid $Group.Uid
			}
			
			If($? -and $Null -ne $RestartSchedule)
			{
				$ScriptInformation += @{Data = "Restart machines automatically"; Value = "Yes"; }

				If($ChkTags -eq $True)
				{
					$ScriptInformation += @{Data = "Restrict to tag"; Value = $RestartSchedule.RestrictToTag; }
				}
				
				$tmp = ""
				If($RestartSchedule.Frequency -eq "Daily")
				{
					$tmp = "Daily"
				}
				ElseIf($RestartSchedule.Frequency -eq "Weekly")
				{
					$tmp = "Every $($RestartSchedule.Day)"
				}
				
				$ScriptInformation += @{Data = "Restart frequency"; Value = $tmp; }
				$ScriptInformation += @{Data = "Begin restart at"; Value = "$($RestartSchedule.StartTime.Hours.ToString("00")):$($RestartSchedule.StartTime.Minutes.ToString("00"))"; }
				
				$xTime = 0
				$tmp = ""
				If($RestartSchedule.RebootDuration -eq 0)
				{
					$tmp = "Restart all machines at once"
				}
				ElseIf($RestartSchedule.RebootDuration -eq 30)
				{
					$tmp = "30 minutes"
				}
				Else
				{
					$xTime = $RestartSchedule.RebootDuration / 60
					$tmp = "$($xTime) hours"
				}
				$ScriptInformation += @{Data = "Restart duration"; Value = $tmp; }
				$xTime = $Null
				$tmp = $Null
				
				$tmp = ""
				If($RestartSchedule.WarningDuration -eq 0)
				{
					$tmp = "Do not send a notification"
					$ScriptInformation += @{Data = "Send notification to users"; Value = $tmp; }
				}
				Else
				{
					$tmp = "$($RestartSchedule.WarningDuration) minutes before user is logged off"
					$ScriptInformation += @{Data = "Send notification to users"; Value = $tmp; }
					$ScriptInformation += @{Data = "Notification message"; Value = $RestartSchedule.WarningMessage; }
				}
				
			}
			Else
			{
				$ScriptInformation += @{Data = "Restart machines automatically"; Value = "No"; }
			}
		}

		If($PwrMgmt1)
		{
			$ScriptInformation += @{Data = "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins"; Value = $xPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins"; Value = $xPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins"; Value = $xOffPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins"; Value = $xOffPeakExtendedDisconnectAction; }
		}
		If($PwrMgmt2)
		{
			$ScriptInformation += @{Data = "Weekday Peak hours"; Value = ""; }
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = ""; Value = "<none>"; }
			}

			$ScriptInformation += @{Data = "Weekend Peak hours"; Value = ""; }
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = ""; Value = "<none>"; }
			}

			$ScriptInformation += @{Data = "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins"; Value = $xPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins"; Value = $xPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During peak hours, when logged off $($Group.PeakLogOffTimeout) mins"; Value = $xPeakLogOffAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins"; Value = $xOffPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins"; Value = $xOffPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when logged off $($Group.OffPeakLogOffTimeout) mins"; Value = $xOffPeakLogOffAction; }
		}
		If($PwrMgmt3)
		{
			$ScriptInformation += @{Data = "Weekday number machines powered on, and when"; Value = ""; }
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
							$ScriptInformation += @{Data = ""; Value = "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"; }
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = ""; Value = "<none>"; }
			}

			$ScriptInformation += @{Data = "Weekend number machines powered on, and when"; Value = ""; }
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
							$ScriptInformation += @{Data = ""; Value = "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"; }
						}
					}
				}
			}
			
			If($val -eq 0)
			{
				$ScriptInformation += @{Data = ""; Value = "<none>"; }
			}
			
			$ScriptInformation += @{Data = "Weekday Peak hours"; Value = ""; }
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = ""; Value = "<none>"; }
			}

			$ScriptInformation += @{Data = "Weekend Peak hours"; Value = ""; }
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$ScriptInformation += @{Data = ""; Value = "$($i.ToString("00")):00"; }
						}
					}
				}
			}

			If($val -eq 0)
			{
				$ScriptInformation += @{Data = ""; Value = "<none>"; }
			}

			$ScriptInformation += @{Data = "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins"; Value = $xPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins"; Value = $xPeakExtendedDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins"; Value = $xOffPeakDisconnectAction; }
			$ScriptInformation += @{Data = "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins"; Value = $xOffPeakExtendedDisconnectAction; }
		}

		$ScriptInformation += @{Data = "Automatic power on for assigned"; Value = $xAutoPowerOnForAssigned; }
		$ScriptInformation += @{Data = "Automatic power on for assigned during peak"; Value = $xAutoPowerOnForAssignedDuringPeak; }
		
		$ScriptInformation += @{Data = "All connections not through NetScaler Gateway"; Value = $xAllConnections; }
		$ScriptInformation += @{Data = "Connections through NetScaler Gateway"; Value = $xNSConnection; }
		$ScriptInformation += @{Data = "Connections meeting any of the following filters"; Value = $xAGFilters[0]; }
		$cnt = -1
		ForEach($tmp in $xAGFilters)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Details: " $Group.Name
		Line 1 "Description`t`t`t`t`t: " $Group.Description
		If(![String]::IsNullOrEmpty($Group.PublishedName))
		{
			Line 1 "Display Name`t`t`t`t`t: " $Group.PublishedName
		}
		Line 1 "Type`t`t`t`t`t`t: " $xDGType
		Line 1 "Set to VDA version`t`t`t`t: " $xVDAVersion
		If($Group.SessionSupport -eq "SingleSession" -and ($xDGType -eq "Static Desktops" -or $xDGType -like "*Random*"))
		{
			Line 1 "Desktops per user`t`t`t`t: " $xMaxDesktops
		}
		Line 1 "Time zone`t`t`t`t`t: " $Group.TimeZone
		Line 1 "Enable Delivery Group`t`t`t`t: " $xEnabled
		Line 1 "Enable Secure ICA`t`t`t`t: " $xSecureICA
		Line 1 "Color Depth`t`t`t`t`t: " $xColorDepth
		Line 1 "Shutdown Desktops After Use`t`t`t: " $xShutdownDesktopsAfterUse
		Line 1 "Turn On Added Machine`t`t`t`t: " $xTurnOnAddedMachine
		Line 1 "Included Users`t`t`t`t`t: " $DGIncludedUsers[0]
		$cnt = -1
		ForEach($tmp in $DGIncludedUsers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 7 "  " $tmp
			}
		}
		
		If($DGExcludedUsers.Count -gt 0)
		{
			Line 1 "Excluded Users`t`t`t`t`t: " $DGExcludedUsers[0]
			$cnt = -1
			ForEach($tmp in $DGExcludedUsers)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 7 "  " $tmp
				}
			}
		}

		If($Group.SessionSupport -eq "MultiSession")
		{
			Line 1 "Give access to unathenticated (anonymous) users`t: " $SFAnonymousUsers
		}

		If($xDeliveryType -ne "Applications" -and $Null -ne $DesktopSettings)
		{
			Line 1 "Desktop Entitlement" ""
			Line 2 "Display name`t`t`t`t: " $DesktopSettings.PublishedName
			Line 2 "Description`t`t`t`t: " $DesktopSettings.Description
			If($CanRestrictToTags)
			{
				Line 2 "Restrict launches to machines with tag`t: " $RestrictedToTag
			}
			Line 2 "Included Users`t`t`t`t: " $DesktopSettingsIncludedUsers[0]
			$cnt = -1
			ForEach($tmp in $DesktopSettingsIncludedUsers)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 7 "" $tmp
				}
			}
			
			If($DesktopSettingsExcludedUsers.Count -gt 0)
			{
				Line 2 "Excluded Users`t`t`t`t: " $DesktopSettingsExcludedUsers[0]
				$cnt = -1
				ForEach($tmp in $DGExcludedUsers)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 7 "" $tmp
					}
				}
			}
			Line 2 "Enable desktop`t`t`t`t: " $DesktopSettings.Enabled
		}
		
		Line 1 "Scopes`t`t`t`t`t`t: " $DGScopes[0]
		$cnt = -1
		ForEach($tmp in $DGScopes)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 7 "  " $tmp
			}
		}
		
		Line 1 "StoreFronts`t`t`t`t`t: " $DGSFServers[0]
		$cnt = -1
		ForEach($tmp in $DGSFServers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 7 "  " $tmp
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession" -and $xDeliveryType -like '*App*')
		{
			Line 1 "Session prelaunch`t`t`t`t: " $xSessionPrelaunch
			If($xSessionPrelaunch -ne "Off")
			{
				Line 1 "Prelaunched session will end in`t`t`t: " $xEndPrelaunchSession
				
				If($xSessionPrelaunchAvgLoad -gt 0)
				{
					Line 1 "When avg load on all machines exceeds (%)`t: " $xSessionPrelaunchAvgLoad
				}
				If($xSessionPrelaunchAnyLoad -gt 0)
				{
					Line 1 "When load on any machines exceeds (%)`t`t: " $xSessionPrelaunchAnyLoad
				}
			}
			Line 1 "Session lingering`t`t`t`t: " $xSessionLinger
			If($xSessionLinger -ne "Off")
			{
				Line 1 "Keep sessions active until after`t`t: " $xEndPrelaunchSession
				
				If($xSessionLingerAvgLoad -gt 0)
				{
					Line 1 "When avg load on all machines exceeds (%)`t: " $xSessionPrelaunchAvgLoad
				}
				If($xSessionLingerAnyLoad -gt 0)
				{
					Line 1 "When load on any machines exceeds (%)`t`t: " $xSessionPrelaunchAnyLoad
				}
			}
		}

		Line 1 "Launch in user's home zone`t`t`t: " $xUsersHomeZone
		
		If($Group.SessionSupport -eq "MultiSession")
		{
			If((Get-BrokerServiceAddedCapability @XDParams1) -contains "MultipleRebootSchedulesPerGroup")
			{
				$ChkTags = $True
				$RestartSchedule = Get-BrokerRebootScheduleV2 @XDParams1 -DesktopGroupUid $Group.Uid
			}
			Else
			{
				$ChkTags = $False
				$RestartSchedule = Get-BrokerRebootSchedule @XDParams1 -DesktopGroupUid $Group.Uid
			}
			
			If($? -and $Null -ne $RestartSchedule)
			{
				Line 1 "Restart machines automatically`t`t`t: " "Yes"

				If($ChkTags -eq $True)
				{
					Line 1 "Restrict to tag`t`t`t`t`t: " $RestartSchedule.RestrictToTag
				}
				
				$tmp = ""
				If($RestartSchedule.Frequency -eq "Daily")
				{
					$tmp = "Daily"
				}
				ElseIf($RestartSchedule.Frequency -eq "Weekly")
				{
					$tmp = "Every $($RestartSchedule.Day)"
				}
				
				Line 1 "Restart frequency`t`t`t`t: " $tmp
				Line 1 "Begin restart at`t`t`t`t: " "$($RestartSchedule.StartTime.Hours.ToString("00")):$($RestartSchedule.StartTime.Minutes.ToString("00"))"
				
				$xTime = 0
				$tmp = ""
				If($RestartSchedule.RebootDuration -eq 0)
				{
					$tmp = "Restart all machines at once"
				}
				ElseIf($RestartSchedule.RebootDuration -eq 30)
				{
					$tmp = "30 minutes"
				}
				Else
				{
					$xTime = $RestartSchedule.RebootDuration / 60
					$tmp = "$($xTime) hours"
				}
				Line 1 "Restart duration`t`t`t`t: " $tmp
				$xTime = $Null
				$tmp = $Null
				
				$tmp = ""
				If($RestartSchedule.WarningDuration -eq 0)
				{
					$tmp = "Do not send a notification"
					Line 1 "Send notification to users`t`t`t: " $tmp
				}
				Else
				{
					$tmp = "$($RestartSchedule.WarningDuration) minutes before user is logged off"
					Line 1 "Send notification to users`t`t`t: " $tmp
					Line 1 "Notification message`t`t`t`t: " $RestartSchedule.WarningMessage
				}
				
			}
			Else
			{
				Line 1 "Restart machines automatically`t`t`t: " "No"
			}
		}
		
		If($PwrMgmt1)
		{
			Line 1 "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins`t: " $xPeakDisconnectAction
			Line 1 "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins`t: " $xPeakExtendedDisconnectAction
			Line 1 "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins: " $xOffPeakDisconnectAction
			Line 1 "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins: " $xOffPeakExtendedDisconnectAction
		}
		If($PwrMgmt2)
		{
			Line 1 "Weekday Peak hours`t`t`t`t:" ""
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							Line 7 "  " "$($i.ToString("00")):00"
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 7 "  "  "<none>"
			}

			Line 1 "Weekend Peak hours`t`t`t`t: " ""
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							Line 7 "  " "$($i.ToString("00")):00"
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 7 "  "  "<none>"
			}

			Line 1 "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins`t: " $xPeakDisconnectAction
			Line 1 "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins`t: " $xPeakExtendedDisconnectAction
			Line 1 "During peak hours, when logged off $($Group.PeakLogOffTimeout) mins`t: " $xPeakLogOffAction
			Line 1 "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins`t: " $xOffPeakDisconnectAction
			Line 1 "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins`t: " $xOffPeakExtendedDisconnectAction
			Line 1 "During off-peak hours, when logged off $($Group.OffPeakLogOffTimeout) mins`t: " $xOffPeakLogOffAction
		}
		If($PwrMgmt3)
		{
			Line 1 "Weekday number machines powered on, and when`t: " ""
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
							Line 7 "  " "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 7 "  "  "<none>"
			}

			Line 1 "Weekend number machines powered on, and when`t: " ""
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
							Line 7 "  " "$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00"
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 7 "  "  "<none>"
			}

			Line 1 "Weekday Peak hours: " ""
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							Line 7 "  " "$($i.ToString("00")):00"
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 7 "  "  "<none>"
			}

			Line 1 "Weekend Peak hours: " ""
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							Line 7 "  " "$($i.ToString("00")):00"
						}
					}
				}
			}

			If($val -eq 0)
			{
				Line 7 "  "  "<none>"
			}

			Line 1 "During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins`t: " $xPeakDisconnectAction
			Line 1 "During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins`t: " $xPeakExtendedDisconnectAction
			Line 1 "During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins: " $xOffPeakDisconnectAction
			Line 1 "During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins: " $xOffPeakExtendedDisconnectAction
		}

		Line 1 "Automatic power on for assigned`t`t`t: " $xAutoPowerOnForAssigned
		Line 1 "Automatic power on for assigned during peak`t: " $xAutoPowerOnForAssignedDuringPeak
		
		Line 1 "All connections not through NetScaler Gateway`t: " $xAllConnections
		Line 1 "Connections through NetScaler Gateway`t`t: " $xNSConnection
		Line 1 "Connections meeting any of the following filters: " $xAGFilters[0]
		$cnt = -1
		ForEach($tmp in $xAGFilters)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 7 "  " $tmp
			}
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Details: " $Group.Name
		$rowdata = @()
		$columnHeaders = @("Description",($htmlsilver -bor $htmlbold),$Group.Description,$htmlwhite)
		If(![String]::IsNullOrEmpty($Group.PublishedName))
		{
			$rowdata += @(,('Display Name',($htmlsilver -bor $htmlbold),$Group.PublishedName,$htmlwhite))
		}
		$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$xDGType,$htmlwhite))
		$rowdata += @(,('Set to VDA version',($htmlsilver -bor $htmlbold),$xVDAVersion,$htmlwhite))
		If($Group.SessionSupport -eq "SingleSession" -and ($xDGType -eq "Static Desktops" -or $xDGType -like "*Random*"))
		{
			$rowdata += @(,('Desktops per user',($htmlsilver -bor $htmlbold),$xMaxDesktops,$htmlwhite))
		}
		$rowdata += @(,('Time zone',($htmlsilver -bor $htmlbold),$Group.TimeZone,$htmlwhite))
		$rowdata += @(,('Enable Delivery Group',($htmlsilver -bor $htmlbold),$xEnabled,$htmlwhite))
		$rowdata += @(,('Enable Secure ICA',($htmlsilver -bor $htmlbold),$xSecureICA,$htmlwhite))
		$rowdata += @(,('Color Depth',($htmlsilver -bor $htmlbold),$xColorDepth,$htmlwhite))
		$rowdata += @(,("Shutdown Desktops After Use",($htmlsilver -bor $htmlbold),$xShutdownDesktopsAfterUse,$htmlwhite))
		$rowdata += @(,("Turn On Added Machine",($htmlsilver -bor $htmlbold),$xTurnOnAddedMachine,$htmlwhite))
		$rowdata += @(,('Included Users',($htmlsilver -bor $htmlbold),$DGIncludedUsers[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $DGIncludedUsers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		
		If($DGExcludedUsers.Count -gt 0)
		{
			$rowdata += @(,('Excluded Users',($htmlsilver -bor $htmlbold), $DGExcludedUsers[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $DGExcludedUsers)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}

		If($Group.SessionSupport -eq "MultiSession")
		{
			$rowdata += @(,('Give access to unathenticated (anonymous) users',($htmlsilver -bor $htmlbold),$SFAnonymousUsers,$htmlwhite))
		}

		If($xDeliveryType -ne "Applications" -and $Null -ne $DesktopSettings)
		{
			$rowdata += @(,("Desktop Entitlement",($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$rowdata += @(,("     Display name",($htmlsilver -bor $htmlbold),$DesktopSettings.PublishedName,$htmlwhite))
			$rowdata += @(,("     Description",($htmlsilver -bor $htmlbold),$DesktopSettings.Description,$htmlwhite))
			If($CanRestrictToTags)
			{
				$rowdata += @(,("     Restrict launches to machines with tag",($htmlsilver -bor $htmlbold),$RestrictedToTag,$htmlwhite))
			}
			$rowdata += @(,("     Included Users",($htmlsilver -bor $htmlbold),$DesktopSettingsIncludedUsers[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $DesktopSettingsIncludedUsers)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,("",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
			
			If($DesktopSettingsExcludedUsers.Count -gt 0)
			{
				$rowdata += @(,("     Excluded Users",($htmlsilver -bor $htmlbold),$DesktopSettingsExcludedUsers[0],$htmlwhite))
				$cnt = -1
				ForEach($tmp in $DGExcludedUsers)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,("",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}
				}
			}
			$rowdata += @(,("     Enable desktop",($htmlsilver -bor $htmlbold),$DesktopSettings.Enabled,$htmlwhite))
		}
		
		$rowdata += @(,('Scopes',($htmlsilver -bor $htmlbold),$DGScopes[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $DGScopes)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		
		$rowdata += @(,('StoreFronts',($htmlsilver -bor $htmlbold),$DGSFServers[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $DGSFServers)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		
		If($Group.SessionSupport -eq "MultiSession" -and $xDeliveryType -like '*App*')
		{
			$rowdata += @(,('Session prelaunch',($htmlsilver -bor $htmlbold),$xSessionPrelaunch,$htmlwhite))
			If($xSessionPrelaunch -ne "Off")
			{
				$rowdata += @(,('Prelaunched session will end in',($htmlsilver -bor $htmlbold),$xEndPrelaunchSession,$htmlwhite))
				
				If($xSessionPrelaunchAvgLoad -gt 0)
				{
					$rowdata += @(,('When avg load on all machines exceeds (%)',($htmlsilver -bor $htmlbold),$xSessionPrelaunchAvgLoad,$htmlwhite))
				}
				If($xSessionPrelaunchAnyLoad -gt 0)
				{
					$rowdata += @(,('When load on any machines exceeds (%)',($htmlsilver -bor $htmlbold),$xSessionPrelaunchAnyLoad,$htmlwhite))
				}
			}
			$rowdata += @(,('Session lingering',($htmlsilver -bor $htmlbold),$xSessionLinger,$htmlwhite))
			If($xSessionLinger -ne "Off")
			{
				$rowdata += @(,('Keep sessions active until after',($htmlsilver -bor $htmlbold),$xEndPrelaunchSession,$htmlwhite))
				
				If($xSessionLingerAvgLoad -gt 0)
				{
					$rowdata += @(,('When avg load on all machines exceeds (%)',($htmlsilver -bor $htmlbold),$xSessionPrelaunchAvgLoad,$htmlwhite))
				}
				If($xSessionLingerAnyLoad -gt 0)
				{
					$rowdata += @(,('When load on any machines exceeds (%)',($htmlsilver -bor $htmlbold),$xSessionPrelaunchAnyLoad,$htmlwhite))
				}
			}
		}
		
		$rowdata += @(,("Launch in user's home zone",($htmlsilver -bor $htmlbold),$xUsersHomeZone,$htmlwhite)) 
			
		If($Group.SessionSupport -eq "MultiSession")
		{
			If((Get-BrokerServiceAddedCapability @XDParams1) -contains "MultipleRebootSchedulesPerGroup")
			{
				$ChkTags = $True
				$RestartSchedule = Get-BrokerRebootScheduleV2 @XDParams1 -DesktopGroupUid $Group.Uid
			}
			Else
			{
				$ChkTags = $False
				$RestartSchedule = Get-BrokerRebootSchedule @XDParams1 -DesktopGroupUid $Group.Uid
			}
			
			If($? -and $Null -ne $RestartSchedule)
			{
				$rowdata += @(,('Restart machines automatically',($htmlsilver -bor $htmlbold),"Yes",$htmlwhite))

				If($ChkTags -eq $True)
				{
					$rowdata += @(,('Restrict to tag',($htmlsilver -bor $htmlbold),$RestartSchedule.RestrictToTag,$htmlwhite))
				}
				
				$tmp = ""
				If($RestartSchedule.Frequency -eq "Daily")
				{
					$tmp = "Daily"
				}
				ElseIf($RestartSchedule.Frequency -eq "Weekly")
				{
					$tmp = "Every $($RestartSchedule.Day)"
				}
				
				$rowdata += @(,('Restart frequency',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,('Begin restart at',($htmlsilver -bor $htmlbold),"$($RestartSchedule.StartTime.Hours.ToString("00")):$($RestartSchedule.StartTime.Minutes.ToString("00"))",$htmlwhite))
				
				$xTime = 0
				$tmp = ""
				If($RestartSchedule.RebootDuration -eq 0)
				{
					$tmp = "Restart all machines at once"
				}
				ElseIf($RestartSchedule.RebootDuration -eq 30)
				{
					$tmp = "30 minutes"
				}
				Else
				{
					$xTime = $RestartSchedule.RebootDuration / 60
					$tmp = "$($xTime) hours"
				}
				$rowdata += @(,('Restart duration',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$xTime = $Null
				$tmp = $Null
				
				$tmp = ""
				If($RestartSchedule.WarningDuration -eq 0)
				{
					$tmp = "Do not send a notification"
					$rowdata += @(,('Send notification to users',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				Else
				{
					$tmp = "$($RestartSchedule.WarningDuration) minutes before user is logged off"
					$rowdata += @(,('Send notification to users',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$rowdata += @(,('Notification message',($htmlsilver -bor $htmlbold),$RestartSchedule.WarningMessage,$htmlwhite))
				}
				
			}
			Else
			{
				$rowdata += @(,('Restart machines automatically',($htmlsilver -bor $htmlbold),"No",$htmlwhite))
			}
		}
		
		If($PwrMgmt1)
		{
			$rowdata += @(,("During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakExtendedDisconnectAction,$htmlwhite))
		}
		If($PwrMgmt2)
		{
			$rowdata += @(,('Weekday Peak hours',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($i.ToString("00")):00",$htmlwhite))
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),"none",$htmlwhite))
			}

			$rowdata += @(,('Weekend Peak hours',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($i.ToString("00")):00",$htmlwhite))
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),"none",$htmlwhite))
			}

			$rowdata += @(,("During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak hours, when logged off $($Group.PeakLogOffTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakLogOffAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when logged off $($Group.OffPeakLogOffTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakLogOffAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when logged off $($Group.OffPeakExtendedLogOffTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakExtendedLogOffAction,$htmlwhite))
		}
		If($PwrMgmt3)
		{
			$rowdata += @(,('Weekday number machines powered on, and when',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00",$htmlwhite))
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),"none",$htmlwhite))
			}

			$rowdata += @(,('Weekend number machines powered on, and when',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PoolSize[$i] -gt 0)
						{
							$val++
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($PwrMgmt.PoolSize[$i].ToString("####0")) - $($i.ToString("00")):00",$htmlwhite))
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),"none",$htmlwhite))
			}

			$rowdata += @(,('Weekday Peak hours',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekdays")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($i.ToString("00")):00",$htmlwhite))
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),"none",$htmlwhite))
			}

			$rowdata += @(,('Weekend Peak hours',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$val = 0
			ForEach($PwrMgmt in $PwrMgmts)
			{
				If($PwrMgmt.DaysOfWeek -eq "Weekend")
				{
					For($i=0;$i -le 24;$i++)
					{
						If($PwrMgmt.PeakHours[$i])
						{
							$val++
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($i.ToString("00")):00",$htmlwhite))
						}
					}
				}
			}

			If($val -eq 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
			}

			$rowdata += @(,("During peak hours, when disconnected $($Group.PeakDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During peak extended hours, when disconnected $($Group.PeakExtendedDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xPeakExtendedDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak hours, when disconnected $($Group.OffPeakDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakDisconnectAction,$htmlwhite))
			$rowdata += @(,("During off-peak extended hours, when disconnected $($Group.OffPeakExtendedDisconnectTimeout) mins",($htmlsilver -bor $htmlbold),$xOffPeakExtendedDisconnectAction,$htmlwhite))
		}
		
		$rowdata += @(,("Automatic power on for assigned",($htmlsilver -bor $htmlbold), $xAutoPowerOnForAssigned,$htmlwhite))
		$rowdata += @(,("Automatic power on for assigned during peak",($htmlsilver -bor $htmlbold), $xAutoPowerOnForAssignedDuringPeak,$htmlwhite))

		$rowdata += @(,('All connections not through NetScaler Gateway',($htmlsilver -bor $htmlbold),$xAllConnections,$htmlwhite))
		$rowdata += @(,('Connections through NetScaler Gateway',($htmlsilver -bor $htmlbold),$xNSConnection,$htmlwhite))
		$rowdata += @(,('Connections meeting any of the following filters',($htmlsilver -bor $htmlbold),$xAGFilters[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xAGFilters)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputDeliveryGroupApplicationDetails 
{
	Param([object] $Group)
	
	$AllApplications = Get-BrokerApplication -AssociatedDesktopGroupUid $Group.Uid
	
	If($? -and $Null -ne $AllApplications)
	{
		$txt = "Applications"
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 $txt
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 4 0 $txt
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $AllApplicationsWordTable = @();
		}
		ElseIf($HTML)
		{
			$rowdata = @()
		}

		ForEach($Application in $AllApplications)
		{
			Write-Verbose "$(Get-Date): `t`t`tAdding Application $($Application.ApplicationName)"

			$xEnabled = "Enabled"
			If($Application.Enabled -eq $False)
			{
				$xEnabled = "Disabled"
			}
			
			$xLocation = "Master Image"
			If($Application.MetadataKeys.Count -gt 0)
			{
				$xLocation = "App-V"
			}
			
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{
				ApplicationName = $Application.ApplicationName; 
				Description = $Application.Description; 
				Location = $xLocation;
				Enabled = $xEnabled; 
				}
				$AllApplicationsWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 1 "Name`t`t: " $Application.ApplicationName
				Line 1 "Description`t: " $Application.Description
				Line 1 "Location`t: " $xLocation
				Line 1 "State`t`t: " $xEnabled
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$Application.ApplicationName,$htmlwhite,
				$Application.Description,$htmlwhite,
				$xLocation,$htmlwhite,
				$xEnabled,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $AllApplicationsWordTable `
			-Columns  ApplicationName,Description,Location,Enabled `
			-Headers  "Name","Description","Location","State" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 175;
			$Table.Columns.Item(2).Width = 170;
			$Table.Columns.Item(3).Width = 100;
			$Table.Columns.Item(4).Width = 55;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'Location',($htmlsilver -bor $htmlbold),
			'State',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("175","170","100","55")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputDeliveryGroupCatalogs 
{
	Param([object] $Group)
	
	$MCs = Get-BrokerDesktop @XDParams1 -DesktopGroupUid $Group.Uid -Property CatalogName
	
	If($? -and $Null -ne $MCs)
	{
		If($MCs.Count -gt 1)
		{
			[array]$MCs = $MCs | Sort -Unique
		}
		
		$txt = "Machine Catalogs"
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 $txt
			[System.Collections.Hashtable[]] $CatalogsWordTable = @();
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 4 0 $txt
			$rowdata = @()
		}

		ForEach($MC in $MCs)
		{
			Write-Verbose "$(Get-Date): `t`t`tAdding catalog $($MC.CatalogName)"

			$Catalog = Get-BrokerCatalog @XDParams1 -Name $MC.CatalogName
			If($? -and $Null -ne $Catalog)
			{
				Switch ($Catalog.AllocationType)
				{
					"Static"	{$xAllocationType = "Permanent"; Break}
					"Permanent"	{$xAllocationType = "Permanent"; Break}
					"Random"	{$xAllocationType = "Random"; Break}
					Default		{$xAllocationType = "Allocation type could not be determined: $($Catalog.AllocationType)"; Break}
				}
				
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{
					Name = $Catalog.Name; 
					Type = $xAllocationType; 
					DesktopsTotal = $Catalog.AssignedCount;
					DesktopsFree = $Catalog.AvailableCount; 
					}
					$CatalogsWordTable += $WordTableRowHash;
				}
				ElseIf($Text)
				{
					Line 1 "Machine Catalog name`t: " $Catalog.Name
					Line 1 "Machine Catalog type`t: " $xAllocationType
					Line 1 "Desktops total`t`t: " $Catalog.AssignedCount
					Line 1 "Desktops free`t`t: " $Catalog.AvailableCount
					Line 0 ""
				}
				ElseIf($HTML)
				{
					$rowdata += @(,(
					$Catalog.Name,$htmlwhite,
					$xAllocationType,$htmlwhite,
					$Catalog.AssignedCount,$htmlwhite,
					$Catalog.AvailableCount,$htmlwhite))
				}
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CatalogsWordTable `
			-Columns  Name,Type,DesktopsTotal,DesktopsFree `
			-Headers  "Machine Catalog name","Machine Catalog type","Desktops total","Desktops free" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 175;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 100;
			$Table.Columns.Item(4).Width = 75;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Machine Catalog name',($htmlsilver -bor $htmlbold),
			'Machine Catalog type',($htmlsilver -bor $htmlbold),
			'Desktops total',($htmlsilver -bor $htmlbold),
			'Desktops free',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("175","150","100","75")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputDeliveryGroupAppDisks
{
	Param([object] $Group)
	
	$txt = "AppDisks"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 $txt
		[System.Collections.Hashtable[]] $CatalogsWordTable = @();
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $txt
		$rowdata = @()
	}

	$GroupAppDisks = $Group.AppDisks
	
	If($GroupAppDisks.Count -gt 0)
	{
		$AppDisks = @()
		ForEach($AppDisk in $GroupAppDisks)
		{
			$Result = Get-AppLibAppDisk @XDParams1 -AppDiskUid $AppDisk.Guid
			
			If($? -and $Null -ne $Result)
			{
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name Name        -Value $Result.AppDiskName
				$obj | Add-Member -MemberType NoteProperty -Name Description -Value $Result.Description
				$obj | Add-Member -MemberType NoteProperty -Name State       -Value $Result.State
				
				$AppDisks += $obj
			}
		}
		
		$AppDisks = $AppDisks | Sort Name
		
		ForEach($AppDisk in $AppDisks)
		{
			Write-Verbose "$(Get-Date): `t`t`tAdding AppDisk $($AppDisk.Name)"

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{
				Name = $AppDisk.Name; 
				Description = $AppDisk.Description; 
				State = $AppDisk.State;
				}
				$CatalogsWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 1 "Name`t`t: " $AppDisk.Name
				Line 1 "Description`t: " $AppDisk.Description
				Line 1 "State`t`t: " $AppDisk.State
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$AppDisk.Name,$htmlwhite,
				$AppDisk.Description,$htmlwhite,
				$AppDisk.State,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CatalogsWordTable `
			-Columns  Name,Description,State `
			-Headers  "Name","Description","State" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'State',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("150","150","150")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt = "There are no AppDisks for $($Group.Name)"
		OutputWarning $txt
	}
}
	
Function OutputDeliveryGroupUtilization
{
	Param([object]$Group)

	#code contributed by Eduardo Molina
	#Twitter: @molikop
	#eduardo@molikop.com
	#www.molikop.com

	$txt = "Delivery Group Utilization Report"
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing Utilization for $($Group.Name)"
		WriteWordLine 3 0 $txt
		WriteWordLine 4 0 "Desktop Group Name: " $Group.Name

		$xEnabled = ""
		If($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $True)
		{
			$xEnabled = "Maintenance Mode"
		}
		ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $True)
		{
			$xEnabled = "Maintenance Mode"
		}
		ElseIf($Group.Enabled -eq $True -and $Group.InMaintenanceMode -eq $False)
		{
			$xEnabled = "Enabled"
		}
		ElseIf($Group.Enabled -eq $False -and $Group.InMaintenanceMode -eq $False)
		{
			$xEnabled = "Disabled"
		}

		$xColorDepth = ""
		If($Group.ColorDepth -eq "FourBit")
		{
			$xColorDepth = "4bit - 16 colors"
		}
		ElseIf($Group.ColorDepth -eq "EightBit")
		{
			$xColorDepth = "8bit - 256 colors"
		}
		ElseIf($Group.ColorDepth -eq "SixteenBit")
		{
			$xColorDepth = "16bit - High color"
		}
		ElseIf($Group.ColorDepth -eq "TwentyFourBit")
		{
			$xColorDepth = "24bit - True color"
		}

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Description"; Value = $Group.Description; }
		$ScriptInformation += @{Data = "User Icon Name"; Value = $Group.PublishedName; }
		$ScriptInformation += @{Data = "Desktop Type"; Value = $Group.DesktopKind; }
		$ScriptInformation += @{Data = "Status"; Value = $xEnabled; }
		$ScriptInformation += @{Data = "Automatic reboots when user logs off"; Value = $Group.ShutdownDesktopsAfterUse; }
		$ScriptInformation += @{Data = "Color Depth"; Value = $xColorDepth; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
		
		Write-Verbose "$(Get-Date): `t`t`tInitializing utilization chart for $($Group.Name)"

		$TempFile =  "$($pwd)\emtempgraph_$(Get-Date -UFormat %Y%m%d_%H%M%S).csv"		
		Write-Verbose "$(Get-Date): `t`t`tGetting utilization data for $($Group.Name)"
		$Results = Get-BrokerDesktopUsage @XDParams2 -DesktopGroupName $Group.Name -SortBy Timestamp | Select-Object Timestamp, InUse

		If($? -and $Null -ne $Results)
		{
			$Results | Export-Csv $TempFile -NoTypeInformation *>$Null

			#Create excel COM object 
			$excel = New-Object -ComObject excel.application 4>$Null

			#Make not visible 
			$excel.Visible  = $False
			$excel.DisplayAlerts  = $False

			#Various Enumerations 
			$xlDirection = [Microsoft.Office.Interop.Excel.XLDirection] 
			$excelChart = [Microsoft.Office.Interop.Excel.XLChartType]
			$excelAxes = [Microsoft.Office.Interop.Excel.XlAxisType]
			$excelCategoryScale = [Microsoft.Office.Interop.Excel.XlCategoryType]
			$excelTickMark = [Microsoft.Office.Interop.Excel.XlTickMark]

			Write-Verbose "$(Get-Date): `t`t`tOpening Excel with temp file $($TempFile)"

			#Add CSV File into Excel Workbook 
			$Null = $excel.Workbooks.Open($TempFile)
			$worksheet = $excel.ActiveSheet
			$Null = $worksheet.UsedRange.EntireColumn.AutoFit()

			#Assumes that date is always on A column 
			$range = $worksheet.Range("A2")
			$selectionXL = $worksheet.Range($range,$range.end($xlDirection::xlDown))
			$Start = @($selectionXL)[0].Text
			$End = @($selectionXL)[-1].Text

			Write-Verbose "$(Get-Date): `t`t`tCreating chart for $($Group.Name)"
			$chart = $worksheet.Shapes.AddChart().Chart 

			$chart.chartType = $excelChart::xlXYScatterLines
			$chart.HasLegend = $false
			$chart.HasTitle = $true
			$chart.ChartTitle.Text = "$($Group.Name) utilization"

			#Work with the X axis for the Date Stamp 
			$xaxis = $chart.Axes($excelAxes::XlCategory)                                     
			$xaxis.HasTitle = $False
			$xaxis.CategoryType = $excelCategoryScale::xlCategoryScale
			$xaxis.MajorTickMark = $excelTickMark::xlTickMarkCross
			$xaxis.HasMajorGridLines = $true
			$xaxis.TickLabels.NumberFormat = "m/d/yyyy"
			$xaxis.TickLabels.Orientation = 48 #degrees to rotate text

			#Work with the Y axis for the number of desktops in use                                               
			$yaxis = $chart.Axes($excelAxes::XlValue)
			$yaxis.HasTitle = $true                                                       
			$yaxis.AxisTitle.Text = "Desktops in use"
			$yaxis.AxisTitle.Font.Size = 12

			$worksheet.ChartObjects().Item(1).copy()
			$word.Selection.PasteAndFormat(13)  #Pastes an Excel chart as a picture

			Write-Verbose "$(Get-Date): `t`t`tClosing excel for $($Group.Name)"
			$excel.Workbooks.Close($false)
			$excel.Quit()

			FindWordDocumentEnd
			WriteWordLine 0 0 ""
			
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($selectionXL)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Chart)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet)){}
			While( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)){}

			If(Test-Path variable:excel)
			{
				Remove-Variable -Name excel 4>$Null
			}

			#If the Excel.exe process is still running for the user's sessionID, kill it
			$SessionID = (Get-Process -PID $PID).SessionId
			(Get-Process 'Excel' -ea 0 | ?{$_.sessionid -eq $Sessionid}) | Stop-Process 4>$Null
			
			Write-Verbose "$(Get-Date): `t`t`tDeleting temp files $($TempFile)"
			Remove-Item $TempFile *>$Null
		}
		ElseIf($? -and $Null -eq $Results)
		{
			$txt = "There is no Utilization data for $($Group.Name)"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Utilization data for $($Group.name)"
			OutputWarning $txt
		}
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputDeliveryGroupTags
{
	Param([object] $Group)

	$txt = "Tags"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 $txt
		[System.Collections.Hashtable[]] $CatalogsWordTable = @();
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $txt
		$rowdata = @()
	}

	$GroupTags = $Group.Tags
	
	If($GroupTags.Count -gt 0)
	{
		$Tags = @()
		ForEach($Tag in $GroupTags)
		{
			$Result = Get-BrokerTag @XDParams1 -Name $Tag
			
			If($? -and $Null -ne $Result)
			{
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name Name        -Value $Result.Name
				$obj | Add-Member -MemberType NoteProperty -Name Description -Value $Result.Description
				$obj | Add-Member -MemberType NoteProperty -Name AppliedTo   -Value "Delivery Group"
				
				$Tags += $obj
			}
		}
		
		$Tags = $Tags | Sort 
		
		ForEach($Tag in $Tags)
		{
			Write-Verbose "$(Get-Date): `t`t`tAdding Tag $($Tag.Name)"

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{
				Name = $Tag.Name; 
				Description = $Tag.Description; 
				AppliedTo = $Tag.AppliedTo;
				}
				$CatalogsWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 1 "Name`t`t: " $Tag.Name
				Line 1 "Description`t: " $Tag.Description
				Line 1 "Applied to`t: " $Tag.AppliedTo
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$Tag.Name,$htmlwhite,
				$Tag.Description,$htmlwhite,
				$Tag.AppliedTo,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CatalogsWordTable `
			-Columns  Name,Description,AppliedTo `
			-Headers  "Name","Description","Applied to" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'Applied to',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("150","150","150")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt = "There are no Tags for $($Group.Name)"
		OutputWarning $txt
	}
}
	
Function OutputDeliveryGroupApplicationGroups
{
	Param([object] $Group)

	$txt = "Application Groups"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 $txt
		[System.Collections.Hashtable[]] $CatalogsWordTable = @();
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $txt
		$rowdata = @()
	}
	
	$ApplicationGroups = Get-BrokerApplicationGroup @XDParams1 -AssociatedDesktopGroupUid $Group.Uid
	
	If($? -and $Null -ne $ApplicationGroups)
	{
		ForEach($ApplicationGroup in $ApplicationGroups)
		{
			Write-Verbose "$(Get-Date): `t`t`tAdding Application Group $($ApplicationGroup.Name)"

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{
				Name = $ApplicationGroup.Name; 
				Description = $ApplicationGroup.Description; 
				Applications = $ApplicationGroup.TotalApplications.ToString();
				}
				$CatalogsWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 1 "Name`t`t: " $ApplicationGroup.Name
				Line 1 "Description`t: " $ApplicationGroup.Description
				Line 1 "Applications`t: " $ApplicationGroup.TotalApplications.ToString()
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$ApplicationGroup.Name,$htmlwhite,
				$ApplicationGroup.Description,$htmlwhite,
				$ApplicationGroup.TotalApplications.ToString(),$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CatalogsWordTable `
			-Columns  Name,Description,Applications `
			-Headers  "Name","Description","Applications" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;
			$Table.Columns.Item(3).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'Applications',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("150","150","150")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf($? -and $Null -eq $ApplicationGroups)
	{
		$txt = "There are no Application Groups for $($Group.Name)"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Application Groups for $($Group.Name)"
		OutputWarning $txt
	}
}
#endregion

#region process application functions
Function ProcessApplications
{
	Write-Verbose "$(Get-Date): Retrieving Applications"
	
	$Global:TotalPublishedApplications = 0
	$Global:TotalAppvApplications = 0
	
	$AllApplications = Get-BrokerApplication @XDParams1 -SortBy Name
	If($? -and $Null -ne $AllApplications)
	{
		OutputApplications $AllApplications
	}
	ElseIf($? -and ($Null -eq $AllApplications))
	{
		$txt = "There are no Applications"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Applications"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputApplications
{
	Param([object]$AllApplications)
	
	Write-Verbose "$(Get-Date): `tProcessing Applications"

	$txt = "Applications"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $AllApplicationsWordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	ForEach($Application in $AllApplications)
	{
		Write-Verbose "$(Get-Date): `t`tAdding Application $($Application.ApplicationName)"

		$xEnabled = "Enabled"
		If($Application.Enabled -eq $False)
		{
			$xEnabled = "Disabled"
		}
		
		$xLocation = "Master Image"
		If($Application.MetadataKeys.Count -gt 0)
		{
			$xLocation = "App-V"
		}

		If($xLocation -eq "Master Image")
		{
			$Global:TotalPublishedApplications++
		}
		Else
		{
			$Global:TotalAppvApplications++
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			ApplicationName = $Application.ApplicationName; 
			Description = $Application.Description; 
			Location = $xLocation;
			Enabled = $xEnabled; 
			}
			$AllApplicationsWordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "Name`t`t: " $Application.ApplicationName
			Line 1 "Description`t: " $Application.Description
			Line 1 "Location`t: " $xLocation
			Line 1 "State`t`t: " $xEnabled
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Application.ApplicationName,$htmlwhite,
			$Application.Description,$htmlwhite,
			$xLocation,$htmlwhite,
			$xEnabled,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $AllApplicationsWordTable `
		-Columns  ApplicationName,Description,Location,Enabled `
		-Headers  "Name","Description","Location","State" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 170;
		$Table.Columns.Item(3).Width = 100;
		$Table.Columns.Item(4).Width = 55;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Location',($htmlsilver -bor $htmlbold),
		'State',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("175","170","100","55")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}

	If($Applications)
	{
		ForEach($Application in $AllApplications)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 2 0 $Application.ApplicationName
			}
			ElseIf($Text)
			{
				Line 0 ""
				Line 0 $Application.ApplicationName
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 $Application.ApplicationName
			}
			
			OutputApplicationDetails $Application
			OutputApplicationSessions $Application
			OutputApplicationAdministrators $Application
		}
	}
}

Function OutputApplicationDetails
{
	Param([object] $Application)
	
	Write-Verbose "$(Get-Date): `t`tApplication details for $($Application.ApplicationName)"
	$txt = "Details"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$xTags = @()
	ForEach($Tag in $Application.Tags)
	{
		$xTags += "$($Tag)"
	}
	$xVisibility = @()
	If($Application.UserFilterEnabled)
	{
		$cnt = -1
		ForEach($tmp in $Application.AssociatedUserFullNames)
		{
			$cnt++
			$xVisibility += "$($tmp) ($($Application.AssociatedUserNames[$cnt]))"
		}
		
	}
	Else
	{
		$xVisibility = {Users inherited from Delivery Group}
	}
	
	$DeliveryGroups = @()
	If($Application.AssociatedDesktopGroupUids.Count -gt 1)
	{
		$cnt = -1
		ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
		{
			$cnt++
			$Results = Get-BrokerDesktopGroup @XDParams1 -Uid $DGUid
			If($? -and $Null -ne $Results)
			{
				$DeliveryGroups += "$($Results.Name) Priority: $($Application.AssociatedDesktopGroupPriorities[$cnt])"
			}
		}
	}
	Else
	{
		ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
		{
			$Results = Get-BrokerDesktopGroup @XDParams1 -Uid $DGUid
			If($? -and $Null -ne $Results)
			{
				$DeliveryGroups += $Results.Name
			}
		}
	}
	
	$RedirectedFileTypes = @()
	$Results = Get-BrokerConfiguredFTA -ApplicationUid $Application.Uid
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$RedirectedFileTypes += $Result.ExtensionName
		}
	}
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Name (for administrator)"; Value = $Application.Name; }
		$ScriptInformation += @{Data = "Name (for user)"; Value = $Application.PublishedName; }
		$ScriptInformation += @{Data = "Description"; Value = $Application.Description; }
		$ScriptInformation += @{Data = "Delivery Group"; Value = $DeliveryGroups[0]; }
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $Group; }
			}
		}
		$ScriptInformation += @{Data = "Folder (for administrators)"; Value = $Application.AdminFolderName; }
		$ScriptInformation += @{Data = "Folder (for user)"; Value = $Application.ClientFolder; }
		$ScriptInformation += @{Data = "Visibility"; Value = $xVisibility[0]; }
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $xVisibility[$cnt]; }
			}
		}
		$ScriptInformation += @{Data = "Application Path"; Value = $Application.CommandLineExecutable; }
		$ScriptInformation += @{Data = "Command line arguments"; Value = $Application.CommandLineArguments; }
		$ScriptInformation += @{Data = "Working directory"; Value = $Application.WorkingDirectory; }
		If($Null -eq $RedirectedFileTypes)
		{
			$ScriptInformation += @{Data = "Redirected file types"; Value = ""; }
		}
		Else
		{
			$tmp1 = ""
			ForEach($tmp in $RedirectedFileTypes)
			{
				$tmp1 += "$($tmp); "
			}
			$ScriptInformation += @{Data = "Redirected file types"; Value = $tmp1; }
			$tmp1 = $Null
		}
		$ScriptInformation += @{Data = "Tags"; Value = $xTags[0]; }
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		
		If($Application.Visible -eq $False)
		{
			$ScriptInformation += @{Data = "Hidden"; Value = "Application is hidden"; }
		}
		
		If((Get-BrokerServiceAddedCapability @XDParams1) -contains "ApplicationUsageLimits")
		{
			
			$tmp = ""
			If($Application.MaxTotalInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxTotalInstances.ToString()
			}
			$ScriptInformation += @{Data = "Maximum concurrent instances"; Value = $tmp; }
			
			$tmp = ""
			If($Application.MaxPerUserInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxPerUserInstances.ToString()
			}
			$ScriptInformation += @{Data = "Maximum instances per user"; Value = $tmp; }
		}
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 175;
		$Table.Columns.Item(2).Width = 325;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 "Name (for administrator)`t: " $Application.Name
		Line 1 "Name (for user)`t`t`t: " $Application.PublishedName
		Line 1 "Description`t`t`t: " $Application.Description
		Line 1 "Delivery Group`t`t`t: " $DeliveryGroups[0]
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $Group
			}
		}
		Line 1 "Folder (for administrators)`t: " $Application.AdminFolderName
		Line 1 "Folder (for user)`t`t: " $Application.ClientFolder
		Line 1 "Visibility`t`t`t: " $xVisibility[0]
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $xVisibility[$cnt]
			}
		}
		Line 1 "Application Path`t`t: " $Application.CommandLineExecutable
		Line 1 "Command line arguments`t`t: " $Application.CommandLineArguments
		Line 1 "Working directory`t`t: " $Application.WorkingDirectory
		$tmp1 = ""
		ForEach($tmp in $RedirectedFileTypes)
		{
			$tmp1 += "$($tmp); "
		}
		Line 1 "Redirected file types`t`t: " $tmp1
		$tmp1 = $Null
		Line 1 "Tags`t`t`t`t: " $xTags[0]
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}

		If($Application.Visible -eq $False)
		{
			Line 1 "Hidden`t`t`t`t: Application is hidden" ""
		}
		
		If((Get-BrokerServiceAddedCapability @XDParams1) -contains "ApplicationUsageLimits")
		{
			
			$tmp = ""
			If($Application.MaxTotalInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxTotalInstances.ToString()
			}
			Line 1 "Maximum concurrent instances`t: " $tmp
			
			$tmp = ""
			If($Application.MaxPerUserInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxPerUserInstances.ToString()
			}
			Line 1 "Maximum instances per user`t: " $tmp
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name (for administrator)",($htmlsilver -bor $htmlbold),$Application.Name,$htmlwhite)
		$rowdata += @(,('Name (for user)',($htmlsilver -bor $htmlbold),$Application.PublishedName,$htmlwhite))
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Application.Description,$htmlwhite))
		$rowdata += @(,('Delivery Group',($htmlsilver -bor $htmlbold),$DeliveryGroups[0],$htmlwhite))
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$Group,$htmlwhite))
			}
		}
		$rowdata += @(,('Folder (for administrators)',($htmlsilver -bor $htmlbold),$Application.AdminFolderName,$htmlwhite))
		$rowdata += @(,('Folder (for user)',($htmlsilver -bor $htmlbold),$Application.ClientFolder,$htmlwhite))
		$rowdata += @(,('Visibility',($htmlsilver -bor $htmlbold),$xVisibility[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xVisibility)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$xVisibility[$cnt],$htmlwhite))
			}
		}
		$rowdata += @(,('Application Path',($htmlsilver -bor $htmlbold),$Application.CommandLineExecutable,$htmlwhite))
		$rowdata += @(,('Command Line arguments',($htmlsilver -bor $htmlbold),$Application.CommandLineArguments,$htmlwhite))
		$rowdata += @(,('Working directory',($htmlsilver -bor $htmlbold),$Application.WorkingDirectory,$htmlwhite))
		$tmp1 = ""
		ForEach($tmp in $RedirectedFileTypes)
		{
			$tmp1 += "$($tmp); "
		}
		$rowdata += @(,('Redirected file types',($htmlsilver -bor $htmlbold),$tmp1,$htmlwhite))
		$tmp1 = $Null
		$rowdata += @(,('Tags',($htmlsilver -bor $htmlbold),$xTags[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xTags)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}

		If($Application.Visible -eq $False)
		{
			$rowdata += @(,('Hidden',($htmlsilver -bor $htmlbold),"Application is hidden",$htmlwhite))
		}

		If((Get-BrokerServiceAddedCapability @XDParams1) -contains "ApplicationUsageLimits")
		{
			$tmp = ""
			If($Application.MaxTotalInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxTotalInstances.ToString()
			}
			$rowdata += @(,('Maximum concurrent instances',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			
			$tmp = ""
			If($Application.MaxPerUserInstances -eq 0)
			{
				$tmp = "Unlimited"
			}
			Else
			{
				$tmp = $Application.MaxPerUserInstances.ToString()
			}
			$rowdata += @(,('Maximum instances per user',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		$msg = ""
		$columnWidths = @("175","325")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputApplicationSessions
{
	Param([object] $Application)
	
	Write-Verbose "$(Get-Date): `t`tApplication sessions for $($Application.BrowserName)"
	$txt = "Sessions"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$Sessions = Get-BrokerSession -ApplicationUid $Application.Uid @XDParams1 -SortBy UserName
	
	If($? -and $Null -ne $Sessions)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $SessionsWordTable = @();
		}
		ElseIf($HTML)
		{
			$rowdata = @()
		}

		#now get the privateappdesktop for each desktopgroup uid
		ForEach($Session in $Sessions)
		{
			#get desktop by Session Uid
			$xMachineName = ""
			$Desktop = Get-BrokerDesktop -SessionUid $Session.Uid @XDParams1
			
			If($? -and $Null -ne $Desktop)
			{
				$xMachineName = $Desktop.MachineName
			}
			Else
			{
				$xMachineName = "Not Found"
			}
			
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{
				UserName = $Session.UserName;
				ClientName= $Session.ClientName;
				MachineName = $xMachineName;
				State = $Session.SessionState;
				ApplicationState = $Session.AppState;
				Protocol = $Session.Protocol;
				}
				$SessionsWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 2 "User Name`t: " $Session.UserName
				Line 2 "Client Name`t: " $Session.ClientName
				Line 2 "Machine Name`t: " $xMachineName
				Line 2 "State`t`t: " $Session.SessionState
				Line 2 "Application State`t`t: " $Session.AppState
				Line 2 "Protocol`t: " $Session.Protocol
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$Session.UserName,$htmlwhite,
				$Session.ClientName,$htmlwhite,
				$xMachineName,$htmlwhite,
				$Session.SessionState,$htmlwhite,
				$Session.AppState,$htmlwhite,
				$Session.Protocol,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $SessionsWordTable `
			-Columns  UserName,ClientName,MachineName,State,ApplicationState,Protocol `
			-Headers  "User Name","Client Name","Machine Name","State","Application State","Protocol" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 135;
			$Table.Columns.Item(2).Width = 85;
			$Table.Columns.Item(3).Width = 135;
			$Table.Columns.Item(4).Width = 50;
			$Table.Columns.Item(5).Width = 50;
			$Table.Columns.Item(6).Width = 55;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'User Name',($htmlsilver -bor $htmlbold),
			'Client Name',($htmlsilver -bor $htmlbold),
			'Machine Name',($htmlsilver -bor $htmlbold),
			'State',($htmlsilver -bor $htmlbold),
			'Application State',($htmlsilver -bor $htmlbold),
			'Protocol',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("135","85","135","50","50","55")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "510"
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf($? -and $Null -eq $Sessions)
	{
		$txt = "There are no Sessions for Application $($Application.ApplicationName)"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Sessions for Application $($Application.ApplicationName)"
		OutputWarning $txt
	}
}

Function OutputApplicationAdministrators
{
	Param([object] $Application)
	
	Write-Verbose "$(Get-Date): `t`tApplication administrators for $($Application.ApplicationName)"
	$txt = "Administrators"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	#get all the delivery groups
	$DeliveryGroups = @()
	ForEach($DGUid in $Application.AssociatedDesktopGroupUids)
	{
		$Results = Get-BrokerDesktopGroup @XDParams1 -Uid $DGUid
		If($? -and $Null -ne $Results)
		{
			$DeliveryGroups += $Results.Name
		}
	}
	
	#now get the administrators for each delivery group
	$Admins = @()
	ForEach($Group in $DeliveryGroups)
	{
		$Results = GetAdmins "DesktopGroup" $Group
		If($? -and $Null -ne $Results)
		{
			$Admins += $Results
		}
	}
	
	If($Null -ne $Admins)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $AdminsWordTable = @();
		}
		ElseIf($HTML)
		{
			$rowdata = @()
		}
		
		ForEach($Admin in $Admins)
		{
			$Tmp = ""
			If($Admin.Enabled)
			{
				$Tmp = "Enabled"
			}
			Else
			{
				$Tmp = "Disabled"
			}
			
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				AdminName = $Admin.Name;
				Role = $Admin.Rights[0].RoleName;
				Status = $Tmp;
				}
				$AdminsWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 1 "Administrator Name`t: " $Admin.Name
				Line 1 "Role`t`t`t: " $Admin.Rights[0].RoleName
				Line 1 "Status`t`t`t: " $tmp
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$Admin.Name,$htmlwhite,
				$Admin.Rights[0].RoleName,$htmlwhite,
				$tmp,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $AdminsWordTable `
			-Columns AdminName, Role, Status `
			-Headers "Administrator Name", "Role", "Status" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 225;
			$Table.Columns.Item(2).Width = 200;
			$Table.Columns.Item(3).Width = 60;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Administrator Name',($htmlsilver -bor $htmlbold),
			'Role',($htmlsilver -bor $htmlbold),
			'Status',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("225","200","60")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "485"
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf($? -and ($Null -eq $Admins))
	{
		$txt = "There are no administrators for $($Group.Name)"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve administrators for $($Group.Name)"
		OutputWarning $txt
	}
	
}
#endregion

#region application group details
Function ProcessApplicationGroupDetails
{
	Write-Verbose "$(Get-Date): `tProcessing Application Groups"

	$txt = "Application Groups"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$ApplicationGroups = Get-BrokerApplicationGroup @XDParams1 -SortBy Name
	
	If($? -and $Null -ne $ApplicationGroups)
	{
		ForEach($AppGroup in $ApplicationGroups)
		{
			Write-Verbose "$(Get-Date): `t`t`tAdding Application Group $($ApplicationGroup.Name)"
#line 11950
			$xEnabled = "No"
			If($AppGroup.Enabled)
			{
				$xEnabled = "Yes"
			}
			
			$xSessionSharing = "Disabled"
			If($AppGroup.SessionSharingEnabled)
			{
				$xSessionSharing = "Enabled"
			}
			
			$xSingleSession = "Disabled"
			If($AppGroup.SingleAppPerSession)
			{
				$xSingleSession = "Enabled"
			}
			
			$DGs = @()
			ForEach($DGUid in $AppGroup.AssociatedDesktopGroupUids)
			{
				$results = Get-BrokerDesktopGroup @XDParams1 -Uid $DGUid
				
				If($? -and $Null -ne $results)
				{
					$DGs += $results.Name
				}
			}
			
			[array]$xTags = $AppGroup.Tags
			
			If($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{Data = "Name"; Value = $AppGroup.Name; }
				$ScriptInformation += @{Data = "Description"; Value = $AppGroup.Description; }
				$ScriptInformation += @{Data = "Applications"; Value = $AppGroup.TotalApplications.ToString(); }
				
				If($Null -eq $AppGroup.Scopes)
				{
					$ScriptInformation += @{Data = "Scopes"; Value = "All"; }
				}
				Else
				{
					$ScriptInformation += @{Data = "Scopes"; Value = "All"; }
					$cnt = -1
					ForEach($tmp in $AppGroup.Scopes)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
				
				$ScriptInformation += @{Data = "Enabled"; Value = $xEnabled; }
				$ScriptInformation += @{Data = "Session sharing"; Value = $xSessionSharing; }
				$ScriptInformation += @{Data = "Single application per session"; Value = $xSingleSession; }
				$ScriptInformation += @{Data = "Restrict launches to machines with tag"; Value = $AppGroup.RestrictToTag; }

				$ScriptInformation += @{Data = "Delivery Groups"; Value = $DGs[0]; }
				$cnt = -1
				ForEach($tmp in $DGs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
				
				$ScriptInformation += @{Data = "Tags"; Value = $xTags[0]; }
				$cnt = -1
				ForEach($tmp in $xTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{Data = ""; Value = $tmp; }
					}
				}
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Name`t`t`t`t`t: " $AppGroup.Name
				Line 1 "Description`t`t`t`t: " $AppGroup.Description
				Line 1 "Applications`t`t`t`t: " $AppGroup.TotalApplications.ToString()
				
				If($Null -eq $AppGroup.Scopes)
				{
					Line 1 "Scopes`t`t`t`t`t: " "All"
				}
				Else
				{
					Line 1 "Scopes`t`t`t`t`t: " "All"
					$cnt = -1
					ForEach($tmp in $AppGroup.Scopes)
					{
						Line 6 "  " $tmp
					}
				}
				
				Line 1 "Enabled`t`t`t`t`t: " $xEnabled
				Line 1 "Session sharing`t`t`t`t: " $xSessionSharing
				Line 1 "Single application per session`t`t: " $xSingleSession
				Line 1 "Restrict launches to machines with tag`t: " $AppGroup.RestrictToTag

				Line 1 "Delivery Groups`t`t`t`t: " $DGs[0]
				$cnt = -1
				ForEach($tmp in $DGs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
				
				Line 1 "Tags`t`t`t`t`t: " $xTags[0]
				$cnt = -1
				ForEach($tmp in $xTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 6 "  " $tmp
					}
				}
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$AppGroup.Name,$htmlwhite)
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$AppGroup.Description,$htmlwhite))
				$rowdata += @(,('Applications',($htmlsilver -bor $htmlbold),$AppGroup.TotalApplications.ToString(),$htmlwhite))
				
				If($Null -eq $AppGroup.Scopes)
				{
					$rowdata += @(,('Scopes',($htmlsilver -bor $htmlbold),"All",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('Scopes',($htmlsilver -bor $htmlbold),"All",$htmlwhite))
					$cnt = -1
					ForEach($tmp in $AppGroup.Scopes)
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}
				}
				
				$rowdata += @(,('Enabled',($htmlsilver -bor $htmlbold),$xEnabled,$htmlwhite))
				$rowdata += @(,('Session sharing',($htmlsilver -bor $htmlbold),$xSessionSharing,$htmlwhite))
				$rowdata += @(,('Single application per session',($htmlsilver -bor $htmlbold),$xSingleSession,$htmlwhite))
				$rowdata += @(,('Restrict launches to machines with tag',($htmlsilver -bor $htmlbold),$AppGroup.RestrictToTag,$htmlwhite))
				
				$rowdata += @(,('Delivery Groups',($htmlsilver -bor $htmlbold),$DGs[0],$htmlwhite))
				$cnt = -1
				ForEach($tmp in $DGs)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}
				}
				
				$rowdata += @(,('Tags',($htmlsilver -bor $htmlbold),$xTags[0],$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xTags)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}
				}

				$msg = ""
				$columnWidths = @("225","275")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 " "
				
			}
		}
	}
	ElseIf($? -and $Null -eq $ApplicationGroups)
	{
		$txt = "There were no Application Groups found"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Application Groups"
		OutputWarning $txt
	}
}
#endregion

#region policy functions
Function ProcessPolicies
{
	$txt = "Policies"
	$txt1 = "Policies in this report may not match the order shown in Studio."
	$txt2 = "See http://blogs.citrix.com/2013/07/15/merging-of-user-and-computer-policies-in-xendesktop-7-0/"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 $txt
		WriteWordLine 0 0 $txt1 "" $Null 8 $False $True	
		WriteWordLine 0 0 $txt2 "" $Null 8 $False $True	
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 $txt1
		Line 0 $txt2
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 $txt
		WriteHTMLLine 0 0 $txt1
		WriteHTMLLine 0 0 $txt2
	}
	Write-Verbose "$(Get-Date): Processing XenDesktop Policies"
	
	ProcessPolicySummary 
	
	If($Policies)
	{
	
		Write-Verbose "$(Get-Date): `tDoes localfarmgpo PSDrive already exist?"
		If(Get-PSDrive localfarmgpo -EA 0)
		{
			Write-Verbose "$(Get-Date): `tRemoving the current localfarmgpo PSDrive"
			Remove-PSDrive localfarmgpo -EA 0 4>$Null
		}
		
		Write-Verbose "$(Get-Date): Creating localfarmgpo PSDrive for Computer policies"
		New-PSDrive localfarmgpo -psprovider citrixgrouppolicy -root \ -controller $AdminAddress -Scope Global *>$Null
		If(Get-PSDrive localfarmgpo -EA 0)
		{
			ProcessCitrixPolicies "localfarmgpo" "Computer"
			Write-Verbose "$(Get-Date): Finished Processing Citrix Site Computer Policies"
			Write-Verbose "$(Get-Date): "
		}
		Else
		{
			Write-Warning "Unable to create the LocalFarmGPO PSDrive on the XenDesktop Controller $($AdminAddress)"
		}

		Write-Verbose "$(Get-Date): Creating localfarmgpo PSDrive for User policies"
		New-PSDrive localfarmgpo -psprovider citrixgrouppolicy -root \ -controller $AdminAddress -Scope Global *>$Null
		If(Get-PSDrive localfarmgpo -EA 0)
		{
			ProcessCitrixPolicies "localfarmgpo" "User"
			Write-Verbose "$(Get-Date): Finished Processing Citrix Site User Policies"
			Write-Verbose "$(Get-Date): "
		}
		Else
		{
			Write-Warning "Unable to create the LocalFarmGPO PSDrive on the XenDesktop Controller $($AdminAddress)"
		}
		
		If($NoADPolicies)
		{
			#don't process AD policies
		}
		Else
		{
			#thanks to the Citrix Engineering Team for helping me solve processing Citrix AD based Policies
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): `tSee if there are any Citrix AD based policies to process"
			$CtxGPOArray = @()
			$CtxGPOArray = GetCtxGPOsInAD
			If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
			{
				Write-Verbose "$(Get-Date): "
				Write-Verbose "$(Get-Date): `tThere are $($CtxGPOArray.Count) Citrix AD based policies to process"
				Write-Verbose "$(Get-Date): "

				[array]$CtxGPOArray = $CtxGPOArray | Sort -unique
				
				ForEach($CtxGPO in $CtxGPOArray)
				{
					Write-Verbose "$(Get-Date): `tCreating ADGpoDrv PSDrive for Computer Policies"
					New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope Global *>$Null
					If(Get-PSDrive ADGpoDrv -EA 0)
					{
						Write-Verbose "$(Get-Date): `tProcessing Citrix AD Policy $($CtxGPO)"
					
						Write-Verbose "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)"
						ProcessCitrixPolicies "ADGpoDrv" "Computer"
						Write-Verbose "$(Get-Date): Finished Processing Citrix AD Computer Policy $($CtxGPO)"
						Write-Verbose "$(Get-Date): "
					}
					Else
					{
						Write-Warning "$($CtxGPO) is not readable by this XenDesktop Controller"
						Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
					}

					Write-Verbose "$(Get-Date): `tCreating ADGpoDrv PSDrive for UserPolicies"
					New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope Global *>$Null
					If(Get-PSDrive ADGpoDrv -EA 0)
					{
						Write-Verbose "$(Get-Date): `tProcessing Citrix AD Policy $($CtxGPO)"
					
						Write-Verbose "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)"
						ProcessCitrixPolicies "ADGpoDrv" "User"
						Write-Verbose "$(Get-Date): Finished Processing Citrix AD User Policy $($CtxGPO)"
						Write-Verbose "$(Get-Date): "
					}
					Else
					{
						Write-Warning "$($CtxGPO) is not readable by this XenDesktop Controller"
						Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
					}
				}
				Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policies"
				Write-Verbose "$(Get-Date): "
			}
			Else
			{
				Write-Verbose "$(Get-Date): There are no Citrix AD based policies to process"
				Write-Verbose "$(Get-Date): "
			}
		}
	}
	Write-Verbose "$(Get-Date): Finished Processing Citrix Policies"
	Write-Verbose "$(Get-Date): "
}

Function ProcessPolicySummary
{
	Write-Verbose "$(Get-Date): `tDoes localfarmgpo PSDrive already exist?"
	If(Get-PSDrive localfarmgpo -EA 0)
	{
		Write-Verbose "$(Get-Date): `tRemoving the current localfarmgpo PSDrive"
		Remove-PSDrive localfarmgpo -EA 0 4>$Null
	}
	Write-Verbose "$(Get-Date): `tRetrieving Site Policies"
	Write-Verbose "$(Get-Date): `t`tCreating localfarmgpo PSDrive"
	New-PSDrive localfarmgpo -psprovider citrixgrouppolicy -root \ -controller $AdminAddress -Scope Global *>$Null

	If(Get-PSDrive localfarmgpo -EA 0)
	{
		$HDXPolicies = Get-CtxGroupPolicy -DriveName localfarmgpo -EA 0 `
		| Select PolicyName, Type, Description, Enabled, Priority `
		| Sort Type, Priority
		
		OutputSummaryPolicyTable $HDXPolicies "localfarmgpo"
	}
	Else
	{
		Write-Warning "Unable to create the LocalFarmGPO PSDrive on the XenDesktop Controller $($AdminAddress)"
	}
	
	If($NoADPolicies)
	{
		#don't process AD policies
	}
	Else
	{
		Write-Verbose "$(Get-Date): "
		Write-Verbose "$(Get-Date): See if there are any Citrix AD based policies to process"
		$CtxGPOArray = @()
		$CtxGPOArray = GetCtxGPOsInAD
		If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
		{
			[array]$CtxGPOArray = $CtxGPOArray | Sort -unique
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): `tThere are $($CtxGPOArray.Count) Citrix AD based policies to process"
			Write-Verbose "$(Get-Date): "
			
			ForEach($CtxGPO in $CtxGPOArray)
			{
				Write-Verbose "$(Get-Date): `tCreating ADGpoDrv PSDrive"
				New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope "Global" *>$Null
				If(Get-PSDrive ADGpoDrv -EA 0)
				{
					Write-Verbose "$(Get-Date): `tProcessing Citrix AD Policy $($CtxGPO)"
				
					Write-Verbose "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)"
					$HDXPolicies = Get-CtxGroupPolicy -DriveName ADGpoDrv -EA 0 `
					| Select PolicyName, Type, Description, Enabled, Priority `
					| Sort Type, Priority
			
					OutputSummaryPolicyTable $HDXPolicies "AD" $CtxGPO
					
					Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policy $($CtxGPO)"
					Write-Verbose "$(Get-Date): "
				}
				Else
				{
					Write-Warning "$($CtxGPO) is not readable by this XenDesktop Controller"
					Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
				}
				Remove-PSDrive ADGpoDrv -EA 0 4>$Null
			}
			Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policies"
			Write-Verbose "$(Get-Date): "
		}
		Else
		{
			Write-Verbose "$(Get-Date): There are no Citrix AD based policies to process"
			Write-Verbose "$(Get-Date): "
		}
	}
}

Function OutputSummaryPolicyTable
{
	Param([object] $HDXPolicies, [string] $xLocation, [string] $ADGPOName = "")
	
	If($xLocation -eq "localfarmgpo")
	{
		$txt = "Site Policies"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt
		}
		ElseIf($Text)
		{
			Line 0 $txt
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt
		}
	}
	ElseIf($xLocation -eq "AD")
	{
		$txt = "Active Directory Policies ($($ADGpoName))"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt
		}
		ElseIf($Text)
		{
			Line 0 $txt
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 $txt
		}
	}

	If($Null -ne $HDXPolicies)
	{
		Write-Verbose "$(Get-Date): `t`t`tPolicies"
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $PoliciesWordTable = @();
		}
		ElseIf($HTML)
		{
			$rowdata = @()
		}

		ForEach($Policy in $HDXPolicies)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{
				Name = $Policy.PolicyName;
				Description = $Policy.Description;
				Enabled= $Policy.Enabled;
				Type = $Policy.Type;
				Priority = $Policy.Priority;
				}
				$PoliciesWordTable += $WordTableRowHash;
			}
			ElseIf($Text)
			{
				Line 2 "Name`t`t: " $Policy.PolicyName
				Line 2 "Description`t: " $Policy.Description
				Line 2 "Enabled`t`t: " $Policy.Enabled
				Line 2 "Type`t`t: " $Policy.Type
				Line 2 "Priority`t: " $Policy.Priority
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$Policy.PolicyName,$htmlwhite,
				$Policy.Description,$htmlwhite,
				$Policy.Enabled,$htmlwhite,
				$Policy.Type,$htmlwhite,
				$Policy.Priority,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $PoliciesWordTable `
			-Columns  Name,Description,Enabled,Type,Priority `
			-Headers  "Name","Description","Enabled","Type","Priority" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 155
			$Table.Columns.Item(2).Width = 185
			$Table.Columns.Item(3).Width = 55;
			$Table.Columns.Item(4).Width = 60;
			$Table.Columns.Item(5).Width = 45;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'Enabled',($htmlsilver -bor $htmlbold),
			'Type',($htmlsilver -bor $htmlbold),
			'Priority',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("155","185","55","60","45")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf($Null -eq $HDXPolicies)
	{
		$txt = "There are no Policies"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Policies"
		OutputWarning $txt
	}
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((gm -Name $topLevel -InputObject $object))
		{
			If((gm -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function ProcessCitrixPolicies
{
	Param([string]$xDriveName, [string]$xPolicyType)

	Write-Verbose "$(Get-Date): `tRetrieving all $($xPolicyType) policy names"

	$Global:TotalComputerPolicies = 0
	$Global:TotalUserPolicies = 0
	$Global:TotalSitePolicies = 0
	$Global:TotalADPolicies = 0
	$Global:TotalADPoliciesNotProcessed = 0
	$Global:TotalPolicies = 0
	$ADPoliciesNotProcessed = @()
	
	If($xDriveName -eq "localfarmgpo")
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "Site Policies"
		}
		ElseIf($Text)
		{
			Line 0 ""
			Line 0 "Site Policies"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "Site Policies"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "Active Directory Policies"
		}
		ElseIf($Text)
		{
			Line 0 ""
			Line 0 "Active Directory Policies"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 0 "Active Directory Policies"
		}
	}
	
	$Policies = Get-CtxGroupPolicy -Type $xPolicyType `
	-DriveName $xDriveName -EA 0 `
	| Select PolicyName, Type, Description, Enabled, Priority `
	| Sort Priority

	If($? -and $Null -ne $Policies)
	{
		ForEach($Policy in $Policies)
		{
			Write-Verbose "$(Get-Date): `tStarted $($Policy.PolicyName) "
			
			If($xDriveName -eq "localfarmgpo")
			{
				$Global:TotalSitePolicies++
			}
			Else
			{
				$Global:TotalADPolicies++
			}
			If($Policy.Type -eq "Computer")
			{
				$Global:TotalComputerPolicies++
			}
			Else
			{
				$Global:TotalUserPolicies++
			}
			$Global:TotalPolicies++
			
			If($MSWord -or $PDF)
			{
				$selection.InsertNewPage()
				If($xDriveName -eq "localfarmgpo")
				{
					WriteWordLine 2 0 "$($Policy.PolicyName) (Site, $($xPolicyType))"
				}
				Else
				{
					WriteWordLine 2 0 "$($Policy.PolicyName) (AD, $($xPolicyType))"
				}
				[System.Collections.Hashtable[]] $ScriptInformation = @()
			
				$ScriptInformation += @{Data = "Description"; Value = $Policy.Description; }
				$ScriptInformation += @{Data = "Enabled"; Value = $Policy.Enabled; }
				$ScriptInformation += @{Data = "Type"; Value = $Policy.Type; }
				$ScriptInformation += @{Data = "Priority"; Value = $Policy.Priority; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 90;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
			ElseIf($Text)
			{
				If($xDriveName -eq "localfarmgpo")
				{
					Line 0 "$($Policy.PolicyName) (Site, $($xPolicyType))"
				}
				Else
				{
					Line 0 "$($Policy.PolicyName) (AD, $($xPolicyType))"
				}
				If(![String]::IsNullOrEmpty($Policy.Description))
				{
					Line 1 "Description`t: " $Policy.Description
				}
				Line 1 "Enabled`t`t: " $Policy.Enabled
				Line 1 "Type`t`t: " $Policy.Type
				Line 1 "Priority`t: " $Policy.Priority
			}
			ElseIf($HTML)
			{
				If($xDriveName -eq "localfarmgpo")
				{
					WriteHTMLLine 2 0 "$($Policy.PolicyName) (Site, $($xPolicyType))"
				}
				Else
				{
					WriteHTMLLine 2 0 "$($Policy.PolicyName) (AD, $($xPolicyType))"
				}
				$rowdata = @()
				$columnHeaders = @("Description",($htmlsilver -bor $htmlbold),$Policy.Description,$htmlwhite)
				$rowdata += @(,('Enabled',($htmlsilver -bor $htmlbold),$Policy.Enabled,$htmlwhite))
				$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$Policy.Type,$htmlwhite))
				$rowdata += @(,('Priority',($htmlsilver -bor $htmlbold),$Policy.Priority,$htmlwhite))

				$msg = ""
				$columnWidths = @("90","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "290"
				WriteHTMLLine 0 0 " "
			}
				

			Write-Verbose "$(Get-Date): `t`tRetrieving all filters"
			$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName `
			-Type $xPolicyType `
			-DriveName $xDriveName -EA 0 `
			| Sort FilterType, FilterName -Unique

			If($? -and $Null -ne $Filters)
			{
				If(![String]::IsNullOrEmpty($filters))
				{
					Write-Verbose "$(Get-Date): `t`tProcessing all filters"
					$txt = "Assigned to"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 $txt
					}
					
					If($MSWord -or $PDF)
					{
						[System.Collections.Hashtable[]] $FiltersWordTable = @();
					}
					ElseIf($HTML)
					{
						$rowdata = @()
					}
					
					ForEach($Filter in $Filters)
					{
						$tmp = ""
						Switch($filter.FilterType)
						{
							"AccessControl"  {$tmp = "Access Control"; Break}
							"BranchRepeater" {$tmp = "Citrix CloudBridge"; Break}
							"ClientIP"       {$tmp = "Client IP Address"; Break}
							"ClientName"     {$tmp = "Client Name"; Break}
							"DesktopGroup"   {$tmp = "Delivery Group"; Break}
							"DesktopKind"    {$tmp = "Delivery GroupType"; Break}
							"DesktopTag"     {$tmp = "Tag"; Break}
							"OU"             {$tmp = "Organizational Unit (OU)"; Break}
							"User"           {$tmp = "User or group"; Break}
							Default {$tmp = "Policy Filter Type could not be determined: $($filter.FilterType)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Name = $filter.FilterName;
							Type= $tmp;
							Enabled = $filter.Enabled;
							Mode = $filter.Mode;
							Value = $filter.FilterValue;
							}
							$FiltersWordTable += $WordTableRowHash;
						}
						ElseIf($Text)
						{
							Line 2 "Name`t: " $filter.FilterName
							Line 2 "Type`t: " $tmp
							Line 2 "Enabled`t: " $filter.Enabled
							Line 2 "Mode`t: " $filter.Mode
							Line 2 "Value`t: " $filter.FilterValue
							Line 2 ""
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$filter.FilterName,$htmlwhite,
							$tmp,$htmlwhite,
							$filter.Enabled,$htmlwhite,
							$filter.Mode,$htmlwhite,
							$filter.FilterValue,$htmlwhite))
						}
					}
					$tmp = $Null
					If($MSWord -or $PDF)
					{
						$Table = AddWordTable -Hashtable $FiltersWordTable `
						-Columns  Name,Type,Enabled,Mode,Value `
						-Headers  "Name","Type","Enabled","Mode","Value" `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 115;
						$Table.Columns.Item(2).Width = 125;
						$Table.Columns.Item(3).Width = 50;
						$Table.Columns.Item(4).Width = 40;
						$Table.Columns.Item(5).Width = 170;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
					}
					ElseIf($HTML)
					{
						$columnHeaders = @(
						'Name',($htmlsilver -bor $htmlbold),
						'Type',($htmlsilver -bor $htmlbold),
						'Enabled',($htmlsilver -bor $htmlbold),
						'Mode',($htmlsilver -bor $htmlbold),
						'Value',($htmlsilver -bor $htmlbold))

						$msg = ""
						$columnWidths = @("115","125","50","40","170")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
						WriteHTMLLine 0 0 " "
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 1 "Assigned to: None"
					}
					ElseIf($Text)
					{
						Line 1 "Assigned to`t`t: None"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 1 "Assigned to: None"
					}
				}
			}
			Else
			{
				If($Policy.PolicyName -eq "Unfiltered")
				{
					$txt = "Unfiltered policy has no filter settings"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "Assigned to"
						WriteWordLine 0 1 $txt
					}
					ElseIf($Text)
					{
						Line 0 "Assigned to"
						Line 1 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 "Assigned to"
						WriteHTMLLine 0 1 $txt
					}
				}
				Else
				{
					$txt = "Unable to retrieve Filter settings"
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 1 $txt
					}
					ElseIf($Text)
					{
						Line 1 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 1 $txt
					}
				}
			}
			
			Write-Verbose "$(Get-Date): `t`tRetrieving all policy settings"
			$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName `
			-Type $Policy.Type `
			-DriveName $xDriveName -EA 0
				
			If($? -and $Null -ne $Settings)
			{
				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $SettingsWordTable = @();
				}
				ElseIf($HTML)
				{
					$rowdata = @()
				}
				
				$First = $True
				ForEach($Setting in $Settings)
				{
					If($First)
					{
						$txt = "Policy settings"
						If($MSWord -or $PDF)
						{
							WriteWordLine 3 0 $txt
						}
						ElseIf($Text)
						{
							Line 1 $txt
						}
						ElseIf($HTML)
						{
							WriteHTMLLine 3 0 $txt
						}
					}
					$First = $False
					
					Write-Verbose "$(Get-Date): `t`tPolicy settings"
					Write-Verbose "$(Get-Date): `t`t`tConnector for Configuration Manager 2012"
					If((validStateProp $Setting AdvanceWarningFrequency State ) -and ($Setting.AdvanceWarningFrequency.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning frequency interval"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AdvanceWarningFrequency.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningFrequency.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningFrequency.Value
						}
					}
					If((validStateProp $Setting AdvanceWarningMessageBody State ) -and ($Setting.AdvanceWarningMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning message box body text"
						$tmpArray = $Setting.AdvanceWarningMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting AdvanceWarningMessageTitle State ) -and ($Setting.AdvanceWarningMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning message box title"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AdvanceWarningMessageTitle.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningMessageTitle.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningMessageTitle.Value
						}
					}
					If((validStateProp $Setting AdvanceWarningPeriod State ) -and ($Setting.AdvanceWarningPeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning time period"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AdvanceWarningPeriod.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningPeriod.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningPeriod.Value 
						}
					}
					If((validStateProp $Setting FinalForceLogoffMessageBody State ) -and ($Setting.FinalForceLogoffMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Final force logoff message box body text"
						$tmpArray = $Setting.FinalForceLogoffMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting FinalForceLogoffMessageTitle State ) -and ($Setting.FinalForceLogoffMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Final force logoff message box title"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FinalForceLogoffMessageTitle.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FinalForceLogoffMessageTitle.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FinalForceLogoffMessageTitle.Value 
						}
					}
					If((validStateProp $Setting ForceLogoffGracePeriod State ) -and ($Setting.ForceLogoffGracePeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff grace period"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ForceLogoffGracePeriod.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ForceLogoffGracePeriod.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ForceLogoffGracePeriod.Value 
						}
					}
					If((validStateProp $Setting ForceLogoffMessageBody State ) -and ($Setting.ForceLogoffMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff message box body text"
						$tmpArray = $Setting.ForceLogoffMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting ForceLogoffMessageTitle State ) -and ($Setting.ForceLogoffMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff message box title"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ForceLogoffMessageTitle.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ForceLogoffMessageTitle.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ForceLogoffMessageTitle.Value 
						}
					}
					If((validStateProp $Setting ImageProviderIntegrationEnabled State ) -and ($Setting.ImageProviderIntegrationEnabled.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Image-managed mode"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ImageProviderIntegrationEnabled.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ImageProviderIntegrationEnabled.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ImageProviderIntegrationEnabled.State 
						}
					}
					If((validStateProp $Setting RebootMessageBody State ) -and ($Setting.RebootMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Reboot message box body text"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.RebootMessageBody.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootMessageBody.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootMessageBody.Value 
						}
					}
					If((validStateProp $Setting AgentTaskInterval State ) -and ($Setting.AgentTaskInterval.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Regular time interval at which the agent task is to run"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AgentTaskInterval.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AgentTaskInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AgentTaskInterval.Value 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tICA"
					If((validStateProp $Setting ClipboardRedirection State ) -and ($Setting.ClipboardRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client clipboard redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClipboardRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardRedirection.State 
						}
					}
					If((validStateProp $Setting ClientClipboardWriteAllowedFormats State ) -and ($Setting.ClientClipboardWriteAllowedFormats.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client clipboard write allowed formats"
						$tmpArray = $Setting.ClientClipboardWriteAllowedFormats.Values
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting ClipboardSelectionUpdateMode State ) -and ($Setting.ClipboardSelectionUpdateMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Clipboard selection update mode"
						$tmp = ""
						Switch ($Setting.ClipboardSelectionUpdateMode.Value)
						{
							"AllUpdatesAllowed"		{$tmp = "Selection changes are updated on both client and host"; Break}
							"AllUpdatesDenied"		{$tmp = "Select changes are not updated on neither client nor host"; Break}
							"UpdateToClientDenied"	{$tmp = "Host selection changes are not updated to client"; Break}
							"UpdateToHostDenied"	{$tmp = "Client selection changes are not updated to host"; Break}
							Default					{$tmp = "Clipboard selection update mode: $($Setting.ClipboardSelectionUpdateMode.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DesktopLaunchForNonAdmins State ) -and ($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop launches"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DesktopLaunchForNonAdmins.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DesktopLaunchForNonAdmins.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DesktopLaunchForNonAdmins.State 
						}
					}
					If((validStateProp $Setting HDXEnlightenedDataTransport State ) -and ($Setting.HDXEnlightenedDataTransport.State -ne "NotConfigured"))
					{
						#7.12
						$txt = "ICA\HDX Enlightened Data Transport (For Evaluation Only)"
						$tmp = ""
						Switch ($Setting.HDXEnlightenedDataTransport.Value)
						{
							"DiagnosticMode"	{$tmp = "Diagnostic mode"; Break}
							"Off"				{$tmp = "Off"; Break}
							"Preferred"			{$tmp = "Preferred"; Break}
							Default {$tmp = "HDX Enlightened Data Transport: $($Setting.HDXEnlightenedDataTransport.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting HDXAdaptiveTransport State ) -and ($Setting.HDXAdaptiveTransport.State -ne "NotConfigured"))
					{
						#7.13
						$txt = "ICA\HDX Adaptive Transport"
						$tmp = ""
						Switch ($Setting.HDXAdaptiveTransport.Value)
						{
							"DiagnosticMode"	{$tmp = "Diagnostic mode"; Break}
							"Off"				{$tmp = "Off"; Break}
							"Preferred"			{$tmp = "Preferred"; Break}
							Default {$tmp = "HDX Adaptive Transport: $($Setting.HDXAdaptiveTransport.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting IcaListenerTimeout State ) -and ($Setting.IcaListenerTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\ICA listener connection timeout"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IcaListenerTimeout.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaListenerTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaListenerTimeout.Value 
						}
					}
					If((validStateProp $Setting IcaListenerPortNumber State ) -and ($Setting.IcaListenerPortNumber.State -ne "NotConfigured"))
					{
						$txt = "ICA\ICA listener port number"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IcaListenerPortNumber.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaListenerPortNumber.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaListenerPortNumber.Value 
						}
					}
					If((validStateProp $Setting NonPublishedProgramLaunching State ) -and ($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured"))
					{
						$txt = "ICA\Launching of non-published programs during client connection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.NonPublishedProgramLaunching.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.NonPublishedProgramLaunching.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.NonPublishedProgramLaunching.Value 
						}
					}
					If((validStateProp $Setting PrimarySelectionUpdateMode State ) -and ($Setting.PrimarySelectionUpdateMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Primary selection update mode"
						$tmp = ""
						Switch ($Setting.PrimarySelectionUpdateMode.Value)
						{
							"AllUpdatesAllowed"		{$tmp = "Selection changes are updated on both client and host"; Break}
							"AllUpdatesDenied"		{$tmp = "Select changes are not updated on neither client nor host"; Break}
							"UpdateToClientDenied"	{$tmp = "Host selection changes are not updated to client"; Break}
							"UpdateToHostDenied"	{$tmp = "Client selection changes are not updated to host"; Break}
							Default					{$tmp = "Clipboard selection update mode: $($Setting.PrimarySelectionUpdateMode.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ReadonlyClipboard State ) -and ($Setting.ReadonlyClipboard.State -ne "NotConfigured"))
					{
						$txt = "ICA\Readonly Clipboard"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ReadonlyClipboard.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ReadonlyClipboard.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ReadonlyClipboard.State 
						}
					}
					If((validStateProp $Setting RestrictClientClipboardWrite State ) -and ($Setting.RestrictClientClipboardWrite.State -ne "NotConfigured"))
					{
						$txt = "ICA\Restrict client clipboard write"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.RestrictClientClipboardWrite.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RestrictClientClipboardWrite.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RestrictClientClipboardWrite.State
						}
					}
					If((validStateProp $Setting RestrictSessionClipboardWrite State ) -and ($Setting.RestrictSessionClipboardWrite.State -ne "NotConfigured"))
					{
						$txt = "ICA\Restrict session clipboard write"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.RestrictSessionClipboardWrite.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RestrictSessionClipboardWrite.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RestrictSessionClipboardWrite.State 
						}
					}
					If((validStateProp $Setting SessionClipboardWriteAllowedFormats State ) -and ($Setting.SessionClipboardWriteAllowedFormats.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session clipboard write allowed formats"
						$tmpArray = $Setting.SessionClipboardWriteAllowedFormats.Values
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					
					Write-Verbose "$(Get-Date): `t`t`tICA\Adobe Flash Delivery\Flash Redirection"
					If((validStateProp $Setting FlashUrlColorList State ) -and ($Setting.FlashUrlColorList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash background color list"
						$Values = $Setting.FlashUrlColorList.Values
						$tmp = ""
						$cnt = 0
						ForEach($Value in $Values)
						{
							If($Null -eq $Value)
							{
								$Value = ''
							}
							$cnt++
							$tmp = "$($Value)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp 
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmp = $Null
						$Values = $Null
					}
					If((validStateProp $Setting FlashBackwardsCompatibility State ) -and ($Setting.FlashBackwardsCompatibility.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash backwards compatibility"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FlashBackwardsCompatibility.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashBackwardsCompatibility.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashBackwardsCompatibility.State 
						}
					}
					If((validStateProp $Setting FlashDefaultBehavior State ) -and ($Setting.FlashDefaultBehavior.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash default behavior"
						$tmp = ""
						Switch ($Setting.FlashDefaultBehavior.Value)
						{
							"Block"		{$tmp = "Block Flash player"; Break}
							"Disable"	{$tmp = "Disable Flash acceleration"; Break}
							"Enable"	{$tmp = "Enable Flash acceleration"; Break}
							Default		{$tmp = "Flash Default behavior could not be determined: $($Setting.FlashDefaultBehavior.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting FlashIntelligentFallback State ) -and ($Setting.FlashIntelligentFallback.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash intelligent fallback"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FlashIntelligentFallback.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashIntelligentFallback.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashIntelligentFallback.State 
						}
					}
					If((validStateProp $Setting FlashLatencyThreshold State ) -and ($Setting.FlashLatencyThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash latency threshold (milliseconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FlashLatencyThreshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashLatencyThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashLatencyThreshold.Value 
						}
					}
					If((validStateProp $Setting FlashServerSideContentFetchingWhitelist State ) -and ($Setting.FlashServerSideContentFetchingWhitelist.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash server-side content fetching URL list"
						$Values = $Setting.FlashServerSideContentFetchingWhitelist.Values
						$tmp = ""
						$cnt = 0
						ForEach($Value in $Values)
						{
							If($Null -eq $Value)
							{
								$Value = ''
							}
							$cnt++
							$tmp = "$($Value)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp 
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmp = $Null
						$Values = $Null
					}
					If((validStateProp $Setting FlashUrlCompatibilityList State ) -and ($Setting.FlashUrlCompatibilityList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = "";
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt
						}
						$Values = $Setting.FlashUrlCompatibilityList.Values
						$tmp = ""
						ForEach($Value in $Values)
						{
							$Items = $Value.Split(' ')
							$Action = $Items[0]
							If($Action -eq "CLIENT")
							{
								$Action = "Render On Client"
							}
							ElseIf($Action -eq "SERVER")
							{
								$Action = "Render On Server"
							}
							ElseIf($Action -eq "BLOCK")
							{
								$Action = "BLOCK           "
							}
							$Url = $Items[1]
							If($Items.Count -eq 3)
							{
								$FlashInstance = $Items[2]
							}
							Else
							{
								$FlashInstance = "Any"
							}
							$tmp = "Action: $($Action)"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = "";
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "" $tmp
							}
							$tmp = "URL Pattern: $($Url)"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = "";
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "" $tmp
							}
							$tmp = "Flash Instance: $($FlashInstance)"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = "";
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "" $tmp
							}
						}
						$Values = $Null
						$Action = $Null
						$Url = $Null
						$FlashInstance = $Null
						$Spc = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting HDXFlashLoadManagement State ) -and ($Setting.HDXFlashLoadManagement.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash video fallback prevention"
						$tmp = ""
						Switch ($Setting.HDXFlashLoadManagement.Value)
						{
							"Small"						{$tmp = "Only small content"; Break}
							"SmallContentWRedirection"	{$tmp = "Only small content with a supported client"; Break}
							"NoServerSide"				{$tmp = "No server side content"; Break}
							Default {$tmp = "Flash video fallback prevention could not be determined: $($Setting.HDXFlashLoadManagement.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting HDXFlashLoadManagementErrorSwf State ) -and ($Setting.HDXFlashLoadManagementErrorSwf.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash video fallback prevention error *.swf"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.HDXFlashLoadManagementErrorSwf.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXFlashLoadManagementErrorSwf.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HDXFlashLoadManagementErrorSwf.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Audio"
					If((validStateProp $Setting AllowRtpAudio State ) -and ($Setting.AllowRtpAudio.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio over UDP real-time transport"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowRtpAudio.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowRtpAudio.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowRtpAudio.State 
						}
					}
					If((validStateProp $Setting AudioPlugNPlay State ) -and ($Setting.AudioPlugNPlay.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio Plug N Play"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AudioPlugNPlay.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioPlugNPlay.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AudioPlugNPlay.State 
						}
					}
					If((validStateProp $Setting AudioQuality State ) -and ($Setting.AudioQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio quality"
						$tmp = ""
						Switch ($Setting.AudioQuality.Value)
						{
							"Low"		{$tmp = "Low - for low-speed connections"; Break}
							"Medium"	{$tmp = "Medium - optimized for speech"; Break}
							"High"		{$tmp = "High - high definition audio"; Break}
							Default		{$tmp = "Audio quality could not be determined: $($Setting.AudioQuality.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ClientAudioRedirection State ) -and ($Setting.ClientAudioRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Client audio redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientAudioRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientAudioRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientAudioRedirection.State 
						}
					}
					If((validStateProp $Setting MicrophoneRedirection State ) -and ($Setting.MicrophoneRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Client microphone redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MicrophoneRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MicrophoneRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MicrophoneRedirection.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Auto Client Reconnect"
					If((validStateProp $Setting AutoClientReconnect State ) -and ($Setting.AutoClientReconnect.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AutoClientReconnect.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoClientReconnect.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoClientReconnect.State 
						}
					}
					If((validStateProp $Setting AutoClientReconnectAuthenticationRequired  State ) -and ($Setting.AutoClientReconnectAuthenticationRequired.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect authentication"
						$tmp = ""
						Switch ($Setting.AutoClientReconnectAuthenticationRequired.Value)
						{
							"DoNotRequireAuthentication" {$tmp = "Do not require authentication"; Break}
							"RequireAuthentication"      {$tmp = "Require authentication"; Break}
							Default {$tmp = "Auto client reconnect authentication could not be determined: $($Setting.AutoClientReconnectAuthenticationRequired.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting AutoClientReconnectLogging State ) -and ($Setting.AutoClientReconnectLogging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect logging"
						$tmp = ""
						Switch ($Setting.AutoClientReconnectLogging.Value)
						{
							"DoNotLogAutoReconnectEvents" {$tmp = "Do Not Log auto-reconnect events"; Break}
							"LogAutoReconnectEvents"      {$tmp = "Log auto-reconnect events"; Break}
							Default {$tmp = "Auto client reconnect logging could not be determined: $($Setting.AutoClientReconnectLogging.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ACRTimeout State ) -and ($Setting.ACRTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ACRTimeout.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ACRTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ACRTimeout.Value 
						}
					}
					If((validStateProp $Setting ReconnectionUiTransparencyLevel State ) -and ($Setting.ReconnectionUiTransparencyLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Reconnection UI transparency level (%)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ReconnectionUiTransparencyLevel.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ReconnectionUiTransparencyLevel.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ReconnectionUiTransparencyLevel.Value 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tICA\Bandwidth"
					If((validStateProp $Setting AudioBandwidthLimit State ) -and ($Setting.AudioBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Audio redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AudioBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AudioBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting AudioBandwidthPercent State ) -and ($Setting.AudioBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Audio redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AudioBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AudioBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting USBBandwidthLimit State ) -and ($Setting.USBBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Client USB device redirection bandwidth limit"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.USBBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.USBBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.USBBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting USBBandwidthPercent State ) -and ($Setting.USBBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Client USB device redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.USBBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.USBBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.USBBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting ClipboardBandwidthLimit State ) -and ($Setting.ClipboardBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Clipboard redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClipboardBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting ClipboardBandwidthPercent State ) -and ($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Clipboard redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClipboardBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting ComPortBandwidthLimit State ) -and ($Setting.ComPortBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\COM port redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ComPortBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComPortBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ComPortBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting ComPortBandwidthPercent State ) -and ($Setting.ComPortBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\COM port redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ComPortBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComPortBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ComPortBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting FileRedirectionBandwidthLimit State ) -and ($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\File redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FileRedirectionBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FileRedirectionBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FileRedirectionBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting FileRedirectionBandwidthPercent State ) -and ($Setting.FileRedirectionBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\File redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FileRedirectionBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FileRedirectionBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FileRedirectionBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting HDXMultimediaBandwidthLimit State ) -and ($Setting.HDXMultimediaBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.HDXMultimediaBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXMultimediaBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HDXMultimediaBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting HDXMultimediaBandwidthPercent State ) -and ($Setting.HDXMultimediaBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.HDXMultimediaBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXMultimediaBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HDXMultimediaBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting LptBandwidthLimit State ) -and ($Setting.LptBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\LPT port redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LptBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LptBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LptBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting LptBandwidthLimitPercent State ) -and ($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\LPT port redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LptBandwidthLimitPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LptBandwidthLimitPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LptBandwidthLimitPercent.Value 
						}
					}
					If((validStateProp $Setting OverallBandwidthLimit State ) -and ($Setting.OverallBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Overall session bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.OverallBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OverallBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OverallBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting PrinterBandwidthLimit State ) -and ($Setting.PrinterBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Printer redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PrinterBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrinterBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PrinterBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting PrinterBandwidthPercent State ) -and ($Setting.PrinterBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Printer redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PrinterBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrinterBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PrinterBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting TwainBandwidthLimit State ) -and ($Setting.TwainBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\TWAIN device redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TwainBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TwainBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting TwainBandwidthPercent State ) -and ($Setting.TwainBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\TWAIN device redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TwainBandwidthPercent.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TwainBandwidthPercent.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Client Sensors\Location"
					If((validStateProp $Setting AllowLocationServices State ) -and ($Setting.AllowLocationServices.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client Sensors\Location\Allow applications to use the physical location of the client device"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowLocationServices.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowLocationServices.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowLocationServices.State 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tICA\Desktop UI"
					If((validStateProp $Setting GraphicsQuality State ) -and ($Setting.GraphicsQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop Composition graphics quality"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.GraphicsQuality.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.GraphicsQuality.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.GraphicsQuality.Value 
						}
					}
					If((validStateProp $Setting AeroRedirection State ) -and ($Setting.AeroRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop Composition Redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AeroRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AeroRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AeroRedirection.State 
						}
					}
					If((validStateProp $Setting DesktopWallpaper State ) -and ($Setting.DesktopWallpaper.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop wallpaper"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DesktopWallpaper.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DesktopWallpaper.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DesktopWallpaper.State 
						}
					}
					If((validStateProp $Setting MenuAnimation State ) -and ($Setting.MenuAnimation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Menu animation"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MenuAnimation.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MenuAnimation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MenuAnimation.State 
						}
					}
					If((validStateProp $Setting WindowContentsVisibleWhileDragging State ) -and ($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\View window contents while dragging"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.WindowContentsVisibleWhileDragging.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WindowContentsVisibleWhileDragging.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.WindowContentsVisibleWhileDragging.State 
						}
					}
			
					Write-Verbose "$(Get-Date): `t`t`tICA\End User Monitoring"
					If((validStateProp $Setting IcaRoundTripCalculation State ) -and ($Setting.IcaRoundTripCalculation.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculation"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculation.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculation.State 
						}
					}
					If((validStateProp $Setting IcaRoundTripCalculationInterval State ) -and ($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculation interval"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculationInterval.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculationInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculationInterval.Value 
						}	
					}
					If((validStateProp $Setting IcaRoundTripCalculationWhenIdle State ) -and ($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculations for idle connections"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculationWhenIdle.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculationWhenIdle.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculationWhenIdle.State 
						}	
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Enhanced Desktop Experience"
					If((validStateProp $Setting EnhancedDesktopExperience State ) -and ($Setting.EnhancedDesktopExperience.State -ne "NotConfigured"))
					{
						$txt = "ICA\Enhanced Desktop Experience\Enhanced Desktop Experience"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnhancedDesktopExperience.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnhancedDesktopExperience.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnhancedDesktopExperience.State 
						}	
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\File Redirection"
					If((validStateProp $Setting AllowFileTransfer State ) -and ($Setting.AllowFileTransfer.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Allow file transfer between desktop and client"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowFileTransfer.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowFileTransfer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowFileTransfer.State 
						}
					}
					If((validStateProp $Setting AutoConnectDrives State ) -and ($Setting.AutoConnectDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Auto connect client drives"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AutoConnectDrives.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoConnectDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoConnectDrives.State 
						}
					}
					If((validStateProp $Setting ClientDriveRedirection State ) -and ($Setting.ClientDriveRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client drive redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientDriveRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientDriveRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientDriveRedirection.State 
						}
					}
					If((validStateProp $Setting ClientFixedDrives State ) -and ($Setting.ClientFixedDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client fixed drives"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientFixedDrives.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientFixedDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientFixedDrives.State 
						}
					}
					If((validStateProp $Setting ClientFloppyDrives State ) -and ($Setting.ClientFloppyDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client floppy drives"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientFloppyDrives.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientFloppyDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientFloppyDrives.State 
						}
					}
					If((validStateProp $Setting ClientNetworkDrives State ) -and ($Setting.ClientNetworkDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client network drives"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientNetworkDrives.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientNetworkDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientNetworkDrives.State 
						}
					}
					If((validStateProp $Setting ClientOpticalDrives State ) -and ($Setting.ClientOpticalDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client optical drives"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientOpticalDrives.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientOpticalDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientOpticalDrives.State 
						}
					}
					If((validStateProp $Setting ClientRemoveableDrives State ) -and ($Setting.ClientRemoveableDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client removable drives"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientRemoveableDrives.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientRemoveableDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientRemoveableDrives.State 
						}
					}
					If((validStateProp $Setting AllowFileDownload State ) -and ($Setting.AllowFileDownload.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Download file from desktop"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowFileDownload.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowFileDownload.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowFileDownload.State 
						}
					}
					If((validStateProp $Setting HostToClientRedirection State ) -and ($Setting.HostToClientRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Host to client redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.HostToClientRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HostToClientRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HostToClientRedirection.State 
						}
					}
					If((validStateProp $Setting ClientDriveLetterPreservation State ) -and ($Setting.ClientDriveLetterPreservation.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Preserve client drive letters"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientDriveLetterPreservation.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientDriveLetterPreservation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientDriveLetterPreservation.State 
						}
					}
					If((validStateProp $Setting ReadOnlyMappedDrive State ) -and ($Setting.ReadOnlyMappedDrive.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Read-only client drive access"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ReadOnlyMappedDrive.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ReadOnlyMappedDrive.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ReadOnlyMappedDrive.State 
						}
					}
					If((validStateProp $Setting SpecialFolderRedirection State ) -and ($Setting.SpecialFolderRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Special folder redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SpecialFolderRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SpecialFolderRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SpecialFolderRedirection.State 
						}
					}
					If((validStateProp $Setting AllowFileUpload State ) -and ($Setting.AllowFileUpload.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Upload file to desktop"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowFileUpload.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowFileUpload.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowFileUpload.State 
						}
					}
					If((validStateProp $Setting AsynchronousWrites State ) -and ($Setting.AsynchronousWrites.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Use asynchronous writes"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AsynchronousWrites.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AsynchronousWrites.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AsynchronousWrites.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Graphics"
					If((validStateProp $Setting AllowVisuallyLosslessCompression State ) -and ($Setting.AllowVisuallyLosslessCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Allow visually lossless compression"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowVisuallyLosslessCompression.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowVisuallyLosslessCompression.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowVisuallyLosslessCompression.State 
						}
					}
					If((validStateProp $Setting DisplayMemoryLimit State ) -and ($Setting.DisplayMemoryLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Display memory limit (KB)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DisplayMemoryLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisplayMemoryLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DisplayMemoryLimit.Value 
						}	
					}
					If((validStateProp $Setting DisplayDegradePreference State ) -and ($Setting.DisplayDegradePreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Display mode degrade preference"
						$tmp = ""
						Switch ($Setting.DisplayDegradePreference.Value)
						{
							"ColorDepth"	{$tmp = "Degrade color depth first"; Break}
							"Resolution"	{$tmp = "Degrade resolution first"; Break}
							Default			{$tmp = "Display mode degrade preference could not be determined: $($Setting.DisplayDegradePreference.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}
					If((validStateProp $Setting DynamicPreview State ) -and ($Setting.DynamicPreview.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Dynamic windows preview"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DynamicPreview.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DynamicPreview.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DynamicPreview.State 
						}	
					}
					If((validStateProp $Setting ImageCaching State ) -and ($Setting.ImageCaching.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Image caching"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ImageCaching.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ImageCaching.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ImageCaching.State 
						}	
					}
					If((validStateProp $Setting LegacyGraphicsMode State ) -and ($Setting.LegacyGraphicsMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Legacy graphics mode"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LegacyGraphicsMode.State;
						}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LegacyGraphicsMode.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LegacyGraphicsMode.State 
						}
					}
					If((validStateProp $Setting MaximumColorDepth State ) -and ($Setting.MaximumColorDepth.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Maximum allowed color depth"
						$tmp = ""
						Switch ($Setting.MaximumColorDepth.Value)
						{
							"BitsPerPixel24"	{$tmp = "24 Bits Per Pixel"; Break}
							"BitsPerPixel32"	{$tmp = "32 Bits Per Pixel"; Break}
							"BitsPerPixel16"	{$tmp = "16 Bits Per Pixel"; Break}
							"BitsPerPixel15"	{$tmp = "15 Bits Per Pixel"; Break}
							"BitsPerPixel8"		{$tmp = "8 Bits Per Pixel"; Break}
							Default				{$tmp = "Maximum allowed color depth could not be determined: $($Setting.MaximumColorDepth.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}
					If((validStateProp $Setting DisplayDegradeUserNotification State ) -and ($Setting.DisplayDegradeUserNotification.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Notify user when display mode is degraded"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DisplayDegradeUserNotification.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisplayDegradeUserNotification.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DisplayDegradeUserNotification.State 
						}	
					}
					If((validStateProp $Setting QueueingAndTossing State ) -and ($Setting.QueueingAndTossing.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Queueing and tossing"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.QueueingAndTossing.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.QueueingAndTossing.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.QueueingAndTossing.State 
						}	
					}
					If((validStateProp $Setting UseHardwareEncodingForVideoCodec State ) -and ($Setting.UseHardwareEncodingForVideoCodec.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Use hardware encoding for video codec"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UseHardwareEncodingForVideoCodec.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UseHardwareEncodingForVideoCodec.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UseHardwareEncodingForVideoCodec.State 
						}	
					}
					If((validStateProp $Setting UseVideoCodecForCompression State ) -and ($Setting.UseVideoCodecForCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Use video codec for compression"
						$tmp = ""
						Switch ($Setting.UseVideoCodecForCompression.Value)
						{
							"UseVideoCodecIfAvailable"	{$tmp = "For the entire screen"; Break}
							"DoNotUseVideoCodec"		{$tmp = "Do not use video codec"; Break}
							"UseVideoCodecIfPreferred"	{$tmp = "Use when preferred"; Break}
							"ActivelyChangingRegions"	{$tmp = "For actively changing regions"; Break}
							Default {$tmp = "Use video codec for compression could not be determined: $($Setting.UseVideoCodecForCompression.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Graphics\Caching"
					If((validStateProp $Setting PersistentCache State ) -and ($Setting.PersistentCache.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Caching\Persistent cache threshold (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PersistentCache.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PersistentCache.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PersistentCache.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Graphics\Framehawk"
					If((validStateProp $Setting EnableFramehawkDisplayChannel State ) -and ($Setting.EnableFramehawkDisplayChannel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Framehawk\Framehawk display channel"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableFramehawkDisplayChannel.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableFramehawkDisplayChannel.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnableFramehawkDisplayChannel.State 
						}
					}
					If((validStateProp $Setting FramehawkDisplayChannelPortRange State ) -and ($Setting.FramehawkDisplayChannelPortRange.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Framehawk\Framehawk display channel port range"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FramehawkDisplayChannelPortRange.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FramehawkDisplayChannelPortRange.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FramehawkDisplayChannelPortRange.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Keep Alive"
					If((validStateProp $Setting IcaKeepAliveTimeout State ) -and ($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keep Alive\ICA keep alive timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IcaKeepAliveTimeout.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaKeepAliveTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaKeepAliveTimeout.Value 
						}
					}
					If((validStateProp $Setting IcaKeepAlives State ) -and ($Setting.IcaKeepAlives.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keep Alive\ICA keep alives"
						$tmp = ""
						Switch ($Setting.IcaKeepAlives.Value)
						{
							"DoNotSendKeepAlives" {$tmp = "Do not send ICA keep alive messages"; Break}
							"SendKeepAlives"      {$tmp = "Send ICA keep alive messages"; Break}
							Default {$tmp = "ICA keep alives could not be determined: $($Setting.IcaKeepAlives.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Local App Access"
					If((validStateProp $Setting AllowLocalAppAccess State ) -and ($Setting.AllowLocalAppAccess.State -ne "NotConfigured"))
					{
						$txt = "ICA\Local App Access\Allow local app access"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowLocalAppAccess.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowLocalAppAccess.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowLocalAppAccess.State 
						}
					}
					If((validStateProp $Setting URLRedirectionBlackList State ) -and ($Setting.URLRedirectionBlackList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Local App Access\URL redirection blacklist"
						$tmpArray = $Setting.URLRedirectionBlackList.Values
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting URLRedirectionWhiteList State ) -and ($Setting.URLRedirectionWhiteList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Local App Access\URL redirection white list"
						$tmpArray = $Setting.URLRedirectionWhiteList.Values
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$TmpArray = $Null
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Mobile Experience"
					If((validStateProp $Setting AutoKeyboardPopUp State ) -and ($Setting.AutoKeyboardPopUp.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Automatic keyboard display"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AutoKeyboardPopUp.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoKeyboardPopUp.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoKeyboardPopUp.State 
						}
					}
					If((validStateProp $Setting MobileDesktop State ) -and ($Setting.MobileDesktop.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Launch touch-optimized desktop"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MobileDesktop.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MobileDesktop.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MobileDesktop.State 
						}
					}
					If((validStateProp $Setting ComboboxRemoting State ) -and ($Setting.ComboboxRemoting.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Remote the combo box"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ComboboxRemoting.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComboboxRemoting.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ComboboxRemoting.State 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tICA\Multimedia"
					If((validStateProp $Setting HTML5VideoRedirection State ) -and ($Setting.HTML5VideoRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\HTML5 video redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.HTML5VideoRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HTML5VideoRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HTML5VideoRedirection.State
						}
					}
					If((validStateProp $Setting VideoQuality State ) -and ($Setting.VideoQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Limit video quality"
						$tmp = ""
						Switch ($Setting.VideoQuality.Value)
						{
							"P1080"			{$tmp = "Maximum Video Quality 1080p/8.5mbps"; Break}
							"P720"			{$tmp = "Maximum Video Quality 720p/4.0mbps"; Break}
							"P480"			{$tmp = "Maximum Video Quality 480p/720kbps"; Break}
							"P380"			{$tmp = "Maximum Video Quality 380p/400kbps"; Break}
							"P240"			{$tmp = "Maximum Video Quality 240p/200kbps"; Break}
							"Unconfigured"	{$tmp = "Not Configured"; Break}
							Default			{$tmp = "Limit video quality could not be determined: $($Setting.VideoQuality.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting MaxSpeexQuality State ) -and ($Setting.MaxSpeexQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Max Speex quality"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MaxSpeexQuality.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MaxSpeexQuality.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MaxSpeexQuality.Value 
						}
					}
					If((validStateProp $Setting MultimediaConferencing State ) -and ($Setting.MultimediaConferencing.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Multimedia conferencing"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultimediaConferencing.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaConferencing.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaConferencing.State 
						}
					}
					If((validStateProp $Setting MultimediaOptimization State ) -and ($Setting.MultimediaOptimization.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Optimization for Windows Media multimedia redirection over WAN"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultimediaOptimization.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaOptimization.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaOptimization.State 
						}
					}
					If((validStateProp $Setting UseGPUForMultimediaOptimization State ) -and ($Setting.UseGPUForMultimediaOptimization.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Use GPU for optimizing Windows Media multimedia redirection over WAN"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UseGPUForMultimediaOptimization.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UseGPUForMultimediaOptimization.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UseGPUForMultimediaOptimization.State 
						}
					}
					If((validStateProp $Setting MultimediaAccelerationEnableCSF State ) -and ($Setting.MultimediaAccelerationEnableCSF.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media client-side content fetching"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultimediaAccelerationEnableCSF.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAccelerationEnableCSF.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAccelerationEnableCSF.State 
						}
					}
					If((validStateProp $Setting VideoLoadManagement State ) -and ($Setting.VideoLoadManagement.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows media fallback prevention"
						$tmp = ""
						Switch ($Setting.VideoLoadManagement.Value)
						{
							"SFSR"	{$tmp = "Play all content"; Break}
							"SFCR"	{$tmp = "Play all content only on client"; Break}
							"CFCR"	{$tmp = "Play only client-accessible content on client"; Break}
							Default	{$tmp = "Windows media fallback prevention could not be determined: $($Setting.VideoLoadManagement.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting MultimediaAcceleration State ) -and ($Setting.MultimediaAcceleration.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultimediaAcceleration.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAcceleration.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAcceleration.State 
						}
					}
					If((validStateProp $Setting MultimediaAccelerationDefaultBufferSize State ) -and ($Setting.MultimediaAccelerationDefaultBufferSize.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection buffer size (seconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultimediaAccelerationDefaultBufferSize.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAccelerationDefaultBufferSize.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAccelerationDefaultBufferSize.Value 
						}
					}
					If((validStateProp $Setting MultimediaAccelerationUseDefaultBufferSize State ) -and ($Setting.MultimediaAccelerationUseDefaultBufferSize.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection buffer size use"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultimediaAccelerationUseDefaultBufferSize.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAccelerationUseDefaultBufferSize.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAccelerationUseDefaultBufferSize.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Multi-Stream Connections"
					If((validStateProp $Setting UDPAudioOnServer State ) -and ($Setting.UDPAudioOnServer.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Audio over UDP"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UDPAudioOnServer.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UDPAudioOnServer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UDPAudioOnServer.State
						}
					}
					If((validStateProp $Setting RtpAudioPortRange State ) -and ($Setting.RtpAudioPortRange.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Audio UDP port range"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.RtpAudioPortRange.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RtpAudioPortRange.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RtpAudioPortRange.Value 
						}
					}
					If((validStateProp $Setting MultiPortPolicy State ) -and ($Setting.MultiPortPolicy.State -ne "NotConfigured"))
					{
						$txt1 = "ICA\MultiStream Connections\Multi-Port Policy\CGP default port"
						$txt2 = "ICA\MultiStream Connections\Multi-Port Policy\CGP default port priority"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt1;
							Value = "Default Port";
							}
							$SettingsWordTable += $WordTableRowHash;

							$WordTableRowHash = @{
							Text = $txt2;
							Value = "High";
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt1,$htmlbold,
							"Default Port",$htmlwhite))

							$rowdata += @(,(
							$txt2,$htmlbold,
							"High",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt1 "Default Port"
							OutputPolicySetting $txt2 "High"
						}
						$txt1 = $Null
						$txt2 = $Null
						[string]$Tmp = $Setting.MultiPortPolicy.Value
						If($Tmp.Length -gt 0)
						{
							$Port1Priority = ""
							$Port2Priority = ""
							$Port3Priority = ""
							[string]$cgpport1 = $Tmp.substring(0, $Tmp.indexof(";"))
							[string]$cgpport2 = $Tmp.substring($cgpport1.length + 1 , ($Tmp.indexof(";")+1))
							[string]$cgpport3 = $Tmp.substring((($cgpport1.length + 1)+($cgpport2.length + 1)) , ($Tmp.indexof(";")+1))
							[string]$cgpport1priority = $cgpport1.substring($cgpport1.length -1, 1)
							[string]$cgpport2priority = $cgpport2.substring($cgpport2.length -1, 1)
							[string]$cgpport3priority = $cgpport3.substring($cgpport3.length -1, 1)
							$cgpport1 = $cgpport1.substring(0, $cgpport1.indexof(","))
							$cgpport2 = $cgpport2.substring(0, $cgpport2.indexof(","))
							$cgpport3 = $cgpport3.substring(0, $cgpport3.indexof(","))
							Switch ($cgpport1priority)
							{
								"0"	{$Port1Priority = "Very High"; Break}
								"2"	{$Port1Priority = "Medium"; Break}
								"3"	{$Port1Priority = "Low"; Break}
								Default	{$Port1Priority = "Unknown"; Break}
							}
							Switch ($cgpport2priority)
							{
								"0"	{$Port2Priority = "Very High"; Break}
								"2"	{$Port2Priority = "Medium"; Break}
								"3"	{$Port2Priority = "Low"; Break}
								Default	{$Port2Priority = "Unknown"; Break}
							}
							Switch ($cgpport3priority)
							{
								"0"	{$Port3Priority = "Very High"; Break}
								"2"	{$Port3Priority = "Medium"; Break}
								"3"	{$Port3Priority = "Low"; Break}
								Default	{$Port3Priority = "Unknown"; Break}
							}
							$txt1 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port1"
							$txt2 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port1 priority"
							$txt3 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port2"
							$txt4 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port2 priority"
							$txt5 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port3"
							$txt6 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port3 priority"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt1;
								Value = $cgpport1;
								}
								$SettingsWordTable += $WordTableRowHash;

								$WordTableRowHash = @{
								Text = $txt2;
								Value = $port1priority;
								}
								$SettingsWordTable += $WordTableRowHash;

								$WordTableRowHash = @{
								Text = $txt3;
								Value = $cgpport2;
								}
								$SettingsWordTable += $WordTableRowHash;

								$WordTableRowHash = @{
								Text = $txt4;
								Value = $port2priority;
								}
								$SettingsWordTable += $WordTableRowHash;

								$WordTableRowHash = @{
								Text = $txt5;
								Value = $cgpport3;
								}
								$SettingsWordTable += $WordTableRowHash;

								$WordTableRowHash = @{
								Text = $txt6;
								Value = $port3priority;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt1,$htmlbold,
								$cgpport1,$htmlwhite))
								
								$rowdata += @(,(
								$txt2,$htmlbold,
								$port1priority,$htmlwhite))
								
								$rowdata += @(,(
								$txt3,$htmlbold,
								$cgpport2,$htmlwhite))
								
								$rowdata += @(,(
								$txt4,$htmlbold,
								$port2priority,$htmlwhite))
								
								$rowdata += @(,(
								$txt5,$htmlbold,
								$cgpport3,$htmlwhite))
								
								$rowdata += @(,(
								$txt6,$htmlbold,
								$port3priority,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt1 $cgpport1
								OutputPolicySetting $txt2 $port1priority
								OutputPolicySetting $txt3 $cgpport2
								OutputPolicySetting $txt4 $port2priority
								OutputPolicySetting $txt5 $cgpport3
								OutputPolicySetting $txt6 $port3priority
							}	
						}
						$Tmp = $Null
						$cgpport1 = $Null
						$cgpport2 = $Null
						$cgpport3 = $Null
						$cgpport1priority = $Null
						$cgpport2priority = $Null
						$cgpport3priority = $Null
						$Port1Priority = $Null
						$Port2Priority = $Null
						$Port3Priority = $Null
						$txt1 = $Null
						$txt2 = $Null
						$txt3 = $Null
						$txt4 = $Null
						$txt5 = $Null
						$txt6 = $Null
					}
					If((validStateProp $Setting MultiStreamPolicy State ) -and ($Setting.MultiStreamPolicy.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Multi-Stream computer setting"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultiStreamPolicy.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultiStreamPolicy.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultiStreamPolicy.State 
						}
					}
					If((validStateProp $Setting MultiStream State ) -and ($Setting.MultiStream.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Multi-Stream user setting"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MultiStream.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultiStream.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultiStream.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Port Redirection"
					If((validStateProp $Setting ClientComPortsAutoConnection State ) -and ($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Auto connect client COM ports"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientComPortsAutoConnection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientComPortsAutoConnection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientComPortsAutoConnection.State 
						}
					}
					If((validStateProp $Setting ClientLptPortsAutoConnection State ) -and ($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Auto connect client LPT ports"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientLptPortsAutoConnection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientLptPortsAutoConnection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientLptPortsAutoConnection.State 
						}
					}
					If((validStateProp $Setting ClientComPortRedirection State ) -and ($Setting.ClientComPortRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Client COM port redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientComPortRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientComPortRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientComPortRedirection.State 
						}
					}
					If((validStateProp $Setting ClientLptPortRedirection State ) -and ($Setting.ClientLptPortRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Client LPT port redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientLptPortRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientLptPortRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientLptPortRedirection.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Printing"
					If((validStateProp $Setting ClientPrinterRedirection State ) -and ($Setting.ClientPrinterRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client printer redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ClientPrinterRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientPrinterRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientPrinterRedirection.State 
						}
					}
					If((validStateProp $Setting DefaultClientPrinter State ) -and ($Setting.DefaultClientPrinter.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Default printer"
						$tmp = ""
						Switch ($Setting.DefaultClientPrinter.Value)
						{
							"ClientDefault" {$tmp = "Set Default printer to the client's main printer"; Break}
							"DoNotAdjust"   {$tmp = "Do not adjust the user's Default printer"; Break}
							Default {$tmp = "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting PrinterAssignments State ) -and ($Setting.PrinterAssignments.State -ne "NotConfigured"))
					{
						If($Setting.PrinterAssignments.State -eq "Enabled")
						{
							$txt = "ICA\Printing\Printer assignments"
							$PrinterAssign = Get-ChildItem -path "$($xDriveName):\User\$($Policy.PolicyName)\Settings\ICA\Printing\PrinterAssignments" 4>$Null
							If($? -and $Null -ne $PrinterAssign)
							{
								$PrinterAssignments = $PrinterAssign.Contents
								ForEach($PrinterAssignment in $PrinterAssignments)
								{
									$Client = @()
									$DefaultPrinter = ""
									$SessionPrinters = @()
									$tmp1 = ""
									$tmp2 = ""
									$tmp3 = ""
									ForEach($Filter in $PrinterAssignment.Filters)
									{
										$Client += "$($Filter); "
									}
									If($PrinterAssignment.SpecificDefaultPrinter -eq "")
									{
										$DefaultPrinter = "<Not set>"
									}
									Else
									{
										$DefaultPrinter = $PrinterAssignment.SpecificDefaultPrinter
									}
									ForEach($SessionPrinter in $PrinterAssignment.SessionPrinters)
									{
										$SessionPrinters += $SessionPrinter
									}
									$tmp1 = "Client Names/IP's: $($Client)"
									$tmp2 = "Default Printer: $($DefaultPrinter)"
									$tmp3 = "Session Printers: $($SessionPrinters)"
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp1;
										}
										$SettingsWordTable += $WordTableRowHash;
										
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp2;
										}
										$SettingsWordTable += $WordTableRowHash;
										
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp3;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp1,$htmlwhite))
										
										$rowdata += @(,(
										"",$htmlbold,
										$tmp2,$htmlwhite))
										
										$rowdata += @(,(
										"",$htmlbold,
										$tmp3,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp1
										OutputPolicySetting "" $tmp2
										OutputPolicySetting "" $tmp3
									}
									$tmp = " "
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
									$tmp1 = $Null
									$tmp2 = $Null
									$tmp3 = $Null
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PrinterAssignments.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PrinterAssignments.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.PrinterAssignments.State 
							}
						}
					}
					If((validStateProp $Setting AutoCreationEventLogPreference State ) -and ($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Printer auto-creation event log preference"
						$tmp = ""
						Switch ($Setting.AutoCreationEventLogPreference.Value)
						{
							"LogErrorsOnly"        {$tmp = "Log errors only"; Break}
							"LogErrorsAndWarnings" {$tmp = "Log errors and warnings"; Break}
							"DoNotLog"             {$tmp = "Do not log errors or warnings"; Break}
							Default {$tmp = "Printer auto-creation event log preference could not be determined: $($Setting.AutoCreationEventLogPreference.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting SessionPrinters State ) -and ($Setting.SessionPrinters.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Session printers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = "";
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt ""
						}
						$valArray = $Setting.SessionPrinters.Values
						$tmp = ""
						ForEach($printer in $valArray)
						{
							$prArray = $printer.Split(',')
							ForEach($element in $prArray)
							{
								If($element.SubString(0, 2) -eq "\\")
								{
									$index = $element.SubString(2).IndexOf('\')
									If($index -ge 0)
									{
										$server = $element.SubString(0, $index + 2)
										$share  = $element.SubString($index + 3)
										$tmp = "Server: $($server)"
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										ElseIf($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										ElseIf($Text)
										{
											OutputPolicySetting "" $tmp
										}
										$tmp = "Shared Name: $($share)"
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										ElseIf($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										ElseIf($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
									$index = $Null
								}
								Else
								{
									$tmp1 = $element.SubString(0, 4)
									$tmp = Get-PrinterModifiedSettings $tmp1 $element
									If(![String]::IsNullOrEmpty($tmp))
									{
										If($MSWord -or $PDF)
										{
											$WordTableRowHash = @{
											Text = "";
											Value = $tmp;
											}
											$SettingsWordTable += $WordTableRowHash;
										}
										ElseIf($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										ElseIf($Text)
										{
											OutputPolicySetting "" $tmp
										}
									}
									$tmp1 = $Null
									$tmp = $Null
								}
							}
						}

						$valArray = $Null
						$prArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting WaitForPrintersToBeCreated State ) -and ($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Wait for printers to be created (server desktop)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.WaitForPrintersToBeCreated.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WaitForPrintersToBeCreated.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.WaitForPrintersToBeCreated.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Printing\Client Printers"
					If((validStateProp $Setting ClientPrinterAutoCreation State ) -and ($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create client printers"
						$tmp = ""
						Switch ($Setting.ClientPrinterAutoCreation.Value)
						{
							"DoNotAutoCreate"    {$tmp = "Do not auto-create client printers"; Break}
							"DefaultPrinterOnly" {$tmp = "Auto-create the client's Default printer only"; Break}
							"LocalPrintersOnly"  {$tmp = "Auto-create local (non-network) client printers only"; Break}
							"AllPrinters"        {$tmp = "Auto-create all client printers"; Break}
							Default {$tmp = "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting GenericUniversalPrinterAutoCreation State ) -and ($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create generic universal printer"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.GenericUniversalPrinterAutoCreation.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.GenericUniversalPrinterAutoCreation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.GenericUniversalPrinterAutoCreation.State 
						}
					}
					If((validStateProp $Setting AutoCreatePDFPrinter State ) -and ($Setting.AutoCreatePDFPrinter.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create PDF Universal Printer"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AutoCreatePDFPrinter.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoCreatePDFPrinter.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoCreatePDFPrinter.State 
						}
					}
					If((validStateProp $Setting ClientPrinterNames State ) -and ($Setting.ClientPrinterNames.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Client printer names"
						$tmp = ""
						Switch ($Setting.ClientPrinterNames.Value)
						{
							"StandardPrinterNames" {$tmp = "Standard printer names"; Break}
							"LegacyPrinterNames"   {$tmp = "Legacy printer names"; Break}
							Default {$tmp = "Client printer names could not be determined: $($Setting.ClientPrinterNames.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DirectConnectionsToPrintServers State ) -and ($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Direct connections to print servers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DirectConnectionsToPrintServers.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DirectConnectionsToPrintServers.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DirectConnectionsToPrintServers.State 
						}
					}
					If((validStateProp $Setting PrinterDriverMappings State ) -and ($Setting.PrinterDriverMappings.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Printer driver mapping and compatibility"
						$array = $Setting.PrinterDriverMappings.Values
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$Items = $element.Split(',')
								$DriverName = $Items[0]
								$Action = $Items[1]
								If($Action -match 'Replace=')
								{
									$ServerDriver = $Action.substring($Action.indexof("=")+1)
									$Action = "Replace "
								}
								Else
								{
									$ServerDriver = ""
									If($Action -eq "Allow")
									{
										$Action = "Allow "
									}
									ElseIf($Action -eq "Deny")
									{
										$Action = "Do not create "
									}
									ElseIf($Action -eq "UPD_Only")
									{
										$Action = "Create with universal driver "
									}
								}
								$tmp = "Driver Name: $($DriverName)"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
								$tmp = "Action: $($Action)"
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
								$tmp = "Settings: "
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
								If($Items.count -gt 2)
								{
									[int]$BeginAt = 2
									[int]$EndAt = $Items.count
									for ($i=$BeginAt;$i -lt $EndAt; $i++) 
									{
										$tmp2 = $Items[$i].SubString(0, 4)
										$tmp = Get-PrinterModifiedSettings $tmp2 $Items[$i]
										If(![String]::IsNullOrEmpty($tmp))
										{
											If($MSWord -or $PDF)
											{
												$WordTableRowHash = @{
												Text = "";
												Value = $tmp;
												}
												$SettingsWordTable += $WordTableRowHash;
											}
											ElseIf($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											ElseIf($Text)
											{
												OutputPolicySetting "" $tmp
											}
										}
									}
								}
								Else
								{
									$tmp = "Unmodified "
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}

								If(![String]::IsNullOrEmpty($ServerDriver))
								{
									$tmp = "Server Driver: $($ServerDriver)"
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
								$tmp = $Null
							}
						}
					}
					If((validStateProp $Setting PrinterPropertiesRetention State ) -and ($Setting.PrinterPropertiesRetention.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Printer properties retention"
						$tmp = ""
						Switch ($Setting.PrinterPropertiesRetention.Value)
						{
							"SavedOnClientDevice"   {$tmp = "Saved on the client device only"; Break}
							"RetainedInUserProfile" {$tmp = "Retained in user profile only"; Break}
							"FallbackToProfile"     {$tmp = "Held in profile only if not saved on client"; Break}
							"DoNotRetain"           {$tmp = "Do not retain printer properties"; Break}
							Default {$tmp = "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetention.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting RetainedAndRestoredClientPrinters State ) -and ($Setting.RetainedAndRestoredClientPrinters.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Retained and restored client printers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.RetainedAndRestoredClientPrinters.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RetainedAndRestoredClientPrinters.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RetainedAndRestoredClientPrinters.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Printing\Drivers"
					If((validStateProp $Setting InboxDriverAutoInstallation State ) -and ($Setting.InboxDriverAutoInstallation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Automatic installation of in-box printer drivers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.InboxDriverAutoInstallation.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.InboxDriverAutoInstallation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.InboxDriverAutoInstallation.State 
						}
					}
					If((validStateProp $Setting UniversalDriverPriority State ) -and ($Setting.UniversalDriverPriority.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Universal driver preference"
						$Values = $Setting.UniversalDriverPriority.Value.Split(';')
						$tmp = ""
						$cnt = 0
						ForEach($Value in $Values)
						{
							If($Null -eq $Value)
							{
								$Value = ''
							}
							$cnt++
							$tmp = "$($Value)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp 
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmp = $Null
						$Values = $Null
					}
					If((validStateProp $Setting UniversalPrintDriverUsage State ) -and ($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Universal print driver usage"
						$tmp = ""
						Switch ($Setting.UniversalPrintDriverUsage.Value)
						{
							"SpecificOnly"       {$tmp = "Use only printer model specific drivers"; Break}
							"UpdOnly"            {$tmp = "Use universal printing only"; Break}
							"FallbackToUpd"      {$tmp = "Use universal printing only if requested driver is unavailable"; Break}
							"FallbackToSpecific" {$tmp = "Use printer model specific drivers only if universal printing is unavailable"; Break}
							Default {$tmp = "Universal print driver usage could not be determined: $($Setting.UniversalPrintDriverUsage.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Printing\Universal Print Server"
					If((validStateProp $Setting UpsEnable State ) -and ($Setting.UpsEnable.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server enable"
						If($Setting.UpsEnable.State)
						{
							$tmp = ""
						}
						Else
						{
							$tmp = "Disabled"
						}
						Switch ($Setting.UpsEnable.Value)
						{
							"UpsEnabledWithFallback"	{$tmp = "Enabled with fallback to Windows' native remote printing"; Break}
							"UpsOnlyEnabled"			{$tmp = "Enabled with no fallback to Windows' native remote printing"; Break}
							Default	{$tmp = "Universal Print Server enable value could not be determined: $($Setting.UpsEnable.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting UpsCgpPort State ) -and ($Setting.UpsCgpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server print data stream (CGP) port"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UpsCgpPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsCgpPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UpsCgpPort.Value 
						}
					}
					If((validStateProp $Setting UpsPrintStreamInputBandwidthLimit State ) -and ($Setting.UpsPrintStreamInputBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server print stream input bandwidth limit (kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UpsPrintStreamInputBandwidthLimit.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsPrintStreamInputBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UpsPrintStreamInputBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting UpsHttpPort State ) -and ($Setting.UpsHttpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server web service (HTTP/SOAP) port"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UpsHttpPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsHttpPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UpsHttpPort.Value 
						}
					}
					If((validStateProp $Setting LoadBalancedPrintServers State ) -and ($Setting.LoadBalancedPrintServers.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Servers for load balancing"
						$array = $Setting.LoadBalancedPrintServers.Values
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting PrintServersOutOfServiceThreshold State ) -and ($Setting.PrintServersOutOfServiceThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Servers out-of-service threshold"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PrintServersOutOfServiceThreshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrintServersOutOfServiceThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PrintServersOutOfServiceThreshold.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Printing\Universal Printing"
					If((validStateProp $Setting EMFProcessingMode State ) -and ($Setting.EMFProcessingMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing EMF processing mode"
						$tmp = ""
						Switch ($Setting.EMFProcessingMode.Value)
						{
							"ReprocessEMFsForPrinter" {$tmp = "Reprocess EMFs for printer"; Break}
							"SpoolDirectlyToPrinter"  {$tmp = "Spool directly to printer"; Break}
							Default {$tmp = "Universal printing EMF processing mode could not be determined: $($Setting.EMFProcessingMode.Value)"; Break}
						}
						 
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ImageCompressionLimit State ) -and ($Setting.ImageCompressionLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing image compression limit"
						$tmp = ""
						Switch ($Setting.ImageCompressionLimit.Value)
						{
							"NoCompression"       {$tmp = "No compression"; Break}
							"LosslessCompression" {$tmp = "Best quality (lossless compression)"; Break}
							"MinimumCompression"  {$tmp = "High quality"; Break}
							"MediumCompression"   {$tmp = "Standard quality"; Break}
							"MaximumCompression"  {$tmp = "Reduced quality (maximum compression)"; Break}
							Default {$tmp = "Universal printing image compression limit could not be determined: $($Setting.ImageCompressionLimit.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting UPDCompressionDefaults State ) -and ($Setting.UPDCompressionDefaults.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing optimization defaults"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = "";
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt "" 
						}
						
						$TmpArray = $Setting.UPDCompressionDefaults.Value.Split(',')
						$tmp = ""
						ForEach($Thing in $TmpArray)
						{
							$TestLabel = $Thing.substring(0, $Thing.indexof("="))
							$TestSetting = $Thing.substring($Thing.indexof("=")+1)
							$TxtLabel = ""
							$TxtSetting = ""
							Switch($TestLabel)
							{
								"ImageCompression"
								{
									$TxtLabel = "Desired image quality:"
									Switch($TestSetting)
									{
										"StandardQuality"	{$TxtSetting = "Standard quality"; Break}
										"BestQuality"	{$TxtSetting = "Best quality (lossless compression)"; Break}
										"HighQuality"	{$TxtSetting = "High quality"; Break}
										"ReducedQuality"	{$TxtSetting = "Reduced quality (maximum compression)"; Break}
									}
								}
								"HeavyweightCompression"
								{
									$TxtLabel = "Enable heavyweight compression:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"ImageCaching"
								{
									$TxtLabel = "Allow caching of embedded images:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"FontCaching"
								{
									$TxtLabel = "Allow caching of embedded fonts:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"AllowNonAdminsToModify"
								{
									$TxtLabel = "Allow non-administrators to modify these settings:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
							}
							$tmp = "$($TxtLabel) $TxtSetting "
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = "";
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "" $tmp
							}
						}
						$TmpArray = $Null
						$tmp = $Null
						$TestLabel = $Null
						$TestSetting = $Null
						$TxtLabel = $Null
						$TxtSetting = $Null
					}
					If((validStateProp $Setting UniversalPrintingPreviewPreference State ) -and ($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing preview preference"
						$tmp = ""
						Switch ($Setting.UniversalPrintingPreviewPreference.Value)
						{
							"NoPrintPreview"        {$tmp = "Do not use print preview for auto-created or generic universal printers"; Break}
							"AutoCreatedOnly"       {$tmp = "Use print preview for auto-created printers only"; Break}
							"GenericOnly"           {$tmp = "Use print preview for generic universal printers only"; Break}
							"AutoCreatedAndGeneric" {$tmp = "Use print preview for both auto-created and generic universal printers"; Break}
							Default {$tmp = "Universal printing preview preference could not be determined: $($Setting.UniversalPrintingPreviewPreference.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DPILimit State ) -and ($Setting.DPILimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing print quality limit"
						$tmp = ""
						Switch ($Setting.DPILimit.Value)
						{
							"Draft"				{$tmp = "Draft (150 DPI)"; Break}
							"LowResolution"		{$tmp = "Low Resolution (300 DPI)"; Break}
							"MediumResolution"	{$tmp = "Medium Resolution (600 DPI)"; Break}
							"HighResolution"	{$tmp = "High Resolution (1200 DPI)"; Break}
							"Unlimited"			{$tmp = "No Limit"; Break}
							Default {$tmp = "Universal printing print quality limit could not be determined: $($Setting.DPILimit.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Security"
					If((validStateProp $Setting MinimumEncryptionLevel State ) -and ($Setting.MinimumEncryptionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Security\SecureICA minimum encryption level" 
						$tmp = ""
						Switch ($Setting.MinimumEncryptionLevel.Value)
						{
							"Unknown"	{$tmp = "Unknown encryption"; Break}
							"Basic"		{$tmp = "Basic"; Break}
							"LogOn"		{$tmp = "RC5 (128 bit) logon only"; Break}
							"Bits40"	{$tmp = "RC5 (40 bit)"; Break}
							"Bits56"	{$tmp = "RC5 (56 bit)"; Break}
							"Bits128"	{$tmp = "RC5 (128 bit)"; Break}
							Default		{$tmp = "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Server Limits"
					If((validStateProp $Setting IdleTimerInterval State ) -and ($Setting.IdleTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Server Limits\Server idle timer interval (milliseconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.IdleTimerInterval.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IdleTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IdleTimerInterval.Value 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tICA\Session Limits"
					If((validStateProp $Setting SessionDisconnectTimer State ) -and ($Setting.SessionDisconnectTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Disconnected session timer"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionDisconnectTimer.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionDisconnectTimer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionDisconnectTimer.State 
						}
					}
					If((validStateProp $Setting SessionDisconnectTimerInterval State ) -and ($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Disconnected session timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionDisconnectTimerInterval.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionDisconnectTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionDisconnectTimerInterval.Value 
						}
					}
					If((validStateProp $Setting SessionConnectionTimer State ) -and ($Setting.SessionConnectionTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session connection timer"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionConnectionTimer.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionConnectionTimer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionConnectionTimer.State 
						}
					}
					If((validStateProp $Setting SessionConnectionTimerInterval State ) -and ($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session connection timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionConnectionTimerInterval.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionConnectionTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionConnectionTimerInterval.Value 
						}
					}
					If((validStateProp $Setting SessionIdleTimer State ) -and ($Setting.SessionIdleTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session idle timer"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionIdleTimer.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionIdleTimer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionIdleTimer.State 
						}
					}
					If((validStateProp $Setting SessionIdleTimerInterval State ) -and ($Setting.SessionIdleTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session idle timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionIdleTimerInterval.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionIdleTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionIdleTimerInterval.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Session Reliability"
					If((validStateProp $Setting SessionReliabilityConnections State ) -and ($Setting.SessionReliabilityConnections.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability connections"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionReliabilityConnections.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityConnections.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityConnections.State 
						}
					}
					If((validStateProp $Setting SessionReliabilityPort State ) -and ($Setting.SessionReliabilityPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability port number"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionReliabilityPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityPort.Value 
						}
					}
					If((validStateProp $Setting SessionReliabilityTimeout State ) -and ($Setting.SessionReliabilityTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SessionReliabilityTimeout.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityTimeout.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Time Zone Control"
					If((validStateProp $Setting LocalTimeEstimation State ) -and ($Setting.LocalTimeEstimation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Estimate local time for legacy clients"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LocalTimeEstimation.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LocalTimeEstimation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LocalTimeEstimation.State 
						}
					}
					If((validStateProp $Setting SessionTimeZone State ) -and ($Setting.SessionTimeZone.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Use local time of client"
						$tmp = ""
						Switch ($Setting.SessionTimeZone.Value)
						{
							"UseServerTimeZone" {$tmp = "Use server time zone"; Break}
							"UseClientTimeZone" {$tmp = "Use client time zone"; Break}
							Default {$tmp = "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\TWAIN Devices"
					If((validStateProp $Setting TwainRedirection State ) -and ($Setting.TwainRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\TWAIN devices\Client TWAIN device redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TwainRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TwainRedirection.State 
						}
					}
					If((validStateProp $Setting TwainCompressionLevel State ) -and ($Setting.TwainCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\TWAIN devices\TWAIN compression level"
						Switch ($Setting.TwainCompressionLevel.Value)
						{
							"None"		{$tmp = "None"; Break}
							"Low"		{$tmp = "Low"; Break}
							"Medium"	{$tmp = "Medium"; Break}
							"High"		{$tmp = "High"; Break}
							Default		{$tmp = "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Bidirectional Content Redirection"
					If((validStateProp $Setting AllowURLRedirection State ) -and ($Setting.AllowURLRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bidirectional Content Redirection\Allow Bidirectional Content Redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AllowURLRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowURLRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowURLRedirection.State 
						}
					}
					If((validStateProp $Setting AllowedClientURLs State ) -and ($Setting.AllowedClientURLs.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bidirectional Content Redirection\Allowed URLs to be redirected to Client"
						$array = $Setting.AllowedClientURLs.Value.Split(';')
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting AllowedVDAURLs State ) -and ($Setting.AllowedVDAURLs.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bidirectional Content Redirection\Allowed URLs to be redirected to VDA"
						$array = $Setting.AllowedVDAURLs.Value.Split(';')
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\USB Devices"
					If((validStateProp $Setting ClientUsbDeviceOptimizationRules State ) -and ($Setting.ClientUsbDeviceOptimizationRules.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device optimization rules"
						$array = $Setting.ClientUsbDeviceOptimizationRules.Values
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting UsbDeviceRedirection State ) -and ($Setting.UsbDeviceRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UsbDeviceRedirection.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UsbDeviceRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UsbDeviceRedirection.State 
						}
					}
					If((validStateProp $Setting UsbDeviceRedirectionRules State ) -and ($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device redirection rules"
						$array = $Setting.UsbDeviceRedirectionRules.Values
						$tmp = $array[0]
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}

						$txt = ""
						$cnt = -1
						ForEach($element in $array)
						{
							$cnt++
							
							If($cnt -ne 0)
							{
								$tmp = "$($element) "
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$array = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting UsbPlugAndPlayRedirection State ) -and ($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB Plug and Play device redirection"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.UsbPlugAndPlayRedirection.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UsbPlugAndPlayRedirection.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UsbPlugAndPlayRedirection.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Visual Display"
					If((validStateProp $Setting PreferredColorDepthForSimpleGraphics State ) -and ($Setting.PreferredColorDepthForSimpleGraphics.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Preferred color depth for simple graphics"
						$tmp = ""
						Switch ($Setting.PreferredColorDepthForSimpleGraphics.Value)
						{
							"ColorDepth24Bit"	{$tmp = "24 bits per pixel"; Break}
							"ColorDepth16Bit"	{$tmp = "16 bits per pixel"; Break}
							"ColorDepth8Bit"	{$tmp = "8 bits per pixel"; Break}
							"Default" {$tmp = "Preferred color depth for simple graphics could not be determined: $($Setting.PreferredColorDepthForSimpleGraphics.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FramesPerSecond State ) -and ($Setting.FramesPerSecond.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Target frame rate (fps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FramesPerSecond.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FramesPerSecond.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FramesPerSecond.Value 
						}
					}
					If((validStateProp $Setting VisualQuality State ) -and ($Setting.VisualQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Visual quality"
						$tmp = ""
						Switch ($Setting.VisualQuality.Value)
						{
							"BuildToLossless"	{$tmp = "Build to Lossless"; Break}
							"AlwaysLossless"	{$tmp = "Always Lossless"; Break}
							"High"				{$tmp = "High"; Break}
							"Medium"			{$tmp = "Medium"; Break}
							"Low"				{$tmp = "Low"; Break}
							"Default" {$tmp = "Visual quality could not be determined: $($Setting.VisualQuality.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Visual Display\Moving Images"
					If((validStateProp $Setting MinimumAdaptiveDisplayJpegQuality State ) -and ($Setting.MinimumAdaptiveDisplayJpegQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Minimum image quality"
						$tmp = ""
						Switch ($Setting.MinimumAdaptiveDisplayJpegQuality.Value)
						{
							"UltraHigh" {$tmp = "Ultra high"; Break}
							"VeryHigh"  {$tmp = "Very high"; Break}
							"High"      {$tmp = "High"; Break}
							"Normal"    {$tmp = "Normal"; Break}
							"Low"       {$tmp = "Low"; Break}
							Default {$tmp = "Minimum image quality could not be determined: $($Setting.MinimumAdaptiveDisplayJpegQuality.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting MovingImageCompressionConfiguration State ) -and ($Setting.MovingImageCompressionConfiguration.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Moving image compression"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MovingImageCompressionConfiguration.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MovingImageCompressionConfiguration.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MovingImageCompressionConfiguration.State 
						}
					}
					If((validStateProp $Setting ProgressiveCompressionLevel State ) -and ($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Progressive compression level"
						$tmp = ""
						Switch ($Setting.ProgressiveCompressionLevel.Value)
						{
							"UltraHigh" {$tmp = "Ultra high"; Break}
							"VeryHigh"  {$tmp = "Very high"; Break}
							"High"      {$tmp = "High"; Break}
							"Normal"    {$tmp = "Normal"; Break}
							"Low"       {$tmp = "Low"; Break}
							"None"      {$tmp = "None"; Break}
							Default {$tmp = "Progressive compression level could not be determined: $($Setting.ProgressiveCompressionLevel.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ProgressiveCompressionThreshold State ) -and ($Setting.ProgressiveCompressionThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Progressive compression threshold value (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProgressiveCompressionThreshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProgressiveCompressionThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProgressiveCompressionThreshold.Value 
						}
					}
					If((validStateProp $Setting TargetedMinimumFramesPerSecond State ) -and ($Setting.TargetedMinimumFramesPerSecond.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Target Minimum Frame Rate (fps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TargetedMinimumFramesPerSecond.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TargetedMinimumFramesPerSecond.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TargetedMinimumFramesPerSecond.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\Visual Display\Still Images"
					If((validStateProp $Setting ExtraColorCompression State ) -and ($Setting.ExtraColorCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Extra Color Compression"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExtraColorCompression.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExtraColorCompression.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExtraColorCompression.State 
						}
					}
					If((validStateProp $Setting ExtraColorCompressionThreshold State ) -and ($Setting.ExtraColorCompressionThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Extra Color Compression Threshold (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExtraColorCompressionThreshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExtraColorCompressionThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExtraColorCompressionThreshold.Value 
						}
					}
					If((validStateProp $Setting ProgressiveHeavyweightCompression State ) -and ($Setting.ProgressiveHeavyweightCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Heavyweight compression"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProgressiveHeavyweightCompression.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProgressiveHeavyweightCompression.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProgressiveHeavyweightCompression.State 
						}
					}
					If((validStateProp $Setting LossyCompressionLevel State ) -and ($Setting.LossyCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Lossy compression level"
						$tmp = ""
						Switch ($Setting.LossyCompressionLevel.Value)
						{
							"None"		{$tmp = "None"; Break}
							"Low"		{$tmp = "Low"; Break}
							"Medium"	{$tmp = "Medium"; Break}
							"High"		{$tmp = "High"; Break}
							Default		{$tmp = "Lossy compression level could not be determined: $($Setting.LossyCompressionLevel.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting LossyCompressionThreshold State ) -and ($Setting.LossyCompressionThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Lossy compression threshold value (Kbps)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LossyCompressionThreshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LossyCompressionThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LossyCompressionThreshold.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tICA\WebSockets"
					If((validStateProp $Setting AcceptWebSocketsConnections State ) -and ($Setting.AcceptWebSocketsConnections.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSocket connections"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.AcceptWebSocketsConnections.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AcceptWebSocketsConnections.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AcceptWebSocketsConnections.State 
						}
					}
					If((validStateProp $Setting WebSocketsPort State ) -and ($Setting.WebSocketsPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockets port number"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.WebSocketsPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WebSocketsPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.WebSocketsPort.Value 
						}
					}
					If((validStateProp $Setting WSTrustedOriginServerList State ) -and ($Setting.WSTrustedOriginServerList.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockets trusted origin server list"
						$tmpArray = $Setting.WSTrustedOriginServerList.Value.Split(",")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $tmpArray)
						{
							$cnt++
							$tmp = "$($Thing)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmpArray = $Null
						$tmp = $Null
					}
					
					Write-Verbose "$(Get-Date): `t`t`tLoad Management"
					If((validStateProp $Setting ConcurrentLogonsTolerance State ) -and ($Setting.ConcurrentLogonsTolerance.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Concurrent logons tolerance"
						If($Setting.ConcurrentLogonsTolerance.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ConcurrentLogonsTolerance.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ConcurrentLogonsTolerance.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ConcurrentLogonsTolerance.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ConcurrentLogonsTolerance.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ConcurrentLogonsTolerance.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ConcurrentLogonsTolerance.State 
							}
						}
					}
					If((validStateProp $Setting CPUUsage State ) -and ($Setting.CPUUsage.State -ne "NotConfigured"))
					{
						$txt = "Load Management\CPU usage"
						If($MSWord -or $PDF)
						{
							If($Setting.CPUUsage.State -eq "Enabled")
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = "Report full load $($Setting.CPUUsage.Value)(%)";
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							Else
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPUUsage.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
						}
						ElseIf($HTML)
						{
							If($Setting.CPUUsage.State -eq "Enabled")
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								"Report full load $($Setting.CPUUsage.Value)(%)",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPUUsage.State,$htmlwhite))
							}
						}
						ElseIf($Text)
						{
							If($Setting.CPUUsage.State -eq "Enabled")
							{
								OutputPolicySetting $txt $Setting.CPUUsage.State 
							}
							Else
							{
								OutputPolicySetting $txt "Report full load $($Setting.CPUUsage.Value)(%)" 
							}
						}
					}
					If((validStateProp $Setting CPUUsageExcludedProcessPriority State ) -and ($Setting.CPUUsageExcludedProcessPriority.State -ne "NotConfigured"))
					{
						$txt = "Load Management\CPU usage excluded process priority"
						If($Setting.CPUUsageExcludedProcessPriority.State -eq "Enabled")
						{
							$tmp = ""
							Switch ($Setting.CPUUsageExcludedProcessPriority.Value)
							{
								"BelowNormalOrLow"	{$tmp = "Below Normal or Low"; Break}
								"Low"				{$tmp = "Low"; Break}
								Default {$tmp = "CPU usage excluded process priority could not be determined: $($Setting.CPUUsageExcludedProcessPriority.Value)"; Break}
							}
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPUUsageExcludedProcessPriority.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPUUsageExcludedProcessPriority.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.CPUUsageExcludedProcessPriority.State 
							}
						}
					}
					If((validStateProp $Setting DiskUsage State ) -and ($Setting.DiskUsage.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Disk usage"
						$tmp = ""
						If($Setting.DiskUsage.State -eq "Enabled")
						{
							$tmp = "Report 75% load (disk queue length): $($Setting.DiskUsage.Value)"
						}
						Else
						{
							$tmp = "Disabled"
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting MaximumNumberOfSessions State ) -and ($Setting.MaximumNumberOfSessions.State -ne "NotConfigured"))
					{
						If($Setting.MaximumNumberOfSessions.State -eq "Enabled")
						{
							$txt = "Load Management\Maximum number of sessions - Limit"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.MaximumNumberOfSessions.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MaximumNumberOfSessions.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.MaximumNumberOfSessions.Value 
							}
						}
						Else
						{
							$txt = "Load Management\Maximum number of sessions"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.MaximumNumberOfSessions.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MaximumNumberOfSessions.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.MaximumNumberOfSessions.Value 
							}
						}
					}
					If((validStateProp $Setting MemoryUsage State ) -and ($Setting.MemoryUsage.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Memory usage"
						$tmp = ""
						If($Setting.MemoryUsage.State -eq "Enabled")
						{
							$tmp = "Report full load (%): $($Setting.MemoryUsage.Value)"
						}
						Else
						{
							$tmp = "Disabled"
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting MemoryUsageBaseLoad State ) -and ($Setting.MemoryUsageBaseLoad.State -ne "NotConfigured"))
					{
						$txt = "Load Management\Memory usage base load"
						$tmp = ""
						If($Setting.MemoryUsageBaseLoad.State -eq "Enabled")
						{
							$tmp = "Report zero load (MBs): $($Setting.MemoryUsageBaseLoad.Value)"
						}
						Else
						{
							$tmp = "Disabled"
						}
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management"
					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Advanced settings"
					If((validStateProp $Setting CEIPEnabled State ) -and ($Setting.CEIPEnabled.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Customer Experience Improvement Program"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.CEIPEnabled.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CEIPEnabled.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.CEIPEnabled.State
						}
					}
					If((validStateProp $Setting DisableDynamicConfig State ) -and ($Setting.DisableDynamicConfig.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Disable automatic configuration"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DisableDynamicConfig.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisableDynamicConfig.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DisableDynamicConfig.State
						}
					}
					If((validStateProp $Setting LogoffRatherThanTempProfile State ) -and ($Setting.LogoffRatherThanTempProfile.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Log off user if a problem is encountered"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogoffRatherThanTempProfile.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogoffRatherThanTempProfile.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogoffRatherThanTempProfile.State
						}
					}
					If((validStateProp $Setting LoadRetries_Part State ) -and ($Setting.LoadRetries_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Number of retries when accessing locked files"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LoadRetries_Part.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LoadRetries_Part.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LoadRetries_Part.Value 
						}
					}
					If((validStateProp $Setting ProcessCookieFiles State ) -and ($Setting.ProcessCookieFiles.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Advanced settings\Process Internet cookie files on logoff"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProcessCookieFiles.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProcessCookieFiles.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProcessCookieFiles.State
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Basic settings"
					If((validStateProp $Setting PSMidSessionWriteBack State ) -and ($Setting.PSMidSessionWriteBack.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Active write back"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSMidSessionWriteBack.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSMidSessionWriteBack.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PSMidSessionWriteBack.State
						}
					}
					If((validStateProp $Setting PSMidSessionWriteBackReg State ) -and ($Setting.PSMidSessionWriteBackReg.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Active write back registry"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSMidSessionWriteBackReg.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSMidSessionWriteBackReg.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PSMidSessionWriteBack.State
						}
					}
					If((validStateProp $Setting ServiceActive State ) -and ($Setting.ServiceActive.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Enable Profile management"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ServiceActive.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ServiceActive.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ServiceActive.State
						}
					}
					If((validStateProp $Setting ExcludedGroups_Part State ) -and ($Setting.ExcludedGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Excluded groups"
						If($Setting.ExcludedGroups_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.ExcludedGroups_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ExcludedGroups_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExcludedGroups_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ExcludedGroups_Part.State
							}
						}
					}
					If((validStateProp $Setting OfflineSupport State ) -and ($Setting.OfflineSupport.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Offline profile support"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.OfflineSupport.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OfflineSupport.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OfflineSupport.State
						}
					}
					If((validStateProp $Setting DATPath_Part State ) -and ($Setting.DATPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Path to user store"
						If($Setting.DATPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.DATPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DATPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.DATPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.DATPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DATPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.DATPath_Part.State
							}
						}
					}
					If((validStateProp $Setting ProcessAdmins State ) -and ($Setting.ProcessAdmins.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Process logons of local administrators"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProcessAdmins.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProcessAdmins.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProcessAdmins.State
						}
					}
					If((validStateProp $Setting ProcessedGroups_Part State ) -and ($Setting.ProcessedGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Basic settings\Processed groups"
						If($Setting.ProcessedGroups_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.ProcessedGroups_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}	
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ProcessedGroups_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ProcessedGroups_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ProcessedGroups_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Cross-Platform settings"
					If((validStateProp $Setting CPUserGroups_Part State ) -and ($Setting.CPUserGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Cross-platform settings user groups"
						If($Setting.CPUserGroups_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.CPUserGroups_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPUserGroups_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPUserGroups_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.CPUserGroups_Part.State
							}
						}
					}
					If((validStateProp $Setting CPEnable State ) -and ($Setting.CPEnable.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Enable cross-platform settings"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.CPEnable.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPEnable.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.CPEnable.State
						}
					}
					If((validStateProp $Setting CPSchemaPathData State ) -and ($Setting.CPSchemaPathData.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Path to cross-platform definitions"
						If($Setting.CPSchemaPathData.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPSchemaPathData.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPSchemaPathData.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.CPSchemaPathData.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPSchemaPathData.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPSchemaPathData.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.CPSchemaPathData.State
							}
						}
					}
					If((validStateProp $Setting CPPathData State ) -and ($Setting.CPPathData.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Path to cross-platform settings store"
						If($Setting.CPPathData.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPPathData.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPPathData.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.CPPathData.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.CPPathData.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.CPPathData.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.CPPathData.State
							}
						}
					}
					If((validStateProp $Setting CPMigrationFromBaseProfileToCPStore State ) -and ($Setting.CPMigrationFromBaseProfileToCPStore.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Cross-Platform settings\Source for creating cross-platform settings"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.CPMigrationFromBaseProfileToCPStore.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPMigrationFromBaseProfileToCPStore.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.CPMigrationFromBaseProfileToCPStore.State
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\File system"
					Write-Verbose "$(Get-Date): `t`t`tProfile Management\File system\Default Exclusions"
					If((validStateProp $Setting DefaultExclusionListSyncDir State ) -and ($Setting.DefaultExclusionListSyncDir.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\Enable Default Exclusion List - directories"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DefaultExclusionListSyncDir.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DefaultExclusionListSyncDir.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DefaultExclusionListSyncDir.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir01 State ) -and ($Setting.ExclusionDefaultDir01.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_internetcache!"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir01.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir01.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir01.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir02 State ) -and ($Setting.ExclusionDefaultDir02.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\Cache"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir02.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir02.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir02.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir03 State ) -and ($Setting.ExclusionDefaultDir03.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\Cache Theme Images"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir03.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir03.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir03.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir04 State ) -and ($Setting.ExclusionDefaultDir04.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\JumpListIcons"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir04.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir04.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir04.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir05 State ) -and ($Setting.ExclusionDefaultDir05.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Google\Chrome\User Data\Default\JumpListIconsOld"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir05.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir05.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir05.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir06 State ) -and ($Setting.ExclusionDefaultDir06.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\GroupPolicy"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir06.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir06.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir06.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir07 State ) -and ($Setting.ExclusionDefaultDir07.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\AppV"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir07.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir07.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir07.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir08 State ) -and ($Setting.ExclusionDefaultDir08.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Messenger"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir08.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir08.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir08.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir09 State ) -and ($Setting.ExclusionDefaultDir09.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Office\15.0\Lync\Tracing"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir09.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir09.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir09.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir10 State ) -and ($Setting.ExclusionDefaultDir10.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\OneNote"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir10.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir10.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir10.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir11 State ) -and ($Setting.ExclusionDefaultDir11.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Outlook"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir11.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir11.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir11.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir12 State ) -and ($Setting.ExclusionDefaultDir12.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Terminal Server Client"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir12.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir12.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir12.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir13 State ) -and ($Setting.ExclusionDefaultDir13.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\UEV"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir13.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir13.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir13.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir14 State ) -and ($Setting.ExclusionDefaultDir14.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows Live"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir14.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir14.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir14.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir15 State ) -and ($Setting.ExclusionDefaultDir15.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows Live Contacts"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir15.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir15.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir15.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir16 State ) -and ($Setting.ExclusionDefaultDir16.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\Application Shortcuts"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir16.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir16.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir16.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir17 State ) -and ($Setting.ExclusionDefaultDir17.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\Burn"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir17.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir17.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir17.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir18 State ) -and ($Setting.ExclusionDefaultDir18.State -ne "NotConfigured"))
					{
						#fixed 3-Mar-2017 thanks to Esther Barthel
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\CD Burning"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir18.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir18.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir18.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir19 State ) -and ($Setting.ExclusionDefaultDir19.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Microsoft\Windows\Notifications"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir19.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir19.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir19.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir20 State ) -and ($Setting.ExclusionDefaultDir20.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Packages"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir20.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir20.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir20.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir21 State ) -and ($Setting.ExclusionDefaultDir21.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Sun"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir21.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir21.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir21.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir22 State ) -and ($Setting.ExclusionDefaultDir22.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localappdata!\Windows Live"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir22.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir22.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir22.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir23 State ) -and ($Setting.ExclusionDefaultDir23.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_localsettings!\Temp"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir23.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir23.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir23.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir24 State ) -and ($Setting.ExclusionDefaultDir24.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Microsoft\AppV\Client\Catalog"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir24.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir24.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir24.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir25 State ) -and ($Setting.ExclusionDefaultDir25.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Sun\Java\Deployment\cache"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir25.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir25.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir25.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir26 State ) -and ($Setting.ExclusionDefaultDir26.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Sun\Java\Deployment\log"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir26.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir26.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir26.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir27 State ) -and ($Setting.ExclusionDefaultDir27.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - !ctx_roamingappdata!\Sun\Java\Deployment\tmp"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir27.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir27.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir27.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir28 State ) -and ($Setting.ExclusionDefaultDir28.State -ne "NotConfigured"))
					{
						$txt = 'Profile Management\File system\Default Exclusions\UPM - $Recycle.Bin'
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir28.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir28.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir28.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir29 State ) -and ($Setting.ExclusionDefaultDir29.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - AppData\LocalLow"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir29.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir29.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir29.State
						}
					}
					If((validStateProp $Setting ExclusionDefaultDir30 State ) -and ($Setting.ExclusionDefaultDir30.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Default Exclusions\UPM - Tracing"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultDir30.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultDir30.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultDir30.State
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tProfile Management\File system\Exclusions"
					If((validStateProp $Setting ExclusionListSyncDir_Part State ) -and ($Setting.ExclusionListSyncDir_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Exclusions\Exclusion list - directories"
						If($Setting.ExclusionListSyncDir_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.ExclusionListSyncDir_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ExclusionListSyncDir_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExclusionListSyncDir_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ExclusionListSyncDir_Part.State
							}
						}
					}
					If((validStateProp $Setting ExclusionListSyncFiles_Part State ) -and ($Setting.ExclusionListSyncFiles_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Exclusions\Exclusion list - files"
						If($Setting.ExclusionListSyncFiles_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.ExclusionListSyncFiles_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ExclusionListSyncFiles_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExclusionListSyncFiles_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ExclusionListSyncFiles_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\File system\Synchronization"
					If((validStateProp $Setting SyncDirList_Part State ) -and ($Setting.SyncDirList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Directories to synchronize"
						If($Setting.SyncDirList_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.SyncDirList_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.SyncDirList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.SyncDirList_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.SyncDirList_Part.State
							}
						}
					}
					If((validStateProp $Setting SyncFileList_Part State ) -and ($Setting.SyncFileList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Files to synchronize"
						If($Setting.SyncFileList_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.SyncFileList_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.SyncFileList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.SyncFileList_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.SyncFileList_Part.State
							}
						}
					}
					If((validStateProp $Setting MirrorFoldersList_Part State ) -and ($Setting.MirrorFoldersList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\File system\Synchronization\Folders to mirror"
						If($Setting.MirrorFoldersList_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.MirrorFoldersList_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.MirrorFoldersList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.MirrorFoldersList_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.MirrorFoldersList_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection"
					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\AppData(Roaming)"
					If((validStateProp $Setting FRAppDataPath_Part State ) -and ($Setting.FRAppDataPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\AppData(Roaming)\AppData(Roaming) path"
						If($Setting.FRAppDataPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRAppDataPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRAppDataPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRAppDataPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRAppDataPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRAppDataPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRAppDataPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRAppData_Part State ) -and ($Setting.FRAppData_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\AppData(Roaming)\Redirection settings for AppData(Roaming)"
						$tmp = ""
						Switch ($Setting.FRAppData_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRAppData_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Common settings"
					If((validStateProp $Setting FRAdminAccess_Part State ) -and ($Setting.FRAdminAccess_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Common settings\Grant administrator access"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FRAdminAccess_Part.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FRAdminAccess_Part.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FRAdminAccess_Part.State
						}
					}
					If((validStateProp $Setting FRIncDomainName_Part State ) -and ($Setting.FRIncDomainName_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Common settings\Include domain name"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.FRIncDomainName_Part.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FRIncDomainName_Part.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FRIncDomainName_Part.State
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Contacts"
					If((validStateProp $Setting FRContactsPath_Part State ) -and ($Setting.FRContactsPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Contacts\Contacts path"
						If($Setting.FRContactsPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRContactsPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRContactsPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRContactsPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRContactsPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRContactsPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRContactsPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRContacts_Part State ) -and ($Setting.FRContacts_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Contacts\Redirection settings for Contacts"
						$tmp = ""
						Switch ($Setting.FRContacts_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRContacts_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Desktop"
					If((validStateProp $Setting FRDesktopPath_Part State ) -and ($Setting.FRDesktopPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Desktop\Desktop path"
						If($Setting.FRDesktopPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRDesktopPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDesktopPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRDesktopPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRDesktopPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDesktopPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRDesktopPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRDesktop_Part State ) -and ($Setting.FRDesktop_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Desktop\Redirection settings for Desktop"
						$tmp = ""
						Switch ($Setting.FRDesktop_Part.Value)
						{
							"RedirectUncPath"	{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRDesktop_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Documents"
					If((validStateProp $Setting FRDocumentsPath_Part State ) -and ($Setting.FRDocumentsPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Documents\Documents path"
						If($Setting.FRDocumentsPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRDocumentsPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDocumentsPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRDocumentsPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRDocumentsPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDocumentsPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRDocumentsPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRDocuments_Part State ) -and ($Setting.FRDocuments_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Documents\Redirection settings for Documents"
						$tmp = ""
						Switch ($Setting.FRDocuments_Part.Value)
						{
							"RedirectUncPath"		{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeHome"	{$tmp = "Redirect to the users' home directory"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRDocuments_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Downloads"
					If((validStateProp $Setting FRDownloadsPath_Part State ) -and ($Setting.FRDownloadsPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Downloads\Downloads path"
						If($Setting.FRDownloadsPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRDownloadsPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDownloadsPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRDownloadsPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRDownloadsPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRDownloadsPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRDownloadsPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRDownloads_Part State ) -and ($Setting.FRDownloads_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Downloads\Redirection settings for Downloads"
						$tmp = ""
						Switch ($Setting.FRDownloads_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRDownloads_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Favorites"
					If((validStateProp $Setting FRFavoritesPath_Part State ) -and ($Setting.FRFavoritesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Favorites\Favorites path"
						If($Setting.FRFavoritesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRFavoritesPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRFavoritesPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRFavoritesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRFavoritesPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRFavoritesPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRFavoritesPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRFavorites_Part State ) -and ($Setting.FRFavorites_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Favorites\Redirection settings for Favorites"
						$tmp = ""
						Switch ($Setting.FRFavorites_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRFavorites_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Links"
					If((validStateProp $Setting FRLinksPath_Part State ) -and ($Setting.FRLinksPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Links\Links path"
						If($Setting.FRLinksPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRLinksPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRLinksPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRLinksPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRLinksPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRLinksPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRLinksPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRLinks_Part State ) -and ($Setting.FRLinks_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Links\Redirection settings for Links"
						$tmp = ""
						Switch ($Setting.FRLinks_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRLinks_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Music"
					If((validStateProp $Setting FRMusicPath_Part State ) -and ($Setting.FRMusicPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Music\Music path"
						If($Setting.FRMusicPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRMusicPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRMusicPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRMusicPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRMusicPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRMusicPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRMusicPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRMusic_Part State ) -and ($Setting.FRMusic_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Music\Redirection settings for Music"
						$tmp = ""
						Switch ($Setting.FRMusic_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeDocuments" {$tmp = "Redirect relative to Documents folder"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRMusic_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Pictures"
					If((validStateProp $Setting FRPicturesPath_Part State ) -and ($Setting.FRPicturesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Pictures\Pictures path"
						If($Setting.FRPicturesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRPicturesPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRPicturesPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRPicturesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRPicturesPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRPicturesPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRPicturesPath_Part.State
							}
						}
					}
					If((validStateProp $Setting FRPictures_Part State ) -and ($Setting.FRPictures_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Pictures\Redirection settings for Pictures"
						$tmp = ""
						Switch ($Setting.FRPictures_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeDocuments" {$tmp = "Redirect relative to Documents folder"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRPictures_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Saved Games"
					If((validStateProp $Setting FRSavedGames_Part State ) -and ($Setting.FRSavedGames_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Saved Games\Redirection settings for Saved Games"
						$tmp = ""
						Switch ($Setting.FRSavedGames_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRSavedGames_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRSavedGamesPath_Part State ) -and ($Setting.FRSavedGamesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Saved Games\Saved Games path"
						If($Setting.FRSavedGamesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRSavedGamesPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSavedGamesPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRSavedGamesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRSavedGamesPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSavedGamesPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRSavedGamesPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Searches"
					If((validStateProp $Setting FRSearches_Part State ) -and ($Setting.FRSearches_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Searches\Redirection settings for Searches"
						$tmp = ""
						Switch ($Setting.FRSearches_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRSearches_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRSearchesPath_Part State ) -and ($Setting.FRSearchesPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Searches\Searches path"
						If($Setting.FRSearchesPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRSearchesPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSearchesPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRSearchesPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRSearchesPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRSearchesPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRSearchesPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Start Menu"
					If((validStateProp $Setting FRStartMenu_Part State ) -and ($Setting.FRStartMenu_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Start Menu\Redirection settings for Start Menu"
						$tmp = ""
						Switch ($Setting.FRStartMenu_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRStartMenu_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRStartMenuPath_Part State ) -and ($Setting.FRStartMenuPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Start Menu\Start Menu path"
						If($Setting.FRStartMenuPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRStartMenuPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRStartMenuPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRStartMenuPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRStartMenuPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRStartMenuPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRStartMenuPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Folder Redirection\Videos"
					If((validStateProp $Setting FRVideos_Part State ) -and ($Setting.FRVideos_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Videos\Redirection settings for Videos"
						$tmp = ""
						Switch ($Setting.FRVideos_Part.Value)
						{
							"RedirectUncPath"			{$tmp = "Redirect to the following UNC path"; Break}
							"RedirectRelativeDocuments" {$tmp = "Redirect relative to Documents folder"; Break}
							Default {$tmp = "AppData(Roaming) path cannot be determined: $($Setting.FRVideos_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FRVideosPath_Part State ) -and ($Setting.FRVideosPath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Folder Redirection\Videos\Videos path"
						If($Setting.FRVideosPath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRVideosPath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRVideosPath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRVideosPath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.FRVideosPath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.FRVideosPath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.FRVideosPath_Part.State
							}
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Log settings"
					If((validStateProp $Setting LogLevel_ActiveDirectoryActions State ) -and ($Setting.LogLevel_ActiveDirectoryActions.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Active Directory actions"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_ActiveDirectoryActions.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_ActiveDirectoryActions.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_ActiveDirectoryActions.State
						}
					}
					If((validStateProp $Setting LogLevel_Information State ) -and ($Setting.LogLevel_Information.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Common information"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_Information.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Information.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Information.State
						}
					}
					If((validStateProp $Setting LogLevel_Warnings State ) -and ($Setting.LogLevel_Warnings.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Common warnings"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_Warnings.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Warnings.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Warnings.State
						}
					}
					If((validStateProp $Setting DebugMode State ) -and ($Setting.DebugMode.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Enable logging"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DebugMode.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DebugMode.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DebugMode.State
						}
					}
					If((validStateProp $Setting LogLevel_FileSystemActions State ) -and ($Setting.LogLevel_FileSystemActions.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\File system actions"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_FileSystemActions.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_FileSystemActions.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_FileSystemActions.State
						}
					}
					If((validStateProp $Setting LogLevel_FileSystemNotification State ) -and ($Setting.LogLevel_FileSystemNotification.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\File system notifications"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_FileSystemNotification.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_FileSystemNotification.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_FileSystemNotification.State
						}
					}
					If((validStateProp $Setting LogLevel_Logoff State ) -and ($Setting.LogLevel_Logoff.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Logoff"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_Logoff.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Logoff.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Logoff.State
						}
					}
					If((validStateProp $Setting LogLevel_Logon State ) -and ($Setting.LogLevel_Logon.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Logon"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_Logon.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_Logon.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_Logon.State
						}
					}
					If((validStateProp $Setting MaxLogSize_Part State ) -and ($Setting.MaxLogSize_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Maximum size of the log file"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.MaxLogSize_Part.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MaxLogSize_Part.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MaxLogSize_Part.Value 
						}
					}
					If((validStateProp $Setting DebugFilePath_Part State ) -and ($Setting.DebugFilePath_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Path to log file"
						If($Setting.DebugFilePath_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.DebugFilePath_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DebugFilePath_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.DebugFilePath_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.DebugFilePath_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.DebugFilePath_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.DebugFilePath_Part.State
							}
						}
					}
					If((validStateProp $Setting LogLevel_UserName State ) -and ($Setting.LogLevel_UserName.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Personalized user information"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_UserName.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_UserName.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_UserName.State
						}
					}
					If((validStateProp $Setting LogLevel_PolicyUserLogon State ) -and ($Setting.LogLevel_PolicyUserLogon.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Policy values at logon and logoff"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_PolicyUserLogon.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_PolicyUserLogon.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_PolicyUserLogon.State
						}
					}
					If((validStateProp $Setting LogLevel_RegistryActions State ) -and ($Setting.LogLevel_RegistryActions.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Registry actions"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_RegistryActions.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_RegistryActions.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_RegistryActions.State
						}
					}
					If((validStateProp $Setting LogLevel_RegistryDifference State ) -and ($Setting.LogLevel_RegistryDifference.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Log settings\Registry differences at logoff"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LogLevel_RegistryDifference.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LogLevel_RegistryDifference.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LogLevel_RegistryDifference.State
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Profile handling"
					If((validStateProp $Setting ProfileDeleteDelay_Part State ) -and ($Setting.ProfileDeleteDelay_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Delay before deleting cached profiles"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProfileDeleteDelay_Part.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProfileDeleteDelay_Part.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProfileDeleteDelay_Part.Value 
						}
					}
					If((validStateProp $Setting DeleteCachedProfilesOnLogoff State ) -and ($Setting.DeleteCachedProfilesOnLogoff.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Delete locally cached profiles on logoff"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DeleteCachedProfilesOnLogoff.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DeleteCachedProfilesOnLogoff.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DeleteCachedProfilesOnLogoff.State
						}
					}
					If((validStateProp $Setting LocalProfileConflictHandling_Part State ) -and ($Setting.LocalProfileConflictHandling_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Local profile conflict handling"
						$tmp = ""
						Switch ($Setting.LocalProfileConflictHandling_Part.Value)
						{
							"Use"		{$tmp = "Use local profile"; Break}
							"Delete"	{$tmp = "Delete local profile"; Break}
							"Rename"	{$tmp = "Rename local profile"; Break}
							Default		{$tmp = "Local profile conflict handling could not be determined: $($Setting.LocalProfileConflictHandling_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting MigrateWindowsProfilesToUserStore_Part State ) -and ($Setting.MigrateWindowsProfilesToUserStore_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Migration of existing profiles"
						$tmp = ""
						Switch ($Setting.MigrateWindowsProfilesToUserStore_Part.Value)
						{
							"All"		{$tmp = "Local and Roaming"; Break}
							"Local"		{$tmp = "Local"; Break}
							"Roaming"	{$tmp = "Roaming"; Break}
							"None"		{$tmp = "None"; Break}
							Default		{$tmp = "Migration of existing profiles could not be determined: $($Setting.MigrateWindowsProfilesToUserStore_Part.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting TemplateProfilePath State ) -and ($Setting.TemplateProfilePath.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Path to the template profile"
						If($Setting.TemplateProfilePath.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.TemplateProfilePath.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.TemplateProfilePath.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.TemplateProfilePath.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.TemplateProfilePath.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.TemplateProfilePath.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.TemplateProfilePath.State
							}
						}
					}
					If((validStateProp $Setting TemplateProfileOverridesLocalProfile State ) -and ($Setting.TemplateProfileOverridesLocalProfile.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Template profile overrides local profile"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TemplateProfileOverridesLocalProfile.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TemplateProfileOverridesLocalProfile.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TemplateProfileOverridesLocalProfile.State
						}
					}
					If((validStateProp $Setting TemplateProfileOverridesRoamingProfile State ) -and ($Setting.TemplateProfileOverridesRoamingProfile.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Template profile overrides roaming profile"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TemplateProfileOverridesRoamingProfile.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TemplateProfileOverridesRoamingProfile.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TemplateProfileOverridesRoamingProfile.State
						}
					}
					If((validStateProp $Setting TemplateProfileIsMandatory State ) -and ($Setting.TemplateProfileIsMandatory.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Profile handling\Template profile used as a Citrix mandatory profile for all logons"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.TemplateProfileIsMandatory.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TemplateProfileIsMandatory.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TemplateProfileIsMandatory.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Registry"
					If((validStateProp $Setting ExclusionList_Part State ) -and ($Setting.ExclusionList_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\Exclusion list"
						If($Setting.ExclusionList_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.ExclusionList_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.ExclusionList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.ExclusionList_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.ExclusionList_Part.State
							}
						}
					}
					If((validStateProp $Setting IncludeListRegistry_Part State ) -and ($Setting.IncludeListRegistry_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\Inclusion list"
						If($Setting.IncludeListRegistry_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.IncludeListRegistry_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.IncludeList_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.IncludeList_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.IncludeList_Part.State
							}
						}
					}
					If((validStateProp $Setting LastKnownGoodRegistry State ) -and ($Setting.LastKnownGoodRegistry.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\NTUSER.DAT backup"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.LastKnownGoodRegistry.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LastKnownGoodRegistry.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LastKnownGoodRegistry.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Registry\Default Exclusions"
					If((validStateProp $Setting DefaultExclusionList State ) -and ($Setting.DefaultExclusionList.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\Enable Default Exclusion list"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.DefaultExclusionList.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DefaultExclusionList.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DefaultExclusionList.State 
						}
					}
					If((validStateProp $Setting ExclusionDefaultReg01 State ) -and ($Setting.ExclusionDefaultReg01.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\UPM - Software\Microsoft\AppV\Client\Integration"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultReg01.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultReg01.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultReg01.State 
						}
					}
					If((validStateProp $Setting ExclusionDefaultReg02 State ) -and ($Setting.ExclusionDefaultReg02.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\UPM - Software\Microsoft\AppV\Client\Publishing"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultReg02.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultReg02.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultReg02.State 
						}
					}
					If((validStateProp $Setting ExclusionDefaultReg03 State ) -and ($Setting.ExclusionDefaultReg03.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Registry\UPM - Software\Microsoft\Speech_OneCore"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ExclusionDefaultReg03.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExclusionDefaultReg03.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExclusionDefaultReg03.State 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tProfile Management\Streamed user profiles"
					If((validStateProp $Setting PSAlwaysCache State ) -and ($Setting.PSAlwaysCache.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Always cache"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSAlwaysCache.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSAlwaysCache.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PSAlwaysCache.State 
						}
					}
					If((validStateProp $Setting PSAlwaysCache_Part State ) -and ($Setting.PSAlwaysCache_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Always cache size"
						If($Setting.PSAlwaysCache_Part.State -eq "Enabled")
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PSAlwaysCache_Part.Value;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PSAlwaysCache_Part.Value,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.PSAlwaysCache_Part.Value 
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PSAlwaysCache_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PSAlwaysCache_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.PSAlwaysCache_Part.State 
							}
						}
					}
					If((validStateProp $Setting PSEnabled State ) -and ($Setting.PSEnabled.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Profile streaming"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSEnabled.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSEnabled.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PSEnabled.State 
						}
					}
					If((validStateProp $Setting PSUserGroups_Part State ) -and ($Setting.PSUserGroups_Part.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Streamed user profile groups"
						If($Setting.PSUserGroups_Part.State -eq "Enabled")
						{
							$tmpArray = $Setting.PSUserGroups_Part.Values
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $tmpArray)
							{
								$cnt++
								$tmp = "$($Thing)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $Setting.PSUserGroups_Part.State;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PSUserGroups_Part.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.PSUserGroups_Part.State 
							}
						}
					}
					If((validStateProp $Setting PSPendingLockTimeout State ) -and ($Setting.PSPendingLockTimeout.State -ne "NotConfigured"))
					{
						$txt = "Profile Management\Streamed user profiles\Timeout for pending area lock files (days)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.PSPendingLockTimeout.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PSPendingLockTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PSPendingLockTimeout.Value 
						}
					}

					Write-Verbose "$(Get-Date): `t`t`tReceiver"
					If((validStateProp $Setting StorefrontAccountsList State ) -and ($Setting.StorefrontAccountsList.State -ne "NotConfigured"))
					{
						$txt = "Receiver\Storefront accounts list"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = "";
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt ""
						}
						$txt = ""
						$tmpArray = $Setting.StorefrontAccountsList.Values
						ForEach($Thing in $TmpArray)
						{
							$cnt++
							$xxx = """$($Thing)"""
							[array]$tmp = $xxx.Split(";").replace('"','')
							$tmp1 = "Name: $($tmp[0])"
							$tmp2 = "URL: $($tmp[1])"
							$tmp3 = "State: $($tmp[2])"
							$tmp4 = "Desc: $($tmp[3])"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp1;
								}
								$SettingsWordTable += $WordTableRowHash
								
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp2;
								}
								$SettingsWordTable += $WordTableRowHash
								
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp3;
								}
								$SettingsWordTable += $WordTableRowHash
								
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp4;
								}
								$SettingsWordTable += $WordTableRowHash
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp1,$htmlwhite))
								
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp2,$htmlwhite))
								
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp3,$htmlwhite))
								
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp4,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp1
								OutputPolicySetting $txt $tmp2
								OutputPolicySetting $txt $tmp3
								OutputPolicySetting $txt $tmp4
							}
							$tmp = " "
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = "";
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								"",$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "" $tmp
							}
							$xxx = $Null
							$tmp = $Null
							$tmp1 = $Null
							$tmp2 = $Null
							$tmp3 = $Null
							$tmp4 = $Null
						}
						$TmpArray = $Null
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tVirtual Delivery Agent Settings"
					If((validStateProp $Setting ControllerRegistrationIPv6Netmask State ) -and ($Setting.ControllerRegistrationIPv6Netmask.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller registration IPv6 netmask"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerRegistrationIPv6Netmask.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerRegistrationIPv6Netmask.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerRegistrationIPv6Netmask.Value 
						}
					}
					If((validStateProp $Setting ControllerRegistrationPort State ) -and ($Setting.ControllerRegistrationPort.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller registration port"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerRegistrationPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerRegistrationPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerRegistrationPort.Value 
						}
					}
					If((validStateProp $Setting ControllerSIDs State ) -and ($Setting.ControllerSIDs.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller SIDs"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerSIDs.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerSIDs.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerSIDs.Value 
						}
					}
					If((validStateProp $Setting Controllers State ) -and ($Setting.Controllers.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controllers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.Controllers.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.Controllers.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.Controllers.Value 
						}
					}
					If((validStateProp $Setting EnableAutoUpdateOfControllers State ) -and ($Setting.EnableAutoUpdateOfControllers.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Enable auto update of Controllers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableAutoUpdateOfControllers.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableAutoUpdateOfControllers.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnableAutoUpdateOfControllers.State 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\HDX3DPro"
					If((validStateProp $Setting EnableLossless State ) -and ($Setting.EnableLossless.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\HDX3DPro\Enable lossless"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableLossless.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableLossless.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnableLossless.State 
						}
					}
					If((validStateProp $Setting ProGraphicsObj State ) -and ($Setting.ProGraphicsObj.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\HDX3DPro\HDX3DPro quality settings"
						$tmp = ""
						$xMin = [math]::floor($Setting.ProGraphicsObj.Value%65536).ToString()
						$xMax = [math]::floor($Setting.ProGraphicsObj.Value/65536).ToString()
						[string]$tmp = "Minimum: $($xMin) Maximum: $($xMax)"
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Verbose "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\Monitoring"
					If((validStateProp $Setting EnableProcessMonitoring State ) -and ($Setting.EnableProcessMonitoring.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\Enable process monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableProcessMonitoring.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableProcessMonitoring.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnableProcessMonitoring.State 
						}
					}
					If((validStateProp $Setting EnableResourceMonitoring State ) -and ($Setting.EnableResourceMonitoring.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Monitoring\Enable resource monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableResourceMonitoring.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableResourceMonitoring.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnableResourceMonitoring.State 
						}
					}
					If((validStateProp $Setting OnlyUseIPv6ControllerRegistration State ) -and ($Setting.OnlyUseIPv6ControllerRegistration.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Only use IPv6 Controller registration"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.OnlyUseIPv6ControllerRegistration.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OnlyUseIPv6ControllerRegistration.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OnlyUseIPv6ControllerRegistration.State 
						}
					}
					If((validStateProp $Setting SiteGUID State ) -and ($Setting.SiteGUID.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Site GUID"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SiteGUID.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SiteGUID.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SiteGUID.Value 
						}
					}
					
					Write-Verbose "$(Get-Date): `t`t`tVirtual IP"
					If((validStateProp $Setting VirtualLoopbackSupport State ) -and ($Setting.VirtualLoopbackSupport.State -ne "NotConfigured"))
					{
						$txt = "Virtual IP\Virtual IP loopback support"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.VirtualLoopbackSupport.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.VirtualLoopbackSupport.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.VirtualLoopbackSupport.State 
						}
					}
					If((validStateProp $Setting VirtualLoopbackPrograms State ) -and ($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured"))
					{
						$txt = "Virtual IP\Virtual IP virtual loopback programs list"
						$tmpArray = $Setting.VirtualLoopbackPrograms.Values
						$array = $Null
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = $txt;
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$WordTableRowHash = @{
									Text = "";
									Value = $tmp;
									}
									$SettingsWordTable += $WordTableRowHash;
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$TmpArray = $Null
						$tmp = $Null
					}
				}
				If($MSWord -or $PDF)
				{
					$Table = AddWordTable -Hashtable $SettingsWordTable `
					-Columns  Text,Value `
					-Headers  "Setting Key","Value"`
					-Format $wdTableLightListAccent3 `
					-NoInternalGridLines `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table -Size 9
					
					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 300;
					$Table.Columns.Item(2).Width = 200;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
				}
				ElseIf($Text)
				{
					Line 0 ""
				}
				ElseIf($HTML)
				{
					If($rowdata.count -gt 0)
					{
						$columnHeaders = @(
						'Setting Key',($htmlsilver -bor $htmlbold),
						'Value',($htmlsilver -bor $htmlbold))

						$msg = ""
						$columnWidths = @("400","300")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				$txt = "Unable to retrieve settings"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 $txt
				}
				ElseIf($Text)
				{
					Line 2 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 1 $txt
				}
			}
			$Filter = $Null
			$Settings = $Null
			Write-Verbose "$(Get-Date): `t`tFinished $($Policy.PolicyName)"
			Write-Verbose "$(Get-Date): "
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Citrix Policy information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results Returned for Citrix Policy information"
	}
	
	$Policies = $Null
	Write-Verbose "$(Get-Date): `tRemoving $($xDriveName) PSDrive"
	Remove-PSDrive $xDriveName -EA 0 4>$Null
	Write-Verbose "$(Get-Date): "
}

Function OutputPolicySetting
{
	Param([string] $outputText, [string] $outputData)

	$xLength = $outputText.Length
	If($outputText.Substring($xLength-2,2) -ne ": ")
	{
		$outputText += ": "
	}
	If($Text)
	{
		Line 2 $outputText $outputData
	}
}

Function Get-PrinterModifiedSettings
{
	Param([string]$Value, [string]$xelement)
	
	[string]$ReturnStr = ""

	Switch ($Value)
	{
		"copi" 
		{
			$txt="Copies: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"coll"
		{
			$txt="Collate: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"scal"
		{
			$txt="Scale (%): "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"colo"
		{
			$txt="Color: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Monochrome"; Break}
					2 {$tmp2 = "Color"; Break}
					Default {$tmp2 = "Color could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"prin"
		{
			$txt="Print Quality: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					-1 {$tmp2 = "150 dpi"; Break}
					-2 {$tmp2 = "300 dpi"; Break}
					-3 {$tmp2 = "600 dpi"; Break}
					-4 {$tmp2 = "1200 dpi"; Break}
					Default {$tmp2 = "Custom...X resolution: $tmp1"; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"yres"
		{
			$txt="Y resolution: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"orie"
		{
			$txt="Orientation: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					"portrait"  {$tmp2 = "Portrait"; Break}
					"landscape" {$tmp2 = "Landscape"; Break}
					Default {$tmp2 = "Orientation could not be determined: $($xelement) ; Break"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"dupl"
		{
			$txt="Duplex: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Simplex"; Break}
					2 {$tmp2 = "Vertical"; Break}
					3 {$tmp2 = "Horizontal"; Break}
					Default {$tmp2 = "Duplex could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"pape"
		{
			$txt="Paper Size: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1   {$tmp2 = "Letter"; Break}
					2   {$tmp2 = "Letter Small"; Break}
					3   {$tmp2 = "Tabloid"; Break}
					4   {$tmp2 = "Ledger"; Break}
					5   {$tmp2 = "Legal"; Break}
					6   {$tmp2 = "Statement"; Break}
					7   {$tmp2 = "Executive"; Break}
					8   {$tmp2 = "A3"; Break}
					9   {$tmp2 = "A4"; Break}
					10  {$tmp2 = "A4 Small"; Break}
					11  {$tmp2 = "A5"; Break}
					12  {$tmp2 = "B4 (JIS)"; Break}
					13  {$tmp2 = "B5 (JIS)"; Break}
					14  {$tmp2 = "Folio"; Break}
					15  {$tmp2 = "Quarto"; Break}
					16  {$tmp2 = "10X14"; Break}
					17  {$tmp2 = "11X17"; Break}
					18  {$tmp2 = "Note"; Break}
					19  {$tmp2 = "Envelope #9"; Break}
					20  {$tmp2 = "Envelope #10"; Break}
					21  {$tmp2 = "Envelope #11"; Break}
					22  {$tmp2 = "Envelope #12"; Break}
					23  {$tmp2 = "Envelope #14"; Break}
					24  {$tmp2 = "C Size Sheet"; Break}
					25  {$tmp2 = "D Size Sheet"; Break}
					26  {$tmp2 = "E Size Sheet"; Break}
					27  {$tmp2 = "Envelope DL"; Break}
					28  {$tmp2 = "Envelope C5"; Break}
					29  {$tmp2 = "Envelope C3"; Break}
					30  {$tmp2 = "Envelope C4"; Break}
					31  {$tmp2 = "Envelope C6"; Break}
					32  {$tmp2 = "Envelope C65"; Break}
					33  {$tmp2 = "Envelope B4"; Break}
					34  {$tmp2 = "Envelope B5"; Break}
					35  {$tmp2 = "Envelope B6"; Break}
					36  {$tmp2 = "Envelope Italy"; Break}
					37  {$tmp2 = "Envelope Monarch"; Break}
					38  {$tmp2 = "Envelope Personal"; Break}
					39  {$tmp2 = "US Std Fanfold"; Break}
					40  {$tmp2 = "German Std Fanfold"; Break}
					41  {$tmp2 = "German Legal Fanfold"; Break}
					42  {$tmp2 = "B4 (ISO)"; Break}
					43  {$tmp2 = "Japanese Postcard"; Break}
					44  {$tmp2 = "9X11"; Break}
					45  {$tmp2 = "10X11"; Break}
					46  {$tmp2 = "15X11"; Break}
					47  {$tmp2 = "Envelope Invite"; Break}
					48  {$tmp2 = "Reserved - DO NOT USE"; Break}
					49  {$tmp2 = "Reserved - DO NOT USE"; Break}
					50  {$tmp2 = "Letter Extra"; Break}
					51  {$tmp2 = "Legal Extra"; Break}
					52  {$tmp2 = "Tabloid Extra"; Break}
					53  {$tmp2 = "A4 Extra"; Break}
					54  {$tmp2 = "Letter Transverse"; Break}
					55  {$tmp2 = "A4 Transverse"; Break}
					56  {$tmp2 = "Letter Extra Transverse"; Break}
					57  {$tmp2 = "A Plus"; Break}
					58  {$tmp2 = "B Plus"; Break}
					59  {$tmp2 = "Letter Plus"; Break}
					60  {$tmp2 = "A4 Plus"; Break}
					61  {$tmp2 = "A5 Transverse"; Break}
					62  {$tmp2 = "B5 (JIS) Transverse"; Break}
					63  {$tmp2 = "A3 Extra"; Break}
					64  {$tmp2 = "A5 Extra"; Break}
					65  {$tmp2 = "B5 (ISO) Extra"; Break}
					66  {$tmp2 = "A2"; Break}
					67  {$tmp2 = "A3 Transverse"; Break}
					68  {$tmp2 = "A3 Extra Transverse"; Break}
					69  {$tmp2 = "Japanese Double Postcard"; Break}
					70  {$tmp2 = "A6"; Break}
					71  {$tmp2 = "Japanese Envelope Kaku #2"; Break}
					72  {$tmp2 = "Japanese Envelope Kaku #3"; Break}
					73  {$tmp2 = "Japanese Envelope Chou #3"; Break}
					74  {$tmp2 = "Japanese Envelope Chou #4"; Break}
					75  {$tmp2 = "Letter Rotated"; Break}
					76  {$tmp2 = "A3 Rotated"; Break}
					77  {$tmp2 = "A4 Rotated"; Break}
					78  {$tmp2 = "A5 Rotated"; Break}
					79  {$tmp2 = "B4 (JIS) Rotated"; Break}
					80  {$tmp2 = "B5 (JIS) Rotated"; Break}
					81  {$tmp2 = "Japanese Postcard Rotated"; Break}
					82  {$tmp2 = "Double Japanese Postcard Rotated"; Break}
					83  {$tmp2 = "A6 Rotated"; Break}
					84  {$tmp2 = "Japanese Envelope Kaku #2 Rotated"; Break}
					85  {$tmp2 = "Japanese Envelope Kaku #3 Rotated"; Break}
					86  {$tmp2 = "Japanese Envelope Chou #3 Rotated"; Break}
					87  {$tmp2 = "Japanese Envelope Chou #4 Rotated"; Break}
					88  {$tmp2 = "B6 (JIS)"; Break}
					89  {$tmp2 = "B6 (JIS) Rotated"; Break}
					90  {$tmp2 = "12X11"; Break}
					91  {$tmp2 = "Japanese Envelope You #4"; Break}
					92  {$tmp2 = "Japanese Envelope You #4 Rotated"; Break}
					93  {$tmp2 = "PRC 16K"; Break}
					94  {$tmp2 = "PRC 32K"; Break}
					95  {$tmp2 = "PRC 32K(Big)"; Break}
					96  {$tmp2 = "PRC Envelope #1"; Break}
					97  {$tmp2 = "PRC Envelope #2"; Break}
					98  {$tmp2 = "PRC Envelope #3"; Break}
					99  {$tmp2 = "PRC Envelope #4"; Break}
					100 {$tmp2 = "PRC Envelope #5"; Break}
					101 {$tmp2 = "PRC Envelope #6"; Break}
					102 {$tmp2 = "PRC Envelope #7"; Break}
					103 {$tmp2 = "PRC Envelope #8"; Break}
					104 {$tmp2 = "PRC Envelope #9"; Break}
					105 {$tmp2 = "PRC Envelope #10"; Break}
					106 {$tmp2 = "PRC 16K Rotated"; Break}
					107 {$tmp2 = "PRC 32K Rotated"; Break}
					108 {$tmp2 = "PRC 32K(Big) Rotated"; Break}
					109 {$tmp2 = "PRC Envelope #1 Rotated"; Break}
					110 {$tmp2 = "PRC Envelope #2 Rotated"; Break}
					111 {$tmp2 = "PRC Envelope #3 Rotated"; Break}
					112 {$tmp2 = "PRC Envelope #4 Rotated"; Break}
					113 {$tmp2 = "PRC Envelope #5 Rotated"; Break}
					114 {$tmp2 = "PRC Envelope #6 Rotated"; Break}
					115 {$tmp2 = "PRC Envelope #7 Rotated"; Break}
					116 {$tmp2 = "PRC Envelope #8 Rotated"; Break}
					117 {$tmp2 = "PRC Envelope #9 Rotated"; Break}
					Default {$tmp2 = "Paper Size could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"form"
		{
			$txt="Form Name: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		"true"
		{
			$txt="TrueType: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Bitmap"; Break}
					2 {$tmp2 = "Download"; Break}
					3 {$tmp2 = "Substitute"; Break}
					4 {$tmp2 = "Outline"; Break}
					Default {$tmp2 = "TrueType could not be determined: $($xelement) "; Break}
				}
			}
			$ReturnStr = "$txt $tmp2"
		}
		"mode" 
		{
			$txt="Printer Model: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"loca" 
		{
			$txt="Location: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		Default {$ReturnStr = "Session printer setting could not be determined: $($xelement) "}
	}
	Return $ReturnStr
}

Function GetCtxGPOsInAD
{
	#thanks to the Citrix Engineering Team for pointers and for Michael B. Smith for creating the function
	#updated 07-Nov-13 to work in a Windows Workgroup environment
	Write-Verbose "$(Get-Date): Testing for an Active Directory environment"
	$root = [ADSI]"LDAP://RootDSE"
	If([String]::IsNullOrEmpty($root.PSBase.Name))
	{
		Write-Verbose "$(Get-Date): `tNot in an Active Directory environment"
		$root = $Null
		$xArray = @()
	}
	Else
	{
		Write-Verbose "$(Get-Date): `tIn an Active Directory environment"
		$domainNC = $root.defaultNamingContext.ToString()
		$root = $Null
		$xArray = @()

		$domain = $domainNC.Replace( 'DC=', '' ).Replace( ',', '.' )
		Write-Verbose "$(Get-Date): `tSearching \\$($domain)\sysvol\$($domain)\Policies"
		$sysvolFiles = @()
		$sysvolFiles = dir -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		If($sysvolFiles.Count -eq 0)
		{
			Write-Verbose "$(Get-Date): `tSearch timed out.  Retrying.  Searching \\ + $($domain)\sysvol\$($domain)\Policies a second time."
			$sysvolFiles = dir -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		}
		ForEach( $file in $sysvolFiles )
		{
			If( -not $file.PSIsContainer )
			{
				#$file.FullName  ### name of the policy file
				If( $file.FullName -like "*\Citrix\GroupPolicy\Policies.gpf" )
				{
					#"have match " + $file.FullName ### name of the Citrix policies file
					$array = $file.FullName.Split( '\' )
					If( $array.Length -gt 7 )
					{
						$gp = $array[ 6 ].ToString()
						$gpObject = [ADSI]( "LDAP://" + "CN=" + $gp + ",CN=Policies,CN=System," + $domainNC )
						If(!$xArray.Contains($gpObject.DisplayName))
						{
							$xArray += $gpObject.DisplayName	### name of the group policy object
						}
					}
				}
			}
		}
	}
	Return ,$xArray
}
#endregion

#region Configuration Logging functions
Function ProcessConfigLogging
{
	#do not show config logging if not Details AND
	# if XenDesktop must be Platinum or Enterprise OR
	# if XenApp must be Platinum or Enterprise

	If($Logging)
	{
		If(($Script:XDSite2.ProductCode -eq "XDT" -and ($Script:XDSite2.ProductEdition -eq "PLT" -or $Script:XDSite2.ProductEdition -eq "ENT")) -or `
		($Script:XDSite2.ProductCode -eq "MPS" -and ($Script:XDSite2.ProductEdition -eq "PLT" -or $Script:XDSite2.ProductEdition -eq "ENT")))
		{
			Write-Verbose "$(Get-Date): Processing Configuration Logging"
			$txt1 = "Logging"
			If($MSword -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 1 0 $txt1
			}
			ElseIf($Text)
			{
				Line 0 $txt1
				Line 0 "For date range $($StartDate) through $($EndDate)"
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 1 0 $txt1
			}
			
			#preferences
			Write-Verbose "$(Get-Date): `tConfiguration Logging Preferences"
			$results = Get-XDLogging @XDParams1
			
			If($? -and $Null -ne $results)
			{
				OutputConfigLogPreferences $results
			}
			Else
			{
				$txt = "Configuration Logging Preferences could not be retrieved."
				OutputWarning $txt
			}
			
			Write-Verbose "$(Get-Date): `tConfiguration Logging Details"
			$ConfigLogItems = Get-LogHighLevelOperation @XDParams2 -Filter {StartTime -ge $StartDate -and EndTime -le $EndDate} -SortBy "-StartTime"
			If($? -and $Null -ne $ConfigLogItems)
			{
				OutputConfigLog $ConfigLogItems
			}
			ElseIf($? -and ($Null -eq $ConfigLogItems))
			{
				$txt = "There are no Configuration Logging actions recorded for $($StartDate) through $($EndDate)."
				OutputWarning $txt
			}
			Else
			{
				$txt = "Configuration Logging information could not be retrieved."
				OutputWarning $txt
			}
			Write-Verbose "$(Get-Date): "
		}
		Else
		{
			$txt = "Not licensed for Configuration Logging"
			OutputWarning $txt
			Write-Verbose "$(Get-Date): "
		}
	}
}

Function OutputConfigLogPreferences 
{
	#2-Mar-2017 Fix bug reported by P. Ewing

	Param([object] $Preferences)
	
	Write-Verbose "$(Get-Date): `t`tOutput Configuration Logging Preferences"
	
	$LogSQLServerPrincipalName = ""
	$LogSQLServerMirrorName = ""
	$LogDatabaseName = ""
	$LogDBs = Get-LogDataStore @XDParams1

	If($? -and ($Null -ne $LogDBs))
	{
		ForEach($LogDB in $LogDBs)
		{
			If($LogDB.DataStore -eq "Logging")
			{
				$tmp = $LogDB.ConnectionString
				$csitems = $tmp.Split(';')
				ForEach($csitem in $csitems)
				{
					$Pair = $csitem.split('=').trimstart()
					Switch ($Pair[0])
					{
						"Server"					{$LogSQLServerPrincipalName = $Pair[1]}
						"Failover Partner"			{$LogSQLServerMirrorName = $Pair[1]}
						"MultiSubnetFailover"		{$LogSQLServerMirrorName = ""}
						"Database"					{$LogDatabaseName = $Pair[1]}
						{$Pair[0] -match "Initial"}	{$LogDatabaseName = $Pair[1]}
					}
				}
			}
		}
	}
	Else
	{
		Write-Warning "Unable to retrieve Configuration Logging Database settings"
	}

	#get database size
	
	$SQLsrv = new-Object Microsoft.SqlServer.Management.Smo.Server("$($LogSQLServerPrincipalName)")
	$db = New-Object Microsoft.SqlServer.Management.Smo.Database
	$db = $SQLsrv.Databases.Item("$($LogDatabaseName)")
	[string]$dbsize = "{0:F2} MB" -f $db.size
	
	If($Preferences.Enabled -eq "Enabled" -or $Preferences.Enabled -eq "Mandatory")
	{
		$PrefEnabled = "Enabled"
	}
	Else
	{
		$PrefEnabled = "Disabled"
	}
	
	If($Preferences.AllowDisconnectedDatabase -eq $True)
	{
		$PrefSecurity = "Selected"
	}
	Else
	{
		$PrefSecurity = "Not Selected"
	}
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Configuration Logging settings"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = $PrefEnabled; Value = ""; }
		$ScriptInformation += @{Data = "   Logging database"; Value = ""; }
		$ScriptInformation += @{Data = "      Database size"; Value = $dbsize; }
		$ScriptInformation += @{Data = "      Server location"; Value = $LogSQLServerPrincipalName; }
		$ScriptInformation += @{Data = "      Database name"; Value = $LogDatabaseName; }
		$ScriptInformation += @{Data = "   Security"; Value = ""; }
		$ScriptInformation += @{Data = "      Allow changes when the database is disconnected"; Value = $PrefSecurity; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 "Configuration Logging settings"
		Line 1 $PrefEnabled
		Line 2 "Logging database"
		Line 3 "Database size`t: " $dbsize
		Line 3 "Server location`t: " $LogSQLServerPrincipalName
		Line 3 "Database name`t: " $LogDatabaseName
		Line 3 "Security"
		Line 4 "Allow changes when the database is disconnected: " $PrefSecurity
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Configuration Logging settings"
		$rowdata = @()
		$columnHeaders = @($PrefEnabled,($htmlsilver -bor $htmlbold),"",$htmlwhite)
		$rowdata += @(,('   Logging database',($htmlsilver -bor $htmlbold),"",$htmlwhite))
		$rowdata += @(,('      Database size',($htmlsilver -bor $htmlbold),$dbsize,$htmlwhite))
		$rowdata += @(,('      Server location',($htmlsilver -bor $htmlbold),$LogSQLServerPrincipalName,$htmlwhite))
		$rowdata += @(,('      Database name',($htmlsilver -bor $htmlbold),$LogDatabaseName,$htmlwhite))
		$rowdata += @(,('   Security',($htmlsilver -bor $htmlbold),"",$htmlwhite))
		$rowdata += @(,('      Allow changes when the database is disconnected',($htmlsilver -bor $htmlbold),$PrefSecurity,$htmlwhite))
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputConfigLog
{
	Param([object] $ConfigLogItems)
	
	Write-Verbose "$(Get-Date): `t`tOutput Configuration Logging Details"
	$txt2 = " For date range $($StartDate) through $($EndDate)"
	If($MSword -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 0 0 $txt2
	}
	ElseIf($Text)
	{
		Line 0 "For date range $($StartDate) through $($EndDate)"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 $txt2
	}
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}
	
	ForEach($Item in $ConfigLogItems)
	{
		$Tmp = $Null
		If($Item.IsSuccessful)
		{
			$Tmp = "Success"
		}
		Else
		{
			$Tmp = "Failed"
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Administrator = $Item.User;
			MainTask = $Item.Text;
			Start = $Item.StartTime;
			End = $Item.EndTime;
			Status = $tmp;
			}

			$ItemsWordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "Administrator`t: " $Item.User
			Line 1 "Main task`t: " $Item.Text
			Line 1 "Start`t`t: " $Item.StartTime
			Line 1 "End`t`t: " $Item.EndTime
			Line 1 "Status`t`t: " $Tmp
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Item.User,$htmlwhite,
			$Item.Text,$htmlwhite,
			$Item.StartTime,$htmlwhite,
			$Item.EndTime,$htmlwhite,
			$Tmp,$htmlwhite
			))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns Administrator, MainTask, Start, End, Status `
		-Headers  "Administrator", "Main task", "Start", "End", "Status" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 120;
		$Table.Columns.Item(2).Width = 210;
		$Table.Columns.Item(3).Width = 60;
		$Table.Columns.Item(4).Width = 60;
		$Table.Columns.Item(5).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Administrator',($htmlsilver -bor $htmlbold),
		'Main task',($htmlsilver -bor $htmlbold),
		'Start',($htmlsilver -bor $htmlbold),
		'End',($htmlsilver -bor $htmlbold),
		'Status',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("120","210","60","60","50")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region site configuration functions
Function ProcessConfiguration
{
	Write-Verbose "$(Get-Date): Process Configuration Settings"
	OutputSiteSettings
	OutputCEIPSetting
	OutputDatastores
	Write-Verbose "$(Get-Date): "
}

Function OutputSiteSettings
{
	Write-Verbose "$(Get-Date): `tSee if StoreFront is installed on the Controller(s)"
	$DefaultStoreFrontAddress = ""
	If(Get-SFIsStoreFrontInstalled @XDParams1)
	{
		$tmp = Get-SFCluster @XDParams1
		If($? -and ($Null -ne $tmp))
		{
			$DefaultStoreFrontAddress = $tmp.Url
		}
		Else
		{
			Write-Warning "Unable to retrieve StoreFront Cluster information"
		}	
	}

	Write-Verbose "$(Get-Date): `tOutput Site Settings"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Configuration"
		WriteWordLine 2 0 "Site Settings"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Site name"; Value = $XDSiteName; }
		$ScriptInformation += @{Data = "Default StoreFront address"; Value = $DefaultStoreFrontAddress; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Configuration"
		Line 0 ""
		Line 0 "Site Settings"
		Line 0 ""
		Line 1 "Site name: " $XDSiteName
		Line 1 "Default StoreFront address: " $DefaultStoreFrontAddress
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Configuration"
		WriteHTMLLine 2 0 "Site Settings"
		$rowdata = @()
		$columnHeaders = @("Site name",($htmlsilver -bor $htmlbold),$XDSiteName,$htmlwhite)
		$rowdata += @(,('Default StoreFront address',($htmlsilver -bor $htmlbold),$DefaultStoreFrontAddress,$htmlwhite))
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputCEIPSetting
{
	Write-Verbose "$(Get-Date): `tProcessing Customer Experience Improvement Program"
	$Results = Get-AnalyticsSite @XDParams1
	If($? -and ($Null -ne $Results))
	{
		If($Results.Enabled)
		{
			$CEIP = "You are currently participating in the Customer Experience Improvement Program"
		}
		Else
		{
			$CEIP = "You are not currently participating in the Customer Experience Improvement Program"
		}
	}
	Else
	{
		Write-Warning "Unable to retrieve CEIP information"
	}	

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Product Support"
		[System.Collections.Hashtable[]] $CEIPWordTable = @();
		$WordTableRowHash = @{ 
		CEIP = $CEIP;
		}
		$CEIPWordTable += $WordTableRowHash;
		
		$Table = AddWordTable -Hashtable $CEIPWordTable `
		-Columns CEIP `
		-Headers "Citrix Customer Experience Improvement Program" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null

		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Product Support"
		Line 0 ""
		Line 1 "Citrix Customer Experience Improvement Program: " $CEIP
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Product Support"
		$rowdata = @()
		$rowdata += @(,(
		$CEIP,$htmlwhite))

		$columnHeaders = @(
		'Citrix Customer Experience Improvement Program',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputDatastores
{
	#2-Mar-2017 Fix bug reported by P. Ewing

	#line starts with server=SQLServerName;
	#only need what is between the = and ;
	Write-Verbose "$(Get-Date): `tRetrieving database connection data"
	Write-Verbose "$(Get-Date): `t`tConfiguration database"
	$ConfigSQLServerPrincipalName = ""
	$ConfigSQLServerMirrorName = ""
	$ConfigDatabaseName = ""
	$ConfigDB = Get-ConfigDBConnection @XDParams1

	If($? -and ($Null -ne $ConfigDB))
	{
		$tmp = $ConfigDB
		$csitems = $tmp.Split(';')
		ForEach($csitem in $csitems)
		{
			$Pair = $csitem.split('=').trimstart()
			Switch ($Pair[0])
			{
				"Server"					{$ConfigSQLServerPrincipalName = $Pair[1]}
				"Failover Partner"			{$ConfigSQLServerMirrorName = $Pair[1]}
				"MultiSubnetFailover"		{$ConfigSQLServerMirrorName = ""}
				"Database"					{$ConfigDatabaseName = $Pair[1]}
				{$Pair[0] -match "Initial"}	{$ConfigDatabaseName = $Pair[1]}
			}
		}
	}
	Else
	{
		Write-Warning "Unable to retrieve Configuration Database settings"
	}

	Write-Verbose "$(Get-Date): `t`tConfiguration Logging database"
	$LogSQLServerPrincipalName = ""
	$LogSQLServerMirrorName = ""
	$LogDatabaseName = ""
	$LogDBs = Get-LogDataStore @XDParams1

	If($? -and ($Null -ne $LogDBs))
	{
		ForEach($LogDB in $LogDBs)
		{
			If($LogDB.DataStore -eq "Logging")
			{
				$tmp = $LogDB.ConnectionString
				$csitems = $tmp.Split(';')
				ForEach($csitem in $csitems)
				{
					$Pair = $csitem.split('=').trimstart()
					Switch ($Pair[0])
					{
						"Server"					{$LogSQLServerPrincipalName = $Pair[1]}
						"Failover Partner"			{$LogSQLServerMirrorName = $Pair[1]}
						"MultiSubnetFailover"		{$LogSQLServerMirrorName = ""}
						"Database"					{$LogDatabaseName = $Pair[1]}
						{$Pair[0] -match "Initial"}	{$LogDatabaseName = $Pair[1]}
					}
				}
			}
		}
	}
	Else
	{
		Write-Warning "Unable to retrieve Configuration Logging Database settings"
	}

	Write-Verbose "$(Get-Date): `t`tMonitoring database"
	$MonitorSQLServerPrincipalName = ""
	$MonitorSQLServerMirrorName = ""
	$MonitorDatabaseName = ""
	$MonitorCollectHotfix = "Disabled"
	$MonitorDataCollection = "Disabled"
	$MonitorDetailedSQL = "Disabled"
	
	$MonitorDBs = Get-MonitorDataStore @XDParams1

	If($? -and ($Null -ne $MonitorDBs))
	{
		ForEach($MonitorDB in $MonitorDBs)
		{
			If($MonitorDB.DataStore -eq "Monitor")
			{
				$tmp = $MonitorDB.ConnectionString
				$csitems = $tmp.Split(';')
				ForEach($csitem in $csitems)
				{
					$Pair = $csitem.split('=').trimstart()
					Switch ($Pair[0])
					{
						"Server"					{$MonitorSQLServerPrincipalName = $Pair[1]}
						"Failover Partner"			{$MonitorSQLServerMirrorName = $Pair[1]}
						"MultiSubnetFailover"		{$MonitorSQLServerMirrorName = ""}
						"Database"					{$MonitorDatabaseName = $Pair[1]}
						{$Pair[0] -match "Initial"}	{$MonitorDatabaseName = $Pair[1]}
					}
				}
			}
		}
		
		$MonitorConfig = $Null
		$MonitorConfig = Get-MonitorConfiguration @XDParams1
		
		If($? -and $Null -ne $MonitorConfig)
		{
			If($MonitorConfig.CollectHotfixDataEnabled)
			{
				$MonitorCollectHotfix = "Enabled"
			}
			Else
			{
				$MonitorCollectHotfix = "Disabled"
			}

			If($MonitorConfig.DataCollectionEnabled)
			{
				$MonitorDataCollection = "Enabled"
			}
			Else
			{
				$MonitorDataCollection = "Disabled"
			}

			If($MonitorConfig.DetailedSqlOutputEnabled)
			{
				$MonitorDetailedSQL = "Enabled"
			}
			Else
			{
				$MonitorDetailedSQL = "Disabled"
			}
		}
	}
	Else
	{
		Write-Warning "Unable to retrieve Monitoring Database settings"
	}

	Write-Verbose "$(Get-Date): `tOutput Datastores"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Datastores"
		[System.Collections.Hashtable[]] $DBsWordTable = @();
		$WordTableRowHash = @{ 
		DataStore = "Site";
		DatabaseName = $ConfigDatabaseName;
		ServerAddress = $ConfigSQLServerPrincipalName;
		MirrorServerAddress = $ConfigSQLServerMirrorName;
		}
		$DBsWordTable += $WordTableRowHash;

		$WordTableRowHash = @{ 
		DataStore = "Logging";
		DatabaseName = $LogDatabaseName;
		ServerAddress = $LogSQLServerPrincipalName;
		MirrorServerAddress = $LogSQLServerMirrorName;
		}
		$DBsWordTable += $WordTableRowHash;

		$WordTableRowHash = @{ 
		DataStore = "Monitoring";
		DatabaseName = $MonitorDatabaseName;
		ServerAddress = $MonitorSQLServerPrincipalName;
		MirrorServerAddress = $MonitorSQLServerMirrorName;
		}
		$DBsWordTable += $WordTableRowHash;

		$Table = AddWordTable -Hashtable $DBsWordTable `
		-Columns DataStore, DatabaseName, ServerAddress, MirrorServerAddress `
		-Headers "Datastore", "Database Name", "Server Address", "Mirror Server Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null

		WriteWordLine 3 0 "Monitoring Database Details"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Collect Hotfix Data"; Value = $MonitorCollectHotfix; }
		$ScriptInformation += @{Data = "Data Collection"; Value = $MonitorDataCollection; }
		$ScriptInformation += @{Data = "Detail SQL Output"; Value = $MonitorDetailedSQL; }
		$ScriptInformation += @{Data = "Full Poll Start Hour"; Value = $MonitorConfig.FullPollStartHour; }
		$ScriptInformation += @{Data = "Resolution Poll Time Hours"; Value = $MonitorConfig.FullPollStartHour; }
		$ScriptInformation += @{Data = "Sync Poll Time Hours"; Value = $MonitorConfig.SyncPollTimeHours; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 3 0 "Groom Retention Settings in Days"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Application Instance"; Value = $MonitorConfig.GroomApplicationInstanceRetentionDays; }
		$ScriptInformation += @{Data = "Deleted"; Value = $MonitorConfig.GroomDeletedRetentionDays; }
		$ScriptInformation += @{Data = "Failures"; Value = $MonitorConfig.GroomFailuresRetentionDays; }
		$ScriptInformation += @{Data = "Load Indexes"; Value = $MonitorConfig.GroomLoadIndexesRetentionDays; }
		$ScriptInformation += @{Data = "Machine Hotfix Log"; Value = $MonitorConfig.GroomMachineHotfixLogRetentionDays; }
		$ScriptInformation += @{Data = "Minute"; Value = $MonitorConfig.GroomMinuteRetentionDays; }
		$ScriptInformation += @{Data = "Sessions"; Value = $MonitorConfig.GroomSessionsRetentionDays; }
		$ScriptInformation += @{Data = "Summaries"; Value = $MonitorConfig.GroomSummariesRetentionDays; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Datastores"
		Line 0 ""
		Line 1 "Datastore`t`t: Site"
		Line 1 "Database Name`t`t: " $ConfigDatabaseName
		Line 1 "Server Address`t`t: " $ConfigSQLServerPrincipalName
		Line 1 "Mirror Server Address`t: " $ConfigSQLServerMirrorName
		Line 0 ""
		Line 1 "Datastore`t`t: Logging"
		Line 1 "Database Name`t`t: " $LogDatabaseName
		Line 1 "Server Address`t`t: " $LogSQLServerPrincipalName
		Line 1 "Mirror Server Address`t: " $LogSQLServerMirrorName
		Line 0 ""
		Line 1 "Datastore`t`t: Monitoring"
		Line 1 "Database Name`t`t: " $MonitorDatabaseName
		Line 1 "Server Address`t`t: " $MonitorSQLServerPrincipalName
		Line 1 "Mirror Server Address`t: " $MonitorSQLServerMirrorName
		Line 0 ""
		Line 1 "Monitoring Database Details"
		Line 1 "Collect Hotfix Data`t`t: " $MonitorCollectHotfix
		Line 1 "Data Collection`t`t`t: " $MonitorDataCollection
		Line 1 "Detail SQL Output`t`t: " $MonitorDetailedSQL
		Line 1 "Full Poll Start Hour`t`t: " $MonitorConfig.FullPollStartHour
		Line 1 "Resolution Poll Time Hours`t: " $MonitorConfig.FullPollStartHour
		Line 1 "Sync Poll Time Hours`t`t: " $MonitorConfig.SyncPollTimeHours
		Line 1 "Groom Retention Settings in Days" 
		Line 2 "Application Instance`t: " $MonitorConfig.GroomApplicationInstanceRetentionDays
		Line 2 "Deleted`t`t`t: " $MonitorConfig.GroomDeletedRetentionDays
		Line 2 "Failures`t`t: " $MonitorConfig.GroomFailuresRetentionDays
		Line 2 "Load Indexes`t`t: " $MonitorConfig.GroomLoadIndexesRetentionDays 
		Line 2 "Machine Hotfix Log`t: " $MonitorConfig.GroomMachineHotfixLogRetentionDays
		Line 2 "Minute`t`t`t: " $MonitorConfig.GroomMinuteRetentionDays
		Line 2 "Sessions`t`t: " $MonitorConfig.GroomSessionsRetentionDays
		Line 2 "Summaries`t`t: " $MonitorConfig.GroomSummariesRetentionDays
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Datastores"
		
		$rowdata = @()

		$rowdata += @(,(
		'Site',$htmlwhite,
		$ConfigDatabaseName,$htmlwhite,
		$ConfigSQLServerPrincipalName,$htmlwhite,
		$ConfigSQLServerMirrorName,$htmlwhite))

		$rowdata += @(,(
		'Logging',$htmlwhite,
		$LogDatabaseName,$htmlwhite,
		$LogSQLServerPrincipalName,$htmlwhite,
		$LogSQLServerMirrorName,$htmlwhite))

		$rowdata += @(,(
		'Monitoring',$htmlwhite,
		$MonitorDatabaseName,$htmlwhite,
		$MonitorSQLServerPrincipalName,$htmlwhite,
		$MonitorSQLServerMirrorName,$htmlwhite))

		$columnHeaders = @(
		'Datastore',($htmlsilver -bor $htmlbold),
		'Database Name',($htmlsilver -bor $htmlbold),
		'Server Address',($htmlsilver -bor $htmlbold),
		'Mirror Server Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "

		WriteHTMLLine 3 0 "Monitoring Database Details"
		$rowdata = @()
		$columnHeaders = @("Collect Hotfix Data",($htmlsilver -bor $htmlbold),$MonitorCollectHotfix,$htmlwhite)
		$rowdata += @(,('Data Collection',($htmlsilver -bor $htmlbold),$MonitorDataCollection,$htmlwhite))
		$rowdata += @(,('Detail SQL Output',($htmlsilver -bor $htmlbold),$MonitorDetailedSQL,$htmlwhite))
		$rowdata += @(,('Full Poll Start Hour',($htmlsilver -bor $htmlbold),$MonitorConfig.FullPollStartHour,$htmlwhite))
		$rowdata += @(,('Resolution Poll Time Hours',($htmlsilver -bor $htmlbold),$MonitorConfig.FullPollStartHour,$htmlwhite))
		$rowdata += @(,('Sync Poll Time Hours',($htmlsilver -bor $htmlbold),$MonitorConfig.SyncPollTimeHours,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
		
		WriteHTMLLine 3 0 "Groom Retention Settings in Days"
		$rowdata = @()
		$columnHeaders = @("Application Instance",($htmlsilver -bor $htmlbold),$MonitorConfig.GroomApplicationInstanceRetentionDays,$htmlwhite)
		$rowdata += @(,('Deleted',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomDeletedRetentionDays,$htmlwhite))
		$rowdata += @(,('Failures',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomFailuresRetentionDays,$htmlwhite))
		$rowdata += @(,('Load Indexes',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomLoadIndexesRetentionDays,$htmlwhite))
		$rowdata += @(,('Machine Hotfix Log',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomMachineHotfixLogRetentionDays,$htmlwhite))
		$rowdata += @(,('Minute',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomMinuteRetentionDays,$htmlwhite))
		$rowdata += @(,('Sessions',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomSessionsRetentionDays,$htmlwhite))
		$rowdata += @(,('Summaries',($htmlsilver -bor $htmlbold),$MonitorConfig.GroomSummariesRetentionDays,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region Administrator, Scope and Roles functions
Function ProcessAdministrators
{
	Write-Verbose "$(Get-Date): Processing Administrators"
	Write-Verbose "$(Get-Date): `tRetrieving Administrator data"
	
	$Global:TotalDeliveryGroupAdmins = 0
	$Global:TotalFullAdmins = 0
	$Global:TotalHelpDeskAdmins = 0
	$Global:TotalHostAdmins = 0
	$Global:TotalMachineCatalogAdmins = 0
	$Global:TotalReadOnlyAdmins = 0
	$Global:TotalCustomAdmins = 0
	
	$Admins = Get-AdminAdministrator @XDParams2 | Sort Name

	If($? -and ($Null -ne $Admins))
	{
		OutputAdministrators $Admins
	}
	ElseIf($? -and ($Null -eq $Admins))
	{
		$txt = "There are no Administrators"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Administrators"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputAdministrators
{
	Param([object] $Admins)
	
	Write-Verbose "$(Get-Date): `tOutput Administrator data"
	
	ForEach($Admin in $Admins)
	{
		Switch ($Admin.Rights.RoleName)
		{
			"Delivery Group Administrator"	{$Global:TotalDeliveryGroupAdmins++}
			"Full Administrator"			{$Global:TotalFullAdmins++}
			"Help Desk Administrator"		{$Global:TotalHelpDeskAdmins++}
			"Host Administrator"			{$Global:TotalHostAdmins++}
			"Machine Catalog Administrator"	{$Global:TotalMachineCatalogAdmins++}
			"Read Only Administrator"		{$Global:TotalReadOnlyAdmins++}
			Default							{$Global:TotalCustomAdmins++}
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Administrators"
		[System.Collections.Hashtable[]] $AdminsWordTable = @();
		ForEach($Admin in $Admins)
		{
			$Tmp = $Null
			If($Admin.Enabled)
			{
				$Tmp = "Enabled"
			}
			Else
			{
				$Tmp = "Disabled"
			}
			$WordTableRowHash = @{Name = $Admin.Name; Scope = $Admin.Rights.ScopeName; Role = $Admin.Rights.RoleName; Status = $Tmp;}

			$AdminsWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $AdminsWordTable `
		-Columns Name, Scope, Role, Status `
		-Headers "Name", "Scope", "Role", "Status" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Administrators"
		Line 0 ""
		ForEach($Admin in $Admins)
		{
			Line 1 "Name`t: " $Admin.Name
			Line 1 "Scope`t: " $Admin.Rights.ScopeName
			Line 1 "Role`t: " $Admin.Rights.RoleName
			Line 1 "Status`t: " -NoNewLine
			If($Admin.Enabled)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 1 0 "Administrators"
		ForEach($Admin in $Admins)
		{
			$xType = ""
			If($Admin.Enabled)
			{
				$xType = "Enabled"
			}
			Else
			{
				$xType = "Disabled"
			}
			$rowdata += @(,(
			$Admin.Name,$htmlwhite,
			$Admin.Rights.ScopeName,$htmlwhite,
			$Admin.Rights.RoleName,$htmlwhite,
			$xType,$htmlwhite))
		}
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Scope',($htmlsilver -bor $htmlbold),
		'Role',($htmlsilver -bor $htmlbold),
		'Status',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function ProcessScopes
{
	Write-Verbose "$(Get-Date): Processing Administrator Scopes"
	$Scopes = Get-AdminScope @XDParams2 -SortBy Name
	
	If($? -and ($Null -ne $Scopes))
	{
		OutputScopes $Scopes
		If($Administrators)
		{
			OutputScopeObjects $Scopes
			OutputScopeAdministrators $Scopes
		}
	}
	ElseIf($? -and ($Null -eq $Scopes))
	{
		$txt = "There are no Administrator Scopes"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Administrator Scopes"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputScopes
{
	Param([object] $Scopes)
	
	Write-Verbose "$(Get-Date): `tOutput Scopes"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Administrative Scopes"
		[System.Collections.Hashtable[]] $ScopesWordTable = @();
		ForEach($Scope in $Scopes)
		{
			$WordTableRowHash = @{ Name = $Scope.Name; Description = $Scope.Description;}

			$ScopesWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $ScopesWordTable `
		-Columns Name, Description `
		-Headers "Name", "Description" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Administrative Scopes"
		Line 0 ""
		ForEach($Scope in $Scopes)
		{
			Line 1 "Name`t`t: " $Scope.Name
			Line 1 "Description`t: " $Scope.Description
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 2 0 "Administrative Scopes"
		ForEach($Scope in $Scopes)
		{
			$rowdata += @(,(
			$Scope.Name,$htmlwhite,
			$Scope.Description,$htmlwhite))
		}
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}

Function OutputScopeObjects
{
	Param([object] $Scopes)
	
	Write-Verbose "$(Get-Date): `t`tOutput Scope Objects"

	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		
		ForEach($Scope in $Scopes)
		{
			WriteWordLine 3 0 "Scope Objects for $($Scope.Name)"

			$Results = GetScopeDG $Scope
			
			If($Results.Count -gt 0)
			{
				[System.Collections.Hashtable[]] $WordTable = @();

				WriteWordLine 4 0 "Delivery Groups"
				
				ForEach($Result in $Results)
				{
					$WordTableRowHash = @{ 
					GroupName = $Result.Name; 
					GroupDesc = $Result.Description; 
					}

					$WordTable += $WordTableRowHash;
				}
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns GroupName, GroupDesc `
				-Headers "Name", "Description" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}

			$Results = GetScopeMC $Scope

			If($Results.Count -gt 0)
			{
				[System.Collections.Hashtable[]] $WordTable = @();

				WriteWordLine 4 0 "Machine Catalogs"
				ForEach($Result in $Results)
				{
					$WordTableRowHash = @{ 
					CatalogName = $Result.Name; 
					CatalogDesc = $Result.Description; 
					}

					$WordTable += $WordTableRowHash;
				}
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns CatalogName, CatalogDesc `
				-Headers "Name", "Description" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}

			$Results = GetScopeHyp $Scope

			If($Results.Count -gt 0)
			{
				[System.Collections.Hashtable[]] $WordTable = @();

				WriteWordLine 4 0 "Hosting"
				ForEach($Result in $Results)
				{
					$WordTableRowHash = @{ 
					HypName = $Result.Name; 
					HypDesc = $Result.Description; 
					}

					$WordTable += $WordTableRowHash;
				}
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns HypName, HypDesc `
				-Headers "Name", "Description" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}

			WriteWordLine 0 0 ""
		}
	}
	ElseIf($Text)
	{
		ForEach($Scope in $Scopes)
		{
			Line 0 "Scope Objects for $($Scope.Name)"

			$Results = GetScopeDG $Scope
			
			If($Results.Count -gt 0)
			{
				Line 0 "Delivery Groups"
				
				ForEach($Result in $Results)
				{
					Line 1 "Name: " $Result.Name
					Line 1 "Description: " $Result.Description
					Line 0 ""
				}
			}

			$Results = GetScopeMC $Scope

			If($Results.Count -gt 0)
			{
				Line 0 "Machine Catalogs"
				ForEach($Result in $Results)
				{
					Line 1 "Name: " $Result.Name
					Line 1 "Description: " $Result.Description
					Line 0 ""
				}
			}

			$Results = GetScopeHyp $Scope

			If($Results.Count -gt 0)
			{
				Line 0 "Hosting"
				ForEach($Result in $Results)
				{
					Line 1 "Name: " $Result.Name
					Line 1 "Description: " $Result.Description
					Line 0 ""
				}
			}

			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		ForEach($Scope in $Scopes)
		{
			WriteHTMLLine 3 0 "Scope Objects for $($Scope.Name)"

			$Results = GetScopeDG $Scope
			
			If($Results.Count -gt 0)
			{
				WriteHTMLLine 4 0 "Delivery Groups"
				$rowdata = @()

				ForEach($Result in $Results)
				{
					$rowdata += @(,(
					$Result.Name,$htmlwhite,
					$Result.Description,$htmlwhite))
				}
				$msg = ""
				$ColumnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $ColumnWidths -tablewidth "500"
				WriteHTMLLine 0 0 " "
			}

			$Results = GetScopeMC $Scope

			If($Results.Count -gt 0)
			{
				WriteHTMLLine 4 0 "Machine Catalogs"
				$rowdata = @()

				ForEach($Result in $Results)
				{
					$rowdata += @(,(
					$Result.Name,$htmlwhite,
					$Result.Description,$htmlwhite))
				}
				$msg = ""
				$ColumnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $ColumnWidths -tablewidth "500"
				WriteHTMLLine 0 0 " "
			}

			$Results = GetScopeHyp $Scope

			If($Results.Count -gt 0)
			{
				WriteHTMLLine 4 0 "Hosting"
				$rowdata = @()

				ForEach($Result in $Results)
				{
					$rowdata += @(,(
					$Result.Name,$htmlwhite,
					$Result.Description,$htmlwhite))
				}
				$msg = ""
				$ColumnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $ColumnWidths -tablewidth "500"
				WriteHTMLLine 0 0 " "
			}
		}
	}
}

Function GetScopeDG
{
	Param([object] $Scope)
	
	$DG = @()
	#get delivery groups
	If($Scope.Name -eq "All")
	{
		$Results = Get-BrokerDesktopGroup @XDParams2 | `
		Select Name, Description, Scopes | `
		Sort Name -unique
	}
	Else
	{
		$Results = Get-BrokerDesktopGroup @XDParams2 | `
		Select Name, Description, Scopes | `
		? {$_.Scopes -like $Scope.Name} | `
		Sort Name -unique
	}
	
	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name Name        -Value $Result.Name
			$obj | Add-Member -MemberType NoteProperty -Name Description -Value $Result.Description
			
			$DG += $obj
		}
	}

	Return ,$DG
}

Function GetScopeMC
{
	Param([object] $Scope)
	
	#get machine catalogs
	$MC = @()
	
	If($Scope.Name -eq "All")
	{
		$Results = Get-BrokerCatalog @XDParams2 | `
		Select Name, Description, Scopes | `
		Sort Name -unique
	}
	Else
	{
		$Results = Get-BrokerCatalog @XDParams2 | `
		Select Name, Description, Scopes | `
		? {$_.Scopes -like $Scope.Name} | `
		Sort Name -unique
	}

	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name Name        -Value $Result.Name
			$obj | Add-Member -MemberType NoteProperty -Name Description -Value $Result.Description
			
			$MC += $obj
		}
	}

	Return ,$MC
}

Function GetScopeHyp
{
	Param([object] $Scope)
	
	#get hypervisor connections
	$Hyp = @()
	
	If($Scope.Name -eq "All")
	{
		$Results = Get-HypScopedObject @XDParams2 | `
		Select ObjectName, Description, ScopeName | `
		Sort ObjectName -unique
	}
	Else
	{
		$Results = Get-HypScopedObject @XDParams2 | `
		Select ObjectName, Description, ScopeName | `
		? {$_.ScopeName -like $Scope.Name} | `
		Sort ObjectName -unique
	}

	If($? -and $Null -ne $Results)
	{
		ForEach($Result in $Results)
		{
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name Name        -Value $Result.ObjectName
			$obj | Add-Member -MemberType NoteProperty -Name Description -Value $Result.Description
			
			$Hyp += $obj
		}
	}

	Return ,$Hyp
}

Function OutputScopeAdministrators 
{
	Param([object] $Scopes)
	Write-Verbose "$(Get-Date): `t`tOutput Scope Administrators"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		
		ForEach($Scope in $Scopes)
		{
			[System.Collections.Hashtable[]] $WordTable = @();
			WriteWordLine 3 0 "Administrators for Scope: $($Scope.Name)"
			$Admins = Get-AdminAdministrator @XDParams1 | ? {$_.Rights.ScopeName -Contains $Scope.Name}
			
			If($? -and $Null -ne $Admins)
			{
				ForEach($Admin in $Admins)
				{
					$xEnabled = "Disabled"
					If($Admin.Enabled)
					{
						$xEnabled = "Enabled"
					}

					$xRoleName = ""
					ForEach($Right in $Admin.Rights)
					{
						If($Right.ScopeName -eq $Scope.Name -or $Right.ScopeName -eq "All")
						{
							$xRoleName = $Right.RoleName
						}
					}
					
					$WordTableRowHash = @{ 
					AdminName = $Admin.Name; 
					Role = $xRoleName; 
					Type = $xEnabled;
					}

					$WordTable += $WordTableRowHash;
				}
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns AdminName, Role, Type `
				-Headers "Administrator Name", "Role", "Status" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 220;
				$Table.Columns.Item(2).Width = 225;
				$Table.Columns.Item(3).Width = 55;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($? -and $Null -eq $Admins)
			{
				WriteWordLine 0 0 "No administrators defined"
				WriteWordLine 0 0 ""
			}
			Else
			{
				WriteWordLine 0 0 "Unable to retrieve administrators"
				WriteWordLine 0 0 ""
			}
		}
	}
	ElseIf($Text)
	{
		ForEach($Scope in $Scopes)
		{
			Line 1 "Administrators for Scope: $($Scope.Name)"
			$Admins = Get-AdminAdministrator @XDParams1 | ? {$_.Rights.ScopeName -Contains $Scope.Name}
			
			If($? -and $Null -ne $Admins)
			{
				ForEach($Admin in $Admins)
				{
					$xEnabled = "Disabled"
					If($Admin.Enabled)
					{
						$xEnabled = "Enabled"
					}

					$xRoleName = ""
					ForEach($Right in $Admin.Rights)
					{
						If($Right.ScopeName -eq $Scope.Name -or $Right.ScopeName -eq "All")
						{
							$xRoleName = $Right.RoleName
						}
					}
					
					Line 2 "Administrator Name`t: " $Admin.Name
					Line 2 "Role`t`t`t: " $xRoleName
					Line 2 "Status`t`t`t: " $xEnabled
					Line 0 ""
				}
			}
			ElseIf($? -and $Null -eq $Admins)
			{
				Line 2 "No administrators defined"
				Line 0 ""
			}
			Else
			{
				Line 2 "Unable to retrieve administrators"
				Line 0 ""
			}
		}
	}
	ElseIf($HTML)
	{
		ForEach($Scope in $Scopes)
		{
			$rowdata = @()
			WriteHTMLLine 3 0 "Administrators for Scope: $($Scope.Name)"
			$Admins = Get-AdminAdministrator @XDParams1 | ? {$_.Rights.ScopeName -Contains $Scope.Name}
			
			If($? -and $Null -ne $Admins)
			{
				ForEach($Admin in $Admins)
				{
					$xEnabled = "Disabled"
					If($Admin.Enabled)
					{
						$xEnabled = "Enabled"
					}

					$xRoleName = ""
					ForEach($Right in $Admin.Rights)
					{
						If($Right.ScopeName -eq $Scope.Name -or $Right.ScopeName -eq "All")
						{
							$xRoleName = $Right.RoleName
						}
					}
					
					$rowdata += @(,(
					$Admin.Name,$htmlwhite,
					$xRoleName,$htmlwhite,
					$xEnabled,$htmlwhite))
				}
				$columnHeaders = @(
				'Administrator Name',($htmlsilver -bor $htmlbold),
				'Role',($htmlsilver -bor $htmlbold),
				'Status',($htmlsilver -bor $htmlbold))

				$msg = ""
				$columnWidths = @("220","225","55")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 " "
			}
			ElseIf($? -and $Null -eq $Admins)
			{
				WriteHTMLLine 0 0 "No administrators defined"
				WriteHTMLLine 0 0 " "
			}
			Else
			{
				WriteHTMLLine 0 0 "Unable to retrieve administrators"
				WriteHTMLLine 0 0 " "
			}
		}
	}
}

Function ProcessRoles
{
	Write-Verbose "$(Get-Date): Processing Administrator Roles"
	$Roles = Get-AdminRole @XDParams2 -SortBy Name

	If($? -and ($Null -ne $Roles))
	{
		OutputRoles $Roles
		If($Administrators)
		{
			OutputRoleDefinitions $Roles
			OutputRoleAdministrators $Roles
		}
	}
	ElseIf($? -and ($Null -eq $Roles))
	{
		$txt = "There are no Administrator Roles"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Administrator Roles"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputRoles
{
	Param([object] $Roles)
	
	Write-Verbose "$(Get-Date): `tOutput Roles"
	If($MSWord -or $PDF)
	{
		If($Administrators)
		{
			$Selection.InsertNewPage()
		}
		WriteWordLine 2 0 "Administrative Roles"
		[System.Collections.Hashtable[]] $RolesWordTable = @();
		ForEach($Role in $Roles)
		{
			$Tmp = $Null
			If($Role.BuiltIn)
			{
				$Tmp = "Built In"
			}
			Else
			{
				$Tmp = "Custom"
			}
			$WordTableRowHash = @{ Role = $Role.Name; Description = $Role.Description; Type = $Tmp;}

			$RolesWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $RolesWordTable `
		-Columns Role, Description, Type `
		-Headers "Role", "Description", "Type" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 300;
		$Table.Columns.Item(3).Width = 50;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Administrative Roles"
		Line 0 ""
		ForEach($Role in $Roles)
		{
			Line 1 "Role`t`t: " $Role.Name
			Line 1 "Description`t: " $Role.Description
			Line 1 "Type`t`t: " -NoNewLine
			If($Role.BuiltIn)
			{
				Line 0 "Built In"
			}
			Else
			{
				Line 0 "Custom"
			}
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 2 0 "Administrative Roles"
		ForEach($Role in $Roles)
		{
			$xType = ""
			If($Role.BuiltIn)
			{
				$xType = "Built In"
			}
			Else
			{
				$xType = "Custom"
			}
			$rowdata += @(,(
			$Role.Name,$htmlwhite,
			$Role.Description,$htmlwhite,
			$xType,$htmlwhite))
		}
		$columnHeaders = @(
		'Role',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("150","300","50")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputRoleDefinitions
{
	Param([object] $Roles)
	Write-Verbose "$(Get-Date): `t`tOutput Role Definitions"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		
		ForEach($Role in $Roles)
		{
			[System.Collections.Hashtable[]] $WordTable = @();
			WriteWordLine 3 0 "Role definition for $($Role.Name)"
			WriteWordLine 0 0 "Details - " $Role.Name
			WriteWordLine 0 0 $Role.Description
			$Permissions = $Role.Permissions
			$Results = GetRolePermissions $Permissions

			$comp = ""
			$x = 0
			ForEach($Result in $Results)
			{
				If($x -eq 0)
				{
					$comp = $Result.Value
					$WordTableRowHash = @{ 
					FolderName = $Result.Value; 
					Permission = $Result.Name; 
					}

					$WordTable += $WordTableRowHash;
				}
				Else
				{
					If($comp -eq $Result.value)
					{
						$WordTableRowHash = @{ 
						FolderName = ""; 
						Permission = $Result.Name; 
						}

						$WordTable += $WordTableRowHash;
					}
					Else
					{
						$comp = $Result.Value
						$WordTableRowHash = @{ 
						FolderName = $Result.Value; 
						Permission = $Result.Name; 
						}

						$WordTable += $WordTableRowHash;
					}
				}
				$x++
			}

			$Table = AddWordTable -Hashtable $WordTable `
			-Columns FolderName, Permission `
			-Headers "Folder Name", "Permissions" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 400;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	ElseIf($Text)
	{
		ForEach($Role in $Roles)
		{
			Line 0 "Role definition for $($Role.Name)"
			Line 0 "Details - " $Role.Name
			Line 0 $Role.Description
			Line 0 ""
			$Permissions = $Role.Permissions
			$Results = GetRolePermissions $Permissions

			ForEach($Result in $Results)
			{
				Line 1 "Folder Name`t: " $Result.Value
				Line 1 "Permission`t: " $Result.Name
				Line 0 ""
			}

			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		ForEach($Role in $Roles)
		{
			$rowdata = @()
			WriteHTMLLine 3 0 "Role definition for $($Role.Name)"
			WriteHTMLLine 0 0 "Details - " $Role.Name
			WriteHTMLLine 0 0 $Role.Description
			$Permissions = $Role.Permissions
			$Results = GetRolePermissions $Permissions

			$comp = ""
			$x = 0
			ForEach($Result in $Results)
			{
				If($x -eq 0)
				{
					$comp = $Result.Value
					$rowdata += @(,(
					$Result.Value,$htmlwhite,
					$Result.Name,$htmlwhite))
				}
				Else
				{
					If($comp -eq $Result.value)
					{
						$rowdata += @(,(
						"",$htmlwhite,
						$Result.Name,$htmlwhite))
					}
					Else
					{
						$comp = $Result.Value
						$rowdata += @(,(
						$Result.Value,$htmlwhite,
						$Result.Name,$htmlwhite))
					}
				}
				$x++
			}

			$columnHeaders = @(
			'Folder Name',($htmlsilver -bor $htmlbold),
			'Permissions',($htmlsilver -bor $htmlbold))

			$msg = ""
			$ColumnWidths = @("100","400")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders	-fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 " "
		}
	}
}

Function GetRolePermissions
{
	Param([object] $Permissions)
	
	$Results = @{}
	
	ForEach($Permission in $Permissions)
	{
		Switch ($Permission)
		{
			"Admin_FullControl"											{$Results.Add("Manage Administrators", "Administrators"); Break}
			"Admin_Read"												{$Results.Add("View Administrators", "Administrators"); Break}
			"Applications_AttachClientHostedApplicationToDesktopGroup"	{$Results.Add("Attach Local Access Application to Delivery Group", "Delivery Groups"); Break}
			"Applications_ChangeMaintenanceMode"						{$Results.Add("Enable/disable maintenance mode of an Application", "Delivery Groups"); Break}
			"Applications_ChangeTags"									{$Results.Add("Edit Application tags", "Delivery Groups"); Break}
			"Applications_ChangeUserAssignment"							{$Results.Add("Change users assigned to an application", "Delivery Groups"); Break}
			"Applications_Create"										{$Results.Add("Create Application", "Delivery Groups"); Break}
			"Applications_CreateFolder"									{$Results.Add("Create Application Folder", "Delivery Groups"); Break}
			"Applications_Delete"										{$Results.Add("Delete Application", "Delivery Groups"); Break}
			"Applications_DetachClientHostedApplicationToDesktopGroup"	{$Results.Add("Detach Local Access Application from Delivery Group", "Delivery Groups"); Break}
			"Applications_EditFolder"									{$Results.Add("Edit Application Folder", "Delivery Groups"); Break}
			"Applications_EditProperties"								{$Results.Add("Edit Application Properties", "Delivery Groups"); Break}
			"Applications_MoveFolder"									{$Results.Add("Move Application Folder", "Delivery Groups"); Break}
			"Applications_Read"											{$Results.Add("View Applications", "Delivery Groups"); Break}
			"Applications_RemoveFolder"									{$Results.Add("Remove Application Folder", "Delivery Groups"); Break}
			"AppV_AddServer"											{$Results.Add("Add App-V publishing server", "App-V"); Break}
			"AppV_DeleteServer"											{$Results.Add("Delete App-V publishing server", "App-V"); Break}
			"AppV_Read"													{$Results.Add("Read App-V servers", "App-V"); Break}
			"Catalog_AddMachines"										{$Results.Add("Add Machines to Machine Catalog", "Machine Catalogs"); Break}
			"Catalog_AddScope"											{$Results.Add("Add Machine Catalog to Scope", "Machine Catalogs"); Break}
			"Catalog_CancelProvTask"									{$Results.Add("Cancel Provisioning Task", "Machine Catalogs"); Break}
			"Catalog_ChangeMachineMaintenanceMode"						{$Results.Add("Enable/disable maintenance mode of a machine via Machine Catalog membership", "Machine Catalogs"); Break}
			"Catalog_ChangeMaintenanceMode"								{$Results.Add("Enable/disable maintenance mode on Desktop via Machine Catalog membership", "Machine Catalogs"); Break}
			"Catalog_ChangeUserAssignment"								{$Results.Add("Change users assigned to a machine", "Machine Catalogs"); Break}
			"Catalog_ConsumeMachines"									{$Results.Add("Allow machines to be consumed by a Delivery Group", "Machine Catalogs"); Break}
			"Catalog_Create"											{$Results.Add("Create Machine Catalog", "Machine Catalogs"); Break}
			"Catalog_Delete"											{$Results.Add("Delete Machine Catalog", "Machine Catalogs"); Break}
			"Catalog_EditProperties"									{$Results.Add("Edit Machine Catalog Properties", "Machine Catalogs"); Break}
			"Catalog_ManageAccounts"									{$Results.Add("Manage Active Directory Accounts", "Machine Catalogs"); Break}
			"Catalog_PowerOperations_RDS"								{$Results.Add("Perform power operations on Windows Server machines via Machine Catalog membership", "Machine Catalogs"); Break}
			"Catalog_PowerOperations_VDI"								{$Results.Add("Perform power operations on Windows Desktop machines via Machine Catalog membership", "Machine Catalogs"); Break}
			"Catalog_Read"												{$Results.Add("View Machine Catalogs", "Machine Catalogs"); Break}
			"Catalog_RemoveMachine"										{$Results.Add("Remove Machines from Machine Catalog", "Machine Catalogs"); Break}
			"Catalog_RemoveScope"										{$Results.Add("Remove Machine Catalog from Scope", "Machine Catalogs"); Break}
			"Catalog_SessionManagement"									{$Results.Add("Perform session management on machines via Machine Catalog membership", "Machine Catalogs"); Break}
			"Catalog_UpdateMasterImage"									{$Results.Add("Perform Machine update", "Machine Catalogs"); Break}
			"Configuration_Read"										{$Results.Add("Read Site Configuration", "Other permissions"); Break}
			"Configuration_Write"										{$Results.Add("Update Site Configuration", "Other permissions"); Break}
			"Controllers_Remove"										{$Results.Add("Remove Delivery Controller", "Controllers"); Break}
			"DesktopGroup_AddApplication" {$Results.Add("Add Application to Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_AddMachines" {$Results.Add("Add Machines to Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_AddScope" {$Results.Add("Add Delivery Group to Scope", "Delivery Groups"); Break}
			"DesktopGroup_ChangeMachineMaintenanceMode" {$Results.Add("Enable/disable maintenance mode of a machine via Delivery Group membership", "Delivery Groups"); Break}
			"DesktopGroup_ChangeMaintenanceMode" {$Results.Add("Enable/disable maintenance mode of a Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_ChangeTags" {$Results.Add("Edit Delivery Group tags", "Delivery Groups"); Break}
			"DesktopGroup_ChangeUserAssignment" {$Results.Add("Change users assigned to a desktop", "Delivery Groups"); Break}
			"DesktopGroup_Create" {$Results.Add("Create Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_Delete" {$Results.Add("Delete Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_EditProperties" {$Results.Add("Edit Delivery Group Properties", "Delivery Groups"); Break}
			"DesktopGroup_PowerOperations_RDS" {$Results.Add("Perform power operations on Windows Server machines via Delivery Group membership", "Delivery Groups"); Break}
			"DesktopGroup_PowerOperations_VDI" {$Results.Add("Perform power operations on Windows Desktop machines via Delivery Group membership", "Delivery Groups"); Break}
			"DesktopGroup_Read" {$Results.Add("View Delivery Groups", "Delivery Groups"); Break}
			"DesktopGroup_RemoveApplication" {$Results.Add("Remove Application from Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_RemoveDesktop" {$Results.Add("Remove Desktop from Delivery Group", "Delivery Groups"); Break}
			"DesktopGroup_RemoveScope" {$Results.Add("Remove Delivery Group from Scope", "Delivery Groups"); Break}
			"DesktopGroup_SessionManagement" {$Results.Add("Perform session management on machines via Delivery Group membership", "Delivery Groups"); Break}
			"Director_ClientDetails_Read" {$Results.Add("View Client Details page", "Director"); Break}
			"Director_ClientHelpDesk_Read" {$Results.Add("View Client Activity Manager page", "Director"); Break}
			"Director_Dashboard_Read" {$Results.Add("View Dashboard page", "Director"); Break}
			"Director_DesktopHardwareInformation_Edit" {$Results.Add("Edit Machine Hardware related Broker machine command properties", "Director"); Break}
			"Director_HDXInformation_Edit" {$Results.Add("Edit HDX related Broker machine command properties", "Director"); Break}
			"Director_HelpDesk_Read" {$Results.Add("View Activity Manager page", "Director"); Break}
			"Director_KillApplication" {$Results.Add("Perform Kill Application running on a machine", "Director"); Break}
			"Director_KillApplication_Edit" {$Results.Add("Edit Kill Application related Broker machine command properties", "Director"); Break}
			"Director_KillProcess" {$Results.Add("Perform Kill Process running on a machine", "Director"); Break}
			"Director_KillProcess_Edit" {$Results.Add("Edit Kill Process related Broker machine command properties", "Director"); Break}
			"Director_LatencyInformation_Edit" {$Results.Add("Edit Latency related Broker machine command properties", "Director"); Break}
			"Director_MachineDetails_Read" {$Results.Add("View Machine Details page", "Director"); Break}
			"Director_MachineMetricValues_Edit" {$Results.Add("Edit Machine metric related Broker machine command properties", "Director"); Break}
			"Director_PersonalizationInformation_Edit" {$Results.Add("Edit Personalization related Broker machine command properties", "Director"); Break}
			"Director_PoliciesInformation_Edit" {$Results.Add("Edit Policies related Broker machine command properties", "Director"); Break}
			"Director_ResetVDisk" {$Results.Add("Perform Reset VDisk operation", "Director"); Break}
			"Director_ResetVDisk_Edit" {$Results.Add("Edit Reset VDisk related Broker machine command properties", "Director"); Break}
			"Director_RoundTripInformation_Edit" {$Results.Add("Edit Roundtrip Time related Broker machine command properties", "Director"); Break}
			"Director_ShadowSession" {$Results.Add("Perform Remote Assistance on a machine", "Director"); Break}
			"Director_ShadowSession_Edit" {$Results.Add("Edit Remote Assistance related Broker machine command properties", "Director"); Break}
			"Director_SliceAndDice_Read" {$Results.Add("View Filters page", "Director"); Break}
			"Director_TaskManagerInformation_Edit" {$Results.Add("Edit Task Manager related Broker machine command properties", "Director"); Break}
			"Director_Trends_Read" {$Results.Add("View Trends page", "Director"); Break}
			"Director_UserDetails_Read" {$Results.Add("View User Details page", "Director"); Break}
			"Director_WindowsSessionId_Edit" {$Results.Add("Edit Windows Sessionid related Broker machine command properties", "Director"); Break}
			"EnvTest" {$Results.Add("Run environment tests", "Other permissions"); Break}
			"Global_Read" {$Results.Add("Read Site Configuration", "Other permissions"); Break}
			"Global_Write" {$Results.Add("Update Site Configuration", "Other permissions"); Break}
			"Hosts_AddScope" {$Results.Add("Add Host Connection to Scope", "Hosts"); Break}
			"Hosts_AddStorage" {$Results.Add("Add storage to Resources", "Hosts"); Break}
			"Hosts_ChangeMaintenanceMode" {$Results.Add("Enable/disable maintenance mode of a Host Connection", "Hosts"); Break}
			"Hosts_Consume" {$Results.Add("Use Host Connection or Resources to Create Catalog", "Hosts"); Break}
			"Hosts_CreateHost" {$Results.Add("Add Host Connection or Resources", "Hosts"); Break}
			"Hosts_DeleteConnection" {$Results.Add("Delete Host Connection", "Hosts"); Break}
			"Hosts_DeleteHost" {$Results.Add("Delete Resources", "Hosts"); Break}
			"Hosts_EditConnectionProperties" {$Results.Add("Edit Host Connection properties", "Hosts"); Break}
			"Hosts_EditHostProperties" {$Results.Add("Edit Resources", "Hosts"); Break}
			"Hosts_Read" {$Results.Add("View Host Connections and Resources", "Hosts"); Break}
			"Hosts_RemoveScope" {$Results.Add("Remove Host Connection from Scope", "Hosts"); Break}
			"Licensing_ChangeLicenseServer"								{$Results.Add("Change licensing server", "Licensing"); Break}
			"Licensing_EditLicensingProperties"							{$Results.Add("Edit product edition", "Licensing"); Break}
			"Licensing_Read"											{$Results.Add("View Licensing", "Licensing"); Break}
			"Logging_Delete"											{$Results.Add("Delete Configuration Logs", "Logging"); Break}
			"Logging_EditPreferences"									{$Results.Add("Edit Logging Preferences", "Logging"); Break}
			"Logging_Read"												{$Results.Add("View Configuration Logs", "Logging"); Break}
			"PerformUpgrade"											{$Results.Add("Perform upgrade", "Other permissions"); Break}
			"Policies_Manage"											{$Results.Add("Manage Policies", "Policies"); Break}
			"Policies_Read"												{$Results.Add("View Policies", "Policies"); Break}
			"Storefront_Create"											{$Results.Add("Create a new StoreFront definition", "StoreFronts"); Break}
			"Storefront_Delete"											{$Results.Add("Delete a StoreFront definition", "StoreFronts"); Break}
			"Storefront_Read"											{$Results.Add("Read StoreFront definitions", "StoreFronts"); Break}
			"Storefront_Update"											{$Results.Add("Update a StoreFront definition", "StoreFronts"); Break}
			"UPM_Reset_Profiles"										{$Results.Add("Reset user profiles", "Director"); Break}
			"UPM_Reset_Profiles_Edit"									{$Results.Add("Edit Reset User Profiles related Broker machine command properties", "Director"); Break}
		}
	}

	$Results = $Results.GetEnumerator() | Sort Value
	Return $Results
}

Function OutputRoleAdministrators 
{
	Param([object] $Roles)
	Write-Verbose "$(Get-Date): `t`tOutput Role Administrators"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		
		ForEach($Role in $Roles)
		{
			[System.Collections.Hashtable[]] $WordTable = @();
			WriteWordLine 3 0 "Administrators for Role: $($Role.Name)"
			$Admins = Get-AdminAdministrator @XDParams1 | ? {$_.Rights.RoleName -Contains $Role.Name}
			
			If($? -and $Null -ne $Admins)
			{
				ForEach($Admin in $Admins)
				{
					$xEnabled = "Disabled"
					If($Admin.Enabled)
					{
						$xEnabled = "Enabled"
					}

					$xScopeName = ""
					ForEach($Right in $Admin.Rights)
					{
						If($Right.RoleName -eq $Role.Name)
						{
							$xScopeName = $Right.ScopeName
						}
					}
					
					$WordTableRowHash = @{ 
					AdminName = $Admin.Name; 
					Scope = $xScopeName; 
					Type = $xEnabled;
					}

					$WordTable += $WordTableRowHash;
				}
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns AdminName, Scope, Type `
				-Headers "Administrator Name", "Scope", "Status" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 220;
				$Table.Columns.Item(2).Width = 225;
				$Table.Columns.Item(3).Width = 55;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($? -and $Null -eq $Admins)
			{
				WriteWordLine 0 0 "No administrators defined"
				WriteWordLine 0 0 ""
			}
			Else
			{
				WriteWordLine 0 0 "Unable to retrieve administrators"
				WriteWordLine 0 0 ""
			}
		}
	}
	ElseIf($Text)
	{
		ForEach($Role in $Roles)
		{
			Line 1 "Administrators for Role: $($Role.Name)"
			$Admins = Get-AdminAdministrator @XDParams1 | ? {$_.Rights.RoleName -Contains $Role.Name}
			
			If($? -and $Null -ne $Admins)
			{
				ForEach($Admin in $Admins)
				{
					$xEnabled = "Disabled"
					If($Admin.Enabled)
					{
						$xEnabled = "Enabled"
					}

					$xScopeName = ""
					ForEach($Right in $Admin.Rights)
					{
						If($Right.RoleName -eq $Role.Name)
						{
							$xScopeName = $Right.ScopeName
						}
					}
					
					Line 2 "Administrator Name`t: " $Admin.Name
					Line 2 "Scope`t`t`t: " $xScopeName
					Line 2 "Status`t`t`t: " $xEnabled
					Line 0 ""
				}
			}
			ElseIf($? -and $Null -eq $Admins)
			{
				Line 2 "No administrators defined"
				Line 0 ""
			}
			Else
			{
				Line 2 "Unable to retrieve administrators"
				Line 0 ""
			}
		}
	}
	ElseIf($HTML)
	{
		ForEach($Role in $Roles)
		{
			$rowdata = @()
			WriteHTMLLine 3 0 "Administrators for Role: $($Role.Name)"
			$Admins = Get-AdminAdministrator @XDParams1 | ? {$_.Rights.RoleName -Contains $Role.Name}
			
			If($? -and $Null -ne $Admins)
			{
				ForEach($Admin in $Admins)
				{
					$xEnabled = "Disabled"
					If($Admin.Enabled)
					{
						$xEnabled = "Enabled"
					}

					$xScopeName = ""
					ForEach($Right in $Admin.Rights)
					{
						If($Right.RoleName -eq $Role.Name)
						{
							$xScopeName = $Right.ScopeName
						}
					}
					
					$rowdata += @(,(
					$Admin.Name,$htmlwhite,
					$xScopeName,$htmlwhite,
					$xEnabled,$htmlwhite))
				}
				$columnHeaders = @(
				'Administrator Name',($htmlsilver -bor $htmlbold),
				'Scope',($htmlsilver -bor $htmlbold),
				'Status',($htmlsilver -bor $htmlbold))

				$msg = ""
				$columnWidths = @("220","225","55")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 " "
			}
			ElseIf($? -and $Null -eq $Admins)
			{
				WriteHTMLLine 0 0 "No administrators defined"
				WriteHTMLLine 0 0 " "
			}
			Else
			{
				WriteHTMLLine 0 0 "Unable to retrieve administrators"
				WriteHTMLLine 0 0 " "
			}
		}
	}
}
#endregion

#region Controllers functions
Function ProcessControllers
{
	Write-Verbose "$(Get-Date): Processing Controllers"
	Write-Verbose "$(Get-Date): `tRetrieving Controller data"
	
	$Global:TotalControllers = 0
	
	$Controllers = Get-BrokerController @XDParams2 -SortBy DNSName

	If($? -and ($Null -ne $Controllers))
	{
		OutputControllers $Controllers
	}
	ElseIf($? -and ($Null -eq $Controllers))
	{
		$txt = "There are no Controllers"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Controllers"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputControllers
{
	Param([object]$Controllers)
	
	Write-Verbose "$(Get-Date): `tOutput Controllers"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Controllers"
		[System.Collections.Hashtable[]] $ControllersWordTable = @();
	}
	ElseIf($Text)
	{
		Line 0 "Controllers"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Controllers"
		$rowdata = @()
	}
	
	ForEach($Controller in $Controllers)
	{
		$Global:TotalControllers++

		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Name = $Controller.DNSName; 
			LastUpdated = $Controller.LastActivityTime; 
			RegisteredDesktops = $Controller.DesktopsRegistered;
			}

			$ControllersWordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "Name`t`t`t: " $Controller.DNSName
			Line 1 "Last updated`t`t: " $Controller.LastActivityTime
			Line 1 "Registered desktops`t: " $Controller.DesktopsRegistered
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Controller.DNSName,$htmlwhite,
			$Controller.LastActivityTime,$htmlwhite,
			$Controller.DesktopsRegistered,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ControllersWordTable `
		-Columns Name, LastUpdated, RegisteredDesktops `
		-Headers "Name", "Last updated", "Registered desktops" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Last updated',($htmlsilver -bor $htmlbold),
		'Registered desktops',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
	
	If($Hardware)
	{
		ForEach($Controller in $Controllers)
		{
			$Script:Selection.InsertNewPage()
			GetComputerWMIInfo $Controller.DNSName
		}
	}
}
#endregion

#region Hosting functions
Function ProcessHosting
{
	#original work on the Hosting was done by Kenny Baldwin
	Write-Verbose "$(Get-Date): Processing Hosting"
	
	$Global:TotalHostingConnections = 0

	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Hosting"
	}
	ElseIf($Text)
	{
		Line 0 "Hosting"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Hosting"
	}

	$vmstorage = @()
	$pvdstorage = @()
	$vmnetwork = @()

	Write-Verbose "$(Get-Date): `tProcessing Hosting Units"
	$HostingUnits = Get-ChildItem @XDParams1 -path 'xdhyp:\hostingunits' 4>$Null
	If($? -and $Null -ne $HostingUnits)
	{
		ForEach($item in $HostingUnits)
		{	
			$Global:TotalHostingConnections++
			ForEach($storage in $item.Storage)
			{	
				$vmstorage += $storage.StoragePath
			}
			ForEach($storage in $item.PersonalvDiskStorage)
			{	
				$pvdstorage += $storage.StoragePath
			}
			ForEach($network in $item.NetworkPath)
			{	
				$vmnetwork += $network
			}
		}
	}
	ElseIf($? -and $Null -eq $HostingUnits)
	{
		$txt = "No Hosting Units found"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Hosting Units"
		OutputWarning $txt
	}

	Write-Verbose "$(Get-Date): `tProcessing Hypervisors"
	$Hypervisors = Get-BrokerHypervisorConnection @XDParams1
	If($? -and $Null -ne $Hypervisors)
	{
		ForEach($Hypervisor in $Hypervisors)
		{
			$hypvmstorage = @()
			$hyppvdstorage = @()
			$hypnetwork = @()
			$capabilities = $Hypervisor.Capabilities -join ', '	
			ForEach($storage in $vmstorage)
			{
				If($storage.Contains($Hypervisor.Name))
				{		
					$hypvmstorage += $storage		
				}
			}
			ForEach($storage in $pvdstorage)
			{
				If($storage.Contains($Hypervisor.Name))
				{
					$hyppvdstorage += $storage		
				}
			}
			ForEach($network in $vmnetwork)
			{
				If($network.Contains($Hypervisor.Name))
				{
					$hypnetwork += $network
				}
			}
			$xStorageName = ""
			ForEach($Unit in $HostingUnits)
			{
				If($Unit.HypervisorConnection.HypervisorConnectionName -eq $Hypervisor.Name)
				{
					$xStorageName = $Unit.HostingUnitName
				}
			}
			$xAddress = ""
			$xHAAddress = @()
			$xUserName = ""
			$xScopes = ""
			$xMaintMode = $False
			$xConnectionType = ""
			$xState = ""
			$xZoneName = ""
			$xPowerActions = @()
			Write-Verbose "$(Get-Date): `tProcessing Hosting Connections"
			$Connections = Get-ChildItem @XDParams1 -path 'xdhyp:\connections' 4>$Null
			
			If($? -and $Null -ne $Connections)
			{
				ForEach($Connection in $Connections)
				{
					If($Connection.HypervisorConnectionName -eq $Hypervisor.Name)
					{
						$xAddress = $Connection.HypervisorAddress[0]
						ForEach($tmpaddress in $Connection.HypervisorAddress)
						{
							$xHAAddress += $tmpaddress
						}
						$xUserName = $Connection.UserName
						ForEach($Scope in $Connection.Scopes)
						{
							$xScopes += $Scope.ScopeName + "; "
						}
						$xScopes += "All"
						$xMaintMode = $Connection.MaintenanceMode
						$xConnectionType = $Connection.ConnectionType
						$xState = $Hypervisor.State
						$xZoneName = $Connection.ZoneName
						$xPowerActions = $Connection.metadata
					}
				}
			}
			ElseIf($? -and $Null -eq $Connections)
			{
				$txt = "No Hosting Connections found"
				OutputWarning $txt
			}
			Else
			{
				$txt = "Unable to retrieve Hosting Connections"
				OutputWarning $txt
			}
			OutputHosting $Hypervisor $xConnectionType $xAddress $xState $xUserName $xMaintMode $xStorageName $xHAAddress $xPowerActions $xScopes $xZoneName
		}
	}
	ElseIf($? -and $Null -eq $Hypervisors)
	{
		$txt = "No Hypervisors found"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve Hypervisors"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date):"
}

Function OutputHosting
{
	Param([object] $Hypervisor, [string] $xConnectionType, [string] $xAddress, [string] $xState, [string] $xUserName, [bool] $xMaintMode, [string] $xStorageName, [array] $xHAAddress, [array]$xPowerActions, [string] $xScopes, [string] $xZoneName)

	$xHAAddress = $xHAAddress | Sort
	
	$xxConnectionType = ""
	Switch ($xConnectionType)
	{
		"XenServer" {$xxConnectionType = "XenServer"; Break}
		"SCVMM"     {$xxConnectionType = "Microsoft System Center Virtual Machine Manager"; Break}
		"vCenter"   {$xxConnectionType = "VMware virtualization"; Break}
		"Custom"    {$xxConnectionType = "Custom"; Break}
		Default     {$xxConnectionType = "Hypervisor Type could not be determined: $($xConnectionType)"; Break}
	}

	$xxState = ""
	If($xState -eq "On")
	{
		$xxState = "Enabled"
	}
	Else
	{
		$xxState = "Disabled"
	}

	$xxMaintMode = ""
	If($xMaintMode)
	{
		$xxMaintMode = "On"
	}
	Else
	{
		$xxMaintMode = "Off"
	}
	
	Write-Verbose "$(Get-Date): `t`t`tOutput $($Hypervisor.Name)"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $Hypervisor.Name
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Connection Name"; Value = $Hypervisor.Name; }
		$ScriptInformation += @{Data = "Type"; Value = $xxConnectionType; }
		$ScriptInformation += @{Data = "Address"; Value = $xAddress; }
		$ScriptInformation += @{Data = "State"; Value = $xxState; }
		$ScriptInformation += @{Data = "Username"; Value = $xUserName; }
		$ScriptInformation += @{Data = "Scopes"; Value = $xScopes; }
		$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xxMaintMode; }
		If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
		{
			$ScriptInformation += @{Data = "Zone"; Value = $xZoneName; }
		}
		$ScriptInformation += @{Data = "Storage resource name"; Value = $xStorageName; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		
		WriteWordLine 4 0 "Advanced"
		$HAtmp = @()
		ForEach($tmpaddress in $xHAAddress)
		{
			$HAtmp += "$($tmpaddress)"
		}
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "High Availability Servers"; Value = $HAtmp[0]; }
		$cnt = -1
		ForEach($tmp in $HATmp)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$ScriptInformation += @{Data = ""; Value = $tmp; }
			}
		}
		$ScriptInformation += @{Data = "Simultaneous actions (all types) [Absolute]"; Value = $xPowerActions[0].Value; }
		$ScriptInformation += @{Data = "Simultaneous actions (all types) [Percentage]"; Value = $xPowerActions[2].Value; }
		$ScriptInformation += @{Data = "Simultaneous Personal vDisk inventory updates [Absolute]"; Value = $xPowerActions[4].Value; }
		$ScriptInformation += @{Data = "Simultaneous Personal vDisk inventory updates [Percentage]"; Value = $xPowerActions[3].Value; }
		$ScriptInformation += @{Data = "Maximum new actions per minute"; Value = $xPowerActions[1].Value; }
		If($xPowerActions.Count -gt 5)
		{
			$ScriptInformation += @{Data = "Connection options"; Value = $xPowerActions[5].Value; }
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 300;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 $Hypervisor.Name
		Line 0 ""
		Line 1 "Connection Name`t`t: " $Hypervisor.Name
		Line 1 "Type`t`t`t: " $xxConnectionType
		Line 1 "Address`t`t`t: " $xAddress
		Line 1 "State`t`t`t: " $xxState
		Line 1 "Username`t`t: " $xUserName
		Line 1 "Scopes`t`t`t: " $xScopes
		Line 1 "Maintenance Mode`t: " $xxMaintMode
		If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
		{
			Line 1 "Zone`t`t`t: " $xZoneName
		}
		Line 1 "Storage resource name`t: " $xStorageName
		Line 0 ""
		
		Line 1 "Advanced"
		Line 2 "High Availability Servers`t`t`t: " $xHAAddress[0]
		$cnt = 0
		ForEach($tmpaddress in $xHAAddress)
		{
			If($cnt -gt 0)
			{
				Line 8 "  " $tmpaddress
			}
			$cnt++
		}
		Line 2 "Simultaneous actions (all types) [Absolute]`t: " $xPowerActions[0].Value
		Line 2 "Simultaneous actions (all types) [Percentage]`t: " $xPowerActions[2].Value
		Line 2 "Simultaneous PvD inventory updates [Absolute]`t: " $xPowerActions[4].Value
		Line 2 "Simultaneous PvD inventory updates [Percentage]`t: " $xPowerActions[3].Value
		Line 2 "Maximum new actions per minute`t`t`t: " $xPowerActions[1].Value
		If($xPowerActions.Count -gt 5)
		{
			Line 2 "Connection options`t`t`t`t: " $xPowerActions[5].Value
		}
		Line 0 ""
		
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $Hypervisor.Name
		$rowdata = @()
		$columnHeaders = @("Connection Name",($htmlsilver -bor $htmlbold),$Hypervisor.Name,$htmlwhite)
		$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$xxConnectionType,$htmlwhite))
		$rowdata += @(,('Address',($htmlsilver -bor $htmlbold),$xAddress,$htmlwhite))
		$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$xxState,$htmlwhite))
		$rowdata += @(,('Username',($htmlsilver -bor $htmlbold),$xUserName,$htmlwhite))
		$rowdata += @(,('Scopes',($htmlsilver -bor $htmlbold),$xScopes,$htmlwhite))
		$rowdata += @(,('Maintenance Mode',($htmlsilver -bor $htmlbold),$xxMaintMode,$htmlwhite))
		If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
		{
			$rowdata += @(,('Zone',($htmlsilver -bor $htmlbold),$xZoneName,$htmlwhite))
		}
		$rowdata += @(,('Storage resource name',($htmlsilver -bor $htmlbold),$xStorageName,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
		
		WriteHTMLLine 4 0 "Advanced"
		$rowdata = @()
		$columnHeaders = @("High Availability Servers",($htmlsilver -bor $htmlbold),$xHAAddress[0],$htmlwhite)
		$cnt = 0
		ForEach($tmpaddress in $xHAAddress)
		{
			If($cnt -gt 0)
			{
				$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmpaddress,$htmlwhite))
			}
			$cnt++
		}
		$rowdata += @(,('Simultaneous actions (all types) [Absolute]',($htmlsilver -bor $htmlbold),$xPowerActions[0].Value,$htmlwhite))
		$rowdata += @(,('Simultaneous actions (all types) [Percentage]',($htmlsilver -bor $htmlbold),$xPowerActions[2].Value,$htmlwhite))
		$rowdata += @(,('Simultaneous Personal vDisk inventory updates [Absolute]',($htmlsilver -bor $htmlbold),$xPowerActions[4].Value,$htmlwhite))
		$rowdata += @(,('Simultaneous Personal vDisk inventory updates [Percentage]',($htmlsilver -bor $htmlbold),$xPowerActions[3].Value,$htmlwhite))
		$rowdata += @(,('Maximum new actions per minute',($htmlsilver -bor $htmlbold),$xPowerActions[1].Value,$htmlwhite))
		If($xPowerActions.Count -gt 5)
		{
			$rowdata += @(,('Connection options',($htmlsilver -bor $htmlbold),$xPowerActions[5].Value,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("300","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
		WriteHTMLLine 0 0 " "
	}
	
	If($Hosting)
	{	
		Write-Verbose "$(Get-Date): `tProcessing Host Administrators"
		$Admins = GetAdmins "Host" $Hypervisor.Name
		
		If($? -and ($Null -ne $Admins))
		{
			OutputAdminsForDetails $Admins
		}
		ElseIf($? -and ($Null -eq $Admins))
		{
			$txt = "There are no administrators for Host $($Hypervisor.Name)"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve administrators for Host $($Hypervisor.Name)"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date): `tProcessing Desktop OS Data"
		$DesktopOSMachines = Get-BrokerMachine @XDParams2 -hypervisorconnectionname $Hypervisor.Name -sessionsupport "SingleSession"

		If($? -and ($Null -ne $DesktopOSMachines))
		{
			[int]$cnt = 0
			If($DesktopOSMachines -is [array])
			{
				$cnt = $DesktopOSMachines.Count
			}
			Else
			{
				If(![String]::IsNullOrEmpty($DesktopOSMachines))
				{
					$cnt = 1
				}
				Else
				{
					$cnt = 0
				}
			}
			
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 4 0 "Desktop OS Machines ($($cnt))"
			}
			ElseIf($Text)
			{
				Line 0 "Desktop OS Machines ($($cnt))"
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 4 0 "Desktop OS Machines ($($cnt))"
			}

			ForEach($Desktop in $DesktopOSMachines)
			{
				OutputDesktopOSMachine $Desktop
			}
		}
		ElseIf($? -and ($Null -eq $DesktopOSMachines))
		{
			$txt = "There are no Desktop OS Machines"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Desktop OS Machines"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date): `tProcessing Server OS Data"
		$ServerOSMachines = Get-BrokerMachine @XDParams2 -hypervisorconnectionname $Hypervisor.Name -sessionsupport "MultiSession"
		
		If($? -and ($Null -ne $ServerOSMachines))
		{
			If($ServerOSMachines -is [array])
			{
				$cnt = $ServerOSMachines.Count
			}
			Else
			{
				If(![String]::IsNullOrEmpty($ServerOSMachines))
				{
					$cnt = 1
				}
				Else
				{
					$cnt = 0
				}
			}

			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 4 0 "Server OS Machines ($($cnt))"
			}
			ElseIf($Text)
			{
				Line 0 ""
				Line 0 "Server OS Machines ($($cnt))"
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 4 0 "Server OS Machines ($($cnt))"
			}
			
			ForEach($Server in $ServerOSMachines)
			{
				OutputServerOSMachine $Server
			}
		}
		ElseIf($? -and ($Null -eq $ServerOSMachines))
		{
			$txt = "There are no Server OS Machines"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Server OS Machines"
			OutputWarning $txt
		}

		Write-Verbose "$(Get-Date): `tProcessing Sessions Data"
		$Sessions = Get-BrokerSession @XDParams1 -hypervisorconnectionname $Hypervisor.Name -SortBy UserName
		If($? -and ($Null -ne $Sessions))
		{
			If($Sessions -is [array])
			{
				$cnt = $Sessions.Count
			}
			Else
			{
				If(![String]::IsNullOrEmpty($Sessions))
				{
					$cnt = 1
				}
				Else
				{
					$cnt = 0
				}
			}

			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 4 0 "Sessions ($($cnt))"
			}
			ElseIf($Text)
			{
				Line 0 ""
				Line 0 "Sessions ($($cnt))"
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 4 0 "Sessions ($($cnt))"
			}
			
			OutputHostingSessions $Sessions
		}
		ElseIf($? -and ($Null -eq $Sessions))
		{
			$txt = "There are no Sessions"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Sessions"
			OutputWarning $txt
		}
	}
}

Function OutputDesktopOSMachine 
{
	Param([object]$Desktop)

	$xName = ""
	$xMaintMode = ""
	$xUserChanges = ""
	
	Write-Verbose "$(Get-Date): `t`t`tOutput desktop $($Desktop.DNSName)"
	If($MSWord -or $PDF)
	{
		If(![String]::IsNullOrEmpty($Desktop.AssociatedUserNames))
		{
			ForEach($AssociatedUserName in $Desktop.AssociatedUserNames)
			{
				$xName += $AssociatedUserName
			}
		}
		If($xName -eq "")
		{
			$xName = "Not assigned"
		}
		If($Desktop.InMaintenanceMode)
		{
			$xMaintMode = "On"
		}
		Else
		{
			$xMaintMode = "Off"
		}
		Switch($Desktop.PersistUserChanges)
		{
			"OnLocal" {$xUserChanges = "On Local"; Break}
			"Discard" {$xUserChanges = "Discard"; Break}
			"OnPvd"   {$xUserChanges = "Personal vDisk"; Break}
			Default   {$xUserChanges = "Unknown: $($Desktop.PersistUserChanges)"; Break}
		}
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Name"; Value = $Desktop.DNSName; }
		$ScriptInformation += @{Data = "Machine Catalog"; Value = $Desktop.CatalogName; }
		$ScriptInformation += @{Data = "Delivery Group"; Value = $Desktop.DesktopGroupName; }
		$ScriptInformation += @{Data = "User"; Value = $xName; }
		$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xMaintMode; }
		$ScriptInformation += @{Data = "Persist User Changes"; Value = $xUserChanges; }
		$ScriptInformation += @{Data = "Power State"; Value = $Desktop.PowerState; }
		$ScriptInformation += @{Data = "Registration State"; Value = $Desktop.RegistrationState; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 "Name`t`t`t: " $Desktop.DNSName
		Line 1 "Machine Catalog`t`t: " $Desktop.CatalogName
		If(![String]::IsNullOrEmpty($Desktop.DesktopGroupName))
		{
			Line 1 "Delivery Group`t`t: " $Desktop.DesktopGroupName
		}
		If(![String]::IsNullOrEmpty($Desktop.AssociatedUserNames))
		{
			ForEach($AssociatedUserName in $Desktop.AssociatedUserNames)
			{
				$xName += $AssociatedUserName
			}
			Line 1 "User`t`t`t: " $xName
		}
		If($Desktop.InMaintenanceMode)
		{
			$xMaintMode = "On"
		}
		Else
		{
			$xMaintMode = "Off"
		}
		Line 1 "Maintenance Mode`t: " $xMaintMode
		Switch($Desktop.PersistUserChanges)
		{
			"OnLocal" {$xUserChanges = "On Local"; Break}
			"Discard" {$xUserChanges = "Discard"; Break}
			"OnPvd"   {$xUserChanges = "Personal vDisk"; Break}
			Default   {$xUserChanges = "Unknown: $($Desktop.PersistUserChanges)"; Break}
		}
		Line 1 "Persist User Changes`t: " $xUserChanges
		Line 1 "Power State`t`t: " $Desktop.PowerState
		Line 1 "Registration State`t: " $Desktop.RegistrationState
		Line 0 ""
	}
	ElseIf($HTML)
	{
		If($Desktop.InMaintenanceMode)
		{
			$xMaintMode = "On"
		}
		Else
		{
			$xMaintMode = "Off"
		}
		Switch($Desktop.PersistUserChanges)
		{
			"OnLocal" {$xUserChanges = "On Local"; Break}
			"Discard" {$xUserChanges = "Discard"; Break}
			"OnPvd"   {$xUserChanges = "Personal vDisk"; Break}
			Default   {$xUserChanges = "Unknown: $($Desktop.PersistUserChanges)"; Break}
		}

		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Desktop.DNSName,$htmlwhite)
		$rowdata += @(,('Machine Catalog',($htmlsilver -bor $htmlbold),$Desktop.CatalogName,$htmlwhite))
		If(![String]::IsNullOrEmpty($Desktop.DesktopGroupName))
		{
			$rowdata += @(,('Delivery Group',($htmlsilver -bor $htmlbold),$Desktop.DesktopGroupName,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($Desktop.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Desktop.AssociatedUserNames)
			{
				If($cnt -eq 0)
				{
					$rowdata += @(,('User',($htmlsilver -bor $htmlbold),$AssociatedUserName,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('User',($htmlsilver -bor $htmlbold),$AssociatedUserName,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Maintenance Mode',($htmlsilver -bor $htmlbold),$xMaintMode,$htmlwhite))
		$rowdata += @(,('Persist User Changes',($htmlsilver -bor $htmlbold),$xUserChanges,$htmlwhite))
		$rowdata += @(,('Power State',($htmlsilver -bor $htmlbold),$Desktop.PowerState,$htmlwhite))
		$rowdata += @(,('Registration State',($htmlsilver -bor $htmlbold),$Desktop.RegistrationState,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputServerOSMachine 
{
	Param([object]$Server)
	
	Write-Verbose "$(Get-Date): `t`t`tOutput server $($Server.DNSName)"
	$xName = ""
	$xMaintMode = ""
	$xUserChanges = ""

	If($MSWord -or $PDF)
	{
		If(![String]::IsNullOrEmpty($Server.AssociatedUserNames))
		{
			ForEach($AssociatedUserName in $Server.AssociatedUserNames)
			{
				$xName += $AssociatedUserName + "`n"
			}
		}
		If($xName -eq "")
		{
			$xName = "Not assigned"
		}
		If($Server.InMaintenanceMode)
		{
			$xMaintMode = "On"
		}
		Else
		{
			$xMaintMode = "Off"
		}
		Switch($Server.PersistUserChanges)
		{
			"OnLocal" {$xUserChanges = "On Local"; Break}
			"Discard" {$xUserChanges = "Discard"; Break}
			"OnPvd"   {$xUserChanges = "Personal vDisk"; Break}
			Default   {$xUserChanges = "Unknown: $($Server.PersistUserChanges)"; Break}
		}
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Name"; Value = $Server.DNSName; }
		$ScriptInformation += @{Data = "Machine Catalog"; Value = $Server.CatalogName; }
		$ScriptInformation += @{Data = "Delivery Group"; Value = $Server.DesktopGroupName; }
		$ScriptInformation += @{Data = "User"; Value = $xName; }
		$ScriptInformation += @{Data = "Maintenance Mode"; Value = $xMaintMode; }
		$ScriptInformation += @{Data = "Persist User Changes"; Value = $xUserChanges; }
		$ScriptInformation += @{Data = "Power State"; Value = $Server.PowerState; }
		$ScriptInformation += @{Data = "Registration State"; Value = $Server.RegistrationState; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 "Name`t`t`t: " $Server.DNSName
		Line 1 "Machine Catalog`t`t: " $Server.CatalogName
		If(![String]::IsNullOrEmpty($Server.DesktopGroupName))
		{
			Line 1 "Delivery Group`t`t: " $Server.DesktopGroupName
		}
		If(![String]::IsNullOrEmpty($Server.AssociatedUserNames))
		{
			ForEach($AssociatedUserName in $Server.AssociatedUserNames)
			{
				$xName += $AssociatedUserName + "`n"
			}
			Line 1 "User`t`t`t: " $xName
		}
		If($Server.InMaintenanceMode)
		{
			$xMaintMode = "On"
		}
		Else
		{
			$xMaintMode = "Off"
		}
		Line 1 "Maintenance Mode`t: " $xMaintMode
		Switch($Server.PersistUserChanges)
		{
			"OnLocal" {$xUserChanges = "On Local"; Break}
			"Discard" {$xUserChanges = "Discard"; Break}
			"OnPvd"   {$xUserChanges = "Personal vDisk"; Break}
			Default   {$xUserChanges = "Unknown: $($Server.PersistUserChanges)"; Break}
		}
		Line 1 "Persist User Changes`t: " $xUserChanges
		Line 1 "Power State`t`t: " $Server.PowerState
		Line 1 "Registration State`t: " $Server.RegistrationState
		Line 0 ""
	}
	ElseIf($HTML)
	{
		If($Server.InMaintenanceMode)
		{
			$xMaintMode = "On"
		}
		Else
		{
			$xMaintMode = "Off"
		}
		Switch($Server.PersistUserChanges)
		{
			"OnLocal" {$xUserChanges = "On Local"; Break}
			"Discard" {$xUserChanges = "Discard"; Break}
			"OnPvd"   {$xUserChanges = "Personal vDisk"; Break}
			Default   {$xUserChanges = "Unknown: $($Server.PersistUserChanges)"; Break}
		}

		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Server.DNSName,$htmlwhite)
		$rowdata += @(,('Machine Catalog',($htmlsilver -bor $htmlbold),$Server.CatalogName,$htmlwhite))
		If(![String]::IsNullOrEmpty($Server.DesktopGroupName))
		{
			$rowdata += @(,('Delivery Group',($htmlsilver -bor $htmlbold),$Server.DesktopGroupName,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($Server.AssociatedUserNames))
		{
			$cnt = -1
			ForEach($AssociatedUserName in $Server.AssociatedUserNames)
			{
				If($cnt -eq 0)
				{
					$rowdata += @(,('User',($htmlsilver -bor $htmlbold),$AssociatedUserName,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('User',($htmlsilver -bor $htmlbold),$AssociatedUserName,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Maintenance Mode',($htmlsilver -bor $htmlbold),$xMaintMode,$htmlwhite))
		$rowdata += @(,('Persist User Changes',($htmlsilver -bor $htmlbold),$xUserChanges,$htmlwhite))
		$rowdata += @(,('Power State',($htmlsilver -bor $htmlbold),$Server.PowerState,$htmlwhite))
		$rowdata += @(,('Registration State',($htmlsilver -bor $htmlbold),$Server.RegistrationState,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputHostingSessions 
{
	Param([object] $Sessions)
	
	ForEach($Session in $Sessions)
	{
		Write-Verbose "$(Get-Date): `t`t`tOutput session $($Session.UserName)"
		#get the private desktop
		#get desktop by Session Uid
		$xMachineName = ""
		$Desktop = Get-BrokerDesktop -SessionUid $Session.Uid @XDParams1
		
		If($? -and $Null -ne $Desktop)
		{
			$xMachineName = $Desktop.MachineName
		}
		Else
		{
			$xMachineName = "Not Found"
		}

		If($Session.SessionSupport -eq "SingleSession")
		{
			$xSessionType = "Single"
		}
		Else
		{
			$xSessionType = "Multi"
		}
		
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{Data = "Current User"; Value = $Session.UserName; }
			$ScriptInformation += @{Data = "Name"; Value = $Session.ClientName; }
			$ScriptInformation += @{Data = "Delivery Group"; Value = $Session.DesktopGroupName; }
			$ScriptInformation += @{Data = "Machine Catalog"; Value = $Session.CatalogName; }
			$ScriptInformation += @{Data = "Brokering Time"; Value = $Session.BrokeringTime; }
			$ScriptInformation += @{Data = "Session State"; Value = $Session.SessionState; }
			$ScriptInformation += @{Data = "Application State"; Value = $Session.AppState; }
			$ScriptInformation += @{Data = "Session Support"; Value = $xSessionType; }
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 "Current User`t`t: " $Session.UserName
			Line 1 "Name`t`t`t: " $Session.ClientName
			Line 1 "Delivery Group`t`t: " $Session.DesktopGroupName
			Line 1 "Machine Catalog`t`t: " $Session.CatalogName
			Line 1 "Brokering Time`t`t: " $Session.BrokeringTime
			Line 1 "Session State`t`t: " $Session.SessionState
			Line 1 "Application State`t: " $Session.AppState
			Line 1 "Session Support`t`t: " $xSessionType
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Current User",($htmlsilver -bor $htmlbold),$Session.UserName,$htmlwhite)
			$rowdata += @(,('Name',($htmlsilver -bor $htmlbold),$Session.ClientName,$htmlwhite))
			$rowdata += @(,('Delivery Group',($htmlsilver -bor $htmlbold),$Session.DesktopGroupName,$htmlwhite))
			$rowdata += @(,('Machine Catalog',($htmlsilver -bor $htmlbold),$Session.CatalogName,$htmlwhite))
			$rowdata += @(,('Brokering Time',($htmlsilver -bor $htmlbold),$Session.BrokeringTime,$htmlwhite))
			$rowdata += @(,('Session State',($htmlsilver -bor $htmlbold),$Session.SessionState,$htmlwhite))
			$rowdata += @(,('Application State',($htmlsilver -bor $htmlbold),$Session.AppState,$htmlwhite))
			$rowdata += @(,('Session Support',($htmlsilver -bor $htmlbold),$xSessionType,$htmlwhite))

			$msg = ""
			$columnWidths = @("150","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
			WriteHTMLLine 0 0 " "
		}
	}
}
#endregion

#region Licensing functions
Function ProcessLicensing
{
	Write-Verbose "$(Get-Date): Processing Licensing"
	
	$Global:Licenses = @()
	OutputLicensingOverview
	
	#get product license info
	Write-Verbose "$(Get-Date): `tRetrieving Licensing data"
	$LSAdminAddress = Get-LicLocation -LicenseServerAddress $Script:XDSite1.LicenseServerName -EA 0 4>$Null
	If($? -and ($Null -ne $LSAdminAddress))
	{
		$LSCertificate = Get-LicCertificate -AdminAddress $LSAdminAddress -EA 0 4>$Null
		If($? -and ($Null -ne $LSCertificate))
		{
			$LicenseAdmins = Get-LicAdministrator -AdminAddress $LSAdminAddress -CertHash $LSCertificate.CertHash -EA 0 4>$Null
			If($? -and ($Null -ne $LicenseAdmins))
			{
				OutputLicenseAdmins $LicenseAdmins
			}
			Else
			{
				$txt = "Unable to retrieve License Administrators"
				OutputWarning $txt
			}

			$ProductLicenses = Get-LicInventory -AdminAddress $LSAdminAddress -CertHash $LSCertificate.CertHash -EA 0 4>$Null
			If($? -and ($Null -ne $ProductLicenses))
			{
				OutputXendesktopLicenses $LSAdminAddress $LSCertificate $ProductLicenses
			}
			Else
			{
				$txt = "Unable to retrieve Product Licenses"
				OutputWarning $txt
			}
			
		}
		Else
		{
			$txt = "Unable to retrieve License Server Certificate"
			OutputWarning $txt
		}
	}
	Else
	{
		$txt = "Unable to retrieve License Server Admin Address"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputLicensingOverview
{
	Write-Verbose "$(Get-Date): `tOutput Licensing Overview"
	$LicenseEditionType = ""
	$LicenseModelType = ""

	If($Script:XDSite2.ProductCode -eq "XDT")
	{
		Switch ($Script:XDSite2.ProductEdition)
		{
			"PLT" 	{$LicenseEditionType = "Platinum Edition"; Break}
			"ENT" 	{$LicenseEditionType = "Enterprise Edition"; Break}
			"APP" 	{$LicenseEditionType = "App Edition"; Break}
			"STD" 	{$LicenseEditionType = "VDI Edition"; Break}
			Default {$LicenseEditionType = "License edition could not be determined: $($Script:XDSite2.ProductEdition)"; Break}
		}
	}
	ElseIf($Script:XDSite2.ProductCode -eq "MPS")
	{
		Switch ($Script:XDSite2.ProductEdition)
		{
			"PLT" 	{$LicenseEditionType = "Platinum Edition"; Break}
			"ENT" 	{$LicenseEditionType = "Enterprise Edition"; Break}
			"ADV" 	{$LicenseEditionType = "Advanced Edition"; Break}
			Default {$LicenseEditionType = "License edition could not be determined: $($Script:XDSite2.ProductEdition)"; Break}
		}
	}

	If($Script:XDSite1.LicenseModel -eq "UserDevice")
	{
		$LicenseModelType = "User/Device"
	}
	Else
	{
		$LicenseModelType = $Script:XDSite1.LicenseModel
	}
	$tmpdate = '{0:yyyy\.MMdd}' -f $Script:XDSite1.LicensingBurnInDate
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Licensing"
		WriteWordLine 2 0 "Licensing Overview"

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "Site"; Value = $Script:XDSite1.Name; }
		$ScriptInformation += @{Data = "Server"; Value = $Script:XDSite1.LicenseServerName; }
		$ScriptInformation += @{Data = "Port"; Value = $Script:XDSite1.LicenseServerPort; }
		$ScriptInformation += @{Data = "Edition"; Value = $LicenseEditionType; }
		$ScriptInformation += @{Data = "License model"; Value = $LicenseModelType; }
		$ScriptInformation += @{Data = "Required SA date"; Value = $tmpdate; }
		$ScriptInformation += @{Data = "XenDesktop license use"; Value = $Script:XDSite1.LicensedSessionsActive; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 125;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Licensing"
		Line 0 "Licensing Overview"
		Line 0 ""
		Line 0 "Site`t`t`t: " $Script:XDSite1.Name
		Line 0 "Server`t`t`t: " $Script:XDSite1.LicenseServerName
		Line 0 "Port`t`t`t: " $Script:XDSite1.LicenseServerPort
		Line 0 "Edition`t`t`t: " $LicenseEditionType
		Line 0 "License model`t`t: " $LicenseModelType
		Line 0 "Required SA date`t: " $tmpdate
		Line 0 "XenDesktop license use`t: " $Script:XDSite1.LicensedSessionsActive
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Licensing"
		WriteHTMLLine 2 0 "Licensing Overview"
		$rowdata = @()
		$columnHeaders = @("Site",($htmlsilver -bor $htmlbold),$Script:XDSite1.Name,$htmlwhite)
		$rowdata += @(,('Server',($htmlsilver -bor $htmlbold),$Script:XDSite1.LicenseServerName,$htmlwhite))
		$rowdata += @(,('Port',($htmlsilver -bor $htmlbold),$Script:XDSite1.LicenseServerPort,$htmlwhite))
		$rowdata += @(,('Edition',($htmlsilver -bor $htmlbold),$LicenseEditionType,$htmlwhite))
		$rowdata += @(,('License model',($htmlsilver -bor $htmlbold),$LicenseModelType,$htmlwhite))
		$rowdata += @(,('Required SA date',($htmlsilver -bor $htmlbold),$tmpdate,$htmlwhite))
		$rowdata += @(,('XenDesktop license use',($htmlsilver -bor $htmlbold),$Script:XDSite1.LicensedSessionsActive,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","125")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "275"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputXenDesktopLicenses
{
	Param([object]$LSAdminAddress, [object]$LSCertificate, [object]$ProductLicenses)
	
	Write-Verbose "$(Get-Date): `tOutput Licenses"
	
	ForEach($Product in $ProductLicenses)
	{
		$LicModel = ""
		If($Product.LicenseProductName -eq $Script:XDSite2.ProductCode)
		{
			If($Product.LicenseModel -eq "UD")
			{
				$LicModel = "U/D"
			}
			Else
			{
				$LicModel = $Product.LicenseModel
			}
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name LicenseProduct	-Value $Product.LicenseProductName
			$obj | Add-Member -MemberType NoteProperty -Name LicenseType	-Value $Product.LicenseEdition
			$obj | Add-Member -MemberType NoteProperty -Name LicenseModel	-Value $LicModel
			$obj | Add-Member -MemberType NoteProperty -Name LicenseCount	-Value $Product.LicensesAvailable
			$Global:Licenses += $obj
		}
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "XenDesktop Licenses"
		[System.Collections.Hashtable[]] $LicensesWordTable = @();
		ForEach($Product in $ProductLicenses)
		{
			If($Product.LicenseProductName -eq $Script:XDSite2.ProductCode)
			{
				$tmpdate1 = '{0:d}' -f $Product.LicenseExpirationDate
				$tmpdate2 = '{0:yyyy\.MMdd}' -f $Product.LicenseSubscriptionAdvantageDate
				$WordTableRowHash = @{ 
				Product = $Product.LocalizedLicenseProductName;
				Mode = $Product.LocalizedLicenseModel;
				ExpirationDate = $tmpdate1;
				SubscriptionAdvantageDate = $tmpdate2;
				Type = $Product.LocalizedLicenseType;
				Quantity = $Product.LicensesAvailable;
				}

				$LicensesWordTable += $WordTableRowHash;
			}
		}
		$Table = AddWordTable -Hashtable $LicensesWordTable `
		-Columns Product, Mode, ExpirationDate, SubscriptionAdvantageDate, Type, Quantity `
		-Headers "Product", "Mode", "Expiration Date", "Subscription Advantage Date", "Type", "Quantity" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 140;
		$Table.Columns.Item(2).Width = 70;
		$Table.Columns.Item(3).Width = 65;
		$Table.Columns.Item(4).Width = 90;
		$Table.Columns.Item(5).Width = 80;
		$Table.Columns.Item(6).Width = 55;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "XenDesktop Licenses"
		Line 0 ""
		ForEach($Product in $ProductLicenses)
		{
			If($Product.LicenseProductName -eq "XDT")
			{
				$tmpdate1 = '{0:d}' -f $Product.LicenseExpirationDate
				$tmpdate2 = '{0:yyyy\.MMdd}' -f $Product.LicenseSubscriptionAdvantageDate
				Line 0 "Product`t`t`t`t: " $Product.LocalizedLicenseProductName
				Line 0 "Mode`t`t`t`t: " $Product.LocalizedLicenseModel
				Line 0 "Expiration Date`t`t`t: " $tmpdate1
				Line 0 "Subscription Advantage Date`t: " $tmpdate2
				Line 0 "Type`t`t`t`t: " $Product.LocalizedLicenseType
				Line 0 "Quantity`t`t`t: " $Product.LicensesAvailable
				Line 0 ""
			}
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "XenDesktop Licenses"
		$rowdata = @()
		ForEach($Product in $ProductLicenses)
		{
			If($Product.LicenseProductName -eq "XDT")
			{
				$tmpdate1 = '{0:d}' -f $Product.LicenseExpirationDate
				$tmpdate2 = '{0:yyyy\.MMdd}' -f $Product.LicenseSubscriptionAdvantageDate
				$rowdata += @(,(
				$Product.LocalizedLicenseProductName,$htmlwhite,
				$Product.LocalizedLicenseModel,$htmlwhite,
				$tmpdate1,$htmlwhite,
				$tmpdate2,$htmlwhite,
				$Product.LocalizedLicenseType,$htmlwhite,
				$Product.LicensesAvailable,$htmlwhite))
			}
		}
		$columnHeaders = @(
		'Product',($htmlsilver -bor $htmlbold),
		'Mode',($htmlsilver -bor $htmlbold),
		'Expiration Date',($htmlsilver -bor $htmlbold),
		'Subscription Advantage Date',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold),
		'Quantity',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("150","125","65","90","80","55")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "565"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputLicenseAdmins
{
	Param([object]$LicenseAdmins)
	
	Write-Verbose "$(Get-Date): `tProcessing License Administrators"

	$txt = "License Administrators"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $AdminsWordTable = @();
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}

	ForEach($Admin in $LicenseAdmins)
	{
		Write-Verbose "$(Get-Date): `t`tAdding Administrator $($Admin.Account)"

		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			AdminName = $Admin.Account; 
			Permissions = $Admin.Permissions; 
			}
			$AdminsWordTable += $WordTableRowHash;
		}
		ElseIf($Text)
		{
			Line 1 "Name`t`t: " $Admin.Account
			Line 1 "Permissions`t: " $Admin.Permissions
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Admin.Account,$htmlwhite,
			$Admin.Permissions,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $AdminsWordTable `
		-Columns  AdminName,Permissions `
		-Headers  "Name","Permissions" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Permissions',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("150","125")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "275"
		WriteHTMLLine 0 0 " "
	}
}

#endregion

#region StoreFront functions
Function ProcessStoreFront
{
	Write-Verbose "$(Get-Date): Processing StoreFront"
	
	$Global:TotalStoreFrontServers = 0
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "StoreFront"
	}
	ElseIf($Text)
	{
		Line 0 "StoreFront"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "StoreFront"
	}
	
	Write-Verbose "$(Get-Date): `tRetrieving StoreFront information"
	$SFInfos = Get-BrokerMachineConfiguration @XDParams1 -Name rs* -SortBy LeafName
	If($? -and ($Null -ne $SFInfos))
	{
		$First = $True
		ForEach($SFInfo in $SFInfos)
		{
			$Global:TotalStoreFrontServers++

			$SFByteArray = $SFInfo.Policy
			Write-Verbose "$(Get-Date): `t`tRetrieving StoreFront server information for $($SFInfo.LeafName)"
			$SFServer = Get-SFStoreFrontAddress -ByteArray $SFByteArray 4>$Null
			If($? -and ($Null -ne $SFServer))
			{
				If($MSWord -or $PDF)
				{
					If(!$First)
					{
						$Selection.InsertNewPage()
					}
					$First = $False
				}
				OutputStoreFront $SFServer $SFInfo
				If($StoreFront)
				{
					If($SFInfo.DesktopGroupUids.Count -gt 0)
					{
						OutputStoreFrontDeliveryGroups $SFInfo
					}
					
					Write-Verbose "$(Get-Date): `t`tProcessing administrators for StoreFront server $($SFServer.Name)"
					$Admins = GetAdmins "Storefront"
					
					If($? -and ($Null -ne $Admins))
					{
						OutputAdminsForDetails $Admins
					}
					ElseIf($? -and ($Null -eq $Admins))
					{
						$txt = "There are no administrators for StoreFront server $($SFServer.Name)"
						OutputWarning $txt
					}
					Else
					{
						$txt = "Unable to retrieve administrators for StoreFront server $($SFServer.Name)"
						OutputWarning $txt
					}
				}
			}
			ElseIf($? -and ($Null -eq $SFServer))
			{
				$txt = "There was no StoreFront Server found for $($SFInfo.LeafName)"
				OutputWarning $txt
			}
			Else
			{
				$txt = "Unable to retrieve StoreFront Server for $($SFInfo.LeafName)"
				OutputWarning $txt
			}
		}
	}
	ElseIf($? -and ($Null -eq $SFInfos))
	{
		$txt = "StoreFront is not configured for this Site"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve StoreFront configuration"
		OutputWarning $txt
	}
	
	Write-Verbose "$(Get-Date): "
}

Function OutputStoreFront
{
	Param([object]$SFServer, [object] $SFInfo)
	
	$DGCnt = $SFInfo.DesktopGroupUids.Count
	
	Write-Verbose "$(Get-Date): `t`t`tOutput StoreFront server $($SFServer.Name)"
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Server: " $SFServer.Name
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "StoreFront Server"; Value = $SFServer.Name; }
		$ScriptInformation += @{Data = "Used by # Delivery Groups"; Value = $DGCnt; }
		$ScriptInformation += @{Data = "URL"; Value = $SFServer.Url; }
		$ScriptInformation += @{Data = "Description"; Value = $SFServer.Description; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Server"
		Line 0 "StoreFront Server`t`t: " $SFServer.Name
		Line 0 "Used by # Delivery Groups`t: " $DGCnt
		Line 0 "URL`t`t`t`t: " $SFServer.Url
		Line 0 "Description`t`t`t: " $SFServer.Description
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Server: " $SFServer.Name
		$rowdata = @()
		$columnHeaders = @("StoreFront Server",($htmlsilver -bor $htmlbold),$SFServer.Name,$htmlwhite)
		$rowdata += @(,('Used by # Delivery Groups',($htmlsilver -bor $htmlbold),$DGCnt,$htmlwhite))
		$rowdata += @(,('URL',($htmlsilver -bor $htmlbold),$SFServer.Url,$htmlwhite))
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$SFServer.Description,$htmlwhite))

		$msg = ""
		$columnWidths = @("150","250")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputStoreFrontDeliveryGroups
{
	Param([object] $SFInfo)
	
	$DeliveryGroups = @()
	ForEach($DGUid in $SFInfo.DesktopGroupUids)
	{
		$Results = Get-BrokerDesktopGroup @XDParams1 -Uid $DGUid
		If($? -and $Null -ne $Results)
		{
			$DeliveryGroups += $Results.Name
		}
	}

	[array]$DeliveryGroups = $DeliveryGroups | Sort
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Delivery Groups"
		[System.Collections.Hashtable[]] $DGWordTable = @();
		ForEach($Group in $DeliveryGroups)
		{
			$WordTableRowHash = @{ 
			DGName = $Group;
			}

			$DGWordTable += $WordTableRowHash;
		}
		$Table = AddWordTable -Hashtable $DGWordTable `
		-Columns DGName `
		-Headers "Delivery Group" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Delivery Groups"
		Line 0 ""
		$cnt = -1
		ForEach($Group in $DeliveryGroups)
		{
			$cnt++
			If($cnt -eq 0)
			{
				Line 1 "Delivery Group: " $Group
			}
			Else
			{
				Line 3 "" $Group
			}
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Delivery Groups"
		$rowdata = @()
		ForEach($Group in $DeliveryGroups)
		{
			$rowdata += @(,(
			$Group,$htmlwhite))
		}

		$columnHeaders = @(
		'Delivery Group',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region AppV functions
Function ProcessAppV
{
	Write-Verbose "$(Get-Date): Processing App-V"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "App-V Publishing"
	}
	ElseIf($Text)
	{
		Line 0 "App-V Publishing"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "App-V Publishing"
	}
	
	Write-Verbose "$(Get-Date): `tRetrieving App-V configuration"
	$AppvConfig = Get-BrokerMachineConfiguration @XDParams1 -Name appv*
	
	If($? -and $Null -ne $AppVConfig)
	{
		Write-Verbose "$(Get-Date): `t`tRetrieving App-V server information"
		$AppVs = Get-CtxAppVServer -ByteArray $Appvconfig[0].Policy -EA 0
		If($? -and ($Null -ne $AppVs))
		{
			ForEach($AppV in $AppVs)
			{
				OutputAppV $AppV
			}
		}
		ElseIf($? -and ($Null -eq $AppVs))
		{
			$txt = "There was no App-V server information found"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve App-V information"
			OutputWarning $txt
		}
	}
	ElseIf($? -and $Null -eq $AppVConfig)
	{
		$txt = "App-V is not configured for this Site"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve App-V configuration"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputAppV
{
	Param([object]$AppV)
	
	Write-Verbose "$(Get-Date): `t`t`tOutput App-V server information"
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "App-V management server"; Value = $AppV.ManagementServer; }
		$ScriptInformation += @{Data = "App-V publishing server"; Value = $AppV.PublishingServer; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "App-V management server: " $Appv.ManagementServer
		Line 0 "App-V publishing server: " $AppV.PublishingServer
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("App-V management server",($htmlsilver -bor $htmlbold),$Appv.ManagementServer,$htmlwhite)
		$rowdata += @(,('App-V publishing server',($htmlsilver -bor $htmlbold),$AppV.PublishingServer,$htmlwhite))

		$msg = ""
		$columnWidths = @("250","250")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region AppDNA functiions
Function ProcessAppDNA
{
	Write-Verbose "$(Get-Date): Processing AppDNA"
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "AppDNA Connection Settings"
	}
	ElseIf($Text)
	{
		Line 0 "AppDNA Connection Settings"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "AppDNA Connection Settings"
	}
	
	Write-Verbose "$(Get-Date): `tRetrieving AppDNA connection settings"
	$AppDNAConfig = Get-AppLibAppDNAConnection @XDParams1
	
	If($? -and $Null -ne $AppDNAConfig)
	{
		Write-Verbose "$(Get-Date): `t`tTesting AppDNA connection. Please wait..."
		$AppDNAConnection = Test-AppLibAppDNAExistingConnection -EA 0
		If($? -and ($Null -ne $AppDNAConnection))
		{
			$AppDNATest = "Passed"
		}
		Else
		{
			$AppDNATest = "Failed"
		}
		OutputAppDNA $AppDNAConfig $AppDNATest
	}
	ElseIf($? -and $Null -eq $AppDNAConfig)
	{
		$txt = "AppDNA is not configured for this Site"
		OutputWarning $txt
	}
	Else
	{
		$txt = "Unable to retrieve AppDNA configuration"
		OutputWarning $txt
	}
	Write-Verbose "$(Get-Date): "
}

Function OutputAppDNA
{
	Param([object]$AppDNAConfig, [string]$AppDNATest)
	
	$AppDNAState = "Disabled"
	If($AppDNAConfig.Enabled)
	{
		$AppDNAState = "Enabled"
	}
	
	Write-Verbose "$(Get-Date): `t`t`tOutput AppDNA connection information"
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{Data = "URL"; Value = $AppDNAConfig.Address; }
		$ScriptInformation += @{Data = "Database"; Value = $AppDNAConfig.Database; }
		$ScriptInformation += @{Data = "User name"; Value = $AppDNAConfig.UserName; }
		$ScriptInformation += @{Data = "State"; Value = $AppDNAState; }
		$ScriptInformation += @{Data = "Test connection"; Value = $AppDNATest; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 250;
		$Table.Columns.Item(2).Width = 250;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 "URL`t`t: " $AppDNAConfig.Address
		Line 1 "Database`t: " $AppDNAConfig.Database
		Line 1 "User name`t: " $AppDNAConfig.UserName
		Line 1 "State`t`t: " $AppDNAState
		Line 1 "Test connection`t: " $AppDNATest
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("URL",($htmlsilver -bor $htmlbold),$AppDNAConfig.Address,$htmlwhite)
		$rowdata += @(,('Database',($htmlsilver -bor $htmlbold),$AppDNAConfig.Database,$htmlwhite))
		$rowdata += @(,('User name',($htmlsilver -bor $htmlbold),$AppDNAConfig.UserName,$htmlwhite))
		$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$AppDNAState,$htmlwhite))
		$rowdata += @(,('Test connection',($htmlsilver -bor $htmlbold),$AppDNATest,$htmlwhite))

		$msg = ""
		$columnWidths = @("250","250")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region zones
Function ProcessZones
{
	Write-Verbose "$(Get-Date): Processing Zones"
	
	$Global:TotalZones = 0

	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Zones"
	}
	ElseIf($Text)
	{
		Line 0 "Zones"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Zones"
	}
	
	#get all zone names
	Write-Verbose "$(Get-Date): `tRetrieving All Zones"
	$Zones = Get-ConfigZone @XDParams1 | Sort Name
	$ZoneMembers = @()
	
	ForEach($Zone in $Zones)
	{
		$Global:TotalZones++
		Write-Verbose "$(Get-Date): `t`tRetrieving Machine Catalogs for Zone $($Zone.Name)"
		$ZoneCatalogs = Get-BrokerCatalog @XDParams2 -ZoneUid $Zone.Uid
		ForEach($ZoneCatalog in $ZoneCatalogs)
		{
			$obj = New-Object -TypeName PSObject
	
			$obj | Add-Member -MemberType NoteProperty -Name MemName -Value $ZoneCatalog.Name
			$obj | Add-Member -MemberType NoteProperty -Name MemDesc -Value $ZoneCatalog.Description
			$obj | Add-Member -MemberType NoteProperty -Name MemType -Value "Machine Catalog"
			$obj | Add-Member -MemberType NoteProperty -Name MemZone -Value $Zone.Name
			
			$ZoneMembers += $obj
		}
		
		Write-Verbose "$(Get-Date): `t`tRetrieving Delivery Controllers for Zone $($Zone.Name)"
		$ZoneControllers = $Zone.ControllerNames
		ForEach($ZoneController in $ZoneControllers)
		{
			$Results = Get-BrokerController -EA 0 | Where {$_.MachineName -Like "*$($ZoneController)"}
			
			If($? -and $Null -ne $Results)
			{
				$obj = New-Object -TypeName PSObject
	
				$obj | Add-Member -MemberType NoteProperty -Name MemName -Value $ZoneController
				$obj | Add-Member -MemberType NoteProperty -Name MemDesc -Value $Results.DNSName
				$obj | Add-Member -MemberType NoteProperty -Name MemType -Value "Delivery Controller"
				$obj | Add-Member -MemberType NoteProperty -Name MemZone -Value $Zone.Name
			
				$ZoneMembers += $obj
			}
		}

		Write-Verbose "$(Get-Date): `t`tRetrieving Host Connections for Zone $($Zone.Name)"
		$ZoneHosts = Get-ChildItem @XDParams1 -path 'xdhyp:\connections' 4>$Null | Where {$_.ZoneUid -eq $Zone.Uid}
		ForEach($ZoneHost in $ZoneHosts)
		{
			$obj = New-Object -TypeName PSObject
			
			$obj | Add-Member -MemberType NoteProperty -Name MemName -Value $ZoneHost.HypervisorConnectionName
			$obj | Add-Member -MemberType NoteProperty -Name MemDesc -Value ""
			$obj | Add-Member -MemberType NoteProperty -Name MemType -Value "Host Connection"
			$obj | Add-Member -MemberType NoteProperty -Name MemZone -Value $Zone.Name
			
			$ZoneMembers += $obj
		}
	}
	
	OutputZoneSiteView $ZoneMembers
	
	OutputPerZoneView $ZoneMembers $Zones
}

Function OutputZoneSiteView
{
	Param([array]$ZoneMembers)
	
	Write-Verbose "$(Get-Date): `tOutput Zone Site View"
	$ZoneMembers = $ZoneMembers | Sort MemName
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Site View"
		[System.Collections.Hashtable[]] $ZoneWordTable = @();

		ForEach($ZoneMember in $ZoneMembers)
		{
			$WordTableRowHash = @{ 
			xName = $ZoneMember.MemName;
			xDesc = $ZoneMember.MemDesc;
			xType = $ZoneMember.MemType;
			xZone = $ZoneMember.MemZone;
			}

			$ZoneWordTable += $WordTableRowHash;
		}

		$Table = AddWordTable -Hashtable $ZoneWordTable `
		-Columns xName, xDesc, xType, xZone `
		-Headers "Name", "Description", "Type", "Zone" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 175;
		$Table.Columns.Item(3).Width = 100;
		$Table.Columns.Item(4).Width = 100;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 0 "Site View"
		Line 0 ""
		ForEach($ZoneMember in $ZoneMembers)
		{
			Line 1 "Name`t`t: " $ZoneMember.MemName
			Line 1 "Description`t: " $ZoneMember.MemDesc
			Line 1 "Type`t`t: " $ZoneMember.MemType
			Line 1 "Zone`t`t: " $ZoneMember.MemZone
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "Site View"
		$rowdata = @()
		ForEach($ZoneMember in $ZoneMembers)
		{
			$rowdata += @(,(
			$ZoneMember.MemName,$htmlwhite,
			$ZoneMember.MemDesc,$htmlwhite,
			$ZoneMember.MemType,$htmlwhite,
			$ZoneMember.MemZone,$htmlwhite))
		}
		
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Description',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold),
		'Zone',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("150","200","150","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "650"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputPerZoneView
{
	Param([array]$ZoneMembers, [object]$Zones)
	
	Write-Verbose "$(Get-Date): `tOutput Per Zone View"
	$ZoneMembers = $ZoneMembers | Sort MemZone, MemName

	ForEach($Zone in $Zones)
	{
		$TmpZoneMembers = $ZoneMembers | Where {$_.MemZone -eq $Zone.Name}
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $Zone.Name
			[System.Collections.Hashtable[]] $ZoneWordTable = @();

			ForEach($ZoneMember in $TmpZoneMembers)
			{
				$WordTableRowHash = @{ 
				xName = $ZoneMember.MemName;
				xDesc = $ZoneMember.MemDesc;
				xType = $ZoneMember.MemType;
				}

				$ZoneWordTable += $WordTableRowHash;
			}

			$Table = AddWordTable -Hashtable $ZoneWordTable `
			-Columns xName, xDesc, xType `
			-Headers "Name", "Description", "Type" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 125;
			$Table.Columns.Item(2).Width = 175;
			$Table.Columns.Item(3).Width = 100;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $Zone.Name
			Line 0 ""
			ForEach($ZoneMember in $TmpZoneMembers)
			{
				Line 1 "Name`t`t: " $ZoneMember.MemName
				Line 1 "Description`t: " $ZoneMember.MemDesc
				Line 1 "Type`t`t: " $ZoneMember.MemType
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 $Zone.Name
			$rowdata = @()
			ForEach($ZoneMember in $TmpZoneMembers)
			{
				$rowdata += @(,(
				$ZoneMember.MemName,$htmlwhite,
				$ZoneMember.MemDesc,$htmlwhite,
				$ZoneMember.MemType,$htmlwhite))
			}
			
			$columnHeaders = @(
			'Name',($htmlsilver -bor $htmlbold),
			'Description',($htmlsilver -bor $htmlbold),
			'Type',($htmlsilver -bor $htmlbold),
			'Zone',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("150","200","150")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 " "
		}
	}
	Write-Verbose "$(Get-Date): "
}
#endregion

#region summary page
Function ProcessSummaryPage
{
	#summary page
	Write-Verbose "$(Get-Date): Create Summary Page"
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "XenDesktop $($Script:XDSiteVersion) $($XDSiteName) Summary Page"
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "XenDesktop $($Script:XDSiteVersion) $($XDSiteName) Summary Page"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "XenDesktop $($Script:XDSiteVersion) $($XDSiteName) Summary Page"
	}

	Write-Verbose "$(Get-Date): `tAdd administrator summary info"
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date): `tAdd Machine Catalog summary info"
		WriteWordLine 0 0 "Machine Catalogs"
		WriteWordLine 0 1 "Total Server OS Catalogs`t: " $Global:TotalServerOSCatalogs
		WriteWordLine 0 1 "Total Desktop OS Catalogs`t: " $Global:TotalDesktopOSCatalogs
		WriteWordLine 0 1 "Total RemotePC Catalogs`t: " $Global:TotalRemotePCCatalogs
		WriteWordLine 0 2 "Total Machine Catalogs`t: " ($Global:TotalServerOSCatalogs+$Global:TotalDesktopOSCatalogs+$Global:TotalRemotePCCatalogs)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd AppDisks summary info"
		WriteWordLine 0 0 "AppDisks"
		WriteWordLine 0 1 "Total AppDisks`t`t`t: " $Global:TotalAppDisks
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd Delivery Group summary info"
		WriteWordLine 0 0 "Delivery Groups"
		WriteWordLine 0 1 "Total Application Groups`t: " $Global:TotalApplicationGroups
		WriteWordLine 0 1 "Total Desktop Groups`t`t: " $Global:TotalDesktopGroups
		WriteWordLine 0 1 "Total Apps & Desktop Groups`t: " $Global:TotalAppsAndDesktopGroups
		WriteWordLine 0 2 "Total Delivery Groups`t: " ($Global:TotalApplicationGroups+$Global:TotalDesktopGroups+$Global:TotalAppsAndDesktopGroups)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd Application summary info"
		WriteWordLine 0 0 "Applications"
		WriteWordLine 0 1 "Total Published Applications`t: " $Global:TotalPublishedApplications
		WriteWordLine 0 1 "Total App-V Applications`t: " $Global:TotalAppvApplications
		WriteWordLine 0 2 "Total Applications`t: " ($Global:TotalPublishedApplications + $Global:TotalAppvApplications)
		WriteWordLine 0 0 ""
		
		If($Policies -eq $True)
		{
			Write-Verbose "$(Get-Date): `tAdd Policy summary info"
			WriteWordLine 0 0 "Policies"
			WriteWordLine 0 1 "Total Computer Policies`t`t: " $Global:TotalComputerPolicies
			WriteWordLine 0 1 "Total User Policies`t`t: " $Global:TotalUserPolicies
			WriteWordLine 0 2 "Total Policies`t`t: " ($Global:TotalComputerPolicies + $Global:TotalUserPolicies)
			WriteWordLine 0 0 ""
			WriteWordLine 0 1 "Site Policies`t`t`t: " $Global:TotalSitePolicies
			
			If($NoADPolicies -eq $False)
			{
				WriteWordLine 0 1 "Citrix AD Policies Processed`t: $($Global:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
				WriteWordLine 0 1 "Citrix AD Policies not Processed`t: " $Global:TotalADPoliciesNotProcessed
			}
			WriteWordLine 0 0 ""
		}
		
		WriteWordLine 0 0 "Administrators"
		WriteWordLine 0 1 "Total Delivery Group Admins`t: " $Global:TotalDeliveryGroupAdmins
		WriteWordLine 0 1 "Total Full Admins`t`t: " $Global:TotalFullAdmins
		WriteWordLine 0 1 "Total Help Desk Admins`t`t: " $Global:TotalHelpDeskAdmins
		WriteWordLine 0 1 "Total Host Admins`t`t: " $Global:TotalHostAdmins
		WriteWordLine 0 1 "Total Machine Catalog Admins`t: " $Global:TotalMachineCatalogAdmins
		WriteWordLine 0 1 "Total Read Only Admins`t`t: " $Global:TotalReadOnlyAdmins
		WriteWordLine 0 1 "Total Custom Admins`t`t: " $Global:TotalCustomAdmins
		WriteWordLine 0 2 "Total Administrators`t: " ($Global:TotalDeliveryGroupAdmins+$Global:TotalFullAdmins+$Global:TotalHelpDeskAdmins+$Global:TotalHostAdmins+$Global:TotalMachineCatalogAdmins+$Global:TotalReadOnlyAdmins+$Global:TotalCustomAdmins)
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd Controller summary info"
		WriteWordLine 0 0 "Controllers"
		WriteWordLine 0 1 "Total Controllers`t`t: " $Global:TotalControllers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd Hosting Connection summary info"
		WriteWordLine 0 0 "Hosting Connections"
		WriteWordLine 0 1 "Total Hosting Connections`t: " $Global:TotalHostingConnections
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd Licensing summary info"
		WriteWordLine 0 0 "Licensing"
		$TotalLicenses = 0
		ForEach($License in $Global:Licenses)
		{
			WriteWordLine 0 1 "$($License.LicenseProduct) $($License.LicenseType) $($License.LicenseModel)`t`t`t: $($License.LicenseCount)"
			$TotalLicenses += $License.LicenseCount
		}
		WriteWordLine 0 2 "Total Licenses`t`t: " $TotalLicenses
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd StoreFront summary info"
		WriteWordLine 0 0 "StoreFront"
		WriteWordLine 0 1 "Total StoreFront Servers`t: " $Global:TotalStoreFrontServers
		WriteWordLine 0 0 ""
		Write-Verbose "$(Get-Date): `tAdd Zone summary info"
		WriteWordLine 0 0 "Zones"
		WriteWordLine 0 1 "Total Zones`t`t`t: " $Global:TotalZones
	}
	ElseIf($Text)
	{
		Write-Verbose "$(Get-Date): `tAdd Machine Catalog summary info"
		Line 0 "Machine Catalogs"
		Line 1 "Total Server OS Catalogs`t: " $Global:TotalServerOSCatalogs
		Line 1 "Total Desktop OS Catalogs`t: " $Global:TotalDesktopOSCatalogs
		Line 1 "Total RemotePC Catalogs`t`t: " $Global:TotalRemotePCCatalogs
		Line 2 "Total Machine Catalogs`t: " ($Global:TotalServerOSCatalogs+$Global:TotalDesktopOSCatalogs+$Global:TotalRemotePCCatalogs)
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd AppDisks summary info"
		Line 0 "AppDisks"
		Line 1 "Total AppDisks`t`t`t: " $Global:TotalAppDisks
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd Delivery Group summary info"
		Line 0 "Delivery Groups"
		Line 1 "Total Application Groups`t: " $Global:TotalApplicationGroups
		Line 1 "Total Desktop Groups`t`t: " $Global:TotalDesktopGroups
		Line 1 "Total Apps & Desktop Groups`t: " $Global:TotalAppsAndDesktopGroups
		Line 2 "Total Delivery Groups`t: " ($Global:TotalApplicationGroups+$Global:TotalDesktopGroups+$Global:TotalAppsAndDesktopGroups)
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd Application summary info"
		Line 0 "Applications"
		Line 1 "Total Published Applications`t: " $Global:TotalPublishedApplications
		Line 1 "Total App-V Applications`t: " $Global:TotalAppvApplications
		Line 2 "Total Applications`t: " ($Global:TotalPublishedApplications + $Global:TotalAppvApplications)
		Line 0 ""
		
		If($Policies -eq $True)
		{
			Write-Verbose "$(Get-Date): `tAdd Policy summary info"
			Line 0 "Policies"
			Line 1 "Total Computer Policies`t`t: " $Global:TotalComputerPolicies
			Line 1 "Total User Policies`t`t: " $Global:TotalUserPolicies
			Line 2 "Total Policies`t`t: " ($Global:TotalComputerPolicies + $Global:TotalUserPolicies)
			Line 0 ""
			Line 1 "Site Policies`t`t`t: " $Global:TotalSitePolicies
			
			If($NoADPolicies -eq $False)
			{
				Line 1 "Citrix AD Policies Processed`t: $($Global:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
				Line 1 "Citrix AD Policies not Processed: " $Global:TotalADPoliciesNotProcessed
			}
			Line 0 ""
		}
		
		Line 0 "Administrators"
		Line 1 "Total Delivery Group Admins`t: " $Global:TotalDeliveryGroupAdmins
		Line 1 "Total Full Admins`t`t: " $Global:TotalFullAdmins
		Line 1 "Total Help Desk Admins`t`t: " $Global:TotalHelpDeskAdmins
		Line 1 "Total Host Admins`t`t: " $Global:TotalHostAdmins
		Line 1 "Total Machine Catalog Admins`t: " $Global:TotalMachineCatalogAdmins
		Line 1 "Total Read Only Admins`t`t: " $Global:TotalReadOnlyAdmins
		Line 1 "Total Custom Admins`t`t: " $Global:TotalCustomAdmins
		Line 2 "Total Administrators`t: " ($Global:TotalDeliveryGroupAdmins+$Global:TotalFullAdmins+$Global:TotalHelpDeskAdmins+$Global:TotalHostAdmins+$Global:TotalMachineCatalogAdmins+$Global:TotalReadOnlyAdmins+$Global:TotalCustomAdmins)
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd Controller summary info"
		Line 0 "Controllers"
		Line 1 "Total Controllers`t`t: " $Global:TotalControllers
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd Hosting Connection summary info"
		Line 0 "Hosting Connections"
		Line 1 "Total Hosting Connections`t: " $Global:TotalHostingConnections
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd Licensing summary info"
		Line 0 "Licensing"
		$TotalLicenses = 0
		ForEach($License in $Global:Licenses)
		{
			Line 1 "$($License.LicenseProduct) $($License.LicenseType) $($License.LicenseModel)`t`t`t: $($License.LicenseCount)"
			$TotalLicenses += $License.LicenseCount
		}
		Line 2 "Total Licenses`t`t: " $TotalLicenses
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd StoreFront summary info"
		Line 0 "StoreFront"
		Line 1 "Total StoreFront Servers`t: " $Global:TotalStoreFrontServers
		Line 0 ""
		Write-Verbose "$(Get-Date): `tAdd Zone summary info"
		Line 0 "Zones"
		Line 1 "Total Zones`t`t`t: " $Global:TotalZones
	}
	ElseIf($HTML)
	{
		Write-Verbose "$(Get-Date): `tAdd Machine Catalog summary info"
		WriteHTMLLine 0 0 "Machine Catalogs"
		WriteHTMLLine 0 1 "Total Server OS Catalogs: " $Global:TotalServerOSCatalogs
		WriteHTMLLine 0 1 "Total Desktop OS Catalogs: " $Global:TotalDesktopOSCatalogs
		WriteHTMLLine 0 1 "Total RemotePC Catalogs: " $Global:TotalRemotePCCatalogs
		WriteHTMLLine 0 2 "Total Machine Catalogs: " ($Global:TotalServerOSCatalogs+$Global:TotalDesktopOSCatalogs+$Global:TotalRemotePCCatalogs)
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add AppDisks summary info"
		WriteHTMLLine 0 0 "AppDisks"
		WriteHTMLLine 0 1 "Total AppDisks: " $Global:TotalAppDisks
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add Delivery Group summary info"
		WriteHTMLLine 0 0 "Delivery Groups"
		WriteHTMLLine 0 1 "Total Application Groups: " $Global:TotalApplicationGroups
		WriteHTMLLine 0 1 "Total Desktop Groups: " $Global:TotalDesktopGroups
		WriteHTMLLine 0 1 "Total Apps & Desktop Groups: " $Global:TotalAppsAndDesktopGroups
		WriteHTMLLine 0 2 "Total Delivery Groups: " ($Global:TotalApplicationGroups+$Global:TotalDesktopGroups+$Global:TotalAppsAndDesktopGroups)
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add Application summary info"
		WriteHTMLLine 0 0 "Applications"
		WriteHTMLLine 0 1 "Total Published Applications: " $Global:TotalPublishedApplications
		WriteHTMLLine 0 1 "Total App-V Applications: " $Global:TotalAppvApplications
		WriteHTMLLine 0 2 "Total Applications: " ($Global:TotalPublishedApplications + $Global:TotalAppvApplications)
		WriteHTMLLine 0 0 ""
		
		If($Policies -eq $True)
		{
			Write-Verbose "$(Get-Date): Add Policy summary info"
			WriteHTMLLine 0 0 "Policies"
			WriteHTMLLine 0 1 "Total Computer Policies: " $Global:TotalComputerPolicies
			WriteHTMLLine 0 1 "Total User Policies: " $Global:TotalUserPolicies
			WriteHTMLLine 0 2 "Total Policies: " ($Global:TotalComputerPolicies + $Global:TotalUserPolicies)
			WriteHTMLLine 0 0 ""
			WriteHTMLLine 0 1 "Site Policies: " $Global:TotalSitePolicies
			
			If($NoADPolicies -eq $False)
			{
				WriteHTMLLine 0 1 "Citrix AD Policies Processed: $($Global:TotalADPolicies)(AD Policies can contain multiple Citrix policies)"
				WriteHTMLLine 0 1 "Citrix AD Policies not Processed: " $Global:TotalADPoliciesNotProcessed
			}
			WriteHTMLLine 0 0 ""
		}
		
		WriteHTMLLine 0 0 "Administrators"
		WriteHTMLLine 0 1 "Total Delivery Group Admins: " $Global:TotalDeliveryGroupAdmins
		WriteHTMLLine 0 1 "Total Full Admins: " $Global:TotalFullAdmins
		WriteHTMLLine 0 1 "Total Help Desk Admins: " $Global:TotalHelpDeskAdmins
		WriteHTMLLine 0 1 "Total Host Admins: " $Global:TotalHostAdmins
		WriteHTMLLine 0 1 "Total Machine Catalog Admins: " $Global:TotalMachineCatalogAdmins
		WriteHTMLLine 0 1 "Total Read Only Admins: " $Global:TotalReadOnlyAdmins
		WriteHTMLLine 0 1 "Total Custom Admins: " $Global:TotalCustomAdmins
		WriteHTMLLine 0 2 "Total Administrators: " ($Global:TotalDeliveryGroupAdmins+$Global:TotalFullAdmins+$Global:TotalHelpDeskAdmins+$Global:TotalHostAdmins+$Global:TotalMachineCatalogAdmins+$Global:TotalReadOnlyAdmins+$Global:TotalCustomAdmins)
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add Controller summary info"
		WriteHTMLLine 0 0 "Controllers"
		WriteHTMLLine 0 1 "Total Controllers: " $Global:TotalControllers
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add Hosting Connection summary info"
		WriteHTMLLine 0 0 "Hosting Connections"
		WriteHTMLLine 0 1 "Total Hosting Connections: " $Global:TotalHostingConnections
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add Licensing summary info"
		WriteHTMLLine 0 0 "Licensing"
		$TotalLicenses = 0
		ForEach($License in $Global:Licenses)
		{
			WriteHTMLLine 0 1 "$($License.LicenseProduct) $($License.LicenseType) $($License.LicenseModel): $($License.LicenseCount)"
			$TotalLicenses += $License.LicenseCount
		}
		WriteHTMLLine 0 2 "Total Licenses: " $TotalLicenses
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add StoreFront summary info"
		WriteHTMLLine 0 0 "StoreFront"
		WriteHTMLLine 0 1 "Total StoreFront Servers: " $Global:TotalStoreFrontServers
		WriteHTMLLine 0 0 ""
		Write-Verbose "$(Get-Date): Add Zone summary info"
		WriteHTMLLine 0 0 "Zones"
		WriteHTMLLine 0 1 "Total Zones: " $Global:TotalZones
	}

	Write-Verbose "$(Get-Date): Finished Create Summary Page"
	Write-Verbose "$(Get-Date): "
}
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	If(!(Check-NeededPSSnapins "Citrix.AdIdentity.Admin.V2",
	"Citrix.Analytics.Admin.V1",
	"Citrix.AppLibrary.Admin.V1",
	"Citrix.AppV.Admin.V1",
	"Citrix.Broker.Admin.V2",
	"Citrix.Common.GroupPolicy",
	"Citrix.Configuration.Admin.V2",
	"Citrix.ConfigurationLogging.Admin.V1",
	"Citrix.DelegatedAdmin.Admin.V1",
	"Citrix.EnvTest.Admin.V1",
	"Citrix.Host.Admin.V2",
	"Citrix.Licensing.Admin.V1",
	"Citrix.MachineCreation.Admin.V2",
	"Citrix.Monitor.Admin.V1",
	"Citrix.Storefront.Admin.V1"))

	# removed "Citrix.Common.Commands" as is it not used and removed in 7.13
	{
		#We're missing Citrix Snapins that we need
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`nMissing Citrix PowerShell Snap-ins Detected, check the console above for more information. 
		`nAre you sure you are running this script against a XenDesktop 7.8 or later Controller? 
		`n`nIf you are running the script remotely, did you install Studio or the PowerShell snapins on $($env:computername)?
		`n`nPlease see the Prerequisites section in the ReadMe file (https://dl.dropboxusercontent.com/u/43555945/XD7_Inventory_V1_ReadMe.rtf).
		`n`nScript will now close."
		Exit
	}

	$Global:DoPolicies = $True
	If(!(Check-LoadedModule "Citrix.GroupPolicy.Commands") -and $Policies -eq $False)
	{
		Write-Warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded `n
		Please see the Prerequisites section in the ReadMe file (https://dl.dropboxusercontent.com/u/43555945/XD7_Inventory_V1_ReadMe.rtf). 
		`nCitrix Policy documentation will not take place"
		Write-Verbose "$(Get-Date): "
		$Global:DoPolicies = $False
	}
	ElseIf(!(Check-LoadedModule "Citrix.GroupPolicy.Commands") -and $Policies -eq $True)
	{
		Write-Error "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded 
		`nPlease see the Prerequisites section in the ReadMe file (https://dl.dropboxusercontent.com/u/43555945/XD7_Inventory_V1_ReadMe.rtf). 
		`n
		`n
		`t`tBecause the Policies parameter was used the script will now close.
		`n
		`n"
		Write-Verbose "$(Get-Date): "
		Exit
	}
	
	If($Policies -eq $False -and $NoPolicies -eq $False -and $NoADPolicies -eq $False)
	{
		#script defaults, so don't process policies
		$Global:DoPolicies = $False
	}
	If($NoPolicies -eq $True)
	{
		#don't process policies
		$Global:DoPolicies = $False
	}
	
	#set value for MaxRecordCount
	$Script:MaxRecordCount = [int]::MaxValue 

	If([String]::IsNullOrEmpty($AdminAddress))
	{
		$AdminAddress = "LocalHost"
	}

	$Script:XDParams1 = @{
	adminaddress = $AdminAddress; 
	EA = 0;
	}

	$Script:XDParams2 = @{
	adminaddress = $AdminAddress; 
	EA = 0;
	MaxRecordCount = $Script:MaxRecordCount;
	}

	# Get Site information
	Write-Verbose "$(Get-Date): Gathering initial Site data"

	$Script:XDSite1 = Get-BrokerSite @XDParams1

	If( !($?) -or $Null -eq $Script:XDSite1)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "XenDesktop Site1 information could not be retrieved.  Script cannot continue"
		Write-Error "cmdlet failed $($error[ 0 ].ToString())"
		AbortScript
	}

	$Script:XDSite2 = Get-ConfigSite @XDParams1

	If( !($?) -or $Null -eq $Script:XDSite2)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "XenDesktop Site2 information could not be retrieved.  Script cannot continue"
		Write-Error "cmdlet failed $($error[ 0 ].ToString())"
		AbortScript
	}

	#changed 18-dec-2016 to allow 32-bit PoSH to get the data in the 64-bit registry location
	#initial idea from WC at Citrix and also from http://stackoverflow.com/questions/630382/how-to-access-the-64-bit-registry-from-a-32-bit-powershell-instance reply from SergVro
	$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
	$subKey =  $key.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Citrix Desktop Delivery Controller")
	$value = $subKey.GetValue("DisplayVersion")
	$Script:XDSiteVersion = $value.Substring(0,4)
	$tmp = $Script:XDSiteVersion.Split(".")
	[int]$MajorVersion = $tmp[0]
	[int]$MinorVersion = $tmp[1]
	Write-Verbose "$(Get-Date): You are running version $($value)"
	Write-Verbose "$(Get-Date): Major version $($MajorVersion)"
	Write-Verbose "$(Get-Date): Minor version $($MinorVersion)"

	#first check to make sure this is a Site between 7.0 and 7.7
	If($MajorVersion -eq 7)
	{
		#this is a XenDesktop 7.x Site, now test to make sure it is less than 7.8
		If($MinorVersion -lt 8)
		{
			Write-Warning "You are running version $($value)"
			Write-Warning "Are the PowerShell Snapins or Studio installed?"
			Write-Warning "This script is designed for XenDesktop 7.8 and later and should not be run on 7.7 and earlier.`n`nScript cannot continue`n"
			AbortScript
		}
	}
	Else
	{
		#this is not a XenDesktop 7.x Site, script cannot proceed
		Write-Warning "You are running version $($value)"
		Write-Warning "Are the PowerShell Snapins or Studio installed?"
		Write-Warning "This script is designed for XenDesktop 7.8 and later and should not be run on other versions of XenDesktop.`n`nScript cannot continue`n"
		AbortScript
	}

	[string]$Script:XDSiteName = $Script:XDSite2.SiteName
	Switch ($Section)
	{
		"Admins"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Administrators Only)"; Break}
		"AppDisks"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (AppDisks Only)"; Break}
		"AppDNA"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (AppDNA Only)"; Break}
		"Apps"			{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Applications Only)"; Break}
		"AppV"			{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (App-V Only"; Break}
		"Catalogs"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Machine Catalogs Only)"; Break}
		"Config"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Configuration Only)"; Break}
		"Controllers"	{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Controllers Only)"; Break}
		"Groups" 		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Delivery Groups Only)"; Break}
		"Hosting"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Hosting Only)"; Break}
		"Licensing"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Licensing Only)"; Break}
		"Logging"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Configuration Logging Only"; Break}
		"Policies"		{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Policies Only)"; Break}
		"StoreFront"	{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (StoreFront Only)"; Break}
		"Zones"			{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site (Zones Only)"; Break}
		"All"			{[string]$Script:Title = "Inventory Report for the $($Script:XDSiteName) Site"; Break}
	}
	Write-Verbose "$(Get-Date): Initial Site data has been gathered"
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($pwd.Path)\XAXDV2InventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime   : $($AddDateTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "AdminAddress   : $($AdminAddress)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Administrators : $($Administrators)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Applications   : $($Applications)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name   : $($Script:CoName)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page     : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "DeliveryGroups : $($DeliveryGroups)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev            : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile   : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "DGUtilization  : $($DeliveryGroupsUtilization)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Filename1      : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2      : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder         : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From           : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Hosting        : $($Hosting)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "HW Inventory   : $($Hardware)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Logging        : $($Logging)" 4>$Null
		If($Logging)
		{
			Out-File -FilePath $SIFile -Append -InputObject "   Start Date   : $($StartDate)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "   End Date     : $($EndDate)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "MachineCatalogs: $($MachineCatalogs)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "NoADPolicies   : $($NoADPolicies)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "NoPolicies     : $($NoPolicies)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Policies       : $($Policies)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML   : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF    : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT   : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD   : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info    : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Section        : $($Section)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Site Name      : $($XDSiteName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port      : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server    : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title          : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To             : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL        : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name      : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "XA/XD Version  : $($Script:XDSiteVersion)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected    : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version   : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture      : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture    : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language  : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version   : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start   : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time   : $($Str)" 4>$Null
	}

	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

SetFileName1andFileName2 "$($Script:XDSiteName)"

If($Section -eq "All" -or $Section -eq "Catalogs")
{
	ProcessMachineCatalogs
}

If($Section -eq "All" -or $Section -eq "AppDisks")
{
	ProcessAppDisks
}

If($Section -eq "All" -or $Section -eq "Groups")
{
	ProcessDeliveryGroups
}

If($Section -eq "All" -or $Section -eq "Apps")
{
	ProcessApplications
	ProcessApplicationGroupDetails
}

If($Section -eq "All" -or $Section -eq "Policies")
{
	If($NoPolicies -or $Global:DoPolicies -eq $False)
	{
		#don't process policies
	}
	Else
	{
		ProcessPolicies
	}
}

If($Section -eq "All" -or $Section -eq "Logging")
{
	ProcessConfigLogging
}

If($Section -eq "All" -or $Section -eq "Config")
{
	ProcessConfiguration
}

If($Section -eq "All" -or $Section -eq "Admins")
{
	ProcessAdministrators
	ProcessScopes
	ProcessRoles
}

If($Section -eq "All" -or $Section -eq "Controllers")
{
	ProcessControllers
}

If($Section -eq "All" -or $Section -eq "Hosting")
{
	ProcessHosting
}

If($Section -eq "All" -or $Section -eq "Licensing")
{
	ProcessLicensing
}

If($Section -eq "All" -or $Section -eq "StoreFront")
{
	ProcessStoreFront
}

If($Section -eq "All" -or $Section -eq "AppV")
{
	ProcessAppV
}

If($Section -eq "All" -or $Section -eq "AppDNA")
{
	ProcessAppDNA
}

If($Section -eq "All" -or $Section -eq "Zones")
{
	If((Get-ConfigServiceAddedCapability @XDParams1) -contains "ZonesSupport")
	{
		ProcessZones
	}
}

If($Section -eq "All")
{
	ProcessSummaryPage
}
#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "Citrix XenDesktop $($Script:XDSiteVersion) Inventory"
$SubjectTitle = "XenDesktop $($Script:XDSiteVersion) Site Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#endregion
# SIG # Begin signature block
# MIIgCgYJKoZIhvcNAQcCoIIf+zCCH/cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfTzXaddzewYPKSQyo+t7uREx
# kPagghtxMIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMzExMTEwMDAwMDAwWjBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtDhXO5EOAXLGH87dg
# +XESpa7cJpSIqvTO9SA5KFhgDPiA2qkVlTJhPLWxKISKityfCgyDF3qPkKyK53lT
# XDGEKvYPmDI2dsze3Tyoou9q+yHyUmHfnyDXH+Kx2f4YZNISW1/5WBg1vEfNoTb5
# a3/UsDg+wRvDjDPZ2C8Y/igPs6eD1sNuRMBhNZYW/lmci3Zt1/GiSw0r/wty2p5g
# 0I6QNcZ4VYcgoc/lbQrISXwxmDNsIumH0DJaoroTghHtORedmTpyoeb6pNnVFzF1
# roV9Iq4/AUaG9ih5yLHa5FcXxH4cDrC0kqZWs72yl+2qp/C3xag/lRbQ/6GW6whf
# GHdPAgMBAAGjYzBhMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBRF66Kv9JLLgjEtUYunpyGd823IDzAfBgNVHSMEGDAWgBRF66Kv9JLL
# gjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAog683+Lt8ONyc3pklL/3
# cmbYMuRCdWKuh+vy1dneVrOfzM4UKLkNl2BcEkxY5NM9g0lFWJc1aRqoR+pWxnmr
# EthngYTffwk8lOa4JiwgvT2zKIn3X/8i4peEH+ll74fg38FnSbNd67IJKusm7Xi+
# fT8r87cmNW1fiQG2SVufAQWbqz0lwcy2f8Lxb4bG+mRo64EtlOtCt/qMHt1i8b5Q
# Z7dsvfPxH2sMNgcWfzd8qVttevESRmCD1ycEvkvOl77DZypoEd+A5wwzZr8TDRRu
# 838fYxAe+o0bJW1sj6W3YQGx0qMmoRBxna3iw/nDmVG3KwcIzi7mULKn+gpFL6Lw
# 8jCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAw
# ZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBS
# b290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/
# DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2
# qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrsk
# acLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/
# 6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE
# 94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8
# np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0w
# azAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUF
# BzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3Js
# ME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823I
# DzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh
# 134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63X
# X0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPA
# JRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC
# /i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG
# /AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBT8wggQnoAMC
# AQICEALKvIFdDaFKh3T2QAUcJiIwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2ln
# bmluZyBDQTAeFw0xNjEwMTgwMDAwMDBaFw0xNzEwMjMxMjAwMDBaMHwxCzAJBgNV
# BAYTAlVTMQswCQYDVQQIEwJUTjESMBAGA1UEBxMJVHVsbGFob21hMSUwIwYDVQQK
# ExxDYXJsIFdlYnN0ZXIgQ29uc3VsdGluZywgTExDMSUwIwYDVQQDExxDYXJsIFdl
# YnN0ZXIgQ29uc3VsdGluZywgTExDMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEAqGo3KHWZmWVSao7Ur+ldBIwwM7v4tM7NQ3X3A9H7DqfGXSVvWvVj5zbc
# zX1yns9Qot1bnrTRLlnimIPJa+GieuEz7ON7jpzQjErmuzJz4HBEfbfAqoVuVmpy
# dsPpxfNqWMQt+0YqeEgYZqoF5mIXK2ACugsQz5e9SMWEsR9Z0s9FQyjEnIKuhQYq
# cLY7y85/CNsH4qgKNoHPfZ+LlPaWFfHCI7XIleLC2QHcLlEe760NDv163eXq6rkC
# tJroHqT4WKeXEEj14nhFNxSp/UUuk004/ju5Pb1gsgOYxkQ94BrixMW9zYghXX2H
# K3JzL8O56djKJuD8em8whmpXAmR6FQIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAU
# WsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFPC0K6tjLci4jiul81bZG+CS
# MocaMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8E
# cDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTIt
# YXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEw
# gYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/
# BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAAwIwqMrUHX/2xnjs13V3ikCzJ+LkAMXu
# z4daOhkO5EdCkE8Cl9nnKtVGEVnC8v2xkUSgDWb9yAoGJfOx8oamS6IA3J1C+lND
# 8cKJwb70FAHzQV+Tyzmwm38VavUC0kc27iE5kfziUOU+UH/bZYwmeo1Z54SiooEB
# atp1RYmvbwE8ATyme/KmYkfbUkYlbfpP0aWGey33sKGiI8ZmWUC4PSDWQ+aXiAWv
# YZQXUiGQTWleWvmhlpSVATho62Db2KuE3hsR8v1wLY3s/WPs0OyhrBD80ExWiX/q
# HoQGTmaBGz0SczPU0sfro1gKghTUr96046UFQQjeybpebrrlMLwcGDCCBmowggVS
# oAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEFBQAwYjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMB4XDTE0
# MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UEBhMCVVMxETAPBgNV
# BAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3RhbXAgUmVzcG9u
# ZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz4II14OD2
# xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmDzm9m7t3LhelfpfnU
# h3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YKZ6O+YZ+u8/0SeHUO
# plsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGIYXIYaLm4fO7m5zQv
# MXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4eMfJBi5GEMiN6ARg
# 27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAIzGvsYkKRrALA76Tw
# iRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIBsjCCAaEGCWCG
# SAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNv
# bS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYA
# IAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkA
# dAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAA
# RABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAA
# UgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAA
# dwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4A
# ZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkA
# bgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMB8GA1Ud
# IwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRhWk0ktkkynUoq
# eRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYyaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmwwdwYIKwYB
# BQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCdJX4bM02yJoFc
# m4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITkWkD73gYBjDf6m7Gd
# JH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2P+fiEUGmvWLZ8Cc9
# OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr849Dp3GdId0UyhVd
# kkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzbXEgnZsijiwoc5ZXa
# rsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DUbuD0FAo6G+OPPcqv
# ao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEw
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKe
# VlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBx
# MevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7
# os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwr
# mdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c
# 9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJ
# AgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUH
# AwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYD
# VR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0
# cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB
# /wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStn
# As0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEF
# BQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKc
# lcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuL
# LZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1f
# wksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFp
# jE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNH
# lXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQDMIID/wIBATCBhjByMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBT
# aWduaW5nIENBAhACyryBXQ2hSod09kAFHCYiMAkGBSsOAwIaBQCgQDAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAjBgkqhkiG9w0BCQQxFgQUQ0Eb3j4LvKBSCW+T
# HmjerLXxppQwDQYJKoZIhvcNAQEBBQAEggEASV3TjaXh+svgiBGxiEwLSqua3Wk9
# y5FLksgs7PI+nJhEWnwnawKX4goF+kl6m59YGZkwOkkDJxDj9mqh6a68+LlRfl9i
# TUlXguITJJTvKRrSJ38wRW7AX1SdFeIHDvfoTnMFUIAKKwodW7ahFSCzhZnVtSkT
# AgGaBGB7jiwbXiZu7lcawlbP1rv8V8rXrYkT1jmFrS/1G9ctmOLUcqH/1TbwtuHt
# rg7MlVzJigJa/WynlwnMb4Kbs3pOSseEoLVmh1cEmo0ziIrbd3+NjSltskYSlM6g
# LjY2I+gAHcBfBJGKrDYNR8/ELqIFEFzUOyr2x9h/qktdr0X2/PWBT6HQa6GCAg8w
# ggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfw
# ZjAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTcwMzA2MTYwOTM1WjAjBgkqhkiG9w0BCQQxFgQU2r3+CyTpJZri
# 7NgD/3MtGMyA/ncwDQYJKoZIhvcNAQEBBQAEggEAHVLgxMAH0uTc4r3wqUB3lmzj
# B4XdzKprR7gHc1t5FWFxGMzWStE4gowPazukUPdZlAqyoqWjr3hkFt4Df6zaTm+U
# JzztavMx9hKbpXwPR1KlyEL0UlzMplKBvX5euscFqa5T2ipsgONLd9+kyYNclbBn
# GQF/tr5y/vI7dSiIRsf+Yu7tI1gpAeH5oXVI7xT1zGBjJHmIjRBUtxNYejCc/HjY
# 1LdxuKqF2BhaG2Lj5FJGxnPhus5r0+4tv/RTKEd40FbQzP+qsft8C6s0XKuZsyRT
# yxnwrKeRweIqWD1s8aQ4rB7zBIVB3X/y7wSKq5LiIxdrve4VLDmR6JEBvY/YpA==
# SIG # End signature block
