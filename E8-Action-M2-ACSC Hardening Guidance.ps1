
# Control Set from : https://www.cyber.gov.au/resources-business-and-government/maintaining-devices-and-systems/system-hardening-and-administration/system-hardening/hardening-microsoft-365-office-2021-office-2019-and-office-2016

#**************************************************************************Flash Content**********************************************************************************************

# Block Flash activation in Office documents
if (!(Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility")) {
    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Force | Out-Null
}

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Name "ActivationFilterOverride" -Value 0 -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Name "Compatibility Flags" -Value 1024 -PropertyType DWORD -Force | Out-Null

#*******************************************************************Loading external content*****************************************************************************************

# Dynamic Data Exchange
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "DataConnectionWarnings" -Value 2 -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "RichDataConnectionWarnings" -Value 2 -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "WorkbookLinkWarnings" -Value 2 -PropertyType DWORD -Force | Out-Null

if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AllowDDE" -Value 0 -PropertyType DWORD -Force | Out-Null

# Always prevent untrusted Microsoft Query files from opening
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "enableblockunsecurequeryfiles" -Value 1 -PropertyType DWORD -Force | Out-Null

# Don't allow Dynamic Data Exchange (DDE) server launch in Excel
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "disableddeserverlaunch" -Value 1 -PropertyType DWORD -Force | Out-Null

# Don't allow Dynamic Data Exchange (DDE) server lookup in Excel
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "disableddeserverlookup" -Value 1 -PropertyType DWORD -Force | Out-Null

# Update automatic links at Open
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\word\options")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\word\options" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\options" -Name "dontupdatelinks" -Value 1 -PropertyType DWORD -Force | Out-Null


#*************************************************************Object Linking and Embedding packages**********************************************************************************

# Excel
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "PackagerPrompt" -Value 2 -PropertyType DWORD -Force | Out-Null

# PowerPoint
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "PackagerPrompt" -Value 2 -PropertyType DWORD -Force | Out-Null

# Word
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "PackagerPrompt" -Value 2 -PropertyType DWORD -Force | Out-Null


#***************************************************************************ActiveX**********************************************************************************************

#  Disable All ActiveX
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\Common\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\Common\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Common\Security" -Name "DisableAllActiveX" -Value 1 -PropertyType DWORD -Force | Out-Null


#*************************************************************************Add-Ins************************************************************************************************
# Disable all application add-ins

# Excel
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "DisableAllAddins" -Value 1 -PropertyType DWORD -Force | Out-Null

# PowerPoint
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "DisableAllAddins" -Value 1 -PropertyType DWORD -Force | Out-Null

# Project
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Project\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Project\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Project\Security" -Name "DisableAllAddins" -Value 1 -PropertyType DWORD -Force | Out-Null

# Visio
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Visio\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Visio\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Visio\Security" -Name "DisableAllAddins" -Value 1 -PropertyType DWORD -Force | Out-Null

# Word
if (!(Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "DisableAllAddins" -Value 1 -PropertyType DWORD -Force | Out-Null



#**************************************************************************Extention Hardening*************************************************************************************

# Extension hardening
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security" -Name "extensionhardening" -Value 2 -PropertyType DWORD -Force | Out-Null


#***********************************************************************File Type Blocking****************************************************************************************
# Excel file block settings
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Force | Out-Null
}

# dBase III / IV files
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "dbasefiles" -Value 2 -PropertyType DWORD -Force | Out-Null
# Dif and Sylk files
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "difandsylkfiles" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 2 macrosheets and add-in files
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl2macros" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 2 worksheets
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl2worksheets" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 3 macrosheets and add-in files
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl3macros" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 3 worksheets
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl3worksheets" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 4 macrosheets and add-in files
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl4macros" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 4 workbooks
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl4workbooks" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 4 worksheets
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl4worksheets" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 95 workbooks
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl95workbooks" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 95-97 workbooks and templates
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl9597workbooksandtemplates" -Value 2 -PropertyType DWORD -Force | Out-Null
# Excel 97-2003 workbooks and templates
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl97addins" -Value 2 -PropertyType DWORD -Force | Out-Null
# Set default file block behavior
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "openinprotectedview" -Value 2 -PropertyType DWORD -Force | Out-Null
# Web pages and Excel 2003 XML spreadsheets
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "htmlandxmlssfiles" -Value 2 -PropertyType DWORD -Force | Out-Null

# PowerPoint file block settings
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security\fileblock")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security\fileblock" -Force | Out-Null
}

# PowerPoint 97-2003 presentations, shows, templates and add-in files
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security\fileblock" -Name "binaryfiles" -Value 2 -PropertyType DWORD -Force | Out-Null
# Set default file block behavior
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security\fileblock" -Name "openinprotectedview" -Value 1 -PropertyType DWORD -Force | Out-Null

# Visio file block settings
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Force | Out-Null
}

# Visio 2000-2002 Binary Drawings, Templates and Stencils
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Name "visio2000files" -Value 2 -PropertyType DWORD -Force | Out-Null
# Visio 2003-2010 Binary Drawings, Templates and Stencils
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Name "visio2003files" -Value 2 -PropertyType DWORD -Force | Out-Null
# Visio 5.0 or earlier Binary Drawings, Templates and Stencils
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Name "visio50andearlierfiles" -Value 2 -PropertyType DWORD -Force | Out-Null

# Word file block settings
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Force | Out-Null
}

# Set default file block behavior - Word
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "openinprotectedview" -Value 2 -PropertyType DWORD -Force | Out-Null



#*************************************************************************Office File Validation*****************************************************************************
# Turn off file validation
# Turn off file validation in Excel
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Excel\security\filevalidation")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Excel\security\filevalidation" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Excel\security\filevalidation" -Name "enableonload" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off file validation in PowerPoint
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Name "enableonload" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off file validation in Word
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Name "enableonload" -Value 1 -PropertyType DWORD -Force | Out-Null


#*************************************************************************Running external programs****************************************************************************

# Run Programs
# Check if the registry path exists
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security" -Name "runprograms" -Value 0 -PropertyType DWORD -Force | Out-Null




#*********************************************************************Protected View*******************************************************************************************

# Always open untrusted database files in Protected View
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "enabledatabasefileprotectedview" -Value 1 -PropertyType DWORD -Force | Out-Null

# Do not open files from the Internet zone in Protected View
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "disableinternetfilesinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Do not open files in unsafe locations in Protected View
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "disableunsafelocationsinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Set document behaviour if file validation fails
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\excel\security\filevalidation")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\excel\security\filevalidation" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\filevalidation" -Name "openinprotectedview" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off Protected View for attachments opened from Outlook
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "disableattachmentsinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Do not open files from the Internet zone in Protected View in PowerPoint
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableinternetfilesinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Do not open files in unsafe locations in Protected View in PowerPoint
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableunsafelocationsinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Turn off Protected View for attachments opened from Outlook in PowerPoint
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Name "openinprotectedview" -Value 1 -PropertyType DWORD -Force | Out-Null

# Add registry key and value for disabling attachments in protected view in PowerPoint
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableattachmentsinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Add registry key and value for disabling internet files in protected view in Word
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableinternetfilesinpv" -Value 0 -PropertyType DWORD -Force | Out-Null

# Do not open files in unsafe locations in Protected View in Word
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableunsafelocationsinpv" -Value 1 -PropertyType DWORD -Force | Out-Null

# Set document behaviour if file validation fails in Word
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Name "openinprotectedview" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off Protected View for attachments opened from Outlook in Word
New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableattachmentsinpv" -Value 0 -PropertyType DWORD -Force | Out-Null



#************************************************************************Trusted documents************************************************************************************


# Turn off trusted documents in Excel
if (!(Test-Path -Path "HKCU:\software\microsoft\Excel\16.0\access\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\Excel\16.0\access\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\Excel\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off Trusted Documents on the network in Excel
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Excel\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Excel\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Excel\security\trusted documents" -Name "disablenetworktrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off trusted documents in PowerPoint
if (!(Test-Path -Path "HKCU:\software\microsoft\PowerPoint\16.0\access\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\PowerPoint\16.0\access\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\PowerPoint\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off Trusted Documents on the network in PowerPoint
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\trusted documents" -Name "disablenetworktrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off trusted documents in Visio
if (!(Test-Path -Path "HKCU:\software\microsoft\Visio\16.0\access\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\Visio\16.0\access\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\Visio\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off Trusted Documents on the network in Visio
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Visio\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Visio\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Visio\security\trusted documents" -Name "disablenetworktrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off trusted documents in Word
if (!(Test-Path -Path "HKCU:\software\microsoft\Word\16.0\access\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\Word\16.0\access\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\Word\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null

# Turn off Trusted Documents on the network in Word
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Word\security\trusted documents")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Word\security\trusted documents" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\trusted documents" -Name "disablenetworktrusteddocuments" -Value 1 -PropertyType DWORD -Force | Out-Null



#*************************************************************************************Hidden markup*********************************************************************************
# Make hidden markup visible in PowerPoint
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\powerpoint\options")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\powerpoint\options" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\options" -Name "markupopensave" -Value 1 -PropertyType DWORD -Force | Out-Null

# Make hidden markup visible in Word
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\Word\options")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\Word\options" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\options" -Name "markupopensave" -Value 1 -PropertyType DWORD -Force | Out-Null



#****************************************************************************************Reporting Information*********************************************************************

# Allow including screenshot with Office Feedback
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common\feedback")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common\feedback" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\feedback" -Name "includescreenshot" -Value 0 -PropertyType DWORD -Force | Out-Null

# Automatically receive small updates to improve reliability
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common" -Name "updatereliabilitydata" -Value 0 -PropertyType DWORD -Force | Out-Null

# Configure the type of diagnostic data sent by Office to Microsoft
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common\clienttelemetry")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common\clienttelemetry" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\clienttelemetry" -Name "sendtelemetry" -Value 1 -PropertyType DWORD -Force | Out-Null

# Disable Opt-in Wizard on first run
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common\general")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common\general" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\general" -Name "shownfirstrunoptin" -Value 1 -PropertyType DWORD -Force | Out-Null

# Enable Customer Experience Improvement Program
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common" -Name "qmenable" -Value 0 -PropertyType DWORD -Force | Out-Null

# Send Office Feedback
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common\feedback")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common\feedback" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\feedback" -Name "enabled" -Value 0 -PropertyType DWORD -Force | Out-Null

# Send personal information
if (!(Test-Path -Path "HKCU:\software\microsoft\office\16.0\common")) {
    New-Item -Path "HKCU:\software\microsoft\office\16.0\common" -Force | Out-Null
}

New-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common" -Name "sendcustomerdata" -Value 0 -PropertyType DWORD -Force | Out-Null