
# Control Set from : https://www.cyber.gov.au/resources-business-and-government/maintaining-devices-and-systems/system-hardening-and-administration/system-hardening/hardening-microsoft-365-office-2021-office-2019-and-office-2016

#********Flash Content************

#Block Flash activation in Office documents
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Name "ActivationFilterOverride" -Force -ErrorAction SilentlyContinue 
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Name "Compatibility Flags" -Force -ErrorAction SilentlyContinue

#*******Loading external content******

#Dynamic Data Exchange
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "DataConnectionWarnings" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "RichDataConnectionWarnings" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "WorkbookLinkWarnings" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AllowDDE" -Force -ErrorAction SilentlyContinue
#Always prevent untrusted Microsoft Query files from opening
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "enableblockunsecurequeryfiles" -Force -ErrorAction SilentlyContinue
#Don’t allow Dynamic Data Exchange (DDE) server launch in Excel
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "disableddeserverlaunch" -Force -ErrorAction SilentlyContinue
#Don’t allow Dynamic Data Exchange (DDE) server lookup in Excel
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "disableddeserverlookup" -Force -ErrorAction SilentlyContinue
# Update automatic links at Open
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\options" -Name "dontupdatelinks" -Force -ErrorAction SilentlyContinue

#*********Object Linking and Embedding packages**************

# Excel
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "PackagerPrompt" -Force -ErrorAction SilentlyContinue
# PowerPoint
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "PackagerPrompt" -Force -ErrorAction SilentlyContinue
# Word
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "PackagerPrompt" -Force -ErrorAction SilentlyContinue


#***********ActiveX****************
# Disable All ActiveX
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Common\Security" -Name "DisableAllActiveX" -Force -ErrorAction SilentlyContinue

#***********Add-Ins****************
# Disable all application add-ins
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "DisableAllAddins" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "DisableAllAddins" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Project\Security" -Name "DisableAllAddins" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Visio\Security" -Name "DisableAllAddins" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "DisableAllAddins" -Force -ErrorAction SilentlyContinue


#***********Extention Hardening****************
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security" -Name "extensionhardening" -Force -ErrorAction SilentlyContinue

#***********File Type Blocking****************
#dBase III / IV files
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "dbasefiles" -Force -ErrorAction SilentlyContinue
#Dif and Sylk files
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "difandsylkfiles" -Force -ErrorAction SilentlyContinue
#Excel 2 macrosheets and add-in files
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl2macros" -Force -ErrorAction SilentlyContinue
#Excel 2 worksheets
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl2worksheets" -Force -ErrorAction SilentlyContinue
#Excel 3 macrosheets and add-in files
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl3macros" -Force -ErrorAction SilentlyContinue
#Excel 3 worksheets
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl3worksheets" -Force -ErrorAction SilentlyContinue
#Excel 4 macrosheets and add-in files
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl4macros" -Force -ErrorAction SilentlyContinue
#Excel 4 workbooks
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl4workbooks" -Force -ErrorAction SilentlyContinue
#Excel 4 worksheets
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl4worksheets" -Force -ErrorAction SilentlyContinue
#Excel 95 workbooks
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl95workbooks" -Force -ErrorAction SilentlyContinue
#Excel 95-97 workbooks and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl9597workbooksandtemplates" -Force -ErrorAction SilentlyContinue
#Excel 97-2003 workbooks and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "xl97addins" -Force -ErrorAction SilentlyContinue
#Set default file block behavior
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "openinprotectedview" -Force -ErrorAction SilentlyContinue
#Web pages and Excel 2003 XML spreadsheets
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\fileblock" -Name "htmlandxmlssfiles" -Force -ErrorAction SilentlyContinue
# PowerPoint 97-2003 presentations, shows, templates and add-in files
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security\fileblock" -Name "binaryfiles" -Force -ErrorAction SilentlyContinue
#Set default file block behavior
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security\fileblock" -Name "openinprotectedview" -Force -ErrorAction SilentlyContinue
# Visio 2000-2002 Binary Drawings, Templates and Stencils
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Name "visio2000files" -Force -ErrorAction SilentlyContinue
#Visio 2003-2010 Binary Drawings, Templates and Stencils
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Name "visio2003files" -Force -ErrorAction SilentlyContinue
#Visio 5.0 or earlier Binary Drawings, Templates and Stencils
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\visio\security\fileblock" -Name "visio50andearlierfiles" -Force -ErrorAction SilentlyContinue
# Set default file block behavior - Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\fileblock" -Name "openinprotectedview" -Force -ErrorAction SilentlyContinue
#Word 2 and earlier binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word2files" -Force -ErrorAction SilentlyContinue
#Word 2000 binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word2000files" -Force -ErrorAction SilentlyContinue
#Word 2003 binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word2003files" -Force -ErrorAction SilentlyContinue
#Word 2007 and later binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word2007files" -Force -ErrorAction SilentlyContinue
#Word 6.0 binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word60files" -Force -ErrorAction SilentlyContinue
#Word 95 binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word95files" -Force -ErrorAction SilentlyContinue
#Word 97 binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "word97files" -Force -ErrorAction SilentlyContinue
#Word XP binary documents and templates
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\word\security\fileblock" -Name "wordxpfiles" -Force -ErrorAction SilentlyContinue

#*************Office File Validation*****************
# Turn off file validation
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Excel\security\filevalidation" -Name "enableonload" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Name "enableonload" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Name "enableonload" -Force -ErrorAction SilentlyContinue

#*************Running external programs****************
# Run Programs
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\security" -Name "runprograms" -Force -ErrorAction SilentlyContinue



#**************Protected View**********************
# Always open untrusted database files in Protected View
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "enabledatabasefileprotectedview" -Force -ErrorAction SilentlyContinue
# Do not open files from the Internet zone in Protected View
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "disableinternetfilesinpv" -Force -ErrorAction SilentlyContinue
# Do not open files in unsafe locations in Protected View
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "disableunsafelocationsinpv" -Force -ErrorAction SilentlyContinue
# Set document behaviour if file validation fails
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\filevalidation" -Name "openinprotectedview" -Force -ErrorAction SilentlyContinue
# Turn off Protected View for attachments opened from Outlook
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\excel\security\protectedview" -Name "disableattachmentsinpv" -Force -ErrorAction SilentlyContinue
# Do not open files from the Internet zone in Protected View in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableinternetfilesinpv" -Force -ErrorAction SilentlyContinue
# Do not open files in unsafe locations in Protected View in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableunsafelocationsinpv" -Force -ErrorAction SilentlyContinue
# Turn off Protected View for attachments opened from Outlook in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Name "openinprotectedview" -Force -ErrorAction SilentlyContinue
# Add registry key and value for disabling attachments in protected view in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableattachmentsinpv" -Force -ErrorAction SilentlyContinue
# Add registry key and value for disabling internet files in protected view in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableinternetfilesinpv" -Force -ErrorAction SilentlyContinue
# Do not open files in unsafe locations in Protected View in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableunsafelocationsinpv" -Force -ErrorAction SilentlyContinue
# Set document behaviour if file validation fails in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Name "openinprotectedview" -Force -ErrorAction SilentlyContinue
# Turn off Protected View for attachments opened from Outlook in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableattachmentsinpv" -Force -ErrorAction SilentlyContinue

#**********Trusted documents***************


# Turn off trusted documents in Excel
Remove-ItemProperty -Path "HKCU:\software\microsoft\Excel\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off Trusted Documents on the network in Excel
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Excel\security\trusted documents" -Name "disablenetworktrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off trusted documents in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\PowerPoint\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off Trusted Documents on the network in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\trusted documents" -Name "disablenetworktrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off trusted documents in Visio
Remove-ItemProperty -Path "HKCU:\software\microsoft\Visio\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off Trusted Documents on the network in Visio
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Visio\security\trusted documents" -Name "disablenetworktrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off trusted documents in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\Word\16.0\access\security\trusted documents" -Name "disabletrusteddocuments" -Force -ErrorAction SilentlyContinue
# Turn off Trusted Documents on the network in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\security\trusted documents" -Name "disablenetworktrusteddocuments" -Force -ErrorAction SilentlyContinue


#************Hidden markup*******************
# Make hidden markup visible in PowerPoint
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\powerpoint\options" -Name "markupopensave" -Force -ErrorAction SilentlyContinue
# Make hidden markup visible in Word
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\Word\options" -Name "markupopensave" -Force -ErrorAction SilentlyContinue


#************Reporting Information*******************

# Allow including screenshot with Office Feedback
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\feedback" -Name "includescreenshot" -Force -ErrorAction SilentlyContinue
# Automatically receive small updates to improve reliability
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common" -Name "updatereliabilitydata" -Force -ErrorAction SilentlyContinue
# Configure the type of diagnostic data sent by Office to Microsoft
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\clienttelemetry" -Name "sendtelemetry" -Force -ErrorAction SilentlyContinue
# Disable Opt-in Wizard on first run
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\general" -Name "shownfirstrunoptin" -Force -ErrorAction SilentlyContinue
# Enable Customer Experience Improvement Program
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common" -Name "qmenable" -Force -ErrorAction SilentlyContinue
# Send Office Feedback
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common\feedback" -Name "enabled" -Force -ErrorAction SilentlyContinue
# Send personal information
Remove-ItemProperty -Path "HKCU:\software\microsoft\office\16.0\common" -Name "sendcustomerdata" -Force -ErrorAction SilentlyContinue