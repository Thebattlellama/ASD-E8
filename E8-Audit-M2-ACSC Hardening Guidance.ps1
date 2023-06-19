# Control Set from : https://www.cyber.gov.au/resources-business-and-government/maintaining-devices-and-systems/system-hardening-and-administration/system-hardening/hardening-microsoft-365-office-2021-office-2019-and-office-2016

#**************************************************************************Flash Content**********************************************************************************************

# Block Flash activation in Office documents
$activationFilterOverride = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Name "ActivationFilterOverride" -ErrorAction SilentlyContinue
$compatibilityFlags = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\COM Compatibility" -Name "Compatibility Flags" -ErrorAction SilentlyContinue

if ($activationFilterOverride -eq $null -or $compatibilityFlags -eq $null) {
    Write-Host "Registry keys not found. Expected values cannot be determined."
}
else {
    Write-Host "Expected value for 'ActivationFilterOverride' is 0. Current value is $activationFilterOverride."
    Write-Host "Expected value for 'Compatibility Flags' is 1024. Current value is $compatibilityFlags."
}


#*******************************************************************Loading external content*****************************************************************************************
# Dynamic Data Exchange
$excelDataConnectionWarnings = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "DataConnectionWarnings" -ErrorAction SilentlyContinue
$excelRichDataConnectionWarnings = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "RichDataConnectionWarnings" -ErrorAction SilentlyContinue
$excelWorkbookLinkWarnings = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "WorkbookLinkWarnings" -ErrorAction SilentlyContinue

$wordAllowDDE = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "AllowDDE" -ErrorAction SilentlyContinue

# Always prevent untrusted Microsoft Query files from opening
$excelEnableBlockUnsecureQueryFiles = Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "enableblockunsecurequeryfiles" -ErrorAction SilentlyContinue

# Don't allow Dynamic Data Exchange (DDE) server launch in Excel
$excelDisableDDEServerLaunch = Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "disableddeserverlaunch" -ErrorAction SilentlyContinue

# Don't allow Dynamic Data Exchange (DDE) server lookup in Excel
$excelDisableDDEServerLookup = Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\excel\security\external content" -Name "disableddeserverlookup" -ErrorAction SilentlyContinue

# Update automatic links at Open
$wordDontUpdateLinks = Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\word\options" -Name "dontupdatelinks" -ErrorAction SilentlyContinue

if (
    $excelDataConnectionWarnings -eq $null -or
    $excelRichDataConnectionWarnings -eq $null -or
    $excelWorkbookLinkWarnings -eq $null -or
    $wordAllowDDE -eq $null -or
    $excelEnableBlockUnsecureQueryFiles -eq $null -or
    $excelDisableDDEServerLaunch -eq $null -or
    $excelDisableDDEServerLookup -eq $null -or
    $wordDontUpdateLinks -eq $null
) {
    Write-Host "Registry keys not found. Expected values cannot be determined."
}
else {
    Write-Host "Expected value for 'DataConnectionWarnings' in Excel is 2. Current value is $excelDataConnectionWarnings."
    Write-Host "Expected value for 'RichDataConnectionWarnings' in Excel is 2. Current value is $excelRichDataConnectionWarnings."
    Write-Host "Expected value for 'WorkbookLinkWarnings' in Excel is 2. Current value is $excelWorkbookLinkWarnings."
    Write-Host "Expected value for 'AllowDDE' in Word is 0. Current value is $wordAllowDDE."
    Write-Host "Expected value for 'enableblockunsecurequeryfiles' in Excel is 1. Current value is $excelEnableBlockUnsecureQueryFiles."
    Write-Host "Expected value for 'disableddeserverlaunch' in Excel is 1. Current value is $excelDisableDDEServerLaunch."
    Write-Host "Expected value for 'disableddeserverlookup' in Excel is 1. Current value is $excelDisableDDEServerLookup."
    Write-Host "Expected value for 'dontupdatelinks' in Word is 1. Current value is $wordDontUpdateLinks."
}

#*************************************************************Object Linking and Embedding packages**********************************************************************************

# Excel
$excelPackagerPrompt = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "PackagerPrompt" -ErrorAction SilentlyContinue

# PowerPoint
$powerPointPackagerPrompt = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "PackagerPrompt" -ErrorAction SilentlyContinue

# Word
$wordPackagerPrompt = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "PackagerPrompt" -ErrorAction SilentlyContinue

if (
    $excelPackagerPrompt -eq $null -or
    $powerPointPackagerPrompt -eq $null -or
    $wordPackagerPrompt -eq $null
) {
    Write-Host "Registry keys not found. Expected values cannot be determined."
}
else {
    Write-Host "Expected value for 'PackagerPrompt' in Excel is 2. Current value is $excelPackagerPrompt."
    Write-Host "Expected value for 'PackagerPrompt' in PowerPoint is 2. Current value is $powerPointPackagerPrompt."
    Write-Host "Expected value for 'PackagerPrompt' in Word is 2. Current value is $wordPackagerPrompt."
}

#***************************************************************************ActiveX**********************************************************************************************
$commonDisableAllActiveX = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\Common\Security" -Name "DisableAllActiveX" -ErrorAction SilentlyContinue

if ($commonDisableAllActiveX -eq $null) {
    Write-Host "Registry key not found. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAllActiveX' in Office Common Security is 1. Current value is $commonDisableAllActiveX."
}

#*************************************************************************Add-Ins************************************************************************************************
$excelDisableAllAddins = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "DisableAllAddins" -ErrorAction SilentlyContinue
$powerPointDisableAllAddins = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "DisableAllAddins" -ErrorAction SilentlyContinue
$projectDisableAllAddins = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Project\Security" -Name "DisableAllAddins" -ErrorAction SilentlyContinue
$visioDisableAllAddins = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Visio\Security" -Name "DisableAllAddins" -ErrorAction SilentlyContinue
$wordDisableAllAddins = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Security" -Name "DisableAllAddins" -ErrorAction SilentlyContinue

if ($excelDisableAllAddins -eq $null) {
    Write-Host "Registry key 'DisableAllAddins' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAllAddins' in Excel is 1. Current value is $excelDisableAllAddins."
}

if ($powerPointDisableAllAddins -eq $null) {
    Write-Host "Registry key 'DisableAllAddins' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAllAddins' in PowerPoint is 1. Current value is $powerPointDisableAllAddins."
}

if ($projectDisableAllAddins -eq $null) {
    Write-Host "Registry key 'DisableAllAddins' not found for Project. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAllAddins' in Project is 1. Current value is $projectDisableAllAddins."
}

if ($visioDisableAllAddins -eq $null) {
    Write-Host "Registry key 'DisableAllAddins' not found for Visio. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAllAddins' in Visio is 1. Current value is $visioDisableAllAddins."
}

if ($wordDisableAllAddins -eq $null) {
    Write-Host "Registry key 'DisableAllAddins' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAllAddins' in Word is 1. Current value is $wordDisableAllAddins."
}

#**************************************************************************Extention Hardening*************************************************************************************

$excelExtensionHardening = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "ExtensionHardening" -ErrorAction SilentlyContinue

if ($excelExtensionHardening -eq $null) {
    Write-Host "Registry key 'ExtensionHardening' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'ExtensionHardening' in Excel is 2. Current value is $excelExtensionHardening."
}


#***********************************************************************File Type Blocking****************************************************************************************
$excelProtectedViewEnabled = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" -Name "EnableDatabaseFileProtectedView" -ErrorAction SilentlyContinue
$excelDisableInternetFilesInPV = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" -Name "DisableInternetFilesInPV" -ErrorAction SilentlyContinue
$excelDisableUnsafeLocationsInPV = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" -Name "DisableUnsafeLocationsInPV" -ErrorAction SilentlyContinue
$excelFileValidationOpenInPV = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\FileValidation" -Name "OpenInProtectedView" -ErrorAction SilentlyContinue
$excelDisableAttachmentsInPV = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView" -Name "DisableAttachmentsInPV" -ErrorAction SilentlyContinue
$ppdisableinternetfilesinpv= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview"-Name "disableinternetfilesinpv" -ErrorAction SilentlyContinue
$ppdisableunsafelocationsinpv= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableunsafelocationsinpv" -ErrorAction SilentlyContinue
$ppopeninprotectedview= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\filevalidation" -Name "openinprotectedview" -ErrorAction SilentlyContinue
$ppdisableattachmentsinpv= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\PowerPoint\security\protectedview" -Name "disableattachmentsinpv" -ErrorAction SilentlyContinue
$Worddisableinternetfilesinpv= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview"-Name "disableinternetfilesinpv" -ErrorAction SilentlyContinue
$Worddisableunsafelocationsinpv= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableunsafelocationsinpv" -ErrorAction SilentlyContinue
$Wordopeninprotectedview= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\Word\security\filevalidation" -Name "openinprotectedview" -ErrorAction SilentlyContinue
$Worddisableattachmentsinpv= Get-ItemPropertyValue -Path "HKCU:\software\microsoft\office\16.0\Word\security\protectedview" -Name "disableattachmentsinpv" -ErrorAction SilentlyContinue


if ($excelProtectedViewEnabled -eq $null) {
    Write-Host "Registry key 'EnableDatabaseFileProtectedView' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'EnableDatabaseFileProtectedView' in Excel is 1. Current value is $excelProtectedViewEnabled."
}

if ($excelDisableInternetFilesInPV -eq $null) {
    Write-Host "Registry key 'DisableInternetFilesInPV' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableInternetFilesInPV' in Excel is 0. Current value is $excelDisableInternetFilesInPV."
}

if ($excelDisableUnsafeLocationsInPV -eq $null) {
    Write-Host "Registry key 'DisableUnsafeLocationsInPV' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableUnsafeLocationsInPV' in Excel is 0. Current value is $excelDisableUnsafeLocationsInPV."
}

if ($excelFileValidationOpenInPV -eq $null) {
    Write-Host "Registry key 'OpenInProtectedView' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'OpenInProtectedView' in Excel is 1. Current value is $excelFileValidationOpenInPV."
}

if ($excelDisableAttachmentsInPV -eq $null) {
    Write-Host "Registry key 'DisableAttachmentsInPV' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAttachmentsInPV' in Excel is 0. Current value is $excelDisableAttachmentsInPV."
}

if ($ppdisableinternetfilesinpv -eq $null) {
    Write-Host "Registry key 'DisableAttachmentsInPV' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'DisableAttachmentsInPV' in PowerPoint is 0. Current value is $ppdisableinternetfilesinpv."
}

if ($ppdisableunsafelocationsinpv -eq $null) {
    Write-Host "Registry key 'disableunsafelocationsinpv' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disableunsafelocationsinpv' in PowerPoint is 0. Current value is $ppdisableunsafelocationsinpv."
}
if ($ppopeninprotectedview -eq $null) {
    Write-Host "Registry key 'openinprotectedview' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'openinprotectedview' in PowerPoint is 1. Current value is $ppopeninprotectedview."
}
if ($ppdisableattachmentsinpv -eq $null) {
    Write-Host "Registry key 'disableattachmentsinpv' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disableattachmentsinpv' in PowerPoint is 0. Current value is $ppdisableattachmentsinpv."
}

if ($Worddisableinternetfilesinpv -eq $null) {
    Write-Host "Registry key 'disableinternetfilesinpv' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disableinternetfilesinpv' in Word is 0. Current value is $Worddisableinternetfilesinpv."
}

if ($Worddisableunsafelocationsinpv -eq $null) {
    Write-Host "Registry key 'disableunsafelocationsinpv' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disableunsafelocationsinpv' in Word is 1. Current value is $Worddisableunsafelocationsinpv."
}

if ($Wordopeninprotectedview -eq $null) {
    Write-Host "Registry key 'openinprotectedview' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'openinprotectedview' in Word is 1. Current value is $Wordopeninprotectedview."
}

if ($Worddisableattachmentsinpv -eq $null) {
    Write-Host "Registry key 'disableattachmentsinpv' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disableattachmentsinpv' in Word is 0. Current value is $Worddisableattachmentsinpv."
}

#************************************************************************Trusted documents************************************************************************************

$excelTrustedDocsPath = "HKCU:\software\microsoft\Excel\16.0\access\security\trusted documents"
$excelNetworkTrustedDocsPath = "HKCU:\software\microsoft\office\16.0\Excel\security\trusted documents"
$powerPointTrustedDocsPath = "HKCU:\software\microsoft\PowerPoint\16.0\access\security\trusted documents"
$powerPointNetworkTrustedDocsPath = "HKCU:\software\microsoft\office\16.0\PowerPoint\security\trusted documents"
$visioTrustedDocsPath = "HKCU:\software\microsoft\Visio\16.0\access\security\trusted documents"
$visioNetworkTrustedDocsPath = "HKCU:\software\microsoft\office\16.0\Visio\security\trusted documents"
$wordTrustedDocsPath = "HKCU:\software\microsoft\Word\16.0\access\security\trusted documents"
$wordNetworkTrustedDocsPath = "HKCU:\software\microsoft\office\16.0\Word\security\trusted documents"

$excelTrustedDocsValue = Get-ItemPropertyValue -Path $excelTrustedDocsPath -Name "disabletrusteddocuments" -ErrorAction SilentlyContinue
$excelNetworkTrustedDocsValue = Get-ItemPropertyValue -Path $excelNetworkTrustedDocsPath -Name "disablenetworktrusteddocuments" -ErrorAction SilentlyContinue
$powerPointTrustedDocsValue = Get-ItemPropertyValue -Path $powerPointTrustedDocsPath -Name "disabletrusteddocuments" -ErrorAction SilentlyContinue
$powerPointNetworkTrustedDocsValue = Get-ItemPropertyValue -Path $powerPointNetworkTrustedDocsPath -Name "disablenetworktrusteddocuments" -ErrorAction SilentlyContinue
$visioTrustedDocsValue = Get-ItemPropertyValue -Path $visioTrustedDocsPath -Name "disabletrusteddocuments" -ErrorAction SilentlyContinue
$visioNetworkTrustedDocsValue = Get-ItemPropertyValue -Path $visioNetworkTrustedDocsPath -Name "disablenetworktrusteddocuments" -ErrorAction SilentlyContinue
$wordTrustedDocsValue = Get-ItemPropertyValue -Path $wordTrustedDocsPath -Name "disabletrusteddocuments" -ErrorAction SilentlyContinue
$wordNetworkTrustedDocsValue = Get-ItemPropertyValue -Path $wordNetworkTrustedDocsPath -Name "disablenetworktrusteddocuments" -ErrorAction SilentlyContinue

if ($excelTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$excelTrustedDocsPath' not found for Excel. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disabletrusteddocuments' in Excel is 1. Current value is $excelTrustedDocsValue."
}

if ($excelNetworkTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$excelNetworkTrustedDocsPath' not found for Excel network. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disablenetworktrusteddocuments' in Excel network is 1. Current value is $excelNetworkTrustedDocsValue."
}

if ($powerPointTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$powerPointTrustedDocsPath' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disabletrusteddocuments' in PowerPoint is 1. Current value is $powerPointTrustedDocsValue."
}

if ($powerPointNetworkTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$powerPointNetworkTrustedDocsPath' not found for PowerPoint network. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disablenetworktrusteddocuments' in PowerPoint network is 1. Current value is $powerPointNetworkTrustedDocsValue."
}

if ($visioTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$visioTrustedDocsPath' not found for Visio. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disabletrusteddocuments' in Visio is 1. Current value is $visioTrustedDocsValue."
}

if ($visioNetworkTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$visioNetworkTrustedDocsPath' not found for Visio network. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disablenetworktrusteddocuments' in Visio network is 1. Current value is $visioNetworkTrustedDocsValue."
}

if ($wordTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$wordTrustedDocsPath' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disabletrusteddocuments' in Word is 1. Current value is $wordTrustedDocsValue."
}

if ($wordNetworkTrustedDocsValue -eq $null) {
    Write-Host "Registry key '$wordNetworkTrustedDocsPath' not found for Word network. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'disablenetworktrusteddocuments' in Word network is 1. Current value is $wordNetworkTrustedDocsValue."
}
#*************************************************************************************Hidden markup*********************************************************************************

$powerPointMarkupPath = "HKCU:\software\microsoft\office\16.0\powerpoint\options"
$wordMarkupPath = "HKCU:\software\microsoft\office\16.0\Word\options"

$powerPointMarkupValue = Get-ItemPropertyValue -Path $powerPointMarkupPath -Name "markupopensave" -ErrorAction SilentlyContinue
$wordMarkupValue = Get-ItemPropertyValue -Path $wordMarkupPath -Name "markupopensave" -ErrorAction SilentlyContinue

if ($powerPointMarkupValue -eq $null) {
    Write-Host "Registry key '$powerPointMarkupPath' not found for PowerPoint. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'markupopensave' in PowerPoint is 1. Current value is $powerPointMarkupValue."
}

if ($wordMarkupValue -eq $null) {
    Write-Host "Registry key '$wordMarkupPath' not found for Word. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'markupopensave' in Word is 1. Current value is $wordMarkupValue."
}

#****************************************************************************************Reporting Information*********************************************************************

$feedbackPath = "HKCU:\software\microsoft\office\16.0\common\feedback"
$updatereliabilityPath = "HKCU:\software\microsoft\office\16.0\common"
$clienttelemetryPath = "HKCU:\software\microsoft\office\16.0\common\clienttelemetry"
$generalPath = "HKCU:\software\microsoft\office\16.0\common\general"
$qmenablePath = "HKCU:\software\microsoft\office\16.0\common"

$includescreenshotValue = Get-ItemPropertyValue -Path $feedbackPath -Name "includescreenshot" -ErrorAction SilentlyContinue
$updatereliabilitydataValue = Get-ItemPropertyValue -Path $updatereliabilityPath -Name "updatereliabilitydata" -ErrorAction SilentlyContinue
$sendtelemetryValue = Get-ItemPropertyValue -Path $clienttelemetryPath -Name "sendtelemetry" -ErrorAction SilentlyContinue
$shownfirstrunoptinValue = Get-ItemPropertyValue -Path $generalPath -Name "shownfirstrunoptin" -ErrorAction SilentlyContinue
$qmenableValue = Get-ItemPropertyValue -Path $qmenablePath -Name "qmenable" -ErrorAction SilentlyContinue
$enabledValue = Get-ItemPropertyValue -Path $feedbackPath -Name "enabled" -ErrorAction SilentlyContinue
$sendcustomerdataValue = Get-ItemPropertyValue -Path $qmenablePath -Name "sendcustomerdata" -ErrorAction SilentlyContinue

if ($includescreenshotValue -eq $null) {
    Write-Host "Registry key '$feedbackPath' not found for including screenshot with Office Feedback. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'includescreenshot' in Office Feedback is 0. Current value is $includescreenshotValue."
}

if ($updatereliabilitydataValue -eq $null) {
    Write-Host "Registry key '$updatereliabilityPath' not found for automatically receiving small updates to improve reliability. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'updatereliabilitydata' for small updates is 0. Current value is $updatereliabilitydataValue."
}

if ($sendtelemetryValue -eq $null) {
    Write-Host "Registry key '$clienttelemetryPath' not found for configuring diagnostic data sent by Office to Microsoft. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'sendtelemetry' in Office is 1. Current value is $sendtelemetryValue."
}

if ($shownfirstrunoptinValue -eq $null) {
    Write-Host "Registry key '$generalPath' not found for disabling Opt-in Wizard on first run. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'shownfirstrunoptin' for Opt-in Wizard is 1. Current value is $shownfirstrunoptinValue."
}

if ($qmenableValue -eq $null) {
    Write-Host "Registry key '$qmenablePath' not found for enabling Customer Experience Improvement Program. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'qmenable' in Office is 0. Current value is $qmenableValue."
}

if ($enabledValue -eq $null) {
    Write-Host "Registry key '$feedbackPath' not found for sending Office Feedback. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'enabled' in Office Feedback is 0. Current value is $enabledValue."
}

if ($sendcustomerdataValue -eq $null) {
    Write-Host "Registry key '$generalPath' not found for sending personal information. Expected value cannot be determined."
}
else {
    Write-Host "Expected value for 'sendcustomerdata' for sending personal information is 0. Current value is $sendcustomerdataValue."
}

