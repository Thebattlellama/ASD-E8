$backupPath = "C:\Windows\RegistryBackup"

# Create the backup directory if it doesn't exist
if (!(Test-Path -Path $backupPath)) {
    New-Item -ItemType Directory -Path $backupPath | Out-Null
}

# Generate a unique backup file name based on the current date and time
$backupFileName = "RegistryBackup_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".reg"
$backupFilePath = Join-Path -Path $backupPath -ChildPath $backupFileName

# Backup the registry
try {
    Write-Host "Backing up the registry to $backupFilePath ..."
    & reg.exe export HKCU $backupFilePath /y
    Write-Host "Registry backup created successfully."
}
catch {
    Write-Host "Error occurred while creating the registry backup:"
    Write-Host $_.Exception.Message
    exit 1  # Exit the script if there was an error
    
}

# Apply the changes to the registry
$officeVersions = @(  # List of Office versions to iterate through
    "16.0"  # Office 2016
    
    # Add more versions if needed
)

$applications = @(  # List of Office applications to iterate through
    "Excel",
    "Word",
    "PowerPoint",
    "Outlook"
    # Add more applications if needed
)

foreach ($version in $officeVersions) {
    foreach ($app in $applications) {
        $registryPath = "HKCU:\SOFTWARE\Microsoft\Office\$version\$app\Security"
        Set-ItemProperty -Path $registryPath -Name "PackagerPrompt" -Value 2 -type Dword
    }
}

Write-Host "Changes applied successfully."
