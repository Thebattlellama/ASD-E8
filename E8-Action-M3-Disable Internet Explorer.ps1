# Disable Internet Explorer
$registryPath = "HKCU:\Software\Microsoft\Internet Explorer\Main"
Set-ItemProperty -Path $registryPath -Name "Start Page" -Value "about:blank"
Set-ItemProperty -Path $registryPath -Name "EnableFirstRunCustomize" -Value 0
Set-ItemProperty -Path $registryPath -Name "RunOnceHasShown" -Value 1
Set-ItemProperty -Path $registryPath -Name "RunOnceComplete" -Value 1

$registryPath = "HKCU:\Software\Policies\Microsoft\Internet Explorer\Main"
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}
Set-ItemProperty -Path $registryPath -Name "DisableFirstRunCustomize" -Value 1

Write-Host "Internet Explorer has been disabled."