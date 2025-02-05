# Ensure the script runs with administrative privileges
$adminCheck = [System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()
if (-Not $adminCheck.IsInRole([System.Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as an administrator"
    exit
}

# Winget command to install Power BI Desktop silently
$packageName = "Microsoft.PowerBI"
$installArgs = "--silent --accept-package-agreements --accept-source-agreements"

Write-Host "Installing Power BI Desktop..."
Start-Process -FilePath "winget" -ArgumentList "install $packageName $installArgs" -NoNewWindow -Wait

# Verify installation
$installed = winget list | Select-String -Pattern "Power BI Desktop"
if ($installed) {
    Write-Host "Power BI Desktop installation successful."
} else {
    Write-Host "Power BI Desktop installation failed."
}
