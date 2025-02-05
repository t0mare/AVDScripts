# Ensure the script runs with administrative privileges
$adminCheck = [System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()
if (-Not $adminCheck.IsInRole([System.Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as an administrator"
    exit
}

# Check if Winget is installed
if (-Not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Winget is not installed. Please install Winget and try again."
    exit
}

# Winget command to install Power BI Desktop silently
$packageName = "Microsoft.PowerBI"
$installArgs = "--silent --accept-package-agreements --accept-source-agreements"

Write-Host "Installing Power BI Desktop..."
try {
    Start-Process -FilePath "winget" -ArgumentList "install $packageName $installArgs" -NoNewWindow -Wait -ErrorAction Stop
    Write-Host "Power BI Desktop installation command executed."
} catch {
    Write-Host "An error occurred during the installation process: $_"
    exit
}

# Verify installation
$installed = winget list | Select-String -Pattern "Power BI Desktop"
if ($installed) {
    Write-Host "Power BI Desktop installation successful."
} else {
    Write-Host "Power BI Desktop installation failed."
}
