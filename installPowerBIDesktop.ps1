# Set Execution Policy (if needed)
Set-ExecutionPolicy Bypass -Scope Process -Force

# Define a custom event log source
$LogSource = "PowerBIInstallScript"
$LogName = "Application"

# Check if the event source exists; if not, create it
if (!(Get-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue)) {
    New-EventLog -LogName $LogName -Source $LogSource
}

# Log the script start
Write-EventLog -LogName $LogName -Source $LogSource -EntryType Information -EventId 1001 -Message "Power BI installation script started."

# Define the URL and the local path for the installer
$InstallerUrl = "https://download.microsoft.com/download/8/8/0/880bca75-79dd-466a-927d-1abf1f5454b0/PBIDesktopSetup_x64.exe"
$InstallerPath = "$env:TEMP\PBIDesktopSetup_x64.exe"

# Download the installer
try {
    Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath
    Write-EventLog -LogName $LogName -Source $LogSource -EntryType Information -EventId 1005 -Message "Power BI installer downloaded successfully."
}
catch {
    Write-EventLog -LogName $LogName -Source $LogSource -EntryType Error -EventId 1006 -Message "❌ Failed to download Power BI installer: $_"
    exit
}

# Run Power BI Installation
try {
    Start-Process -FilePath $InstallerPath -ArgumentList "/quiet /norestart" -Wait -NoNewWindow

    # Verify Installation Path
    $PowerBI = Get-AppxPackage -Name "Microsoft.MicrosoftPowerBIDesktop"
    if ($PowerBI) {
        Write-EventLog -LogName $LogName -Source $LogSource -EntryType Information -EventId 1002 -Message "✅ Power BI installed successfully."
    } 
    else {
        Write-EventLog -LogName $LogName -Source $LogSource -EntryType Warning -EventId 1003 -Message "⚠️ Power BI installation not found."
    }
}
catch {
    Write-EventLog -LogName $LogName -Source $LogSource -EntryType Error -EventId 1004 -Message "❌ Power BI installation failed: $_"
}
finally {
    # Clean up the installer
    if (Test-Path $InstallerPath) {
        Remove-Item -Path $InstallerPath -Force
    }
}
