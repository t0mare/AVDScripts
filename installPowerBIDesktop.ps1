# Define Event Log Source
$eventSource = "PowerBI-Deploy"

# Check if Event Log Source exists; create it if needed
if ([System.Diagnostics.EventLog]::SourceExists($eventSource) -eq $false) {
    New-EventLog -LogName Application -Source $eventSource
}

# Log the start of installation
Write-EventLog -LogName Application -Source $eventSource -EntryType Information -EventId 100 -Message "Starting Power BI Desktop installation."

# Ensure Chocolatey is installed
if (!(Get-Command choco -ErrorAction SilentlyContinue)) {
    Write-EventLog -LogName Application -Source $eventSource -EntryType Warning -EventId 101 -Message "Chocolatey not found. Attempting to install."

    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    try {
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-EventLog -LogName Application -Source $eventSource -EntryType Information -EventId 102 -Message "Chocolatey installation completed successfully."
    } catch {
        Write-EventLog -LogName Application -Source $eventSource -EntryType Error -EventId 103 -Message "Chocolatey installation failed: $_"
        Exit 1
    }
}

# Log Power BI Desktop installation start
Write-EventLog -LogName Application -Source $eventSource -EntryType Information -EventId 104 -Message "Installing Power BI Desktop using Chocolatey."

# Install Power BI Desktop silently
try {
    choco install powerbi -y --limit-output
    Write-EventLog -LogName Application -Source $eventSource -EntryType Information -EventId 105 -Message "Power BI Desktop installation completed successfully."
} catch {
    Write-EventLog -LogName Application -Source $eventSource -EntryType Error -EventId 106 -Message "Power BI Desktop installation failed: $_"
    Exit 1
}

# Verify Installation
if (Test-Path "C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe") {
    Write-EventLog -LogName Application -Source $eventSource -EntryType Information -EventId 107 -Message "Power BI Desktop verification successful. Installation complete."
    Exit 0
} else {
    Write-EventLog -LogName Application -Source $eventSource -EntryType Error -EventId 108 -Message "Power BI Desktop executable not found. Installation may have failed."
    Exit 1
}
