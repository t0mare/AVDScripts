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

# Run Power BI Installation using WinGet
try {
    Start-Process -NoNewWindow -Wait -FilePath "winget.exe" -ArgumentList "install --id 9NTXR16HNW1T --source msstore --silent --accept-package-agreements"

    # Verify Installation Path
    if (Test-Path "C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe") {
        Write-EventLog -LogName $LogName -Source $LogSource -EntryType Information -EventId 1002 -Message "✅ Power BI installed successfully."
    } else {
        Write-EventLog -LogName $LogName -Source $LogSource -EntryType Warning -EventId 1003 -Message "⚠️ Power BI installation completed, but executable not found."
    }
}
catch {
    Write-EventLog -LogName $LogName -Source $LogSource -EntryType Error -EventId 1004 -Message "❌ Power BI installation failed: $_"
}
