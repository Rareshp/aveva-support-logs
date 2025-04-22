# Set the source and destination folders
$baseFolder = "$env:TEMP"

# Create a sanitized timestamp by removing colons and spaces
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$destinationFolder = Join-Path -Path $baseFolder -ChildPath "aveva_logfiles_$timestamp"
$zipFilePath = Join-Path -Path $baseFolder -ChildPath "aveva_logfiles_$timestamp.zip"

$archestraLogsFolder = "C:\ProgramData\ArchestrA\Logger\LogExports"
$plantSCADAFolder = Get-ChildItem -Path "$env:ProgramData\AVEVA Plant SCADA *\Logs" -Directory -ErrorAction SilentlyContinue | Select-Object -First 1

# Ensure the destination folder exists
if (-Not (Test-Path -Path $destinationFolder)) {
    New-Item -Path $destinationFolder -ItemType Directory | Out-Null
}

# Get the current date in the format MMDDYYYY
$currentDate = Get-Date
$month = $currentDate.Month
$day = $currentDate.Day
$year = $currentDate.Year

# Remove leading zeros from month and day
$month = $month.ToString()
$day = $day.ToString()


Write-Output "ℹ️ Script created by rares.plesea@ace-iss.com"
Write-Output ""

#####################################################################################################################
# ↪️ Export installed updates
Write-Output "↪️ Export installed updates"
Get-HotFix |
    Sort-Object InstalledOn |
    Select-Object -Property Caption, Description, HotFixID, InstalledOn |
    Export-Csv -Path "$destinationFolder\$($env:COMPUTERNAME)_InstalledUpdates_$timestamp.csv" -NoTypeInformation

# Export installed applications sorted by Name
Write-Output "↪️ Export installed applications (be patient)"
# Retrieve installed applications, avoiding empty or null values
$installedApps = Get-WmiObject -Query "Select Name, Version from Win32_Product" |
    Where-Object { $_.Name -ne $null -and $_.Version -ne $null } |  # Filter out null values
    Sort-Object Name | 
    Select-Object -Property Name, Version

# Export the list to CSV with proper headers
$installedApps | Export-Csv -Path "$destinationFolder\$($env:COMPUTERNAME)_InstalledApps_$timestamp.csv" -NoTypeInformation

# Export system information
Write-Output "↪️ Export system information"
systeminfo | Out-File -FilePath "$destinationFolder\$($env:COMPUTERNAME)_Systeminfo_$timestamp.txt"

# Export list of all services
Write-Output "↪️ Export list of all services"
Get-Service | Out-File -FilePath "$destinationFolder\$($env:COMPUTERNAME)_Services_$timestamp.txt"

# Export ports information with active connections
Write-Output "↪️ Export ports information with active connections"
netstat -a -b | Out-File -FilePath "$destinationFolder\$($env:COMPUTERNAME)_Ports_$timestamp.txt"

# Export System event logs
Write-Output "↪️ Export System event logs"
wevtutil epl System "$destinationFolder\$($env:COMPUTERNAME)_System_event_log_$timestamp.evtx"

# Export Application event logs
Write-Output "↪️ Export Application event logs"
wevtutil epl Application "$destinationFolder\$($env:COMPUTERNAME)_Application_event_log_$timestamp.evtx"

Start-Sleep -Seconds 1

#####################################################################################################################

# LogViewer
Write-Output ""
Write-Output "ℹ️ An application will open. Go to Action > Messages > Export. Then click OK."
Write-Output "ℹ️ When you are done come back to this PowerShell window and press Enter to continue."
Start-Process "C:\Program Files (x86)\Common Files\ArchestrA\aaLogViewer.exe"
Read-Host "Press Enter to continue"

# Construct the filename
$logFileName = "LogExport$($month)$($day)$($year).aaLGX"

# Build the full source path
$archestraLogsourcePath = Join-Path -Path $archestraLogsFolder -ChildPath $logFileName

# Check if the file exists, then copy it
if (Test-Path -Path $archestraLogsFolder) {
    Copy-Item -Path $archestraLogsourcePath -Destination $destinationFolder
    Write-Output "ℹ️ aaLGX file copied successfully to $destinationFolder"
} else {
    Write-Output "⚠️ aaLGX file not found: $archestraLogsFolder"
}

#####################################################################################################################

if ($plantSCADAFolder) {
    Write-Output ""
    Write-Output "ℹ️ Found Plant SCADA logs folder"
    $targetFolder = Join-Path -Path $destinationFolder -ChildPath $plantSCADAFolder.Parent.Name
    # Check if the folder exists, and create it if not
    if (-Not (Test-Path -Path $targetFolder)) {
        New-Item -Path $targetFolder -ItemType Directory | Out-Null
        # Write-Output "Created folder: $targetFolder"
    }
    Write-Output "🔎 Scanning folder: $($plantSCADAFolder.FullName)"

    # Get all .txt, .dat, and .log files in the folder (non-recursive)
    $files = Get-ChildItem -Path "$($plantSCADAFolder.FullName)\*" -Include *.txt, *.dat, *.log -File -ErrorAction SilentlyContinue

    if ($files) {
        Write-Output "   Copying files from $($plantSCADAFolder.FullName) to $targetFolder"
        foreach ($file in $files) {
            Copy-Item -Path $file.FullName -Destination "$targetFolder" -Force
        }
    } else {
        Write-Output "⚠️ No .txt, .dat, or .log files found in $($plantSCADAFolder.FullName)."
    }

}
# else {
#     Write-Output "⚠️ No matching Plant SCADA folders found."
# }

#####################################################################################################################

# Archive the destination folder into a ZIP file

# if (Test-Path -Path $zipFilePath) {
#     Remove-Item -Path $zipFilePath -Force
# }

# Notify user of completion
Write-Output "☑️ All diagnostics have been saved to $destinationFolder"

# Write-Output "ℹ️ Archiving $destinationFolder into $zipFilePath"
# Compress-Archive -Path $destinationFolder -DestinationPath $zipFilePath

# Write-Output "ℹ️ Archive created with path: $zipFilePath"

Write-Output ""
Write-Output "ℹ️ You must send this file to support."

# Open the $destinationFolder folder
Start-Process -FilePath $archestraLogsFolder
Start-Process -FilePath $destinationFolder
