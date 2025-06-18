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


# Function to get installed applications from the registry
function Get-InstalledApps {
    $installedApps = @()

    # Registry paths for installed applications
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $registryPaths) {
        $apps = Get-ItemProperty $path | Where-Object { $_.DisplayName -and $_.DisplayName -ne "" }
        $installedApps += $apps | Select-Object DisplayName, DisplayVersion
    }

    return $installedApps
}

Write-Output "i Script created by rares.plesea@ace-iss.com"
Write-Output ""

#####################################################################################################################
# → Export installed updates
Write-Output "→ Export installed updates"
Get-HotFix |
    Sort-Object InstalledOn |
    Select-Object -Property Caption, Description, HotFixID, InstalledOn |
    Export-Csv -Path "$destinationFolder\$($env:COMPUTERNAME)_InstalledUpdates_$timestamp.csv" -NoTypeInformation

# Retrieve installed applications
$installedApps = Get-InstalledApps | Sort-Object DisplayName

# Export the list to CSV with proper headers
$installedApps | Export-Csv -Path "$destinationFolder\$($env:COMPUTERNAME)_InstalledApps_$timestamp.csv" -NoTypeInformation

# Export system information
Write-Output "→ Export system information"
systeminfo | Out-File -FilePath "$destinationFolder\$($env:COMPUTERNAME)_Systeminfo_$timestamp.txt"

# Export list of all services
Write-Output "→ Export list of all services"
Get-Service | Out-File -FilePath "$destinationFolder\$($env:COMPUTERNAME)_Services_$timestamp.txt"

# Export ports information with active connections
Write-Output "→ Export ports information with active connections"
netstat -a -b | Out-File -FilePath "$destinationFolder\$($env:COMPUTERNAME)_Ports_$timestamp.txt"

# Export System event logs
Write-Output "→ Export System event logs"
wevtutil epl System "$destinationFolder\$($env:COMPUTERNAME)_System_event_log_$timestamp.evtx"

# Export Application event logs
Write-Output "→ Export Application event logs"
wevtutil epl Application "$destinationFolder\$($env:COMPUTERNAME)_Application_event_log_$timestamp.evtx"

Start-Sleep -Seconds 1

#####################################################################################################################

Write-Output ""

# LogViewer
$logViewer = "C:\Program Files (x86)\Common Files\ArchestrA\aaLogViewer.exe"
if (Test-Path $logViewer) {
    Write-Output "i An application will open; it has AVEVA logs. Go to Action > Messages > Export. Then click OK."
    Write-Output "i When you are done come back to this PowerShell window and press Enter to continue."
    Start-Process $logViewer
    Read-Host "Press Enter to continue"

    # Construct the filename
    $logFileName = "LogExport$($month)$($day)$($year).aaLGX"

    # Build the full source path
    $archestraLogsourcePath = Join-Path -Path $archestraLogsFolder -ChildPath $logFileName

    # Check if the file exists, then copy it
    if (Test-Path -Path $archestraLogsFolder) {
        Copy-Item -Path $archestraLogsourcePath -Destination $destinationFolder
        Write-Output "i aaLGX file copied successfully to $destinationFolder"
    } else {
        Write-Output "! aaLGX file not found: $archestraLogsFolder"
    }
}

# Common Service Portal 
if (Test-Path -Path "C:\Program Files (x86)\AVEVA\Platform Common Services\Portal") {
    # Check if "Microsoft Edge WebView2 Runtime" is in the exported CSV
    $csvInstalledApps = Import-Csv -Path "$destinationFolder\$($env:COMPUTERNAME)_InstalledApps_$timestamp.csv"
    $appName = "Microsoft Edge WebView2 Runtime"

    Write-Output ""
    Write-Output "i An application should open. Go to Sidebar > Troubleshooing (2nd icon) > Scan > Export."

    if ($csvInstalledApps.DisplayName -contains $appName) {
        Write-Output "   i AVEVA Common Service Portal prerequisite, $appName, found."
    } else {
        Write-Output "   ! AVEVA Common Service Portal may not open on your system. You must install $appName and try again."
        Write-Output "   i AVEVA Common Service Portal is found in Start Menu > AVEVA Folder > Common Services Portal."
    }

    Write-Output "i Save the json file, and then copy it to the folder that will open soon."
    # it is easier to start this process as it has some extra options; the direc exe does not open properly
    Start-Process "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\AVEVA\Common Services Portal"
    Read-Host "Press Enter to continue"
}

#####################################################################################################################

if ($plantSCADAFolder) {
    Write-Output ""
    Write-Output "i Found Plant SCADA logs folder"
    $targetFolder = Join-Path -Path $destinationFolder -ChildPath $plantSCADAFolder.Parent.Name
    # Check if the folder exists, and create it if not
    if (-Not (Test-Path -Path $targetFolder)) {
        New-Item -Path $targetFolder -ItemType Directory | Out-Null
        # Write-Output "Created folder: $targetFolder"
    }
    Write-Output "i Scanning folder: $($plantSCADAFolder.FullName)"

    # Get all .txt, .dat, and .log files in the folder (non-recursive)
    $files = Get-ChildItem -Path "$($plantSCADAFolder.FullName)\*" -Include *.txt, *.dat, *.log -File -ErrorAction SilentlyContinue

    if ($files) {
        Write-Output "   Copying files from $($plantSCADAFolder.FullName) to $targetFolder"
        foreach ($file in $files) {
            Copy-Item -Path $file.FullName -Destination "$targetFolder" -Force
        }
    } else {
        Write-Output "! No .txt, .dat, or .log files found in $($plantSCADAFolder.FullName)."
    }

}
# else {
#     Write-Output "! No matching Plant SCADA folders found."
# }

#####################################################################################################################

# Archive the destination folder into a ZIP file

# if (Test-Path -Path $zipFilePath) {
#     Remove-Item -Path $zipFilePath -Force
# }

# Notify user of completion
Write-Output ""
Write-Output "v Diagnostics have been saved to $destinationFolder"

# Write-Output "i Archiving $destinationFolder into $zipFilePath"
# Compress-Archive -Path $destinationFolder -DestinationPath $zipFilePath

# Write-Output "i Archive created with path: $zipFilePath"

Write-Output ""
Write-Output "i You must send the archived folder to support."

# Open the $destinationFolder folder
Start-Process -FilePath $archestraLogsFolder
Start-Process -FilePath $destinationFolder
