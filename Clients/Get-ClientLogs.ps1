
$LogShare = "\\wtokiosksccm01\ClientLogs$"

# Get path for SCCM client Log files
$LogPath = Get-ItemProperty -Path "HKLM:\Software\Microsoft\CCM\Logging\@Global"
$Log = $LogPath.LogDirectory

# Create folders on the CM client 
If (Test-Path -Path "$env:temp\CMlogs") {
    Remove-item "$env:temp\CMlogs"
}
New-Item -Path "$env:temp\CMLogs" -ItemType Directory -Force
Copy-Item -Path "$log\*" -Destination "$env:temp\CMlogs" -Force

# Create a .zip archive with sccm logs
Compress-Archive -Path "$env:temp\CMLogs\*" -CompressionLevel Optimal -DestinationPath "$env:temp\CMLogs.zip"

# Copy zipped logfile to servershare
$ComputerLogShare = $LogShare + "\" + $env:Computername
Write-host $ComputerLogShare
New-Item -Path $ComputerLogShare -ItemType Directory -Force
Copy-Item -Path "$env:temp\CMlogs.zip" -Destination $ComputerLogShare -Force

#Cleanup temporary files and folders from CM client
Remove-Item -Path "$env:temp\CMlogs" -Recurse
Remove-item -Path "$env:temp\CMlogs.zip"