<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER SiteCode

.NOTES

.EXAMPLE
#>

# ----------------------------------------------------------------------------------------------------------------------------------------
# Define Parameters
# ----------------------------------------------------------------------------------------------------------------------------------------
#region parameters

[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]
param(
    [Parameter(Mandatory=$true)]
    [string[]]
    $SiteServer,

    [Parameter(Mandatory=$true)]
    [string[]]
    $SiteCode,

    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $(Split-Path $_) -PathType 'Container'})] 
    [string[]]
    $CollectionsXML
)

#endregion

# ----------------------------------------------------------------------------------------------------------------------------------------
# Declare Variables
# ----------------------------------------------------------------------------------------------------------------------------------------
#region variables

[string]$ScriptName = $($((Split-Path -Path $MyInvocation.MyCommand.Definition -Leaf)).Replace(".ps1",""))
[string]$ScriptPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
[string]$LogFile = "$($ScriptPath)\$($ScriptName).log"

#endregion

# ----------------------------------------------------------------------------------------------------------------------------------------
# Function Section
# ----------------------------------------------------------------------------------------------------------------------------------------
#region function

Function Write-ToLog ([string]$File, [string]$Message) {
    # Simple function that writes a log file with ConfigMgr format to be correctly processed by CMTrace

    $DateTime = $(Get-Date -Format "MM-dd-yyyy HH:mm:ss.fff")
    $MessageLog = "$Message~~ `$`$<$ScriptName><$DateTime><thread=0 (0x0)>"
    $MessageScr = "$DateTime - $Message"
    Write-Verbose -Message $MessageScr
    Write-Host -Object $MessageScr
    Out-File -FilePath $File -Encoding ascii -InputObject $MessageLog -Append
}

Function Initialize-CMModule {
    # Loads ConfigMgr module if not already loaded, stops script execution if it fails loading

    if (-not (Get-Module ConfigurationManager)) {
        # Module not loaded
        Write-ToLog -File $LogFile -Message "INFO: Configuration Manager Module not loaded. Loading"
        try {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop

            # If the PS-Drive it's not auto-created, try to create it
            if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer
            }
            Set-Location -Path "$($SiteCode):"
            Write-ToLog -File $LogFile -Message "INFO: Configuration Manager Module loaded"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: Error loading PowerShell module. Stopping script. Error message: $($_.Exception.Message)"
            Remove-Module -Name ConfigurationManager -Force -ErrorAction SilentlyContinue
            exit
        }
    } else {
        # Module already loaded
        try {
            # If the PS-Drive it's not auto-created, try to create it
            if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer
            }
            Set-Location -Path "$($SiteCode):"
            Write-ToLog -File $LogFile -Message "INFO: Configuration Manager Module already loaded"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: Error creating or accessing the PS-Drive. Error message: $($_.Exception.Message)"
            Remove-Module -Name ConfigurationManager -Force -ErrorAction SilentlyContinue
            exit
        }
        
    }
}

#endregion

# ----------------------------------------------------------------------------------------------------------------------------------------
# Script Process
# ----------------------------------------------------------------------------------------------------------------------------------------
#region process

# Load CM Module
Initialize-CMModule

# Import XML Data
[xml]$OpCollections = Get-Content $CollectionsXML

foreach ($col in $OpCollections.Collections.Collection) {
    $ColName = $col.name

    # Check if collection already exist
    if ((Get-CMDeviceCollection -Name $ColName).Name -eq $ColName) {
        Write-ToLog -File $LogFile -Message "Collection $($ColName) already exist. Skipping"
        continue
    }

    # Check and create the folder path
    $RootFolder = "$($SiteCode):\DeviceCollection"
    $ColPath = $col.folderpath
    $FolderPath = "$($RootFolder)\$($ColPath)"

    if (!(Test-Path $FolderPath)) {
        if ($ColPath -like "*\*") {
            $SplitPath = $ColPath.split("\")
            foreach ($item in $SplitPath) {
                if (!(Test-Path "$($RootFolder)\$($item)")) {
                    New-Item -Path $RootFolder -Name $item -ItemType Directory
                }
                $RootFolder = "$($RootFolder)\$($item)"
            }
        } else {
            New-Item -Path $RootFolder -Name $ColPath -ItemType Directory
        }
        Write-ToLog -File $LogFile -Message "Created collection folder $($FolderPath)"
    }


}



#endregion