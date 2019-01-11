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
    [string]
    $SiteServer,

    [Parameter(Mandatory=$true)]
    [string]
    $SiteCode,

    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $(Split-Path $_) -PathType 'Container'})] 
    [string]
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

[int]$DefaultRecurCount = 7
[string]$DefaultRecurInterval = "Days"
[string]$DefaultDescription = "Operational Collection"
[int]$SchedAddHours = 6

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

Function Add-FolderPath ([string]$RootFolder, [string]$Path) {
    if ($Path -like "*\*") {
        $SplitPath = $Path.split("\")
        foreach ($item in $SplitPath) {
            if (!(Test-Path "$($RootFolder)\$($item)")) {
                New-Item -Path $RootFolder -Name $item -ItemType Directory
            }
            $RootFolder = "$($RootFolder)\$($item)"
        }
    } else {
        New-Item -Path $RootFolder -Name $Path -ItemType Directory
    }
    Write-ToLog -File $LogFile -Message "Created collection folder $($FolderPath)"
}

Function Add-Collection ([string]$ColName, [string]$ColLimiting, [string]$ColDescription, [string]$ColRecurInterval, [int]$ColRecurCount) {
    $Schedule = New-CMSchedule -RecurInterval $ColRecurInterval -RecurCount $ColRecurCount -Start (Get-Date).AddHours($SchedAddHours)
    try {
        New-CMDeviceCollection -Name $ColName -LimitingCollectionName $ColLimiting -Comment $ColDescription -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
        Write-ToLog -File $LogFile -Message "Created collection $($ColName)"
    }
    catch {
        Write-ToLog -File $LogFile -Message "ERROR. Error creating collection $($ColName). Error message: $($_.Exception.Message)"
    }
}

Function Add-CollectionQuery ([string]$ColName, [string]$Query) {
    try {
        Add-CMDeviceCollectionQueryMembershipRule -CollectionName $ColName -RuleName $ColName -QueryExpression $Query
        Write-ToLog -File $LogFile -Message "Added collection membership query to collection $($ColName)"
    }
    catch {
        Write-ToLog -File $LogFile -Message "ERROR. Could not add collection membership query to collection $($ColName). Error message: $($_.Exception.Message)"
    }
}

Function Add-CollectionInclude ([string]$ColName, [string]$Include) {
    if ((Get-CMDeviceCollection -Name $ColName).Name -eq $ColName) {
        try {
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $ColName -IncludeCollectionName $Include
            Write-ToLog -File $LogFile -Message "Added include membership rule to collection $($ColName)"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR. Could not add include membership rule to collection $($ColName). Error message: $($_.Exception.Message)"
        }
    } else {
        Write-ToLog -File $LogFile -Message "ERROR. Include membership collection $($Include) doesn't exist. Couldn't be added to collection $($ColName)"
    }
}

Function Add-CollectionExclude ([string]$ColName, [string]$Exclude) {
    if ((Get-CMDeviceCollection -Name $ColName).Name -eq $ColName) {
        try {
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $ColName -ExcludeCollectionName $Exclude
            Write-ToLog -File $LogFile -Message "Added exclude membership rule to collection $($ColName)"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR. Could not add exclude membership rule to collection $($ColName). Error message: $($_.Exception.Message)"
        }
    } else {
        Write-ToLog -File $LogFile -Message "ERROR. Exclude membership collection $($Exclude) doesn't exist. Couldn't be added to collection $($ColName)"
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
        Add-FolderPath -RootFolder $RootFolder -Path $ColPath
    }

    # Gather from the XML the necessary data to create the collection
    if ($null -ne $col.limiting) {
        $ColLimiting = $col.limiting
        if ($null -eq (Get-CMDeviceCollection -Name $ColLimiting).Name) {
            Write-ToLog "ERROR. Limiting collection does not exist for collection $($ColName)"
            continue
        }
    } else {
        Write-ToLog -File $LogFile -Message "ERROR. Limiting collection is missing for collection $($ColName)"
        continue
    }
    
    if ($null -ne $col.description) {
        $ColDescription = $col.description
    } else {
        $ColDescription = $DefaultDescription
    }

    if ($null -ne $col.recurcount) {
        $ColRecurCount = $col.recurcount
    } else {
        $ColRecurCount = $DefaultRecurCount
    }

    if ($null -ne $col.recurinterval) {
        $ColRecurInterval = $col.recurinterval
    } else {
        $ColRecurInterval = $DefaultRecurInterval
    }

    # Create the empty collection
    Add-Collection -ColName $ColName -ColLimiting $ColLimiting -ColDescription $ColDescription -ColRecurInterval $ColRecurInterval -ColRecurCount $ColRecurCount

    # Move the collection to their corresponding folder
    Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $ColName)

    # Check for query membership rules for the collection
    $ColQueries = $col.query

    if ($ColQueries.Length -gt 0) {
        foreach ($Query in $ColQueries) {
            Add-CollectionQuery -ColName $ColName -Query $Query
        }
    }

    # Check for include membership rules for the collection
    $ColInclude = $col.include

    if ($ColInclude.Length -gt 0) {
        foreach ($Include in $ColInclude) {
            Add-CollectionInclude -ColName $ColName -Include $Include
        }
    }
   
    # Check for exclude membership rules for the collection
    $ColExclude = $col.exclude

    if ($ColExclude.Length -gt 0) {
        foreach ($Exclude in $ColExclude) {
            Add-CollectionExclude -ColName $ColName -Exclude $Exclude
        }
    }



}



#endregion