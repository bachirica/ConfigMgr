<#
.SYNOPSIS
    Create ConfigMgr operational collections based on a XML definition file

.DESCRIPTION
    Creates ConfigMgr collections imported from an XML file
    See XML file for additional definition information

.PARAMETER SiteServer
    [Mandatory] ConfigMgr Site Server FQDN

.PARAMETER SiteCode
    [Mandatory] ConfigMgr Site Code (3 characters)

.PARAMETER CollectionsXML
    [Mandatory] XML file path containing collection definition

.NOTES
    Author: Bernardo Achirica (@bachirica)
    Version: 1.0
    Date: 2019.01.11
    References: Idea based on Mark Allen's script (https://github.com/markhallen/configmgr/tree/master/New-CMOperationalCollections)

.EXAMPLE
    .\Add-OperationalCollections.ps1 -SiteServer mysccmserver.mydomain.local -SiteCode PR1 -CollectionsXML .\OperationalCollections.xml
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
    $CollectionsXML,

    [Parameter(Mandatory=$false)]
    [switch]
    $Maintain
)

#endregion

# ----------------------------------------------------------------------------------------------------------------------------------------
# Declare Variables
# ----------------------------------------------------------------------------------------------------------------------------------------
#region variables

# Default values used for some parameters if they're not specified in the XML
[int]$DefaultRecurCount = 7
[string]$DefaultRecurInterval = "Days"
[string]$DefaultDescription = "Operational Collection"

# $SchedAddHours adds some extra hours to the collection evaluation schedule
# If a collection refreshes every 7 days, it'll do it always at the same time the collection was created (probably during office hours)
# Adding a few hours can allow you to schedule those refresh cycles outside office hours
[int]$SchedAddHours = 6 

# Log file location definition
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

function Get-DefaultIfNull ([string]$XMLValue, [string]$Default) {
    if ($null -ne $XMLValue) {
        return $XMLValue 
    } else {
        return $Default
    }
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

function Compare-HashTables ([hashtable]$hashXML, [hashtable]$hashCol, [hashtable]$hashNotInXML, [hashtable]$hashNotInCol) {
    foreach ($h in $hashXML.Keys) {
        $query = "$($hashXML.Item($h))"
        if (!($hashCol.ContainsValue($query))) {
            $hashNotInCol.Add(${h}, $hashXML.Item($h))
        }
    }
    
    foreach ($h in $hashCol.Keys) {
        $query = "$($hashCol.Item($h))"
        if (!($hashXML.ContainsValue($query))) {
            $hashNotInXML.Add(${h}, $hashCol.Item($h))
        }
    }    
}

function Reset-Collection ([string]$ColName, $ColXML) {
    $CMCol = Get-CMDeviceCollection -Name $ColName

    # Check limiting collection
    if ($CMCol.LimitToCollectionName -ne $ColXML.limiting) {
        try {
            Set-CMCollection -InputObject $CMCol -LimitingCollectionName $ColXML.limiting
            Write-ToLog -File $LogFile -Message "$($ColName). Corrected limiting collection to $($ColXML.limiting)"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not correct limiting collection to $($ColXML.limiting). Error message: $($_.Exception.Message)"
        }
    }

    # Check description
    if ($CMCol.Comment -ne $ColXML.description) {
        try {
            Set-CMCollection -InputObject $CMCol -Comment $ColXML.description
            Write-ToLog -File $LogFile -Message "$($ColName). Corrected collection comment"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not correct collection comment. Error message: $($_.Exception.Message)"
        }
    }

    # Check refresh type
    if ($CMCol.RefreshType -ne 2) {
        try {
            Set-CMCollection -InputObject $CMCol -RefreshType 2
            Write-ToLog -File $LogFile -Message "$($ColName). Corrected refresh type"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not correct refresh type. Error message: $($_.Exception.Message)"
        }
    }

    # Check refresh schedule
    $ColRecurCount = Get-DefaultIfNull -XMLValue $col.recurcount -Default $DefaultRecurCount
    $ColRecurInterval = Get-DefaultIfNull -XMLValue $col.recurinterval -Default $DefaultRecurInterval
    $ToCorrect = $false

    switch ($ColRecurInterval) {
        "Minutes" {
            if ($CMCol.RefreshSchedule.MinuteSpan -ne $ColRecurCount) {
                $ToCorrect = $true
            }
        }
        "Hours" {
            if ($CMCol.RefreshSchedule.HourSpan -ne $ColRecurCount) {
                $ToCorrect = $true
            }
        }
        "Days" {
            if ($CMCol.RefreshSchedule.DaySpan -ne $ColRecurCount) {
                $ToCorrect = $true
            }
        }
    }

    if ($ToCorrect) {
        $Schedule = New-CMSchedule -RecurInterval $ColRecurInterval -RecurCount $ColRecurCount -Start (Get-Date).AddHours($SchedAddHours)

        try {
            Set-CMCollection -InputObject $CMCol -RefreshSchedule $Schedule
            Write-ToLog -File $LogFile -Message "$($ColName). Corrected refresh schedule"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not correct refresh schedule. Error message: $($_.Exception.Message)"
        }
    }

    # Remove all direct memberships
    try {
        Get-CMDeviceCollectionDirectMembershipRule -CollectionName $ColName | ForEach-Object {Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $ColName -ResourceId $_.ResourceID -Force}
        Write-ToLog -File $LogFile -Message "$($ColName). Removed direct membership rules"
    }
    catch {
        Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not remove direct membership rules. Error message: $($_.Exception.Message)"
    }

    # Check collection queries
    $XMLQueries = $ColXML.query
    $i = 1
    $hashXML = @{}
    $hashCol = @{}
    $hashNotInCol = @{} # Elements in XML and not in the collection (to be added)
    $hashNotInXML = @{} # Elements in collection and not in the XML (to be removed)

    if ($XMLQueries.Length -gt 0) {
        foreach ($Query in $XMLQueries) {
            $hashXML.Add("$($i)", "$($Query)")
            $i++
        }
    }

    $ColQueries = Get-CMCollectionQueryMembershipRule -CollectionName $ColName
    if ($ColQueries.Count -gt 0) {
        foreach ($Query in $ColQueries) {
            $hashCol.Add("$($Query.RuleName)", "$($Query.QueryExpression)")
        }
    }

    Compare-HashTables -hashXML $hashXML -hashCol $hashCol -hashNotInXML $hashNotInXML -hashNotInCol $hashNotInCol

    if ($hashNotInXML.Count -gt 0) {
        foreach ($h in $hashNotInXML.Keys) {
            try {
                Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $ColName -RuleName $h -Force
                Write-ToLog -File $LogFile -Message "$($ColName). Removed query membership rule"
            }
            catch {
                Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not remove query membership rule. Error message: $($_.Exception.Message)"
            }
        }
    }

    if ($hashNotInCol.Count -gt 0) {
        foreach ($h in $hashNotInCol.Keys) {
            try {
                Add-CMDeviceCollectionQueryMembershipRule -CollectionName $ColName -RuleName $ColName -QueryExpression "$($hashNotInCol.Item($h))"
                Write-ToLog -File $LogFile -Message "$($ColName). Added query membership rule"
            }
            catch {
                Write-ToLog -File $LogFile -Message "ERROR. $($ColName). Could not add query membership rule. Error message: $($_.Exception.Message)"
            }
        }
    }

    # Check include membership rules
    
    # Check exclude membership rules
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
        if ($Maintain) {
            Write-ToLog -File $LogFile -Message "$($ColName). Collection already exist. Maintenance enabled"
            Reset-Collection -ColName $ColName -ColXML $col
            continue
        } else {
            Write-ToLog -File $LogFile -Message "$($ColName). Collection already exist. Skipping, maintenance not enabled"
            continue
        }
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

    $ColDescription = Get-DefaultIfNull -XMLValue $col.description -Default $DefaultDescription
    $ColRecurCount = Get-DefaultIfNull -XMLValue $col.recurcount -Default $DefaultRecurCount
    $ColRecurInterval = Get-DefaultIfNull -XMLValue $col.recurinterval -Default $DefaultRecurInterval

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