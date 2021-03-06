<#
.SYNOPSIS
    Create ConfigMgr operational collections based on a XML definition file

.DESCRIPTION
    Creates and maintains ConfigMgr collections imported from an XML file
    See XML file for additional definition information

.PARAMETER SiteServer
    [Mandatory] ConfigMgr Site Server FQDN

.PARAMETER SiteCode
    [Mandatory] ConfigMgr Site Code (3 characters)

.PARAMETER CollectionsXML
    [Mandatory] XML file path containing collection definition

.PARAMETER Maintain
    [Optional] Enables maintenance mode to correct collections that deviate from the XML definition

.NOTES
    Author: Bernardo Achirica (@bachirica)
    Version: 2.1
    Date: 2019.12.31
    References: Idea based on Mark Allen's script (https://github.com/markhallen/configmgr/tree/master/New-CMOperationalCollections)

    Version History:
    1.0 - 2019.01.11: Initial release
    2.0 - 2019.01.17: Added Maintain parameter to correct collections that deviate from the XML definition
    2.1 - 2019.12.31: Added the option to handle RefreshType in the XML definition
                      Consolidate logging lines under INFO, WARN, ERROR

.EXAMPLE
    .\Add-OperationalCollections.ps1 -SiteServer mysccmserver.mydomain.local -SiteCode PR1 -CollectionsXML .\OperationalCollections.xml

    Creates the collections found in the XML. Already existing collections are not modified

.EXAMPLE
    .\Add-OperationalCollections.ps1 -SiteServer mysccmserver.mydomain.local -SiteCode PR1 -CollectionsXML .\OperationalCollections.xml -Maintain

    Creates the collections found in the XML. Already existing collections are corrected if they deviate from the XML definition
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
[string]$DefaultRefreshType = "Periodic"
[string]$DefaultDescription = "Operational Collection"

# $SchedAddHours adds some extra hours to the collection evaluation schedule
# If a collection refreshes every 7 days, it'll do it always at the same time the collection was created (probably during office hours)
# Adding a few hours can allow you to schedule those refresh cycles outside office hours
[int]$SchedAddHours = 0

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
    Write-ToLog -File $LogFile -Message "INFO: Created collection folder $($FolderPath)"
}

function Get-DefaultIfNull ([string]$XMLValue, [string]$Default) {
    if (($null -ne $XMLValue) -and ($XMLValue -ne "")) {
        return $XMLValue 
    } else {
        return $Default
    }
}

Function Add-Collection ([string]$ColName, [string]$ColLimiting, [string]$ColDescription, [string]$ColRecurInterval, [int]$ColRecurCount, [string]$ColRefreshType) {
    $Schedule = New-CMSchedule -RecurInterval $ColRecurInterval -RecurCount $ColRecurCount -Start (Get-Date).AddHours($SchedAddHours)
    try {
        New-CMDeviceCollection -Name $ColName -LimitingCollectionName $ColLimiting -Comment $ColDescription -RefreshSchedule $Schedule -RefreshType $ColRefreshType | Out-Null
        Write-ToLog -File $LogFile -Message "INFO: $($ColName). Collection created"
    }
    catch {
        Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Error creating collection. Error message: $($_.Exception.Message)"
    }
}

Function Add-CollectionQuery ([string]$ColName, [string]$Query) {
    try {
        Add-CMDeviceCollectionQueryMembershipRule -CollectionName $ColName -RuleName $ColName -QueryExpression $Query
        Write-ToLog -File $LogFile -Message "INFO: $($ColName). Added collection membership query to collection"
    }
    catch {
        Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not add collection membership query. Error message: $($_.Exception.Message)"
    }
}

Function Add-CollectionInclude ([string]$ColName, [string]$Include) {
    if ((Get-CMDeviceCollection -Name $ColName).Name -eq $ColName) {
        try {
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $ColName -IncludeCollectionName $Include
            Write-ToLog -File $LogFile -Message "INFO: $($ColName). Added include membership rule"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not add include membership rule. Error message: $($_.Exception.Message)"
        }
    } else {
        Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Include membership collection $($Include) doesn't exist"
    }
}

Function Add-CollectionExclude ([string]$ColName, [string]$Exclude) {
    if ((Get-CMDeviceCollection -Name $ColName).Name -eq $ColName) {
        try {
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $ColName -ExcludeCollectionName $Exclude
            Write-ToLog -File $LogFile -Message "INFO: $($ColName). Added exclude membership rule"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not add exclude membership rule. Error message: $($_.Exception.Message)"
        }
    } else {
        Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Exclude membership collection $($Exclude) doesn't exist"
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
            Write-ToLog -File $LogFile -Message "WARN: $($ColName). Corrected limiting collection to $($ColXML.limiting)"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not correct limiting collection to $($ColXML.limiting). Error message: $($_.Exception.Message)"
        }
    }

    # Check description
    $ColDescription = Get-DefaultIfNull -XMLValue $ColXML.description -Default $DefaultDescription

    if ($CMCol.Comment -ne $ColDescription) {
        try {
            Set-CMCollection -InputObject $CMCol -Comment $ColDescription
            Write-ToLog -File $LogFile -Message "WARN: $($ColName). Corrected collection comment"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not correct collection comment. Error message: $($_.Exception.Message)"
        }
    }

    # Check refresh type
    $ColRefreshType = Get-DefaultIfNull -XMLValue $col.refreshtype -Default $DefaultRefreshType

    switch ($ColRefreshType) {
        "None" { $intRefreshType = 0 }
        "Manual" { $intRefreshType = 1 }
        "Periodic" { $intRefreshType = 2 }
        "Continuous" { $intRefreshType = 4 }
        "Both" { $intRefreshType = 6 }
        #Default { $intRefreshType = 2 }
    }

    if ($CMCol.RefreshType -ne $intRefreshType) {
        try {
            Set-CMCollection -InputObject $CMCol -RefreshType $ColRefreshType
            Write-ToLog -File $LogFile -Message "WARN: $($ColName). Corrected refresh type to $($ColRefreshType)"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not correct refresh type. Error message: $($_.Exception.Message)"
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
            Write-ToLog -File $LogFile -Message "WARN: $($ColName). Corrected refresh schedule"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not correct refresh schedule. Error message: $($_.Exception.Message)"
        }
    }

    # Remove all direct memberships
    $ColDirects = Get-CMDeviceCollectionDirectMembershipRule -CollectionName $ColName

    if ($ColDirects.Count -gt 0) {
        try {
            Get-CMDeviceCollectionDirectMembershipRule -CollectionName $ColName | ForEach-Object {Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $ColName -ResourceId $_.ResourceID -Force}
            Write-ToLog -File $LogFile -Message "WARN: $($ColName). Removed direct membership rules"
        }
        catch {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not remove direct membership rules. Error message: $($_.Exception.Message)"
        }
    }
    
    # Check collection queries
    $XMLQueries = $ColXML.query
    $ColQueries = Get-CMCollectionQueryMembershipRule -CollectionName $ColName

    if (($XMLQueries.Length -gt 0) -or ($ColQueries.Count -gt 0)) {
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
                    Write-ToLog -File $LogFile -Message "WARN: $($ColName). Removed query membership rule"
                }
                catch {
                    Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not remove query membership rule. Error message: $($_.Exception.Message)"
                }
            }
        }

        if ($hashNotInCol.Count -gt 0) {
            foreach ($h in $hashNotInCol.Keys) {
                try {
                    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $ColName -RuleName $ColName -QueryExpression "$($hashNotInCol.Item($h))"
                    Write-ToLog -File $LogFile -Message "WARN: $($ColName). Added query membership rule"
                }
                catch {
                    Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not add query membership rule. Error message: $($_.Exception.Message)"
                }
            }
        }
    }

    # Check include membership rules
    $XMLIncludes = $ColXML.include
    $ColIncludes = Get-CMCollectionIncludeMembershipRule -CollectionName $ColName

    if (($XMLIncludes.Length -gt 0) -or ($ColIncludes.Count -gt 0)) {
        $i = 1
        $hashXML = @{}
        $hashCol = @{}
        $hashNotInCol = @{} # Elements in XML and not in the collection (to be added)
        $hashNotInXML = @{} # Elements in collection and not in the XML (to be removed)

        if ($XMLIncludes.Length -gt 0) {
            foreach ($Include in $XMLIncludes) {
                $hashXML.Add("$($i)", "$($Include)")
                $i++
            }
        }

        if ($ColIncludes.Count -gt 0) {
            foreach ($Include in $ColIncludes) {
                $hashCol.Add("$($Include.IncludeCollectionID)", "$($Include.RuleName)")
            }
        }

        Compare-HashTables -hashXML $hashXML -hashCol $hashCol -hashNotInXML $hashNotInXML -hashNotInCol $hashNotInCol

        if ($hashNotInXML.Count -gt 0) {
            foreach ($h in $hashNotInXML.Keys) {
                try {
                    Remove-CMDeviceCollectionIncludeMembershipRule -CollectionName $ColName -IncludeCollectionName $hashNotInXML.Item($h) -Force
                    Write-ToLog -File $LogFile -Message "WARN: $($ColName). Removed include membership rule ($($hashNotInXML.Item($h)))"
                }
                catch {
                    Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not remove include membership rule ($($hashNotInXML.Item($h))). Error message: $($_.Exception.Message)"
                }
            }
        }

        if ($hashNotInCol.Count -gt 0) {
            foreach ($h in $hashNotInCol.Keys) {
                try {
                    Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $ColName -IncludeCollectionName $hashNotInCol.Item($h)
                    Write-ToLog -File $LogFile -Message "WARN: $($ColName). Added include membership rule ($($hashNotInCol.Item($h)))"
                }
                catch {
                    Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not add include membership rule ($($hashNotInCol.Item($h))). Error message: $($_.Exception.Message)"
                }
            }
        }
    }

    # Check exclude membership rules
    $XMLExcludes = $ColXML.exclude
    $ColExcludes = Get-CMCollectionExcludeMembershipRule -CollectionName $ColName

    if (($XMLExcludes.Length -gt 0) -or ($ColExcludes.Count -gt 0)) {
        $i = 1
        $hashXML = @{}
        $hashCol = @{}
        $hashNotInCol = @{} # Elements in XML and not in the collection (to be added)
        $hashNotInXML = @{} # Elements in collection and not in the XML (to be removed)

        if ($XMLExcludes.Length -gt 0) {
            foreach ($Exclude in $XMLExcludes) {
                $hashXML.Add("$($i)", "$($Exclude)")
                $i++
            }
        }

        if ($ColExcludes.Count -gt 0) {
            foreach ($Exclude in $ColExcludes) {
                $hashCol.Add("$($Exclude.ExcludeCollectionID)", "$($Exclude.RuleName)")
            }
        }

        Compare-HashTables -hashXML $hashXML -hashCol $hashCol -hashNotInXML $hashNotInXML -hashNotInCol $hashNotInCol

        if ($hashNotInXML.Count -gt 0) {
            foreach ($h in $hashNotInXML.Keys) {
                try {
                    Remove-CMDeviceCollectionExcludeMembershipRule -CollectionName $ColName -ExcludeCollectionName $hashNotInXML.Item($h) -Force
                    Write-ToLog -File $LogFile -Message "WARN: $($ColName). Removed exclude membership rule ($($hashNotInXML.Item($h)))"
                }
                catch {
                    Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not remove exclude membership rule ($($hashNotInXML.Item($h))). Error message: $($_.Exception.Message)"
                }
            }
        }

        if ($hashNotInCol.Count -gt 0) {
            foreach ($h in $hashNotInCol.Keys) {
                try {
                    Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $ColName -ExcludeCollectionName $hashNotInCol.Item($h)
                    Write-ToLog -File $LogFile -Message "WARN: $($ColName). Added exclude membership rule ($($hashNotInCol.Item($h)))"
                }
                catch {
                    Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Could not add exclude membership rule ($($hashNotInCol.Item($h))). Error message: $($_.Exception.Message)"
                }
            }
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
        if ($Maintain) {
            #Write-ToLog -File $LogFile -Message "INFO: $($ColName). Collection already exist. Maintenance enabled"
            Reset-Collection -ColName $ColName -ColXML $col
            continue
        } else {
            Write-ToLog -File $LogFile -Message "INFO: $($ColName). Collection already exist. Skipping, maintenance not enabled"
            continue
        }
    } else {
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
                Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Limiting collection does not exist"
                continue
            }
        } else {
            Write-ToLog -File $LogFile -Message "ERROR: $($ColName). Limiting collection is missing in XML"
            continue
        }

        $ColDescription = Get-DefaultIfNull -XMLValue $col.description -Default $DefaultDescription
        $ColRecurCount = Get-DefaultIfNull -XMLValue $col.recurcount -Default $DefaultRecurCount
        $ColRecurInterval = Get-DefaultIfNull -XMLValue $col.recurinterval -Default $DefaultRecurInterval
        $ColRefreshType = Get-DefaultIfNull -XMLValue $col.refreshtype -Default $DefaultRefreshType

        # Create the empty collection
        Add-Collection -ColName $ColName -ColLimiting $ColLimiting -ColDescription $ColDescription -ColRecurInterval $ColRecurInterval -ColRecurCount $ColRecurCount -ColRefreshType $ColRefreshType

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
}
#endregion