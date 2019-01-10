<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER SiteCode

.NOTES

.EXAMPLE
#>

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

