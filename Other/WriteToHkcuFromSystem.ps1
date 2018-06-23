PARAM(

    [Parameter(Mandatory=$true)]
    [ValidatePattern('\.reg$')]
    [string]$RegFile,
	
    [switch]$CurrentUser,
    [switch]$AllUsers,
    [switch]$DefaultProfile
)

function Get-TempRegFilePath {
    (Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ([guid]::NewGuid().Guid)) + '.reg'
}


function Write-Registry {
    param($RegFileContents, $UserSid)
    
    $TempRegFile = Get-TempRegFilePath
    $regFileContents = $regFileContents -replace 'HKEY_CURRENT_USER', "HKEY_USERS\$userSid"
    $regFileContents | Out-File -FilePath $TempRegFile
    
    $p = Start-Process -FilePath C:\Windows\regedit.exe -ArgumentList @('/s', $TempRegFile) -PassThru
    do { Start-Sleep -Seconds 1 } while (-not $p.HasExited)
    
    Remove-Item -Path $TempRegFile -Force
}


function Write-RegistryWithHiveLoad {
    param($RegFileContents, $DatFilePath)

    $hiveName = 'x_' + (($datFilePath -split '\\')[-2]).ToUpper() 
    try {
        C:\Windows\System32\reg.exe load "HKU\$hiveName" $DatFilePath
        if($LASTEXITCODE -ne 0) { throw 'Error loading the DAT file' }
    
        $TempRegFile = Get-TempRegFilePath
        $regFileContents = $regFileContents -replace 'HKEY_CURRENT_USER', "HKEY_USERS\$hiveName"
        $regFileContents | Out-File -FilePath $TempRegFile

        $p = Start-Process -FilePath C:\Windows\regedit.exe -ArgumentList @('/s', $TempRegFile) -PassThru
        do { Start-Sleep -Seconds 1 } while (-not $p.HasExited)

        C:\Windows\System32\reg.exe unload "HKU\$hiveName"

        Remove-Item -Path $TempRegFile -Force
    } catch {
        Write-Verbose $_.Exception.Message -Verbose
    }
}


if (-not (Test-Path -Path $RegFile)) {
    throw "RegFile $RegFile doesn't exist. Aborted operation."
}
else {

    # Read the .reg file contents:
    $regFileContents = Get-Content -Path $RegFile -ReadCount 0

    # For the current logged on user only:
    if ($CurrentUser) {
        $explorers = Get-WmiObject -Class Win32_Process -Filter "Name='Explorer.exe'"
        $explorers | ForEach-Object {
            $owner = $_.GetOwner()
            if ($owner.ReturnValue -eq 0) {
                $user = "{0}\{1}" -f $owner.Domain, $owner.User
                $oUser = New-Object -TypeName System.Security.Principal.NTAccount($user)
                $sid = $oUser.Translate([System.Security.Principal.SecurityIdentifier]).Value
                Write-Verbose ('Writing registry values for user: {0}' -f $user) -Verbose
                Write-Registry -RegFileContents $regFileContents -UserSid $sid
            }
        }
    }

    # For the default profile (future users):
    if ($DefaultProfile) {
        Write-Verbose ('Writing registry values for the DEFAULT PROFILE') -Verbose
        Write-RegistryWithHiveLoad -RegFileContents $regFileContents -DatFilePath C:\Users\Default\NTUSER.DAT
    }

    # For all users that have profiles on the machine:
    if ($AllUsers) {
        dir -Path (Split-Path -Path $env:Public -Parent) -Exclude Public, Default | ForEach-Object {
            Write-Verbose ('Writing registry values for user: {0}' -f $_.Name) -Verbose
            $datFilePath = Join-Path -Path $_.FullName -ChildPath NTUSER.DAT
            Write-RegistryWithHiveLoad -RegFileContents $regFileContents -DatFilePath $datFilePath
        }
    }

}


