Write-Output "Quering Win32_TPM WMI object..."
$TPM = Get-WmiObject -Class "Win32_Tpm" -Namespace "ROOT\CIMV2\Security\MicrosoftTpm"

Write-Output "Clearing TPM ownership....."
$tmp = $TPM.SetPhysicalPresenceRequest(5)
if ($tmp.ReturnValue -eq 0) {
    Write-Output "Successfully cleared the TPM chip. A reboot is required."
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $TSEnv.Value("NeedRebootTpmClear") = "YES"
    exit 0
} else {
    Write-Warning "Failed to clear TPM ownership. Exiting..."
    Stop-Transcript
    exit 0
}