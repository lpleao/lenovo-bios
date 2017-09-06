try
{
    $tpm=get-wmiobject -class Win32_Tpm -namespace root\cimv2\security\microsofttpm
    $tpm.DisableAutoprovisioning()
    $tpm.SetPhysicalPresenceRequest(16)
    $tpm.SetPhysicalPresenceRequest(18)
    $tpm.SetPhysicalPresenceRequest(21)
    exit 0
}
catch
{
    exit 1
}
