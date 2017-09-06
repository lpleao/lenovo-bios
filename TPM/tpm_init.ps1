try
{
    Enable-TpmAutoProvisioning
    Initialize-Tpm -AllowClear -AllowPhysicalPresence
    exit 0
}
catch
{
    exit 1
}
