# Intune Detection Script - RemoteFX USB Device Redirection
# Exit 0 = Compliant (no remediation needed)
# Exit 1 = Non-compliant (remediation needed)

$regPath_USB   = "HKLM:\System\CurrentControlSet\Control\Class\{36fc9e60-c465-11cf-8056-444553540000}"
$regPath_FLT   = "HKLM:\System\CurrentControlSet\Services\TsUsbFlt"
$regPath_hubg  = "HKLM:\System\CurrentControlSet\Services\usbhub\hubg"

try {
    # Check UpperFilters contains TsUSBFlt
    $upperFilters = (Get-ItemProperty -Path $regPath_USB -ErrorAction Stop).UpperFilters
    if ($upperFilters -notcontains "TsUSBFlt") {
        Write-Output "Non-compliant: TsUSBFlt missing from UpperFilters"
        exit 1
    }

    # Check BootFlags = 4
    $bootFlags = (Get-ItemProperty -Path $regPath_FLT -ErrorAction Stop).BootFlags
    if ($bootFlags -ne 4) {
        Write-Output "Non-compliant: BootFlags is not set to 4"
        exit 1
    }

    # Check EnableDiagnosticMode
    if (-not (Test-Path -Path $regPath_hubg)) {
        Write-Output "Non-compliant: Registry path usbhub\hubg does not exist"
        exit 1
    }
    $diagMode = (Get-ItemProperty -Path $regPath_hubg -ErrorAction Stop).EnableDiagnosticMode
    if ($diagMode -ne 0) {
        Write-Output "Non-compliant: EnableDiagnosticMode is not set correctly"
        exit 1
    }

    Write-Output "Compliant: RemoteFX USB Device Redirection is correctly configured"
    exit 0
}
catch {
    Write-Output "Non-compliant: Error during detection - $_"
    exit 1
}
