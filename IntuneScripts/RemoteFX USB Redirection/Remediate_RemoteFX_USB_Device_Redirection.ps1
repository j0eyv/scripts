# Intune Remediation Script - RemoteFX USB Device Redirection
# Exit 0 = Remediation successful
# Exit 1 = Remediation failed

$regPath_USB   = "HKLM:\System\CurrentControlSet\Control\Class\{36fc9e60-c465-11cf-8056-444553540000}"
$regPath_FLT   = "HKLM:\System\CurrentControlSet\Services\TsUsbFlt"
$regPath_hubg  = "HKLM:\System\CurrentControlSet\Services\usbhub\hubg"

try {
    # Set UpperFilters to include TsUSBFlt
    $upperFilters = (Get-ItemProperty -Path $regPath_USB -ErrorAction SilentlyContinue).UpperFilters
    if ($upperFilters) {
        if ($upperFilters -notcontains "TsUSBFlt") {
            $upperFilters = $upperFilters + "TsUSBFlt"
            Set-ItemProperty -Path $regPath_USB -Name "UpperFilters" -Value $upperFilters -Type MultiString -ErrorAction Stop
        }
    }
    else {
        Set-ItemProperty -Path $regPath_USB -Name "UpperFilters" -Value @('TsUSBFlt') -Type MultiString -ErrorAction Stop
    }

    # Set BootFlags = 4
    Set-ItemProperty -Path $regPath_FLT -Name "BootFlags" -Value 4 -Type DWord -ErrorAction Stop

    # Ensure usbhub\hubg path exists and set EnableDiagnosticMode
    if (-not (Test-Path -Path $regPath_hubg)) {
        New-Item -Path $regPath_hubg -Force -ErrorAction Stop | Out-Null
    }
    Set-ItemProperty -Path $regPath_hubg -Name "EnableDiagnosticMode" -Value 0 -Type DWord -ErrorAction Stop

    Write-Output "Remediation successful: RemoteFX USB Device Redirection has been configured"
    exit 0
}
catch {
    Write-Output "Remediation failed: $_"
    exit 1
}
