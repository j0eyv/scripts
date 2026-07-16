<#
.SYNOPSIS
    Intune Proactive Remediation - Detection script
    Checks HKLM\SOFTWARE\Policies\Microsoft\Windows\AAD -> AutoAcceptSsoPermission (DWORD) = 1

.NOTES
    Exit 0 = Compliant (no remediation needed)
    Exit 1 = Non-compliant (remediation will run)
#>

$RegPath   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AAD'
$ValueName = 'AutoAcceptSsoPermission'
$Expected  = 1

try {
    $current = Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction Stop |
        Select-Object -ExpandProperty $ValueName

    if ($current -eq $Expected) {
        Write-Output "Compliant: $ValueName = $current"
        exit 0
    }
    else {
        Write-Output "Non-compliant: $ValueName = $current (expected $Expected)"
        exit 1
    }
}
catch {
    Write-Output "Non-compliant: $ValueName not found at $RegPath"
    exit 1
}
