<#
.SYNOPSIS
    Intune Proactive Remediation - Remediation script
    Sets HKLM\SOFTWARE\Policies\Microsoft\Windows\AAD -> AutoAcceptSsoPermission (DWORD) = 1

.NOTES
    Exit 0 = Remediation succeeded
    Exit 1 = Remediation failed
#>

$RegPath   = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AAD'
$ValueName = 'AutoAcceptSsoPermission'
$Value     = 1

try {
    if (-not (Test-Path -Path $RegPath)) {
        New-Item -Path $RegPath -Force -ErrorAction Stop | Out-Null
    }

    New-ItemProperty -Path $RegPath -Name $ValueName -Value $Value -PropertyType DWord -Force -ErrorAction Stop | Out-Null

    # Verify
    $current = Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction Stop |
        Select-Object -ExpandProperty $ValueName

    if ($current -eq $Value) {
        Write-Output "Remediation succeeded: $ValueName = $current"
        exit 0
    }
    else {
        Write-Output "Remediation failed: $ValueName = $current (expected $Value)"
        exit 1
    }
}
catch {
    Write-Output "Remediation failed: $($_.Exception.Message)"
    exit 1
}
