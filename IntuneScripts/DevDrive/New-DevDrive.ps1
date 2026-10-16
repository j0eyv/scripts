<#
.SYNOPSIS
    Creates a persistent Dev Drive for unattended Intune / PSADT deployment.
.DESCRIPTION
    Creates a 55 GB dynamically expanding VHDX by default, mounted as V:.
    Requires elevated, 64-bit Windows PowerShell and Windows Dev Drive support.
    Does not require the Hyper-V PowerShell module.

    Retries reuse the specified VHDX. Only an unformatted partition is formatted;
    an existing formatted volume must be ReFS and pass fsutil devdrv query.
    Existing volumes are never reformatted or deleted.

    Invoke with & from PSADT so terminating errors reach its error handler.
    When launched in a child PowerShell process, configure PSADT to check its
    exit code. Logs and the startup mount script are stored in DrivePath.
    The 64-bit registry detection marker is written only after verification
    and successful registration of the SYSTEM startup task.
.PARAMETER DrivePath
    Absolute local directory containing drive.vhdx and deployment logs.
    Creating a new VHDX requires an ASCII-only path for the DiskPart script.
.PARAMETER DriveSize
    Maximum VHDX size in MB, GB or TB. Minimum 50 GB. Default: 55GB.
.PARAMETER DriveLetter
    Drive letter to assign. Default: V.
.EXAMPLE
    & "$dirFiles\New-DevDrive.ps1"
.EXAMPLE
    .\New-DevDrive.ps1 -DrivePath "C:\DevDrive" -DriveSize 100GB -DriveLetter V
#>
[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z]:\\')]
    [string]$DrivePath = 'C:\ProgramData\Custom Dev Drive',
    [ValidatePattern('^\d+(\.\d+)?(MB|GB|TB)$')]
    [string]$DriveSize = '55GB',
    [ValidatePattern('^[D-Zd-z]$')]
    [string]$DriveLetter = 'V'
)

$ErrorActionPreference = 'Stop'
$DriveLetter = $DriveLetter.ToUpperInvariant()
$DrivePath = [System.IO.Path]::GetFullPath($DrivePath)
$vhdPath = Join-Path $DrivePath 'drive.vhdx'
$logPath = Join-Path $DrivePath 'install.log'
$diskpartTempFile = $null
$logInitialized = $false

function Write-DevDriveLog {
    param([string]$Message)
    Add-Content -LiteralPath $logPath -Value ('{0:u} {1}' -f (Get-Date), $Message) -ErrorAction Stop
    Write-Output $Message
}

try {
    if (-not [Environment]::Is64BitProcess) {
        throw 'Run this installer in 64-bit Windows PowerShell, not a 32-bit host.'
    }
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Run this installer elevated or as SYSTEM.'
    }
    if (-not (Get-Command Format-Volume -ErrorAction Stop).Parameters.ContainsKey('DevDrive')) {
        throw 'This Windows installation does not support Format-Volume -DevDrive. Update Windows before installing.'
    }
    if ($DrivePath.Contains('"') -or $DrivePath.Contains("`r") -or $DrivePath.Contains("`n")) {
        throw 'DrivePath cannot contain quotes or newline characters.'
    }
    $sizeMatch = [regex]::Match($DriveSize, '^(\d+(?:\.\d+)?)(MB|GB|TB)$', 'IgnoreCase')
    $size = [double]::Parse($sizeMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
    $multiplier = switch ($sizeMatch.Groups[2].Value.ToUpperInvariant()) {
        'MB' { 1 }
        'GB' { 1024 }
        'TB' { 1048576 }
    }
    $sizeInMB = [long][Math]::Ceiling($size * $multiplier)
    if ($sizeInMB -lt 51200) {
        throw 'A Dev Drive must be at least 50 GB.'
    }

    New-Item -Path $DrivePath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    $logInitialized = $true
    Write-DevDriveLog "Starting as $($identity.Name); VHDX=$vhdPath; drive=${DriveLetter}:."

    if (-not (Test-Path -LiteralPath $vhdPath -PathType Leaf)) {
        if ($vhdPath -match '[^\x00-\x7F]') {
            throw 'Creating a VHDX requires an ASCII-only DrivePath so the DiskPart command file preserves the path.'
        }
        if (Get-PSDrive -PSProvider FileSystem | Where-Object Name -eq $DriveLetter) {
            throw "Drive ${DriveLetter}: is already in use. No VHDX was created."
        }
        $diskpartTempFile = Join-Path $DrivePath ('diskpart_{0}.txt' -f [guid]::NewGuid())
        # DiskPart requires a plain-text command file, not UTF-16 with a BOM.
        'create vdisk file="{0}" maximum={1} type=expandable' -f $vhdPath, $sizeInMB |
            Set-Content -LiteralPath $diskpartTempFile -Encoding ASCII -ErrorAction Stop
        $diskpartOutput = & "$env:SystemRoot\System32\diskpart.exe" /s $diskpartTempFile 2>&1
        $diskpartExitCode = $LASTEXITCODE
        Write-DevDriveLog ($diskpartOutput -join [Environment]::NewLine)
        if ($diskpartExitCode -ne 0 -or -not (Test-Path -LiteralPath $vhdPath -PathType Leaf)) {
            throw "DiskPart failed to create the VHDX (exit code $diskpartExitCode). Output: $($diskpartOutput -join [Environment]::NewLine). See $logPath."
        }
    }

    $image = Get-DiskImage -ImagePath $vhdPath -ErrorAction Stop
    if (-not $image.Attached) {
        Mount-DiskImage -ImagePath $vhdPath -ErrorAction Stop | Out-Null
    }
    # Resolve the disk through this image; never format a drive letter alone.
    $disk = Get-DiskImage -ImagePath $vhdPath -ErrorAction Stop | Get-Disk -ErrorAction Stop
    if (@($disk).Count -ne 1) {
        throw 'The VHDX did not resolve to exactly one disk.'
    }
    $letterPartitions = @(Get-Partition -ErrorAction Stop | Where-Object DriveLetter -eq $DriveLetter)
    if ($letterPartitions | Where-Object DiskNumber -ne $disk.Number) {
        throw "Drive ${DriveLetter}: belongs to another disk. Nothing will be formatted."
    }
    if (-not $letterPartitions.Count -and
        (Get-PSDrive -PSProvider FileSystem | Where-Object Name -eq $DriveLetter)) {
        throw "Drive ${DriveLetter}: is already in use by another filesystem drive."
    }
    if ($disk.PartitionStyle -eq 'RAW') {
        Initialize-Disk -Number $disk.Number -PartitionStyle GPT -ErrorAction Stop | Out-Null
    }
    $partitions = @(Get-Partition -DiskNumber $disk.Number -ErrorAction Stop |
        Where-Object Type -ne 'Reserved')
    if ($partitions.Count -gt 1) {
        throw 'The VHDX contains multiple data partitions. Refusing to modify it.'
    }
    $partitionCreated = $false
    if ($partitions.Count -eq 0) {
        $partition = New-Partition -DiskNumber $disk.Number -UseMaximumSize -DriveLetter $DriveLetter -ErrorAction Stop
        $partitionCreated = $true
    }
    else {
        $partition = $partitions[0]
        if ($partition.Type -notin @('Basic', 'IFS')) {
            throw "Unexpected partition type '$($partition.Type)'. Refusing to modify it."
        }
        if ($partition.DriveLetter -and $partition.DriveLetter -ne $DriveLetter) {
            throw "The VHDX is already assigned to $($partition.DriveLetter):. Refusing to change it."
        }
        if ($partition.DriveLetter -ne $DriveLetter) {
            $partition | Set-Partition -NewDriveLetter $DriveLetter -ErrorAction Stop
            $partition = Get-Partition -DiskNumber $disk.Number -PartitionNumber $partition.PartitionNumber -ErrorAction Stop
        }
    }
    $volume = $partition | Get-Volume -ErrorAction Stop
    if (@($volume).Count -ne 1) {
        throw 'The VHDX partition did not resolve to exactly one volume.'
    }
    if ($volume.FileSystemType -eq 'RAW' -or
        ($partitionCreated -and $volume.FileSystemType -eq 'Unknown')) {
        Write-DevDriveLog 'Formatting the unformatted VHDX partition as a Dev Drive.'
        $partition | Format-Volume -FileSystem ReFS -DevDrive -NewFileSystemLabel 'Dev Drive' `
            -Force -Confirm:$false -ErrorAction Stop | Out-Null
    }
    elseif ($volume.FileSystemType -ne 'ReFS') {
        throw "The VHDX already contains $($volume.FileSystemType). Existing data will not be reformatted."
    }
    $volume = $partition | Get-Volume -ErrorAction Stop
    if ($volume.FileSystemType -ne 'ReFS' -or $volume.DriveLetter -ne $DriveLetter) {
        throw 'The expected ReFS volume and drive letter were not established.'
    }
    $queryOutput = & "$env:SystemRoot\System32\fsutil.exe" devdrv query "${DriveLetter}:" 2>&1
    $queryExitCode = $LASTEXITCODE
    Write-DevDriveLog ($queryOutput -join [Environment]::NewLine)
    if ($queryExitCode -ne 0) {
        throw "Dev Drive verification failed (exit code $queryExitCode). Check Windows support and Dev Drive policy. Existing volumes will not be reformatted."
    }

    $mountScriptPath = Join-Path $DrivePath 'devdrivetask.ps1'
    $mountScript = @'
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ImagePath,
    [Parameter(Mandatory = $true)][string]$DriveLetter
)
$ErrorActionPreference = 'Stop'
try {
    $image = Get-DiskImage -ImagePath $ImagePath -ErrorAction Stop
    if (-not $image.Attached) {
        Mount-DiskImage -ImagePath $ImagePath -ErrorAction Stop | Out-Null
    }
    $disk = Get-DiskImage -ImagePath $ImagePath -ErrorAction Stop | Get-Disk -ErrorAction Stop
    $partitions = @(Get-Partition -DiskNumber $disk.Number -ErrorAction Stop | Where-Object Type -ne 'Reserved')
    if ($partitions.Count -ne 1) {
        throw 'Expected exactly one data partition in the Dev Drive VHDX.'
    }
    $partition = $partitions[0]
    if ($partition.DriveLetter -ne $DriveLetter) {
        if ($partition.DriveLetter -or
            (Get-Partition -ErrorAction Stop | Where-Object DriveLetter -eq $DriveLetter) -or
            (Get-PSDrive -PSProvider FileSystem | Where-Object Name -eq $DriveLetter)) {
            throw "Cannot assign ${DriveLetter}:; a conflicting drive letter is in use."
        }
        $partition | Set-Partition -NewDriveLetter $DriveLetter -ErrorAction Stop
    }
}
catch {
    Add-Content -LiteralPath (Join-Path $PSScriptRoot 'tasklog.txt') -Value (
        '{0:u} {1}' -f (Get-Date), $_.ToString()
    ) -ErrorAction Stop
    throw
}
'@
    Set-Content -LiteralPath $mountScriptPath -Value $mountScript -Encoding UTF8 -ErrorAction Stop
    $taskAction = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -ImagePath "{1}" -DriveLetter "{2}"' -f
            $mountScriptPath, $vhdPath, $DriveLetter)
    $taskTrigger = New-ScheduledTaskTrigger -AtStartup
    $taskPrincipal = New-ScheduledTaskPrincipal -UserId 'S-1-5-18' -LogonType ServiceAccount -RunLevel Highest
    $taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
    Register-ScheduledTask -TaskName 'Mount Dev Drive' -Description 'Mount Dev Drive after reboot' `
        -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -Settings $taskSettings `
        -Force -ErrorAction Stop | Out-Null

    New-Item -Path 'HKLM:\Software\CustomDeviceManagement\DevDrive' -Force -ErrorAction Stop | Out-Null
    New-ItemProperty -Path 'HKLM:\Software\CustomDeviceManagement\DevDrive' -Name 'DevDriveDetection' -Value 1 `
        -PropertyType DWord -Force -ErrorAction Stop | Out-Null
    Write-DevDriveLog 'Dev Drive installation completed successfully.'
}
catch {
    $failure = $_
    if ($logInitialized) {
        try {
            Write-DevDriveLog "ERROR: $($failure.Exception.Message)"
        }
        catch {
            Write-Warning "Could not write the installation error to ${logPath}: $($_.Exception.Message)"
        }
    }
    Write-Error -Message "Dev Drive installation failed: $($failure.Exception.Message)" -ErrorAction Continue
    throw $failure
}
finally {
    if ($diskpartTempFile -and (Test-Path -LiteralPath $diskpartTempFile)) {
        Remove-Item -LiteralPath $diskpartTempFile -Force -ErrorAction Stop
    }
}
