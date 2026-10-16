# Dev Drive scripts

These scripts create and remove a persistent Windows Dev Drive backed by a VHDX. They are intended for elevated, unattended deployment, including Intune and PSADT workflows.

## `New-DevDrive.ps1`

Creates or reuses `drive.vhdx`, prepares a ReFS Dev Drive, verifies it with `fsutil devdrv query`, and registers a startup scheduled task to mount the VHDX and assign its drive letter after reboot. It writes a detection marker to `HKLM:\Software\CustomDeviceManagement\DevDrive` only after verification and task registration succeed.

By default, the VHDX and logs are stored in `C:\ProgramData\Custom Dev Drive`, with a maximum size of 55 GB and drive letter `V:`. The size is dynamically expanding; the VHDX file grows as data is written.

### Requirements and safeguards

- Run in elevated, 64-bit Windows PowerShell on a Windows version that supports `Format-Volume -DevDrive`.
- A new Dev Drive must be at least 50 GB. `DriveSize` accepts MB, GB, or TB values.
- When creating a new VHDX, `DrivePath` must contain only ASCII characters because it is passed through a DiskPart command file. Existing VHDXs can be reused.
- Existing formatted volumes are not reformatted: they must already be ReFS and pass Dev Drive verification. The script refuses conflicting drive letters and unexpected or multiple partitions.
- Invoke the script directly with `&` from PSADT so terminating errors reach its error handler. If running it in a child PowerShell process, configure the deployment to check the process exit code.

### Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `DrivePath` | `C:\ProgramData\Custom Dev Drive` | Absolute local directory for `drive.vhdx`, logs, and the startup mount script. |
| `DriveSize` | `55GB` | Maximum dynamically expanding VHDX size; minimum 50 GB. |
| `DriveLetter` | `V` | Drive letter to assign. Must be an available letter from D through Z. |

Example:

```powershell
& ".\New-DevDrive.ps1" -DrivePath "C:\DevDrive" -DriveSize 100GB -DriveLetter V
```

The installer writes `install.log`, generates `devdrivetask.ps1` for startup mounting, and may write `tasklog.txt` if the startup task encounters an error.

## `Remove-DevDrive.ps1`

Creates a temporary DiskPart command file, selects `C:\ProgramData\Custom Dev Drive\drive.vhdx`, and detaches it. It then attempts to remove the drive directory, unregisters the `Mount Dev Drive` scheduled task, and removes the Dev Drive detection registry value and key.

The removal script has no parameters and is hard-coded to the installer's default `DrivePath`. It does **not** use recursive deletion; because the directory normally contains the VHDX and logs, its directory removal may fail, leaving those files in place. The VHDX is detached, not explicitly deleted. Run elevated, and check the command output and remaining files before considering cleanup complete.

## Typical use

1. Run `New-DevDrive.ps1` elevated in 64-bit Windows PowerShell, supplying parameters if needed.
2. Confirm that the selected drive letter is available and that `install.log` reports successful completion.
3. To remove the default installation, run `Remove-DevDrive.ps1` elevated and verify the scheduled task, registry marker, and remaining files afterward.

## PSADT
The installer and uninstaller scripts are PSADT proof and have been tested with PSADT 4.1.8