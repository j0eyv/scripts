#Source: https://gist.github.com/lawndoc/ea03e2ee0f4d64162669d1b5e997ec77
$DrivePath="C:\ProgramData\Custom Dev Drive"
$FileName = "${env:LOCALAPPDATA}\Temp\diskpart_devdrive.txt"
Write-Output "select vdisk file='${DrivePath}\drive.vhdx'" | Out-File -Encoding ascii -FilePath $FileName
Write-Output "detach vdisk" | Out-File -Encoding ascii -Append -FilePath $FileName
diskpart /s $FileName
rm $DrivePath
Unregister-ScheduledTask -TaskName "Mount Dev Drive"
Remove-ItemProperty -Path "HKLM:\Software\CustomDeviceManagement\DevDrive" -Name "DevDriveDetection" -Force
Remove-Item -Path "HKLM:\Software\CustomDeviceManagement\DevDrive" -Force