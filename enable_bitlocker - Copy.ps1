#Enabling Encryption
Enable-Bitlocker -MountPoint $env:SystemDrive -UsedSpaceOnly -RecoveryPasswordProtector

#Getting Recovery Key GUID
$RecoveryKeyGUID = (Get-BitLockerVolume -MountPoint $env:SystemDrive).keyprotector | where {$_.Keyprotectortype -eq 'RecoveryPassword'} | Select-Object -ExpandProperty KeyProtectorID

#Backing up the Recovery to AD.
$volume = Get-BitLockerVolume -MountPoint "C:"
Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $volume.KeyProtector[1].KeyProtectorId