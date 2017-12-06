function Start-OneDriveBackup
{
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory = $true)]
        [string]        
        $InstanceId,
        [Parameter(Mandatory = $true)]
        [string]        
        $PasswordFilePath = 'C:\Scripts\OneDrive Backup\Password.txt',
        [Parameter(Mandatory = $true)]
        [string]        
        $VeeamConfig = 'C:\Scripts\OneDrive Backup\VAFW-Weekly-Backup-Config.xml',
        [Parameter(Mandatory = $true)]
        [string]        
        $Include = 'E:\'
    )

    function Dismount-USBHardDisk
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $InstanceId
        )

        Disable-PnpDevice -InstanceId $InstanceId -Confirm:$false
        
        $USBDeviceStatus = (Get-PnpDevice -InstanceId $InstanceId).Status
        
        if ($USBDevice.Status -eq 'Error')
        {
            return $true
        }
        else
        {
            return $false
        }    
    }
    
    function Get-USBHardDiskStatus
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $InstanceId
        )

        $USBDeviceStatus = (Get-PnpDevice -InstanceId $InstanceId).Status
        
        return $USBDeviceStatus
    }

    function Lock-USBHardDisk
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $DriveLetter
        )

        if ((Get-BitLockerVolume $DriveLetter).LockStatus -eq 'Locked')
        {
            return $true
        }
        else 
        {
            if ((Lock-BitLocker -MountPoint $DriveLetter).LockStatus -eq 'Locked')
            {
                return $true
            }
            else 
            {
                return $false    
            }
            
        }
    }

    function Mount-USBHardDisk
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $InstanceId
        )

        Enable-PnpDevice -InstanceId $InstanceId -Confirm:$false
        
        if ((Get-USBHardDiskStatus -InstanceId $InstanceId) -eq 'Ok')
        {
            return $true
        }
        else
        {
            return $false
        }
    }

    function Get-USBHardDiskVolumeLetter
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $InstanceId
        )

        $Disk = Get-Disk | Where-Object {$_.Path -match ($InstanceId -split "\\")[-1]}
        $DriveLetter = (Get-Partition -DiskNumber $Disk.DiskNumber).DriveLetter

        return $DriveLetter
    }

    function Start-WindowsServerBackup
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $DriveLetter,
            [Parameter(Mandatory = $true)]
            [string]
            $Include
        )

        $BackupTarget = $DriveLetter + ":\"
        wbadmin start backup -backupTarget:$BackupTarget -include:$Include -vssFull -quiet
    }

    class BackupOutput
    {
        [boolean]
        $USBHardDiskMounted
        [string]
        $DriveLetter
        [boolean]
        $USBHardDiskUnlocked
        [boolean]
        $BackupSuccessful
        [boolean]
        $USBHardDiskLocked
        [boolean]
        $USBHardDiskDismounted        
    }

    $BackupOutput = New-Object BackupOutput

    # Mount the USB disk if needed
    if ((Get-USBHardDiskStatus -InstanceId $InstanceId) -eq 'OK')
    {
        $BackupOutput.USBHardDiskMounted = $true
    }
    else
    {
        $BackupOutput.USBHardDiskMounted = Mount-USBHardDisk -InstanceId $InstanceId
    }

    # Get USB disk drive letter
    $DriveLetter = Get-USBHardDiskVolumeLetter -InstanceId $InstanceId
    $BackupOutput.DriveLetter = $DriveLetter

    # Unlock the USB disk
    Unlock-BitLocker -MountPoint $DriveLetter -Password $password

    # Perform the backup
    Start-WindowsServerBackup -DriveLetter $DriveLetter -Include $Include

    # Lock the USB disk
    $BackupOutput.USBHardDiskLocked = Lock-USBHardDisk -DriveLetter $DriveLetter

    # Dismount the USB disk after the backup completes
    $BackupOutput.USBHardDiskDismounted = Dismount-USBHardDisk -InstanceId $InstanceId

}


# Bug with Dismount-USBHardDisk (Generic failure)

# Create function to unlock the USB drive