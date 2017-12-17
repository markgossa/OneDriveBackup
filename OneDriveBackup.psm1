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
        $PasswordFilePath,
        [Parameter(Mandatory = $true)]
        [string]        
        $Include,
        [Parameter(Mandatory = $true)]
        [string]        
        $Sender,
        [Parameter(Mandatory = $true)]
        [string]        
        $Recipient,
        [Parameter(Mandatory = $true)]
        [string]        
        $SmtpServer,
        [Parameter(Mandatory = $true)]
        [string]        
        $BackupName                    
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
        
        if ($USBDeviceStatus -eq 'Error')
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
            if ((Lock-BitLocker -MountPoint $DriveLetter -ForceDismount).LockStatus -eq 'Locked')
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

    function Unlock-USBHardDisk
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [string]
            $DriveLetter,
            [Parameter(Mandatory = $true)]
            [securestring]
            $Password
        )

        if ((Get-BitLockerVolume $DriveLetter).LockStatus -eq 'Ok')
        {
            return $true
        }
        else 
        {        
            if ((Unlock-BitLocker -MountPoint $DriveLetter -Password $password).LockStatus -eq 'Unlocked')
            {
                return $true
            }
            else 
            {
                return $false    
            }
        }
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

    $Password = Get-Content $PasswordFilePath | ConvertTo-SecureString 

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
    $BackupOutput.USBHardDiskUnlocked = Unlock-USBHardDisk -DriveLetter $DriveLetter -Password $Password

    # Perform the backup
    . "C:\Program Files\Veeam\Endpoint Backup\Veeam.EndPoint.Manager.exe" /backup

    # Lock the USB disk
    $BackupOutput.USBHardDiskLocked = Lock-USBHardDisk -DriveLetter $DriveLetter

    # Dismount the USB disk after the backup completes
    $BackupOutput.USBHardDiskDismounted = Dismount-USBHardDisk -InstanceId $InstanceId

    # Send notification email
    if ($BackupOutput.USBHardDiskLocked -eq $true -and $BackupOutput.USBHardDiskLocked -eq $true)
    {
        Send-MailMessage -To $Recipient -From $Sender -SmtpServer $SmtpServer -Subject "[Success] $BackupName - Disk locked and unmounted"
    }
    elseif ($BackupOutput.USBHardDiskLocked -eq $false)
    {
        Send-MailMessage -To $Recipient -From $Sender -SmtpServer $SmtpServer -Subject "[Warning] $BackupName - Disk not unmounted"
    }
    elseif ($BackupOutput.USBHardDiskDismounted -eq $false)
    {
        Send-MailMessage -To $Recipient -From $Sender -SmtpServer $SmtpServer -Subject "[Fail] $BackupName - Disk not locked"
    }
}
