<# 
.SYNOPSIS 
	Reports Rogue Active Sync Devices (devices still syncing after account is disabled)
.DESCRIPTION 
    Behavior has been observed that an active sync device may still be able to send/receive email even when the associated AD account is disabled. This can be problematic, especially considering the context may be an employee exit process.
	This script uses the Get-ActiveSyncDeviceStatistics cmdlet to check when a devices associated with disabled accounts sync, should that timestamp be greater than the accounts last modified date (assuming that date is when the account is disabled) it will create an email alert
.NOTES 
    File Name  : ReportRogueActiveSyncDevices.ps1
    Author     : Brenton keegan - brenton.keegan@gmail.com 
    Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/Exchange-PowerShell-Tools
    License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	ReportRogueActiveSyncDevices -dn "OU=Employees,DC=CONTOSO,DC=COM" -c "cas.contoso.com" -To "alerts@contoso.com" -From "Alerts@contoso.com" -smtp "smtp.contoso.com" -i 60
#> 


#imports
import-module activedirectory


Function ReportRogueActiveSyncDevices
{
	
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[alias("dn")]
		[string]$dnToMonitor,

		[parameter(Mandatory=$true)]
		[alias("c")]
		[string]$casServer,

		[parameter(Mandatory=$true)]
		[alias("To")]
		[string]$emailRecipient,
		
		[parameter(Mandatory=$true)]
		[alias("From")]
		[string]$emailSender,
		
		[parameter(Mandatory=$true)]
		[alias("smtp")]
		[string]$emailServer,
		
		[parameter(Mandatory=$false)]
		[alias("Subject")]
		[string]$emailSubject="ActiveSync Alert",
		
		[parameter(Mandatory=$true)]
		[alias("i")]
		[int]$monitorInterval
	)
	
	$existingSession = get-pssession
	if(!($existingSession))
	{
		$session = New-PSSession -configurationname Microsoft.Exchange -connectionURI http://$casServer/PowerShell
		Import-PSSession $session
	}
		
	while ($true)
	{
		$inactiveUsers = get-aduser -filter * -searchbase $dnToMonitor -property whenchanged | where {$_.Enabled -eq $false}
		$rogueDevices = $false
		$objects = @()
		foreach($inactiveUser in $inactiveUsers)
		{
			
			$devices = Get-ActiveSyncDeviceStatistics -mailbox "$($inactiveUser.SamAccountName)"
			foreach($device in $devices)
			{
				
				$deviceSyncVariance = $device.LastSuccessSync.AddMinutes(30)
				If($inactiveUser.whenchanged -lt $deviceSyncVariance)
				{
					#since detecting when an account went inactive is based on when the account was last changed, a 30 minute buffer is added as the last change might not be the moment the account was disable
					#30 minutes may account for a sysadmin executing a exit process on an account.
					$checkInterval = $device.LastSuccessSync.AddSeconds($monitorInterval)
					$currentTime = get-date
					if($checkInterval -ge $currentTime)
					{
						$userStats = New-Object PSObject        
						$userStats | add-member Noteproperty User $inactiveUser.SamAccountName 
						$userStats | add-member Noteproperty UserLastChanged $inactiveUser.whenchanged 
						$userStats | add-member Noteproperty LastSuccessfulSync $device.LastSuccessSync 
						$userStats | add-member Noteproperty FriendlyName $device.DeviceFriendlyName
						$objects += $userStats 
						$rogueDevices = $true
					}
				}
			}
		}
		If($rogueDevices)
		{
			$emailBody = $objects | out-string
			Send-MailMessage -To $emailRecipient -Subject $emailSubject -smtp $emailServer -From $emailSender -body $emailBody
		}
		sleep $monitorInterval
		
	}
}
