<# 
.SYNOPSIS 
	Sends an email alert when the specified mailbox is accessed by a user of a different name.
.DESCRIPTION 
    Sends an email alert when the specified mailbox is accessed by a user of a different name. Based on results from exchange cmdlet "Get-LogonStatistics"
.NOTES 
    File Name  : MailboxAccessAlerter.ps1
    Author     : Brenton keegan - brenton.keegan@gmail.com 
    Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/Exchange-PowerShell-Tools
    License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	MailboxAccessAlerter -c casServer -m mbServer -t "User Mailbox" -i 600 -r "notify@contoso.com" -smtp "mail.contoso.com" -f "notify@contoso.com"
#>
Function MailboxAccessAlerter
{
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[alias("c")]
		[string]$casServer,
		
		[parameter(Mandatory=$true)]
		[alias("m")]
		[string]$mbServer,
		
		[parameter(Mandatory=$true)]
		[alias("t")]
		[string]$mailboxToMonitor,
		
		[parameter(Mandatory=$true)]
		[alias("i")]
		[int]$monitorInterval,
	
		[parameter(Mandatory=$true)]
		[alias("r")]
		[string]$alertRecipient,
		
		[parameter(Mandatory=$true)]
		[alias("smtp")]
		[string]$smtpServer,
		
		[parameter(Mandatory=$true)]
		[alias("f")]
		[string]$smtpSender
		
		
	)

	while ($true)
	{
		$mismatchedUsers = Get-ExchangeUserMailboxMismatch -c $casServer -m $mbServer
		foreach ($mismatchedUser in $mismatchedUsers)
		{
			if($mismatchedUser.Username -eq $mailboxToMonitor)
			{
				$checkInterval = $mismatchedUser.LastAccessTime.AddSeconds($monitorInterval)
				$currentTime = Get-Date
				if($checkInterval -ge $currentTime)
				{
					
					$emailSubject = "User $($mismatchedUser.Windows2000Account) accessed $($mismatchedUser.Username)"
					$emailString = $mismatchedUser | select Username,Windows2000Account,LastAccessTime | FL * | Out-String
					Send-MailMessage -To $alertRecipient -Subject $emailSubject -smtpServer $smtpServer -From $smtpSender -body $emailString
				}
			}
		}
		sleep $monitorInterval
	}
}
