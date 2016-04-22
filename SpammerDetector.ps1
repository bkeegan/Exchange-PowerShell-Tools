<# 
.SYNOPSIS 
	Attempts to detetct a spammer based on frequency of emails.
.DESCRIPTION 
	Uses the cmdlet get-messagetrackinglog to determine how many emails each user has sent within a specified interval. 
	If the total number of recipients of emails sent within the specified interval an email notification will be send.
.NOTES 
    File Name  : SpammerDetector.ps1
    Author     : Brenton keegan - brenton.keegan@gmail.com 
    Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/Exchange-PowerShell-Tools
    License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	SpammerDetector -c "mail.contoso.com"  -n 100 -i 10 -r "alerts@contoso.com" -smtp "mail.contoso.com" -f "alerts@contoso.com"
#> 

Function SpammerDetector
{
	
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[alias("c")]
		[string]$casServer,
		
		[parameter(Mandatory=$true)]
		[alias("n")]
		[string]$numberofEmails,
		
		[parameter(Mandatory=$true)]
		[alias("i")]
		[string]$interval,
		
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
	
	$existingSession = get-pssession
	if(!($existingSession))
	{
		$session = New-PSSession -configurationname Microsoft.Exchange -connectionURI http://$casServer/PowerShell
		Import-PSSession $session
	}
	
	
	$intervalTimeSpan = New-Timespan -minutes $interval
	$currentTime = Get-Date
	$startTime = $currentTime.Subtract($intervalTimeSpan)
	$allMailboxes = Get-Mailbox -resultsize Unlimited
	foreach($mailbox in $allMailboxes)
	{
		$totalCount = 0
		$messages = Get-MessageTrackingLog -sender $mailbox.PrimarySmtpAddress -start $startTime -end $currentTime | where {($_.EventID -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER") -or ($_.EventID -eq "TRANSFER" -and $_.Source -eq "ROUTING" -and $_.SourceContext -eq "Resolver") }
		Foreach($message in $messages)
		{
			$totalCount += $message.RecipientCount
		}
		if($totalCount -ge $numberofEmails)
		{
			$emailString = $messages | select Sender, Recipients, MessageSubject,Timestamp | FL | Out-String
			$emailSubject = "Potential Internal Spammer - $($mailbox.PrimarySmtpAddress)"
			Send-MailMessage -To $alertRecipient -Subject $emailSubject -smtpServer $smtpServer -From $smtpSender -body $emailString
		}
	}
}
