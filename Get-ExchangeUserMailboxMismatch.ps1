<# 
.SYNOPSIS 
	Retrives mailbox logon statistics and returns entries where the username of the mailbox doesn't match the user accounts username that is connected to it.
.DESCRIPTION 
    	Retrives mailbox logon statistics and returns entries where the username of the mailbox doesn't match the user accounts username that is connected to it. This may indicate that another user is viewing email in another individual's mailbox
.NOTES 
    	File Name  : Get-ExchangeUserMailboxMismatch.ps1
    	Author     : Brenton keegan - brenton.keegan@gmail.com 
    	Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/Exchange-PowerShell-Tools
    	License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	Get-ExchangeUserMailboxMismatch  -c "casServerName" -m "mailboxServerName"
.EXAMPLE 
#> 


#imports
import-module activedirectory

Function Get-ExchangeUserMailboxMismatch 
{
	
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[alias("c")]
		[string]$casServer,
		
		[parameter(Mandatory=$true)]
		[alias("m")]
		[string]$mbServer
		
	)
	
	$session = New-PSSession -configurationname Microsoft.Exchange -connectionURI http://$casServer/PowerShell
	Import-PSSession $session 
	
	
	$connectedExUsers = Get-LogonStatistics -server $mbServer| where {$_.ClientMode -ne "Cached" } | where {$_.ClientName -eq $casServer}

	Foreach($connectedExUser in $connectedExUsers)
	{

		$adUser = $connectedExUser.Windows2000Account -replace "+.\\",""
		$adUser = get-aduser $adUser -property displayname
		if ($connectedExUser.Username -ne $adUser.displayname)
		{
			$connectedExUser
		}		
	}
}
