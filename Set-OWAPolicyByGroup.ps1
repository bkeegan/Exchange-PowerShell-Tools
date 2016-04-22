<# 
.SYNOPSIS 
	This will assign an OWA Mailbox Policy to all members of a specified AD group.
.DESCRIPTION 
	In EMC (Exchange Management Console), OWA Mailbox policies are applied on a per-user basis. This script will iterate thru an Active Directory group and apply a specified OWA Mailbox policy to each member of this group.
	This cmdlet requires the Active Directory PowerShell components to be installed and the ability to open a WinRM session to the Exchange CAS server.
.NOTES 
    File Name  : Set-OWAPolicyByGroup.ps1 
    Author     : Brenton Keegan - brenton.keegan@gmail.com 
    Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/Exchange-PowerShell-Tools
    License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	Set-OWAPolicyByGroup -g "TestGroup" -p "OWA Test Policy" -cas "excas01.domain.com"
	This command will connect to excas01.domain.com and apply the OWA Mailbox policy named "OWA Test Policy" to every user in the group "TestGroup"
#> 

function Set-OWAPolicyByGroup
{
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[alias("g")]
		$adGroup,
		
		[parameter(Mandatory=$true)]
		[alias("p")]
		$owaPolicy,

		[parameter(Mandatory=$true)]
		[alias("cas")]
		$casServer
	)
	
	#imports
	$session = New-PSSession -configurationname Microsoft.Exchange -connectionURI http://$casServer/PowerShell
	Import-PSSession $session 

	get-adgroupmember -identity $adGroup | foreach {$_.SamAccountName
		Set-CASMailbox -Identity $_.SamAccountName -OwaMailboxPolicy:$owaPolicy 
	} | Out-Null
}
