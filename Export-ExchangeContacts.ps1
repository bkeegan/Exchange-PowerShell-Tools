<# 
.SYNOPSIS 
	Exports Exchange contacts
.DESCRIPTION 
	Exports exchange contacts in PST or CSV format. Script must be run as an administrator. Export to CSV requires the presence of Outlook as the Outlook API is required to read the PST exported from Exchange.
.NOTES 
    File Name  : Export-ExchangeContacts.ps1
    Author     : Brenton keegan - brenton.keegan@gmail.com 
    Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/Exchange-PowerShell-Tools
    License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	Export-ExchangeContacts -mb "User1" -cas "mail.contoso.com" -re -csv

#> 
Function Export-ExchangeContacts
{
	[cmdletbinding()]

	Param
	(

		[parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[alias("mb")]
		[string]$mailbox,
		
		[parameter(Mandatory=$true)]
		[alias("cas")]
		[string]$casServer,
		
		[parameter(Mandatory=$false)]
		[alias("re")]
		[switch]$removeExportRequest,
		
		[parameter(Mandatory=$false)]
		[alias("csv")]
		[switch]$exportToCSV
		
	)
	
	
	$existingSession = get-pssession
	if(!($existingSession))
	{
		$session = New-PSSession -configurationname Microsoft.Exchange -connectionURI http://$casServer/PowerShell
		Import-PSSession $session
	}
	
	
	#variable init
	[string]$dateStamp = Get-Date -UFormat "%Y%m%d_%H%M%S"
	$tempFolder = get-item env:temp
	$computerName = get-item env:computername
	$computerName = $computerName.Value
	$shareName = "$dateStamp-exContacts"
	$sharePath = "$($tempFolder.Value)\$dateStamp-exContacts"
	$UNCPath = "\\$computerName\$shareName"
	

	#create share to deposit export
	New-Item $sharePath -type directory | Out-Null
	New-SmbShare -Name $shareName -path $sharePath -FullAccess "Everyone" | Out-Null
	$acl = Get-Acl $sharePath
	$permission = "Everyone","FullControl","ContainerInherit, ObjectInherit", "None", "Allow"
	$rule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
	$acl.SetAccessRule($rule)
	$acl | Set-Acl $sharePath

	#begin export request
	$exportRequest =  New-MailboxExportRequest -Mailbox $mailbox -ExcludeDumpster -IncludeFolders "#Contacts#" -filepath $UNCPath\Contacts.pst
	
	#check when export request is completed and verify status. Throw error if result lists "failed"
	#delete export request if specified (only on successful export)
	$requestNotCompleted = $true
	while ($requestNotCompleted)
	{
		$requestResults = Get-mailboxexportrequest $exportRequest.RequestGuid
		if($requestResults.Status -eq "Completed")
		{
			$requestNotCompleted = $false
			if($removeExportRequest)
			{
				Remove-MailboxExportRequest -Identity $exportRequest.RequestGuid -Confirm:$false
			}
			Write-host "Contacts successfully exported. Contacts are located: $sharePath"
		}
		if($requestResults.Status -eq "Failed")
		{
			$requestNotCompleted = $false
			Throw "Export request not successful. Get-MailboxExportRequest returned status of 'Failed'. Export Request NOT removed."
		}
	}
	#delete share
	Get-Smbshare | where {$_.Path -eq $sharePath} | remove-smbshare -confirm:$false
	
	if($exportToCSV)
	{

		Add-type -assembly Microsoft.Office.Interop.Outlook 
		$olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]  
        $outlookAPI = new-object -comobject outlook.application
		$mapiNamespace = $outlookAPI.GetNameSpace('MAPI')
		$mapiNamespace.AddStore("$sharePath\Contacts.pst")
		$pstFile = $mapiNamespace.Stores | where {$_.FilePath -eq "$sharePath\Contacts.pst"}
		$pstRoot = $pstFile.GetRootFolder()  
		$pstFolders = $pstRoot.folders
		foreach($pstFolder in $pstFolders)
		{
			if($pstFolder.DefaultMessageClass -eq "IPM.Contact")
			{
				$pstFolder.Items | Export-CSV "$sharePath\Contacts.csv" -NoTypeInformation
				$mapiNamespace.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$mapiNamespace,($pstFolder))

			}
		}
		$outlookAPI.Quit()
	}
}
