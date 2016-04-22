# Exchange-PowerShell-Tools

Contains a colleciton of scripts and cmdlets to assist in Exchange server administration and monitoring. All scripts currently tested on Exchange 2010.

1. **Export-ExchangeContacts.ps1** : Cmdlet to export contacts stored on exchange for a particular user. Supports exporting to .pst or .csv (using the Outlook com object)

2. **SpammerDectector.ps1** : monitoring script designed to be run on a frequent interval. Sends an email alert when an exchange user is sending a specified number of emails in a specified time interval. Specify an interval that might be unrealistic except in a mail abuse situation. 

3. ReportRogueActiveSyncDevices.ps1 : Checks accounts disabled in Active Directory that have ActiveSync devices making active connections later than the accounts last modified date. 

4. MailboxAccessAlerter.ps1 : Sends an alert when a mailbox is accessed by a different user than the one the mailbox is associated with. 

5. Get-MailboxItemPerDay.ps1 : Calculates total amount of email items + deleted items and divides by age of inbox in days to calcuate average amount of emails (or mailbox items) received per day. 

6. Get-ExchangeUserMailboxMismatch.ps1 : gets logon statistics where the user account accessing a mailbox does not match the user associated with it. 
