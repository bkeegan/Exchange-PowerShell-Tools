# Exchange-PowerShell-Tools

Contains a colleciton of scripts and cmdlets to assist in Exchange server administration and monitoring. All scripts currently tested on Exchange 2010.

1. Export-ExchangeContacts.ps1 : Cmdlet to export contacts stored on exchange for a particular user. Supports exporting to .pst or .csv (using the Outlook com object)

2. SpammerDectector.ps1 : monitoring script designed to be run on a frequent interval. Sends an email alert when an exchange user is sending a specified number of emails in a specified time interval. Specify an interval that might be unrealistic except in a mail abuse situation. 
