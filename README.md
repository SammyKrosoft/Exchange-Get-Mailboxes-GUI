# Exchange-Get-Mailboxes-GUI
GUI to search mailboxes in an Exchange 2010, 2013, 2016 or Exchange Online (O365) environments

## Important note

Requires to be run from a PowerShell console with Exchange tools loaded, which can be an Exchange Management Shell window or a Powershell window from where you imported an Exchange session, see my TechNet blog post for a summary on how to do this:

* <a href="https://blogs.technet.microsoft.com/samdrey/2018/04/06/how-to-load-remote-powershell-session-on-exchange-2010-2013-2016-exchange-online-o365-2/" target="_blank"> How-to – Load Remote Exchange PowerShell Session on Exchange 2010, 2013, 2016, Exchange Online (O365) – which ports do you need</a>

* [Connect to Office 365 PowerShell](https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell)

* [How To–Load Exchange Management Shell into PowerShell ISE](https://blogs.technet.microsoft.com/samdrey/2017/12/17/how-to-load-exchange-management-shell-into-powershell-ise-2/)


## Screenshots - because a picture is worth 1000 words...

### First window when launching the tool
![screenshot1](DocResources/image0.jpg)

### After a sample Get-Mailbox which name includes "user" string
![screenshot2](DocResources/image1.jpg)

### If you select "Unlimited" for the Resultsize (max number of mailboxes to search) that is greater than 1000, you get a warning asking you if you want to continue
![screenshot3](DocResources/image-Question-LotsOfItems.jpg)

### Selecting mailboxes in the grid, notice the "Action on selected" button that becomes active
![screenshot4](DocResources/image-SelectForAction.jpg)

### After selecting some mailboxes in the grid, calling the "List Mailbox Features" action in the drop-down list
![screenshot5](DocResources/image-Action-ListMbxFeatures.jpg)

### Anoter action possible, calling the Single Item Recovery and mailbox dumpster limits for the selected mailboxes
![screenshot6](DocResources/image-Action-SingleItemRecoveryStatus.jpg)

### On most actions, you can copy the list in Windows clipboard (will be CSV Formatted) for further analyis, reporting or documentation about your mailboxes
![screenshot7](DocResources/image-copyToClipBoard.jpg)

### More to come...

