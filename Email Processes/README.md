# Email Processes

A collection of processes, utilities, scripts, etc. that are used to process (do things with) email in a mailbox.

| Name | Description | Tech Used |
| --- | --- | --- |
| [MailboxFolder-to-SP-List-Simple.ps1](MailboxFolder-to-SP-List-Simple.ps1) | For each item in a mailbox folder add an item to a SharePoint list. Login to the Graph using delegated interactive OAuth and Implicit OAuth to additionally get a SharePoint access token. | Graph API<br/>MSAL.PS<br/>REST<br/>PowerShell<br/> Interactive OAuth<br/> Implicit OAuth<br/>PnP PowerShell <img width=500/>|
| [MailboxFolder-to-SP-List-Advanced.ps1](MailboxFolder-to-SP-List-Advanced.ps1) | For each item in a mailbox folder extract attachments to a SharePoint library according to to a plain text JSON snippet in the email body. This sample will process items in multiple threads to increase speed and log information to a CSV file. | Graph API<br/>MSAL.PS<br/>REST<br/>PowerShell<br/> Interactive OAuth<br/> Implicit OAuth<br/> Windows Integrated OAuth<br/> PnP PowerShell<br/> Multi-Threading<br/> ScriptBlock<br/> Logging |