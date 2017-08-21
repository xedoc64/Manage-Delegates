# Manage-Delegates
PowerShell script which can add,remove or list delegates from a exchange mailbox.

##Examples

List all delegates with the permissions with impersonisation
.\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -Mode "List"
or
.\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate

Add a delegate to a mailbox with impersonisation only with reviewer rights to the inbox
.\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer"

Remove a delegate
.\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -DelegateToRemove "davids@contoso.com"
Note, currently the script is not asking for a confirmation. It simply removes the delegate and log the permissions to the log file. 

## Requirements

- ImpersonisationRights: if you wish to change delegates on other mailboxes
- EWS Managed API 2.2
- GlobalModules from Thomas Stensitzki (for logging and console output, min Version 2.1) => https://github.com/Apoc70/GlobalFunctions
- Working autodiscover, certificate from a trusted ca or certificate from an enterprise ca


Modules:
GlobalFunctions module from Thomas Stensitzki (min. Version 2.1): https://github.com/Apoc70/GlobalFunctions
