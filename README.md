# Manage-Delegates
PowerShell script which can add,remove or list delegates from a exchange mailbox.

## Requirements

- ImpersonisationRights: if you wish to change delegates on other mailboxes
- EWS Managed API 2.2
- GlobalModules from Thomas Stensitzki (for logging and console output, min Version 2.1) => https://github.com/Apoc70/GlobalFunctions


Modules:
GlobalFunctions module from Thomas Stensitzki (min. Version 2.1): https://github.com/Apoc70/GlobalFunctions


Parameters:
  * Identity (string):
   Mailbox which you would like to alter.

  * Credentials (PSCredentials):
   Credentials for the service. You can also use -UseDefaultCredentials for passing the credentials of the session

  * UseDefaultCredentials (bool. Default: $false):
    Use this parameter if you want to connect with the session credentials

  * Url (string):
   Url to connect to. If no set the script use autodiscover

  * IgnoreSSLCertificate (bool):
   If set the script ignores SSL certificate errors

  * Impersonate (bool. Default: $false):
   Pass this if you would like alter other mailboxes than yours

  * DelegateToRemove (string array):
   Delegate(s) which should be removed.

  * DelegateToSet (string array):
   Delegate(s) which should be added or altered.

  * CalendarPermissions, ContactsPermissions, InboxPermissions, JournalPermissions, NotesPermissions, TasksPermissions (string. Default: "None"):  
   Define the permissions. There is a validate set for each parameter, you can use TAB. If a parameter is not set the default value "None" is used for this permission.  
   So if you would like to alter a delegate which have already permissions you have to pass the old permissions also (or the will be set to "None").

  * ReceiveCopiesOfMeetingMessages (bool. Default: $false):
   Should the delegate(s) receive copies of the meeting messages?

  * CanViewPrivateItems (bool. Default: $false):
   Should the delegate(s) see items that marked as private?

  * WriteOnConsole (bool. Default: $false):
   Log some output which will be written in the log file to the command line.

  * NoConfirm (bool. Default: $false):
   If set to $true the script will not be asking for confirmation at removing delegates.

  * IgnoreSSLCertificate (bool. Default: $false)
   Ignore SSL certificate errors when connecting to an EWS endpoint


Examples:  

1. List all delegates with the permissions with impersonisation  
  .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -UseDefaultCredentials $true

2. Add a delegate to a mailbox with impersonation only with reviewer rights to the inbox using the session credentials  
  .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer" -UseDefaultCredentials $true

3. Add a delegate to a mailbox with impersonation only with reviewer rights to the inbox using a specific url,ignore ssl certificate errors and session credentials  
  .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer" -UseDefaultCredentials $true -Url "https://Exchangeserver1/ews/exchange.asmx" -IgnoreSSLCertificate $true

4. Add a delegate to a mailbox with impersonation only with reviewer rights to the inbox not using the session credentials  
  .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer" -Credentials (Get-Credentials)
  
5. Remove a delegate  
  .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -DelegateToRemove "davids@contoso.com"