# Manage-Delegates
PowerShell script which can add,remove or list delegates from a exchange mailbox.

## Available Parameters
- Identity
Type: String
Format: mail address (primary smtp address)
The mailbox you whish to alter

- Impersonate
Type: switch
Use if you need to impersonate (alter other mailboxes than yours)

- Mode
Type: string
Allowed valu´e: "List","Remove","Set"
Default: "List"
Define the mode of the script

- DelegateToRemove
Type: string
Format: mail address (primary smtp address)
Define the delegate you want to remove

- DelegateToSet
Type: string
Format: mail address (primary smtp address)
Define the delegate you want to set or add

- CalendarPermissions
Type: string
Allowed value: "None","Reviewer","Author","Editor"
Default: "None"
Set the permission for the calendar

- ContactsPermissions
Type: string
Allowed value: "None","Reviewer","Author","Editor"
Default: "None"
Set the permission for the contacts folder

- InboxPermissions
Type: string
Allowed value: "None","Reviewer","Author","Editor"
Default: "None"
Set the permission for the inbox

- JournalPermissions
Type: string
Allowed value: "None","Reviewer","Author","Editor"
Default: "None"
Set the permission for the journal

-  NotesPermissions
Type: string
Allowed value: "None","Reviewer","Author","Editor"
Default: "None"
Set the permission for the notes folder

-  TasksPermissions
Type: string
Allowed value: "None","Reviewer","Author","Editor"
Default: "None"
Set the permission for the notes folder

-  ReceiveCopiesOfMeetingMessages
Type: boolean
Allowed value: $true,$false
Default: $false
Activate the function, that the delegate receives all meeting messages

-  CanViewPrivateItems
Type: boolean
Allowed value: $true,$false
Default: $false
Is the delegate allowed to view private items?

-  WriteOnConsole
Type: switch
If set the (some) logging output will also be written to the console

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

ImpersonisationRights: if you wish to change delegates on other mailboxes
EWS Managed API 2.2
GlobalModules from Thomas Stensitzki (for logging and console output, min Version 2.1) => https://github.com/Apoc70/GlobalFunctions
Working autodiscover, trusted certificate or certificate from an enterprise ca


Modules:
GlobalFunctions module from Thomas Stensitzki (min. Version 2.1): https://github.com/Apoc70/GlobalFunctions
