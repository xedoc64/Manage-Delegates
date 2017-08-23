<#
    .Synopsis
    PowerShell script which can manage delegates for an exchange mailbox

    .Description

    With this script you can add, remove or list deletgates for an exchange mailbox

    Author: Torsten Schlopsnies
    
    Version 1.0 2017-08-18
    
    .NOTES 
    Requirements 
    - EWS 2.2 installed
    - Impersonisation rights
    - Exchange 2013 (tested with CU17) or higher (works for Exchange 2010 maybe also)
    - GlobalModules from Thomas Stensitzki (for logging and console output, min Version 2.1) => https://github.com/Apoc70/GlobalFunctions
	- Working autodiscover, trusted certificate or certificate from an enterprise ca

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial release


    Based on a article from Glen: http://gsexdev.blogspot.de/2012/03/ews-managed-api-and-powershell-how-to.html

    .PARAMETER Identity
    Type: String
    Format: mail address (primary smtp address)
    The mailbox you whish to alter

    .PARAMETER Impersonate
    Type: switch
    Use if you need to impersonate (alter other mailboxes than yours)

    .PARAMETER Mode
    Type: string
    Allowed valu´e: "List","Remove","Set"
    Default: "List"
    Define the mode of the script

    .PARAMETER DelegateToRemove
    Type: string
    Format: mail address (primary smtp address)
    Define the delegate you want to remove

    .PARAMETER DelegateToSet
    Type: string
    Format: mail address (primary smtp address)
    Define the delegate you want to set or add

    .PARAMETER CalendarPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the calendar

    .PARAMETER ContactsPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the contacts folder

    .PARAMETER InboxPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the inbox

    .PARAMETER JournalPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the journal

    .PARAMETER NotesPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the notes folder

    .PARAMETER TasksPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the notes folder

    .PARAMETER ReceiveCopiesOfMeetingMessages
    Type: boolean
    Allowed value: $true,$false
    Default: $false
    Activate the function, that the delegate receives all meeting messages

    .PARAMETER CanViewPrivateItems
    Type: boolean
    Allowed value: $true,$false
    Default: $false
    Is the delegate allowed to view private items?

    .PARAMETER WriteOnConsole
    Type: switch
    If set the (some) logging output will also be written to the console

    .EXAMPLE
    List all delegates with the permissions with impersonisation
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -Mode "List"
    or
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate

    .EXAMPLE
    Add a delegate to a mailbox with impersonisation only with reviewer rights to the inbox
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer"

    .EXAMPLE
    Remove a delegate
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -DelegateToRemove "davids@contoso.com"
    Note, currently the script is not asking for a confirmation. It simply removes the delegate and log the permissions to the log file. 
#>

 [CmdletBinding(DefaultParameterSetName="List")]
 Param(
    [string]$Identity,
    [switch]$Impersonate,
    [ValidateSet("List","Remove","Set")][string]$Mode = "List",
    [Parameter(ParameterSetName="Remove")][string]$DelegateToRemove,
    [Parameter(ParameterSetName="Set")][string]$DelegateToSet,
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$CalendarPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$ContactsPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$InboxPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$JournalPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$NotesPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$TasksPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet($true,$false)][bool]$ReceiveCopiesOfMeetingMessages = $false,
    [Parameter(ParameterSetName="Set")][ValidateSet($true,$false)][bool]$CanViewPrivateItems = $false,
    [switch]$WriteOnConsole = $true
 )

 # Import global modules
Import-Module GlobalFunctions
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

# Create a log folder
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

# loading the ews.dll
try 
{
  if($env:ExchangeInstallPath -ne '') {
    # Use local Exchange install path, if available
    $dllpath = "$($env:ExchangeInstallPath)\bin\Microsoft.Exchange.WebServices.dll"
  }
  else {
    # Use EWS managed API install path
    $dllpath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
  }
  
  [void][Reflection.Assembly]::LoadFile($dllpath)
}
catch {
  $logger.Write('Error on loading the EWS dll. Please check the path or install the EWS Managed API!',1,$WriteOnConsole)
  $logger.Write('Script aborted',1,$WriteOnConsole)
  exit(1)
}

# function to connect to a mailbox
function Create-Service {
  Param(
    [string]$Identity,
    [switch]$Impersonate
  )
  try {
    #Create a service reference
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $Service.AutodiscoverUrl($Identity)
    $enumSmtpAddress = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress

    if ($Impersonate) {
        # Set Impersonation
        $Service.ImpersonatedUserId =  New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId($enumSmtpAddress,$Identity) 
        $logger.Write("Web service for user $($Identity) created.")
    }
  }
  catch [Exception] {
    $logger.Write("Failed creating the exchange web service for $($Identity). Check mailbox name and impersonisation rights.",1,$WriteOnConsole)
    $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
    return
  }
  return $Service
}

function Get-Delegates {
    Param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [string]$Identity
    )
    $delegates = $null
    try {
        $delegates = $service.GetDelegates($Identity,$true)
    } catch [Exception] {
        $logger.Write("Failed to get delegates for user $($Identity)",1,$WriteOnConsole)
        $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
        return
    }
    return $delegates
}

function Remove-Delegate {
    Param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [string]$Identity,
        [string]$Delegate
    )
    try {
        $null = $service.RemoveDelegates($Identity,$Delegate)
        $logger.Write("$($Delegate) removed from $($Identity) successfully.",0,$true)
    } catch [Exception] {
        $logger.Write("Failed to remove the delegate $($Delegate) from user $($Identity)",1,$WriteOnConsole)
        $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
    }
}

function Set-Delegate
{
  Param(
    [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
    [string]$Identity,
    [string]$Delegate,
    [Microsoft.Exchange.WebServices.Data.DelegateInformation]$Delegates,
    [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]$CalendarFolderPermissionLevel,
    [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]$ContactsFolderPermissionLevel,
    [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]$InboxFolderPermissionLevel,
    [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]$JournalFolderPermissionLevel,
    [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]$NotesFolderPermissionLevel,
    [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]$TasksFolderPermissionLevel,
    [bool]$ReceiveCopiesOfMeetingMessages,
    [bool]$ViewPrivateItems,
    [switch]$Add
  )
  
  $logger.Write("Set the delegate $($Delegate) to mailbox with follwing permissions:",0,$WriteOnConsole)
  $logger.Write("Calendar: $($CalendarFolderPermissionLevel)",0,$WriteOnConsole)
  $logger.Write("Contacts: $($ContactsFolderPermissionLevel)",0,$WriteOnConsole)
  $logger.Write("Inbox: $($InboxFolderPermissionLevel)",0,$WriteOnConsole)
  $logger.Write("Journal: $($JournalFolderPermissionLevel)",0,$WriteOnConsole)
  $logger.Write("Notes: $($NotesFolderPermissionLevel)",0,$WriteOnConsole)
  $logger.Write("Tasks: $($TasksFolderPermissionLevel)",0,$WriteOnConsole)
  $logger.Write("Receive copies of meeting messages: $($ReceiveCopiesOfMeetingMessages)",0,$WriteOnConsole)
  $logger.Write("View private items: $($ViewPrivateItems)",0,$WriteOnConsole)
  
  if ($Add) {
    try {
      $logger.Write("Mode: Adding to mailbox",0,$WriteOnConsole)
      $dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($Delegate)
      $dgUser.ViewPrivateItems = $ViewPrivateItems
      $dgUser.ReceiveCopiesOfMeetingMessages = $ReceiveCopiesOfMeetingMessages
      $dgUser.Permissions.CalendarFolderPermissionLevel = $CalendarFolderPermissionLevel
      $dgUser.Permissions.ContactsFolderPermissionLevel = $ContactsFolderPermissionLevel
      $dgUser.Permissions.InboxFolderPermissionLevel = $InboxFolderPermissionLevel
      $dgUser.Permissions.JournalFolderPermissionLevel = $JournalFolderPermissionLevel
      $dgUser.Permissions.NotesFolderPermissionLevel = $NotesFolderPermissionLevel
      $dgUser.Permissions.TasksFolderPermissionLevel = $TasksFolderPermissionLevel
      $dgArray = new-object Microsoft.Exchange.WebServices.Data.DelegateUser[] 1
      $dgArray[0] = $dgUser
      $null = $Service.AddDelegates($Identity,$delegates.MeetingRequestsDeliveryScope,$dgArray)
      $logger.Write("$($Delegate) successfully added",0,$WriteOnConsole) 
    } catch [Exception] {
      $logger.Write("Failed to add the delegate $($Delegate) to mailbox $($Identity)",1,$WriteOnConsole)
      $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
    }
  }
  else {
    try {
      $logger.Write("Mode: Altering the permissions",0,$WriteOnConsole)
      $dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($Delegate)
      $dgUser.ViewPrivateItems = $ViewPrivateItems
      $dgUser.ReceiveCopiesOfMeetingMessages = $ReceiveCopiesOfMeetingMessages
      $dgUser.Permissions.CalendarFolderPermissionLevel = $CalendarFolderPermissionLevel
      $dgUser.Permissions.ContactsFolderPermissionLevel = $ContactsFolderPermissionLevel
      $dgUser.Permissions.InboxFolderPermissionLevel = $InboxFolderPermissionLevel
      $dgUser.Permissions.JournalFolderPermissionLevel = $JournalFolderPermissionLevel
      $dgUser.Permissions.NotesFolderPermissionLevel = $NotesFolderPermissionLevel
      $dgUser.Permissions.TasksFolderPermissionLevel = $TasksFolderPermissionLevel
      $dgArray = new-object Microsoft.Exchange.WebServices.Data.DelegateUser[] 1
      $dgArray[0] = $dgUser
      $null = $Service.UpdateDelegates($Identity,$delegates.MeetingRequestsDeliveryScope,$dgArray)
      $logger.Write("$($Delegate) successfully altered",0,$WriteOnConsole)   
    } catch [Exception] {
      $logger.Write("Failed to set the permissions for the delegate $($Delegate) to mailbox $($Identity)",1,$WriteOnConsole)
      $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
    }
    
  }
}

function Match-ToPermission
{
  Param(
    [string]$Permission
  )
  
  switch($Permission) {
    "None" {
      return [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None
    }
    "Editor" {
      return [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Editor
    }
    "Author" {
      [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Author
    }
    "Reviewer" {
      [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Reviewer
    }
    default: {
      return [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None
    }
  }
}



# create the service
$Service = Create-Service -Identity $Identity -Impersonate $Impersonate

if ($service -ne $null) {
    switch ($Mode) {
        "List" {
            $DelegateList = Get-Delegates -Service $service -Identity $Identity
            if (($DelegateList -ne $null) -and ($DelegateList.DelegateUserResponses.Count -ge 1)) {
                $logger.Write("Total delegates: $($DelegateList.DelegateUserResponses.Count)",0,$WriteOnConsole)
                foreach ($User in $DelegateList.DelegateUserResponses) {
                    $logger.Write("Delegate: $($User.DelegateUser.UserId.PrimarySmtpAddress)",0,$WriteOnConsole)
                    $logger.Write("Permissions...",0,$true)
                    $logger.Write("Calendar: $($User.DelegateUser.Permissions.CalendarFolderPermissionLevel)",0,$WriteOnConsole)
                    $logger.Write("Contacts: $($User.DelegateUser.Permissions.ContactsFolderPermissionLevel)",0,$WriteOnConsole)
                    $logger.Write("Inbox: $($User.DelegateUser.Permissions.InboxFolderPermissionLevel)",0,$WriteOnConsole)
                    $logger.Write("Journal: $($User.DelegateUser.Permissions.JournalFolderPermissionLevel)",0,$WriteOnConsole)
                    $logger.Write("Notes: $($User.DelegateUser.Permissions.NotesFolderPermissionLevel)",0,$WriteOnConsole)
                    $logger.Write("Tasks: $($User.DelegateUser.Permissions.TasksFolderPermissionLevel)",0,$WriteOnConsole)
                    $logger.Write("Receive copies of meeting messages: $($User.DelegateUser.ReceiveCopiesOfMeetingMessages)",0,$WriteOnConsole)
                    $logger.Write("View private items: $($User.DelegateUser.ViewPrivateItems)",0,$WriteOnConsole)
                }
            } else {
                $logger.Write("No delegates are set for $($Identity)",0,$WriteOnConsole)
            }
        }
        "Remove" {
            [bool]$exist = $false
            # Checking if is already a delegate
            $DelegateList = Get-Delegates -service $service -Identity $Identity
            if (($DelegateList -ne $null) -and ($DelegateList.DelegateUserResponses.Count -ge 1)) {
                foreach($User in $DelegateList.DelegateUserResponses) {
                    # is the smtp address is matching?
                    if($User.DelegateUser.UserId.PrimarySmtpAddress.ToLower() -eq $DelegateToRemove.ToLower()){
                        $logger.Write("Delegate $($DelegateToRemove) found. Try to removing now",0,$WriteOnConsole)
                        $logger.Write("Permissions...",0)
                        $logger.Write("Calendar: $($User.DelegateUser.Permissions.CalendarFolderPermissionLevel)")
                        $logger.Write("Contacts: $($User.DelegateUser.Permissions.ContactsFolderPermissionLevel)")
                        $logger.Write("Inbox: $($User.DelegateUser.Permissions.InboxFolderPermissionLevel)")
                        $logger.Write("Journal: $($User.DelegateUser.Permissions.JournalFolderPermissionLevel)")
                        $logger.Write("Notes: $($User.DelegateUser.Permissions.NotesFolderPermissionLevel)")
                        $logger.Write("Tasks: $($User.DelegateUser.Permissions.TasksFolderPermissionLevel)")
                        $logger.Write("Receive copies of meeting messages: $($User.DelegateUser.ReceiveCopiesOfMeetingMessages)",0)
                        $logger.Write("View private items: $($User.DelegateUser.ViewPrivateItems)",0)
                        $exist = $true
                        Remove-Delegate -Service $service -Identity $Identity -Delegate $DelegateToRemove
                    }  
                }
            }
            if (-not ($exist)) {
                $logger.Write("$($DelegateToRemove) is not a delegate of $($Identity)",0,$WriteOnConsole)
            }
        }
        "Set" {
            # Checking if is already a delegate
            $delegatelist = Get-Delegates -service $service -Identity $Identity
            if ($delegatelist -ne $null) {
                # Are there any delegates set?
                if ($delegatelist.DelegateUserResponses.Count -eq 0) {
                    # The user isn't delegate, so we need to add the user
                    Set-Delegate -Service $service -Identity $Identity -Delegate $DelegateToSet `
                        -CalendarFolderPermissionLevel (Match-ToPermission -Permission $CalendarPermissions) -ContactsFolderPermissionLevel (Match-ToPermission -Permission $ContactsPermissions) `
                        -InboxFolderPermissionLevel (Match-ToPermission -Permission $InboxPermissions) -JournalFolderPermissionLevel (Match-ToPermission -Permission $JournalPermissions) `
                        -NotesFolderPermissionLevel (Match-ToPermission -Permission $NotesPermissions) -TasksFolderPermissionLevel (Match-ToPermission -Permission $TasksPermissions) `
                        -ReceiveCopiesOfMeetingMessages $ReceiveCopiesOfMeetingMessages -ViewPrivateItems $CanViewPrivateItems -Delegates $delegatelist -Add
                        break
                }
                else {
                    foreach($User in $delegatelist.DelegateUserResponses) {
                         if($User.DelegateUser.UserId.PrimarySmtpAddress.ToLower() -eq $DelegateToSet.ToLower()){
                            # The user is already delegate, only set the permissions
                            Set-Delegate -Service $service -Identity $Identity -Delegate $DelegateToSet `
                              -CalendarFolderPermissionLevel (Match-ToPermission -Permission $CalendarPermissions) -ContactsFolderPermissionLevel (Match-ToPermission -Permission $ContactsPermissions) `
                              -InboxFolderPermissionLevel (Match-ToPermission -Permission $InboxPermissions) -JournalFolderPermissionLevel (Match-ToPermission -Permission $JournalPermissions) `
                              -NotesFolderPermissionLevel (Match-ToPermission -Permission $NotesPermissions) -TasksFolderPermissionLevel (Match-ToPermission -Permission $TasksPermissions) `
                              -ReceiveCopiesOfMeetingMessages $ReceiveCopiesOfMeetingMessages -ViewPrivateItems $CanViewPrivateItems -Delegates $delegatelist
                         }
                         else {
                            # The user isn't delegate, so we need to add the user
                            Set-Delegate -Service $service -Identity $Identity -Delegate $DelegateToSet `
                              -CalendarFolderPermissionLevel (Match-ToPermission -Permission $CalendarPermissions) -ContactsFolderPermissionLevel (Match-ToPermission -Permission $ContactsPermissions) `
                              -InboxFolderPermissionLevel (Match-ToPermission -Permission $InboxPermissions) -JournalFolderPermissionLevel (Match-ToPermission -Permission $JournalPermissions) `
                              -NotesFolderPermissionLevel (Match-ToPermission -Permission $NotesPermissions) -TasksFolderPermissionLevel (Match-ToPermission -Permission $TasksPermissions) `
                              -ReceiveCopiesOfMeetingMessages $ReceiveCopiesOfMeetingMessages -ViewPrivateItems $CanViewPrivateItems -Delegates $delegatelist -Add
                         }
                    }
                }
            }
        }
    }
}