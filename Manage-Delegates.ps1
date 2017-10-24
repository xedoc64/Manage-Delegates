<#
    .Synopsis
    PowerShell script which can manage delegates for an exchange mailbox

    .Description

    With this script you can add, remove or list deletgates for an exchange mailbox

    Author: Torsten Schlopsnies
    
    Version 1.5 2017-09-08
    
    .NOTES 
    Requirements 
    - EWS 2.2 installed
    - Impersonitsation rights
    - Exchange 2013 (tested with CU17) or higher (runs maybe for Exchange 2010 also)
    - GlobalFunctions from Thomas Stensitzki (for logging and console output) => https://github.com/Apoc70/GlobalFunctions

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial release
    1.1     fixed behaviour when altering delegates
    1.5:	
            - fixed various bugs (regarding impersonation, logical). 
            - Removed "Mode" switch
            - added possibility to pass credentials and url
            - SSL certtification errors can be ignored now
            - You can remove/set multiple users to one mailbox
            - Instead of the switches from earlier versions you now have to pass $true or $false

    Based on a article from Glen: http://gsexdev.blogspot.de/2012/03/ews-managed-api-and-powershell-how-to.html

    .PARAMETER Identity
    Type: String
    Format: mail address (primary smtp address)
    The mailbox you whish to alter

    .PARAMETER Credentials
    Type: PSCredentials
    Credentials which will be used to create the service.

    .PARAMETER UseDefaultCredentials
    Type: boolean
    Default: $false
    On domain joined computer you can use the session credentials. With this switch these will be passed to the service.

    .PARAMETER Url
    Type: string
    Format: https://servername/ews/exchange.asmx
    To connect to a specific server set this parameter. If not set autodiscover will be used.

    .PARAMETER IgnoreSSLCertificate
    Type: boolean
    Default: $false
    If set invalid SSL certificates will be ignored.

    .PARAMETER Impersonate
    Type: boolean
    Default: $false
    Use if you need to impersonate (alter other mailboxes than yours).

    .PARAMETER DelegateToRemove
    Type: string
    Format: mail address (primary smtp address)
    Define the delegate you want to remove.

    .PARAMETER DelegateToSet
    Type: string
    Format: mail address (primary smtp address)
    Define the delegate you want to set or add.

    .PARAMETER CalendarPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the calendar.

    .PARAMETER ContactsPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the contacts folder.

    .PARAMETER InboxPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the inbox.

    .PARAMETER JournalPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the journal.

    .PARAMETER NotesPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the notes folder.

    .PARAMETER TasksPermissions
    Type: string
    Allowed value: "None","Reviewer","Author","Editor"
    Default: "None"
    Set the permission for the notes folder.

    .PARAMETER ReceiveCopiesOfMeetingMessages
    Type: boolean
    Allowed value: $true,$false
    Default: $false
    Activate the function, that the delegate receives all meeting messages.

    .PARAMETER CanViewPrivateItems
    Type: boolean
    Allowed value: $true,$false
    Default: $false
    Is the delegate allowed to view private items?.

    .PARAMETER WriteOnConsole
    Type: boolean
    Default: $false
    If set the (some) logging output will also be written to the console.

    .PARAMETER NoConfirm
    Type: boolean
    Default: $false
    Will be used for removing delegates. Should be set to $false if script is running as a task. Note: Before removing delegates run the script to list existing delegates with the permission

    .EXAMPLE
    List all delegates with the permissions with impersonisation
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -UseDefaultCredentials $true

    .EXAMPLE
    Add a delegate to a mailbox with impersonation only with reviewer rights to the inbox using the session credentials
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer" -UseDefaultCredentials $true

    .EXAMPLE
    Add a delegate to a mailbox with impersonation only with reviewer rights to the inbox using a specific url,ignore ssl certificate errors and session credentials
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer" -UseDefaultCredentials $true -Url "https://Exchangeserver1/ews/exchange.asmx" -IgnoreSSLCertificate $true

    Add a delegate to a mailbox with impersonation only with reviewer rights to the inbox not using the session credentials
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate $true -DelegateToSet "davids@contoso.com" -InboxPermissions "Reviewer" -Credentials (Get-Credentials)
    .EXAMPLE
    Remove a delegate
    .\Manage-Delegates.p1 -Identity "mollyc@contoso.com" -Impersonate -DelegateToRemove "davids@contoso.com"

#>

Param(
    [Parameter(Mandatory=$true)][string]$Identity,
    [System.Management.Automation.PSCredential]$Credentials,
    [string]$Url,
    [bool]$IgnoreSSLCertificate = $false,
    [bool]$UseDefaultCredentials = $false,
    [bool]$Impersonate = $false,
    [Parameter(ParameterSetName="Set")][string[]]$DelegateToRemove,
    [Parameter(ParameterSetName="Set")][string[]]$DelegateToSet,
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$CalendarPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$ContactsPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$InboxPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$JournalPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$NotesPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet("None","Reviewer","Author","Editor")][string]$TasksPermissions = "None",
    [Parameter(ParameterSetName="Set")][ValidateSet($true,$false)][bool]$ReceiveCopiesOfMeetingMessages = $false,
    [Parameter(ParameterSetName="Set")][ValidateSet($true,$false)][bool]$CanViewPrivateItems = $false,
    [bool]$WriteOnConsole = $false,
    [bool]$Confirm = $true
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
  if ($env:ExchangeInstallPath -ne $null -and $env:ExchangeInstallPath -ne '') {
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
function New-Service {
  Param(
    [string]$ID,
    [bool]$Impersonation = $false,
    [System.Management.Automation.PSCredential]$Credentials,
    [bool]$IgnoreSSL = $false,
    [string]$Url,
    [bool]$DefaultCredentials = $false
  )
  try {
    #Create a service reference
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    if ($IgnoreSSLCertificate) {
        Set-IgnoreSSLCerfiticates
    }
    # using default credentials or specified ones
    if ($UseDefaultCredentials) {
        $Service.UseDefaultCredentials = $true
    }
    else {
        try {
            $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().Password.ToString())
            $Service.Credentials = $creds
        }
        catch [Exception] {
            $logger.Write("Failed to set the credentials fot the webservice. Please check the entered credentials",1,$WriteOnConsole)
            $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
            return $null
          }        
    }
    # check if we need to connect to a specific server
    if ($Url -ne $null -and $Url -ne "") {
        $uri = [System.Uri]$URl
        $Service.Url = $uri
    }
    else {
        $Service.AutodiscoverUrl($ID, {$true})
    }
    $enumSmtpAddress = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress
    if ($Impersonate) {
        # Set Impersonation
        $Service.ImpersonatedUserId =  New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId($enumSmtpAddress,$ID) 
        $logger.Write("Web service for user $($ID) created.")
    }
    else {
        $logger.Write("Web service for user $($ID) created.")
    }
  }
  catch [Exception] {
    $logger.Write("Failed creating the exchange web service for $($ID). Check mailbox name and impersonisation rights.",1,$WriteOnConsole)
    $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
    return $null
  }
  return $Service
}

function Set-IgnoreSSLCerfiticates {
Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@

    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
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
        return $null
    }
    return $delegates
}

function Remove-Delegate {
    Param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [string]$Identity,
        [string[]]$DelegateList
    )
    try {
        $null = $service.RemoveDelegates($Identity,$DelegateList)
        for ($i = 0;$i -le $DelegateList.Count-1;$i++) {
            $logger.Write("$($DelegateList[$i]) removed from $($Identity) successfully.",0,$true)
        }
    } 
    catch [Exception] {
        for ($i = 0;$i -le $DelegateList.Count-1;$i++) {
            $logger.Write("Failed to remove the delegate $($DelegateList[$i]) from user $($Identity)",1,$WriteOnConsole)
            $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
        }
        
    }
}

function Set-Delegate
{
    Param(
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [string]$Identity,
        [string[]]$Delegates,
        [Microsoft.Exchange.WebServices.Data.DelegateInformation]$DelegateList,
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
  
    if ($Add) {
        $logger.Write("Mode: Adding to mailbox",0,$WriteOnConsole)
    }
    else {
        $logger.Write("Mode: Altering delegates",0,$WriteOnConsole)
    }

    for ($i = 0; $i -le $Delegates.Count-1;$i++) {
        $logger.Write("Set the delegate $($Delegates[$i]) to mailbox with follwing permissions:",0,$WriteOnConsole)
        $logger.Write("Calendar: $($CalendarFolderPermissionLevel)",0,$WriteOnConsole)
        $logger.Write("Contacts: $($ContactsFolderPermissionLevel)",0,$WriteOnConsole)
        $logger.Write("Inbox: $($InboxFolderPermissionLevel)",0,$WriteOnConsole)
        $logger.Write("Journal: $($JournalFolderPermissionLevel)",0,$WriteOnConsole)
        $logger.Write("Notes: $($NotesFolderPermissionLevel)",0,$WriteOnConsole)
        $logger.Write("Tasks: $($TasksFolderPermissionLevel)",0,$WriteOnConsole)
        $logger.Write("Receive copies of meeting messages: $($ReceiveCopiesOfMeetingMessages)",0,$WriteOnConsole)
        $logger.Write("View private items: $($ViewPrivateItems)",0,$WriteOnConsole)
    }
    
    [Microsoft.Exchange.WebServices.Data.DelegateUser[]]$dgArray = @()
    for ($i = 0;$i -le $Delegates.Count-1;$i++) {
        $dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($Delegates[$i])
        $dgUser.ViewPrivateItems = $ViewPrivateItems
        $dgUser.ReceiveCopiesOfMeetingMessages = $ReceiveCopiesOfMeetingMessages
        $dgUser.Permissions.CalendarFolderPermissionLevel = $CalendarFolderPermissionLevel
        $dgUser.Permissions.ContactsFolderPermissionLevel = $ContactsFolderPermissionLevel
        $dgUser.Permissions.InboxFolderPermissionLevel = $InboxFolderPermissionLevel
        $dgUser.Permissions.JournalFolderPermissionLevel = $JournalFolderPermissionLevel
        $dgUser.Permissions.NotesFolderPermissionLevel = $NotesFolderPermissionLevel
        $dgUser.Permissions.TasksFolderPermissionLevel = $TasksFolderPermissionLevel 
        $dgArray += $dgUser
    }

    if ($Add) {
        try {
            $null = $Service.AddDelegates($Identity,$DeleGateList.MeetingRequestsDeliveryScope,$dgArray)
            for ($i = 0;$i -le $DelegateList.Count-1;$i++) {
                $logger.Write("$($Delegates[$i]) added to $($Identity) successfully.",0,$true)
            }
        } 
        catch [Exception] {
            for ($i = 0;$i -le $DelegateList.Count-1;$i++) {
                $logger.Write("Failed to add $($Delegates[$i]) to mailbox $($Identity)",1,$WriteOnConsole)
                $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
            }     
        }
    }
    else {
        try {
        $null = $Service.UpdateDelegates($Identity,$DeleGateList.MeetingRequestsDeliveryScope,$dgArray)
        for ($i = 0;$i -le $DelegateList.Count-1;$i++) {
                $logger.Write("$($Delegates[$i]) for mailbox $($Identity) altered successfully.",0,$true)
            }
        } 
        catch [Exception] {
        for ($i = 0;$i -le $DelegateList.Count-1;$i++) {
                $logger.Write("Failed to alter $($Delegates[$i]) for mailbox $($Identity)",1,$WriteOnConsole)
                $logger.Write("Exception: $($_.Exception.GetType().FullName) - Message: $($_.Exception.Message)")
        }
        }
    }
}

function Get-Permission
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
        return [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Author
        }
        "Reviewer" {
        return [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Reviewer
        }
        default: {
        return [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None
        }
    }
}

function Request-Choice {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Caption)
    $choices =  [System.Management.Automation.Host.ChoiceDescription[]]@("&Yes","&No")
    [int]$defaultChoice = 1

    $choiceReturn = $Host.UI.PromptForChoice($Caption, "", $choices, $defaultChoice)

    return $choiceReturn   
}

# create the service
$Service = New-Service -Id $Identity -Impersonation $Impersonate -Credentials $Credentials -DefaultCredentials $UseDefaultCredentials -Url $Url -IgnoreSSL $IgnoreSSLCertificate

if ($service -ne $null) {
    # first check, if DelegateToRemove or DelegateTo got any entries
    if ($DelegateToRemove.Count -eq 0 -and $DelegateToSet.Count -eq 0) {
        $DelegateList = $null
        $DelegateList = Get-Delegates -Service $service -Identity $Identity
    
        if (($DelegateList -ne $null) -and ($DelegateList.DelegateUserResponses.Count -ge 1)) {
            $logger.Write("Total delegates: $($DelegateList.DelegateUserResponses.Count)",0,$WriteOnConsole)
            foreach ($User in $DelegateList.DelegateUserResponses) {
                $logger.Write("Delegate: $($User.DelegateUser.UserId.PrimarySmtpAddress)",0,$WriteOnConsole)
                $logger.Write("Permissions...",0, $WriteOnConsole)
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

    # check if any value is set fpr remove or set/adding a delegate
    if ($DelegateToRemove.Count -ge 1 -or $DelegateToSet.Count -ge 1) {
        # Variables
        [string[]]$ToRemove = @()
        [hashtable]$FoundDelegates = @{}
        [string[]]$DelegateAddList = @()
        [string[]]$DelegateToAlterList = @()
        
        $DelegateList = Get-Delegates -service $service -Identity $Identity
        if (($DelegateList -ne $null) -and ($DelegateList.DelegateUserResponses.Count -ge 1)) {
            for ($i = 0; $i -le $delegatelist.DelegateUserResponses.Count-1; $i++) {
                $FoundDelegates.Add("Delegate $($i)",$DelegateList.DelegateUserResponses[$i].DelegateUser.UserId.PrimarySmtpAddress.ToLower())
            }
        }
        
        # is there any user to remove?
        if ($DelegateToRemove.Count -ge 1)  {
            for($i = 0; $i -le $DelegateToRemove.Count-1;$i++) {
                if($FoundDelegates.ContainsValue($DelegateToRemove[$i])) {
                    if ($Confirm) {
                        if ((Request-Choice -Caption "$($DelegateToRemove[$i]) found. Do you want to remove the delegate?") -eq 0) {
                            $ToRemove += $DelegateToRemove[$i]
                            $logger.Write("Delegate $($DelegateToRemove[$i]) found. Try to removing now",0,$WriteOnConsole)
                        }
                        else {
                            $logger.Write("Delegate $($DelegateToRemove[$i]) found but deletion was not approved by the user.",0,$WriteOnConsole)
                        }
                    }
                    else {
                        $ToRemove += $DelegateToRemove[$i]
                        $logger.Write("Delegate $($DelegateToRemove[$i]) found. Try to removing now",0,$WriteOnConsole)
                    }
                }
                else {
                    $logger.Write("$($DelegateToRemove[$i]) is not a delegate of $($Identity)",0,$WriteOnConsole)
                }  
            }
            if ($ToRemove.Count -ge 1) {
                Remove-Delegate -Service $service -Identity $Identity -DelegateList $ToRemove
            }
        }

        if ($DelegateToSet.Count -ge 1) {
            If ($FoundDelegates.Count -ge 1) {
                foreach($User in $DelegateToSet) {
                    if ($FoundDelegates.ContainsValue($User)) {
                        # User is delegate, we need to alter him
                        $DelegateToAlterList += $User
                    }
                    else {
                        # New delegate to add
                        $DelegateToAddList += $User
                    }
                }
            }
            else
            {
                # we have no delegates on the mailbox, so add all to the $DelegateAddList
                $DelegateAddList = $DelegateToSet
            }
    
            if (($DelegateToAddList -ne $null) -or ($DelegateToAlterList -ne $null)) {
                # First we adding the delegates
                if ($DelegateAddList.Count -ge 1) {
                    Set-Delegate -Service $service -Identity $Identity -Delegates $DelegateToSet `
                        -CalendarFolderPermissionLevel (Get-Permission -Permission $CalendarPermissions) `
                        -ContactsFolderPermissionLevel (Get-Permission -Permission $ContactsPermissions) `
                        -InboxFolderPermissionLevel (Get-Permission -Permission $InboxPermissions) -JournalFolderPermissionLevel (Get-Permission -Permission $JournalPermissions) `
                        -NotesFolderPermissionLevel (Get-Permission -Permission $NotesPermissions) -TasksFolderPermissionLevel (Get-Permission -Permission $TasksPermissions) `
                        -ReceiveCopiesOfMeetingMessages $ReceiveCopiesOfMeetingMessages -ViewPrivateItems $CanViewPrivateItems -DelegateList $DelegateList -Add
                }
                
                # The we alter delegates
                if ($DelegateToAlterList.Count -ge 1) {
                    Set-Delegate -Service $service -Identity $Identity -Delegates $DelegateToSet `
                        -CalendarFolderPermissionLevel (Get-Permission -Permission $CalendarPermissions) `
                        -ContactsFolderPermissionLevel (Get-Permission -Permission $ContactsPermissions) `
                        -InboxFolderPermissionLevel (Get-Permission -Permission $InboxPermissions) -JournalFolderPermissionLevel (Get-Permission -Permission $JournalPermissions) `
                        -NotesFolderPermissionLevel (Get-Permission -Permission $NotesPermissions) -TasksFolderPermissionLevel (Get-Permission -Permission $TasksPermissions) `
                        -ReceiveCopiesOfMeetingMessages $ReceiveCopiesOfMeetingMessages -ViewPrivateItems $CanViewPrivateItems -DelegateList $DelegateList
                }
            }    
        }       
    }
}