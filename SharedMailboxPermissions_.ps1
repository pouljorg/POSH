[CmdletBinding()]
#Paramters
PARAM
(
        [String]$ExchangeServerFQDN = 'http://kbg-exch.su.int/powershell'
    ,   [String]$UserPrincipalName = 'lmn@superusers.dk'
    ,   [string]$GroupNamePrefix = "exch-"                                # Set the prefix for the groups that control the shared mailboxes. Note: Make sure you only user lower case!
    ,   [string]$GroupNameSuffix = "-fullaccess"                          # Set the suffix Note: Make sure you only use lower case!
    ,   [Switch]$Online
    ,   [string]$LogfilePath = "c:\Logs\"                                 # Logfile path
    ,   [Parameter(Mandatory=$true,HelpMessage="Indtast Ldap path IE. 'OU=Salg,OU=Sjælland,DC=Su,DC=Int'")]
        [string]$LDAP_Distinguished_Name_OU                               # OU where the security groups resides
)

<#
    #  Index
    01 Variables
    02 Functions
    03 Check Variables/Parameters
    04 Connect to Exchange server
    05 Find Security groups from suffix/preffix
    06 Loop through groups -match security groups with mail groups, set permissions pr grp
#>

#region 01 Variables
$Logfile = $logfilepath + "SharedMailBoxes-log.csv"                              # Log file bør laves som csv/object - har ændre funktion
$LogID = (Get-Date).ToString("yyyyMMddHHmmss")

if ( -not(Test-Path -LiteralPath $LogfilePath))
{
    try{        
        New-Item -Path $LogfilePath -ItemType Directory -ErrorAction Stop | Out-Null 
        Write-Verbose -Message "$LogfilePath Created to store logfiles"
        Write-log -Message "$LogfilePath Created to store logfiles" -Path $LogfilePath -LogID $LogID
    }
    catch
    {
        Write-Host -ForegroundColor Red "$LogfilePath does not exist, and can't be created"
        Write-Verbose -Message "Unable to create location: $LogfilePath for storing logfiles"
        Write-log -Message "Unable to create location: $LogfilePath for storing logfiles" -Path $LogfilePath -LogID $LogID
    }
}
#endregion

#region 02 Functions
function Write-log
{
    [CmdLetBinding()]
    Param(
            [Parameter(Mandatory=$true,Position =0)]
            [string]$Message,
            [Parameter(Mandatory=$true,Position =1)]
            [string]$Path,
            [Parameter(Mandatory=$true,Position =2)]
            [string]$LogID
        )

    $Date = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    #$Date = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") -- if UTC Time Registration is preferred
    
    $Hash = [ordered]@{
    'Date' = $Date
    'Message' = $Message
    'LogID' = "Log" + $LogID
    }

    $Log = New-Object -TypeName PsObject -Property $Hash
    $Log | Export-Csv -LiteralPath $Path -Append -Encoding Default -Delimiter ';' -NoTypeInformation -Force
}

# Checking Logfile before run
# If logfile > 10Mb - zip and archive log
$logcheck = Get-Item -Path $Logfile -ErrorAction SilentlyContinue
if ($logcheck.length -gt 10MB) #10485760)
{
    $Date = Get-Date
    $ArchiveLogPath = "$logfilepath\Archive\$($Date.ToString("yyyy-MM-dd"))"
    
    Write-Verbose -Message 'Archiving Logfile'
    Write-log -Message "Log Archived @ $ArchiveLogPath\$($Date.ToString("yyyy-MM-dd HH_mm_ss")).zip" -Path $Logfile -LogID $LogID
        
    if (-not (Test-Path -LiteralPath $ArchiveLogPath))
    {
        New-Item -Path $ArchiveLogPath -ItemType Directory -Force
    }

    Compress-Archive -LiteralPath $Logfile -DestinationPath "$ArchiveLogPath\$($Date.ToString("yyyy-MM-dd HH_mm_ss")).zip" -Force
    Remove-Item -LiteralPath $Logfile
}
#endregion

#region 03 Check Variables/Parameters
Write-log -Message 'Script Start' -Path $Logfile -LogID $LogID
Write-Verbose -Message 'Script Start'

#Test if OU is valid
try
{
    # 
    Test-Path -LiteralPath "AD:$LDAP_Distinguished_Name_OU" #returns True if exists and False If Not
    [adsi]::Exists("LDAP://$LDAP_Distinguished_Name_OU") | Out-Null
    Write-Verbose -Message "OU found"
}
catch 
{
    Write-Verbose -Message "OU not found"
    Write-Verbose -Message "Script End"
    Write-log -Message "$LDAP_Distinguished_Name_OU not found" -Path $Logfile -LogID $LogID
    Write-log -Message "Script End" -Path $Logfile -LogID $LogID
    break;
}

# Check if dependency modules exist
Write-Verbose -Message 'Checking modul dependency'

if ((Get-Module -ListAvailable -Name 'ActiveDirectory').count -ne 1)
{
    Write-Host -ForegroundColor Yellow -BackgroundColor Black 'Missing Module ActiveDirectory'
    Write-log -Message "Could not find dependency modules" -Path $Logfile -LogID $LogID
    Write-log -Message "Script End" -Path $Logfile -LogID $LogID
    break;
}
#endregion

#region 04 Connect to Exchange server
If ($Online) 
{
    # Connect to Exchange (Online)
    # Kan laves som cred object, men nok bare anbefale at køre via kerb/service account
    $Cred = Get-Credential
    
    # Ændre $ExchangeServerFQDN til Exchangeonline uri
    $ExchangeServerFQDN = 'https://outlook.office365.com/powershell-liveid'
    try
    {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeServerFQDN -Credential $Cred -Authentication Basic -AllowRedirection -Name 'MailSession' -ErrorAction Stop
        set-ExecutionPolicy remotesigned -Force
        Import-PSSession $Session
        Import-Module MSOnline -Verbose:$false | Out-Null
        Connect-MsolService -Credential $Cred

        Write-log -Message "Connected to $ServerFQDN" -Path $Logfile -LogID $LogID
    }
    Catch
    {
        Write-Host -ForegroundColor Red 'Could not connect to Exchange Online'
        Write-log -Message 'Could not connect to Exchange Online' -Path $Logfile -LogID $LogID
        Write-log -Message "Script End" -Path $Logfile -LogID $LogID
        Write-Verbose -Message "Script End"

        Get-PSSession -Name 'MailSession' | Remove-PSSession
        break;
    }
    
} 
Else 
{
    # Connect to Exchange (On-Premesis)
    $ServerFQDN = $ExchangeServerFQDN.split('/')[2]
    try
    {
        $Session = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri $ExchangeServerFQDN -Authentication Kerberos -Name 'MailSession'
        Import-PSSession -Session $Session | Out-Null
        Write-Verbose -Message "Connected to $ServerFQDN"
        Write-log -Message "Connected to $ServerFQDN" -Path $Logfile -LogID $LogID
    }
    Catch
    {
        Write-Host -ForegroundColor Red "Could not connect to $ServerFQDN"
        Write-log -Message "Could not connect to $ServerFQDN" -Path $Logfile -LogID $LogID
        Write-log -Message "Script End" -Path $Logfile -LogID $LogID
        Write-Verbose -Message "Script End"

        Get-PSSession -Name 'MailSession' | Remove-PSSession
        break;
    }
}
#endregion

#region 05 Find Security groups from suffix/preffix
# Find all groups used for setting permissions on Shared Mailboxes
$GroupNameSearchString = $GroupnamePrefix + "*" + $GroupNameSuffix
write-Verbose -Message "Searching for $GroupNameSearchString"

try
{
    $Groups = Get-ADGroup -Filter "name -like '$GroupNameSearchString'" -SearchBase $GroupsOU -ResultSetSize $null -ErrorAction Stop
}
catch
{
    Write-Host -ForegroundColor Red "No Groups matching $GroupNameSearchString found"
    Write-log -Message "No Groups matching $GroupNameSearchString found" -Path $Logfile -LogID $LogID
    Write-log -Message "Script End" -Path $Logfile -LogID $LogID

    Get-PSSession -Name 'MailSession' | Remove-PSSession
    break;
}
#endregion

#region 06 Loop through groups -match security groups with mail groups, set permissions pr grp
# Iterate through all Groups used for handling FullAccess & Send-As permissions for Shared Mailboxes
foreach ($Group in $Groups)
{
    $ExistingDelegates = New-Object -TypeName System.Collections.Generic.List[String]   # MailboxUsers with permission on Shared Mailbox
    $RequiredDelegates = New-Object -TypeName System.Collections.Generic.List[String]   # MailboxUSers required to have permission on Current Shared Mailbox
    $SharedDelegates = @()
    $SharedMBX = $null
    $SharedName = $Group.name.tolower() -replace $GroupNamePrefix, "" -replace $GroupNameSufffix, ""

    write-Verbose $SharedName
  
    # Retrieve the shared mailbox matching Current Group
    if ( $Online) 
    {
        $SharedMBX = Get-Mailbox -RecipientTypeDetails shared -Filter "displayname -eq '$SharedName'" -ErrorAction SilentlyContinue
    } 
    Else 
    {
        $SharedMBX = Get-Mailbox -RecipientTypeDetails sharedMailbox -Filter "displayname -eq '$SharedName'" -ErrorAction SilentlyContinue
    } 


    # Verify if the Shared Mailbox exists or not. If it doesnt, log it and skip.
    if ($SharedMBX -eq $Null)
    { 
        Write-log -message "Shared mailbox $($SharedName) does not exist"  -Path $Logfile -LogID $LogID
    }
    else
    {
        Write-log -message "Checking for Group: $($Group.name) and Shared Mailbox: $($SharedMBX.displayname)" -Path $Logfile -LogID $LogID
            
        # Shared Mailbox exists, get the current delegates
        $SharedDelegates =  Get-Mailbox -identity $SharedMBX.SamAccountName | 
                                Get-MailboxPermission |
                                where { 
                                          ($_.AccessRights -eq 'FullAccess')       -and 
                                          ($_.IsInherited -eq $false)              -and 
                                          -not ($_.User -like 'NT AUTHORITY\SELF') 
                                      }
        
        # Iterate through all delegates of the Shared Mailbox associated with the Group
        # Also, make sure that if the membership is empty, do not throw an error
        
        if ($SharedDelegates.count -lt 1)
        {
            $ExistingDelegates.Add("empty")
        }
        else
        {
            $SharedDelegates.user | foreach { $ExistingDelegates.Add( $_.split('\')[-1] ) }
        }
            
        # Iterate the Members of the Current Shared Permission Group
        $Groupmembers = Get-ADGroupMember $group | Select -ExpandProperty samaccountname
        
        if ($SharedDelegates.count -lt 1)
        {
            $RequiredDelegates.Add("empty")
        }
        else
        {
            $Groupmembers | foreach { $RequiredDelegates.Add($_) }
        }
            
        # Find the differences between the Shared Mailbox delegates and the Group members
        # Store the result in the two Arraylists ($SharedDelegatesToAdd and $SharedDelegatesToRemove)
        $SharedDelegateToRemove = Compare-Object -ReferenceObject $RequiredDelegates -DifferenceObject $ExistingDelegates | 
                where { $_.sideindicator -eq '=>' } | Select -ExpandProperty InputObject
        $SharedDelegateToAdd = Compare-Object -ReferenceObject $RequiredDelegates -DifferenceObject $ExistingDelegates | 
                where { $_.sideindicator -eq '<=' } | Select -ExpandProperty InputObject
        
        # Report
        $SharedDelegateToAdd | foreach { Write-log -Message "Delegate Add: $_" -Path $Logfile -LogID $LogID }
        $SharedDelegateToRemove | foreach { Write-log -Message "Delegate Remove: $_" -Path $Logfile -LogID $LogID }
        
        # REMOVE Permissions for users no longer permitted to access Shared Mailbox
        foreach ($Delegate in $SharedDelegateToRemove)
        {
            $User = Get-ADUser -Identity $Delegate
            Get-Mailbox -identity $SharedMBX | Remove-MailboxPermission -AccessRights fullaccess -User $User.UserPrincipalName -Confirm:$false
            
            If ($Online) 
            {
                Remove-RecipientPermission -identity $SharedMBX -Trustee $User.UserPrincipalName -AccessRights SendAs -Confirm:$false
            } 
            Else 
            {
                Remove-AdPermission -identity $SharedMBX -User $User.UserPrincipalName -ExtendedRights 'Send As' -Confirm:$false
            }
        }  
        
        # ADD Permissions for new users requiring access Shared Mailbox
        foreach ($Delegate in $SharedDelegateToAdd)
        {
            $User = Get-ADUser -Identity $Delegate
            Get-Mailbox -identity $SharedMBX | Add-MailboxPermission -AccessRights fullaccess -User $User.UserPrincipalName -AutoMapping $true -Confirm:$false
            
            If ($Online) 
            {
                Add-RecipientPermission -identity $SharedMBX -Trustee $User.UserPrincipalName -AccessRights SendAs -Confirm:$false
            } 
            else 
            {
                Add-AdPermission -identity $SharedMBX -User $User.UserPrincipalName -ExtendedRights 'Send As' -Confirm:$false
            }
        }
    }
}
#endregion 

Write-log -Message "Script Ended" -Path $Logfile -LogID $LogID
Get-PSSession -Name 'MailSession' | Remove-PSSession