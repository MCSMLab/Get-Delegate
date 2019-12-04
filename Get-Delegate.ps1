<#
.SYNOPSIS
This script lists delegates for all mailbox folders.

.DESCRIPTION
This script lists delegates for all mailbox folders. Works on Exchange 2010+

.NOTES
Get-Delegate
v1.5
12/4/2019
By Nathan O'Bryan
nathan@mcsmlab.com
http://www.mcsmlab.com

Change Log
1.0 -
1.1 - Added status bar
1.5 - Added auto detection for EMS. If script isn't run in EMS it will ask you to enter FQDN of Exchange server for remote PS connection

.LINK
https://github.com/MCSMLab/Get-Delegate/blob/master/Get-Delegate.ps1

.EXAMPLE
.\Get-Delegates.ps1
#>

Clear-Host

If ($exscripts)
{
    Write-Host 'Exchange Management Shell loaded'
}

Else
{   
    $ExchangeServer = Read-Host "Enter the FQDN of your Exchange server"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session
}

Set-ADServerSettings -ViewEntireForest $True
$UserMailboxes = Get-Mailbox -RecipientTypeDetails 'UserMailbox' -ResultSize Unlimited

ForEach ($UserMailbox in $UserMailboxes)
{
    $i = $i+1
    Write-Progress -Activity "Reviewing Mailbox Permissions" -Status "For $UserMailbox" -PercentComplete ($i/$UserMailboxes.count*100)
    $Mailbox = "" + $UserMailbox.PrimarySmtpAddress
    $MailboxName = "" + $UserMailbox.Name
    $Folders = Get-MailboxFolderStatistics $Mailbox | ForEach-Object {$_.FolderPath} | ForEach-Object {$_.Replace(“/”,”\”)}
    $MailboxName >> .\DelegateReport.txt
    
    ForEach ($Folder in $Folders)
    {
        $FolderPath = $Mailbox + ":" + $Folder
        $Permissions = Get-MailboxFolderPermission -Identity $FolderPath -ErrorAction SilentlyContinue
        $Permissions = $Permissions | Where-Object { ($_.User -NotLike "Default") -And ($_.User -NotLike "Anonymous") -And ($_.AccessRights -NotLike "None") -And ($_.AccessRights -NotLike "Owner") }
        $Permissions | Select-Object User, FolderName, AccessRights | Format-Table  >> .\DelegateReport.txt
    }
}
