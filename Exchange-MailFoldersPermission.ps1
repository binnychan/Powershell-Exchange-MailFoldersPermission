<#
.SYNOPSIS

Mailbox folder permission to the delegates for user and resource mailboxes

.DESCRIPTION

Important task of the Exchange admin to assign the folder permission to the delegates
when new delegates added to the generic mailbox and Resource mailboxes.
the script simplify the task and eliminate the manual errors

.NOTES
Author : Kamaraj Ulaganathan
Email: kamaraj0926@outlook.com
Requires: PowerShell Version 1.0

.UPDATE
Author : Binny Chan 
Email  : binnychan(at)gmail.com
Date   : 15/09/2021

#>

Write-host “

Mailbox folder Permission
——————————–

1.Assign Folder permission to Single folder

2.Assign Folder Permission to All folders(includes user created,default,recoverable mailbox folders)

3.Assign Folder permission only to the default folders(inbox,calendar,….)

4.Assign Folder permission only to the user created folders

5.Assign Folder permission only to the default mail folders(inbox,sent items,….)

6.Check Folder permission to All folder (Everyone)

7.Check Folder permission to All folder (1 user)

8.Remove Folder permission to All folders (1 user)

9.Exit ” -ForeGround “Cyan”

$option = Read-host “Choose the Option”

switch ($option)
{

1 {

$Mailbox = Read-Host “Enter Mailbox ID “

$Folder = Read-Host “Enter the FOLDER NAME ( Examplles : Inbox,calendar…)”

$delegate = Read-Host “Enter Delegate ID “

$Permission = Read-Host “Enter Type of Permission(Author, Editor, Owner, Reviewer, none)”

$foldername = $Mailbox + “:\” + $folder

If ($folder -ne “”)

{

Add-MailboxFolderPermission $foldername -User $delegate -AccessRights $Permission -confirm:$true

}

Else

{ Write-Host ” Please Enter Folder name ” -ForeGround “red”}

;break}

2
{

$Mailbox = Read-Host “Enter Mailbox ID ”

$delegate = Read-Host “Enter Delegate ID “

$Permission = Read-Host “Enter Type of Permission(Author, Editor, Owner, Reviewer, none)”

$AllFolders = Get-MailboxFolderStatistics $Mailbox | Where { $_.FolderPath.ToLower().StartsWith(“/“) -eq $True }

ForEach($folder in $AllFolders)

{

$foldername = $Mailbox + “:” + $folder.FolderPath.Replace(“/”,”\”)

Add-MailboxFolderPermission $foldername -User $delegate -AccessRights $Permission -confirm:$true

}

;Break}

3 {

$Mailbox = Read-Host “Enter Mailbox ID ”

$delegate = Read-Host “Enter Delegate ID “

$Permission = Read-Host “Enter Type of Permission(Author, Editor, Owner, Reviewer, none)”

$Default = Get-MailboxFolderStatistics $mailbox | ?{$_.foldertype -ne “user created” -and $_.foldertype -ne “Recoverableitemsroot” -and $_.foldertype -ne “RecoverableItemsDeletions” -and $_.foldertype -ne “RecoverableItemspurges” -and $_.foldertype -ne “RecoverableItemsversions” -and $_.foldertype -ne “syncissues” -and $_.foldertype -ne “conflicts” -and $_.foldertype -ne “localfailures” -and $_.foldertype -ne “serverfailures” -and $_.foldertype -ne “RssSubscription” -and $_.foldertype -ne “JunkEmail” -and $_.foldertype -ne “CommunicatorHistory” -and $_.foldertype -ne “conversationactions”}

ForEach($folder in $default)

{

$foldername = $Mailbox + “:” + $folder.FolderPath.Replace(“/”,”\”)

Add-MailboxFolderPermission $foldername -User $delegate -AccessRights $Permission -confirm:$true

}

;break}

4 {

$Mailbox = Read-Host “Enter Mailbox ID ”

$delegate = Read-Host “Enter Delegate ID “

$Permission = Read-Host “Enter Type of Permission(Author, Editor, Owner, Reviewer, none)”

$Default = Get-MailboxFolderStatistics $mailbox | ?{$_.foldertype -eq “user created”}

ForEach($folder in $default)

{

$foldername = $Mailbox + “:” + $folder.FolderPath.Replace(“/”,”\”)

Add-MailboxFolderPermission $foldername -User $delegate -AccessRights $Permission -confirm:$true

}

;break}

5 {

$Mailbox = Read-Host “Enter Mailbox ID ”

$delegate = Read-Host “Enter Delegate ID “

$Permission = Read-Host “Enter Type of Permission(Author, Editor, Owner, Reviewer, none)”

$Default = Get-MailboxFolderStatistics $mailbox | ?{$_.foldertype -eq “Inbox” -or $_.foldertype -eq “Drafts” -or $_.foldertype -eq “SentItems” -or $_.foldertype -eq “DeletedItems” -or $_.foldertype -eq “Archive”}

ForEach($folder in $default)

{

$foldername = $Mailbox + “:” + $folder.FolderPath.Replace(“/”,”\”)

Add-MailboxFolderPermission $foldername -User $delegate -AccessRights $Permission -confirm:$true

}

;break}

6 {

$Mailbox = Read-Host “Enter Mailbox ID ”

$exclusions = @("/Sync Issues",
                "/Sync Issues/Conflicts",
                "/Sync Issues/Local Failures",
                "/Sync Issues/Server Failures",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions"
                )

$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)
$output = @()

foreach ($mailboxfolder in $mailboxfolders)
{
    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace("\Top of Information Store","\")
    }
    $identity = "$($mailbox):$folder"
    try
    {
        $folderusers = Get-MailboxFolderPermission -Identity $identity -ErrorAction SilentlyContinue
        foreach ($folderuser in $folderusers)
	{
		$obj = [PSCustomObject]@{Folder=$identity;User=$folderuser.User;Permission=$folderuser.AccessRights}
		$output += $obj
	}
    }
    catch
    {
	$obj = [PSCustomObject]@{Folder=$identity;User=$user;Permission=$_.Exception.Message}
        $output += $obj
    }
}
$output | ogv -Title ("Mailbox Permission Report (" + $Mailbox + ") @ " + (Get-Date -Format "yyyyMMdd HH:mm K"))

;break}

7 {

$Mailbox = Read-Host “Enter Mailbox ID ”

$delegate = Read-Host “Enter Delegate ID “

$exclusions = @("/Sync Issues",
                "/Sync Issues/Conflicts",
                "/Sync Issues/Local Failures",
                "/Sync Issues/Server Failures",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions"
                )

$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

foreach ($mailboxfolder in $mailboxfolders)
{
    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace("\Top of Information Store","\")
    }
    $identity = "$($mailbox):$folder"

    Write-Host -NoNewline -ForegroundColor Cyan "($delegate)"
    Write-Host -NoNewline " $identity "

    if (Get-MailboxFolderPermission -Identity $identity -User $delegate -ErrorAction SilentlyContinue)
    {
        try
        {
            Write-Host -ForegroundColor Green (Get-MailboxFolderPermission -Identity $identity -User $delegate).AccessRights
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }
    }else{
	Write-Host -ForegroundColor Yellow "Not found"
    }
}

;break}

8 {

$Mailbox = Read-Host “Enter Mailbox ID ”

$delegate = Read-Host “Enter Delegate ID “

$exclusions = @("/Sync Issues",
                "/Sync Issues/Conflicts",
                "/Sync Issues/Local Failures",
                "/Sync Issues/Server Failures",
                "/Recoverable Items",
                "/Deletions",
                "/Purges",
                "/Versions"
                )

$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

foreach ($mailboxfolder in $mailboxfolders)
{
    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace("\Top of Information Store","\")
    }
    $identity = "$($mailbox):$folder"

    Write-Host -NoNewline -ForegroundColor Cyan "($delegate)"
    Write-Host -NoNewline " $identity "

    if (Get-MailboxFolderPermission -Identity $identity -User $delegate -ErrorAction SilentlyContinue)
    {
        try
        {
            Write-Host -NoNewLine -ForegroundColor Green (Get-MailboxFolderPermission -Identity $identity -User $delegate).AccessRights
	        Remove-MailboxFolderPermission -Identity $identity -User $delegate -Confirm:$false -ErrorAction STOP
	        Write-Host -ForegroundColor Green " Removed!"
        }
        catch
        {
            Write-Warning $_.Exception.Message
        }
    }else{
	Write-Host -ForegroundColor Yellow "Not found"
    }
}

;break}

9 {

}
}
