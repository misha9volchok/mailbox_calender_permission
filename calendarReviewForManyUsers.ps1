# Connect to 365 Admin center
Connect-ExchangeOnline 

$EmailAdressess = Get-Content "User-Route"

# input security group name

$group_name = Read-Host -Prompt 'Enter the Security group name '

# Add permission

function add_permission{ param ($EmailAdressess)
    foreach ($email in $EmailAdressess){
        Write-Host "================================================================"
        Write-Host $email
        Write-Host "----------------------------------------------------------------"
        try{
            Add-MailboxFolderPermission -Identity ${email}:\Calendar -User $group_name -AccessRights Reviewer -ErrorAction Stop
            }
        catch {
            Write-Host "No such Email account"-ForegroundColor DarkGray -BackgroundColor cyan
        }
    }
}

# Set a new permission

function set_permission{ param ($EmailAdressess)
    foreach ($email in $EmailAdressess){
        Write-Host "================================================================"
        Write-Host $email
        Write-Host "----------------------------------------------------------------"
        try{
            Set-MailboxFolderPermission -Identity ${email}:\Calendar -User $group_name -AccessRights Reviewer -ErrorAction Stop
            }
        catch {
            Write-Host "No such Email account"-ForegroundColor DarkGray -BackgroundColor cyan
        }
    }
}

# Get current permission

function get_permission{ param ($EmailAdressess)
    foreach ($email in $EmailAdressess){
        Write-Host "================================================================"
        Write-Host $email
        Write-Host "----------------------------------------------------------------"
        try{
            Get-MailboxFolderPermission -Identity ${email}:\Calendar -User $group_name -ErrorAction Stop
            }
        catch {
            Write-Host "No Permission found for mailbox"-ForegroundColor DarkGray -BackgroundColor cyan
        }
    }
}


# Option menu
Write-Host "Enter 1 to Add Reviewer Permission"
Write-Host "Enter 2 to change to Reviewer Permission"
Write-Host "Enter 3 to view current permission"
Write-Host "Enter 4 to Exit"

# Switch case for options
$choice = Read-Host -Prompt 'waiting for input'

switch ($choice)
 {
   1 {add_permission($EmailAdressess); Break}
   2 {set_permission($EmailAdressess); Break}
   3 {get_permission($EmailAdressess); Break}
   4 {Break}
 }