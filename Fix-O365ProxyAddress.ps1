# Fix mailboxes that don't have address policy inheritance enabled so don't get a hybrid proxy (@mail.onmicrosoft) email address.
# Record the current primary email address, enable email policy inheritance, disable it again then reset the primary address.
 
function fix-o365address {
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]$user
    )
 
    $mailbox=get-mailbox $user
    $emailaddress=$mailbox.windowsemailaddress.address
    "Fixing $user"
    if ($emailaddress) {
        set-mailbox $user -EmailAddressPolicyEnabled $true
        start-sleep 1
        set-mailbox $user -EmailAddressPolicyEnabled $false
        start-sleep 1
        set-mailbox $user -WindowsEmailAddress $emailaddress
    } else {
        "No email address"
    }
}
 
$users=Get-MailBox -Filter {emailaddresses -notlike "*mail.onmicrosoft*"}
foreach ($user in $users) {
    $fixname=$user.name
    fix-o365address $fixname
}