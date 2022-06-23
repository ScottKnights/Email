# Repair recipients that don't have address policy inheritance enabled so don't get a hybrid proxy (@mail.onmicrosoft) email address.
# Record the current primary email address, enable email policy inheritance, disable it again then reset the primary address.

# ============================================================================
#region Functions
# ============================================================================
function Repair-Mailbox {
    Param(
        [string]$recipient
    )

    $mailbox=Get-Mailbox $recipient
    $emailaddress=$mailbox.windowsemailaddress.address
    "Repairing Mailbox $recipient"
    if ($emailaddress) {
        Set-Mailbox $recipient -EmailAddressPolicyEnabled $true
        start-sleep 1
        Set-Mailbox $recipient -EmailAddressPolicyEnabled $false
        start-sleep 1
        Set-Mailbox $recipient -WindowsEmailAddress $emailaddress
    } else {
        "No email address"
    }
}

function Repair-DynamicGroup {
    Param(
        [string]$recipient
    )

    $group=Get-DynamicDistributionGroup $recipient
    $emailaddress=$group.windowsemailaddress.address
    "Repairing Dynamic Group $recipient"
    if ($emailaddress) {
        Set-DynamicDistributionGroup $recipient -EmailAddressPolicyEnabled $true
        start-sleep 1
        Set-DynamicDistributionGroup $recipient -EmailAddressPolicyEnabled $false
        start-sleep 1
        Set-DynamicDistributionGroup $recipient -WindowsEmailAddress $emailaddress
    } else {
        "No email address"
    }
}

function Repair-DistributionGroup {
    Param(
        [string]$recipient
    )

    $group=Get-DistributionGroup $recipient
    $emailaddress=$group.windowsemailaddress.address
    "Repairing Distribution Group $recipient"
    if ($emailaddress) {
        Set-DistributionGroup $recipient -EmailAddressPolicyEnabled $true
        start-sleep 1
        Set-DistributionGroup $recipient -EmailAddressPolicyEnabled $false
        start-sleep 1
        Set-DistributionGroup $recipient -WindowsEmailAddress $emailaddress
    } else {
        "No email address"
    }
}

function Repair-Contact {
    Param(
        [string]$recipient
    )

    $contact=Get-MailContact $recipient
    $emailaddress=$contact.windowsemailaddress.address
    "Repairing Contact $recipient"
    if ($emailaddress) {
        Set-MailContact $recipient -EmailAddressPolicyEnabled $true
        start-sleep 1
        Set-MailContact $recipient -EmailAddressPolicyEnabled $false
        start-sleep 1
        Set-MailContact $recipient -WindowsEmailAddress $emailaddress
    } else {
        "No email address"
    }
}
#endregion Functions

# ============================================================================
#region Execute
# ============================================================================
$recipients=Get-recipient -Filter {emailaddresses -notlike "*mail.onmicrosoft*"} -resultsize unlimited
foreach ($recipient in $recipients) {
	$Repairname=$recipient.name
	$type=$recipient.recipienttype
	if ($type -like "*User*") {
		Repair-mailbox $Repairname
	} elseif ($type -eq "DynamicDistributionGroup") {
		Repair-DynamicGroup $Repairname
	} elseif ($type -like "*group*") {
		Repair-DistributionGroup $Repairname
	} elseif ($type -eq "MailContact") {
		Repair-Contact $Repairname
	} else {
		"Unknown Recipient type $Repairname"
	}
}

#endregion Execute
