#script to look at mailbox forwarding and inbox rules to audit users sending email outside the domain

Connect-ExchangeOnline

$mailboxForwards = Get-Mailbox -Filter {ForwardingSmtpAddress -ne $null} | select name,delivertomailboxandforward,forwardingsmtpaddress

$mailboxes = Get-Mailbox -ResultSize unlimited -Filter {userprincipalname -like "*@<STAFFDOMAIN>"}
Foreach($mailbox in $mailboxes){
  Write-Host "Checking on $($mailbox.userprincipalname)"
  Get-InboxRule -Mailbox $mailbox.userprincipalname
}