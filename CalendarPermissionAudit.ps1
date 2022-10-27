#script to audit user calendars to see the 'default' permissions that users in the domain get to the folder
#users could potentially be revealing PII/Sensitive info in their calendar items and exposing it to students as well

Connect-ExchangeOnline

$staffmailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object -FilterScript {$_.userprincipalname -like "*@<STAFF DOMAIN>"}

#$staffmailboxes = $staffmailboxes | Select-Object -First 5
$cachedmailboxes = $staffmailboxes
$array = @()

foreach ($staffmailbox in $cachedmailboxes)
{
  $staffmailbox.UserPrincipalName
  $nonstandardcalendarpermissions = Get-MailboxFolderPermission -Identity "$($staffmailbox.userprincipalname):\Calendar" -User 'Default' | Where-Object -FilterScript {$_.accessrights -notlike "*AvailabilityOnly*"}

  foreach ($nonstandardcalendarpermission in $nonstandardcalendarpermissions)
  {
    $nonstandardobject = New-Object -TypeName PSObject
    $nonstandardobject | Add-Member -MemberType NoteProperty -Name userprincipalname -Value $staffmailbox.UserPrincipalName
    $nonstandardobject | Add-Member -MemberType NoteProperty -Name foldername -Value $nonstandardcalendarpermission.FolderName
    $nonstandardobject | Add-member -MemberType NoteProperty -Name accessrights -Value $nonstandardcalendarpermission.AccessRights
    $array += $nonstandardobject
  }
}
$array

$array | Export-Csv -Path C:\scripts\calendarpermissionaudit_<DISTRICT>_(Get-Date -Format MMddyy).csv -NoTypeInformation