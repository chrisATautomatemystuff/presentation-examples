<#
  Tattletale script to reveal 'primary users' (40 hours over 2 week timeframe device usage) who are only assigned A1 licensing (no windows licensing)...
#>
$sqlQuery = @"

    SELECT
      RV.Netbios_Name0,
      RV.User_Domain0,
      RV.User_Name0,
      CS.UserName0,
      SCUM.TopConsoleUser0,
      CDR.CurrentLogonUser
    FROM
      dbo.v_R_System_Valid RV
      LEFT OUTER JOIN
        dbo.v_GS_COMPUTER_SYSTEM CS
      ON
        RV.ResourceID = CS.ResourceID
      LEFT OUTER JOIN
        dbo.v_GS_SYSTEM_CONSOLE_USAGE_MAXGROUP SCUM
      ON
        RV.ResourceID = SCUM.ResourceID
      LEFT OUTER JOIN
        dbo.v_CombinedDeviceResources CDR
      ON
        RV.ResourceID = CDR.MachineID
    WHERE
      RV.Netbios_Name0 NOT LIKE 'V-%'
    ORDER BY
      RV.User_Domain0
    DESC

"@

$devices = [System.Collections.ArrayList]@(Invoke-Sqlcmd -ServerInstance "<SERVER INSTANCE>" -Database "<DB NAME>" -Query $sqlQuery)

$devices | Measure-Object

$array = @()

foreach ($device in $devices)
{
  $hostname = $device.NetBios_Name0
  $domain = $device.User_Domain0
  $username = $device.User_Name0
  
  switch ($domain)
  {
    <DOMAINSHORTNAME> {
      $domainName = '<DOMAINNAME>'
      $a1members = Get-ADGroupMember -Server <FQDNSERVERNAME> -Identity 'O365_A1_Faculty' | Select-Object -ExpandProperty SamAccountName}
    <DOMAINSHORTNAME2> {
      $domainName = '<DOMAINNAME2>'
      $a1members = Get-ADGroupMember -Server <FQDNSERVERNAME2> -Identity 'O365_A1_Faculty' | Select-Object -ExpandProperty SamAccountName}
    <DOMAINSHORTNAME3> {
      $domainName = '<DOMAINNAME3>'
      $a1members = Get-ADGroupMember -Server <FQDNSERVERNAME3> -Identity 'O365_A1_Faculty' | Select-Object -ExpandProperty SamAccountName}
    default {$domainName = 'INVESTIGATE'}
  }

  if ($domainName -ne 'INVESTIGATE')
  {
    $domainController = (Get-ADDomain -Identity $domainName).PDCEmulator
    
    Write-Host "Checking on $username from $domainName"

    $aduser = Get-ADUser -Server $domainController -Identity $username

    if($aduser.Enabled -eq $false)
    {
      Write-Host "$username is disabled." -ForegroundColor Red
    }

    if ($a1members -contains $aduser.SamAccountName)
    {
      $object = New-Object -TypeName PSObject
      $object | Add-Member -MemberType NoteProperty -Name USER -Value $username
      $object | Add-Member -MemberType NoteProperty -Name DOMAIN -Value $domainName
      $object | Add-Member -MemberType NoteProperty -Name HOSTNAME -Value $hostname
      Write-Host "$username from $domain is in A1 and should likely be licensed as they are a primary user on $hostname." -ForegroundColor Cyan

      $array += $object
    }
  }
}

$array | Export-Csv -Path C:\scripts\primary_a1members.csv -NoTypeInformation