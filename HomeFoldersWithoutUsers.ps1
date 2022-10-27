$folders = Get-ChildItem -Path \\<FQDN FILE SERVER>\<SHARE PATH> -Directory | select name
$array = @()

foreach ($folder in $folders)
{
  $foldername = $folder | select -ExpandProperty name
  try{
    Get-aduser -Identity $foldername | ?{$_.enabled -eq $true}
  }
  catch{
    Write-Host $_ -ForegroundColor Yellow
    $o = New-Object PSObject
    $o | Add-Member -MemberType NoteProperty -Name 'foldername' -Value $foldername
    $array += $o
  }
}
$array | Export-Csv -Path C:\scripts\FOLDERS_WITH_NO_USERS_$(Get-Date -Format MMddyy).csv -NoTypeInformation
