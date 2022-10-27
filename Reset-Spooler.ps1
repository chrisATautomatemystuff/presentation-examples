function Reset-Spooler
{
    Param([Parameter(Mandatory=$true)]$PrintServerFQDN)
    Invoke-Command -ComputerName $PrintServerFQDN -ScriptBlock {Stop-Service -Name Spooler -Force}
    Invoke-Command -ComputerName $PrintServerFQDN -ScriptBlock {Remove-Item C:\Windows\System32\spool\PRINTERS\*.*}
    Invoke-Command -ComputerName $PrintServerFQDN -ScriptBlock {Start-Service -Name Spooler}

}