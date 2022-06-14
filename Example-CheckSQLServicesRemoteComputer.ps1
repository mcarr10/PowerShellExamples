# Powershell script to that returns services named like SQL
# 2022-06-11

$ComputerName = read-host -prompt "Enter computer name"
get-service -ComputerName $ComputerName | Where-object {$_.name -like '*sql*' -or $_.name -like '*report*'}
read-host "Press enter to continue"