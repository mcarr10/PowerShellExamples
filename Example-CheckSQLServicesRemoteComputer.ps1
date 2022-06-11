#
#

$ComputerName = read-host -prompt "Enter computer name"
get-service -ComputerName $ComputerName | Where-object {$_.name -like '*sql*'}
read-host "Press enter to continue"