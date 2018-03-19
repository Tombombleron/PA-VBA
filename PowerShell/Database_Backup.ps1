$currDate = Get-Date -UFormat "%Y-%m-%d"

Copy-Item -Path Z:/Accounts/Accounts/Dir1/Dir2/DataBase/database.accdb `
-Destination Z:/Accounts/Accounts/Dir1/Dir2/Database/Backup/database_${currDate}.accdb
