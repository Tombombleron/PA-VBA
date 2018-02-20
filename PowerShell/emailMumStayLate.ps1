$excuses = ("I've found a bug in one of my reports.", "I've been asked to stay late by a cardholder.", "I've been given an adhoc report to run.", "I've got an automation script which I really need to finish.","I want to get overtime.", "I have nothing to do at home, so I see no reason to leave.")
$rand = Get-Random -Maximum 5 -Minimum 0
Start-Process Outlook

$o = New-Object -com Outlook.Application


$mail = $o.CreateItem(0)

$mail.subject = “I'll be home later than usual“

$mail.body = “I've not left work yet because “ + $excuses[$rand] +" I'll see you later."

#for multiple email, use semi-colon ; to separate
$mail.To = “maureenjgreen@hotmail.co.uk"
$mail.Cc = "robertogreen@hotmail.co.uk"
$mail.Send()

#$o.Quit()
