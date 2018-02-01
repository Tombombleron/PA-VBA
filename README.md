# PA-VBA
This is a collection of sub procedures and functions which I've created and implemented whilst administrating the Corporate Credit Card program of a well-known conglomerate. The majority of the code is related to automating the sending of notification emails from MS Access and Excel.

There are several PERSONAL.XLSB sub procedures included which I also run on a daily basis.

## Access

The Access folder consists of macros from a custom database I created specifically for the role; all the macros are run from forms in an executable-only version of the database.

## Excel

Excel is used largely for ad-hoc reporting, and the generation and sending of a set of emails with attachments each morning.

### send_ssc_emails

This script eliminated \~90 minutes of work each morning. The original process included:
- Manually creating emails with concatenated strings from a CSV for subjects,
- Attaching a file with the same name as one record in the CSV,
- Sending them to the same email address.

The script takes between 3 and 5 minutes to run on its own now. The reason it takes so long is a result of the awkward way of sending each email to comply with the receiving server's 'anti-spam' protocols.

## Outlook

This sub procedure is called by a rule which runs on all incoming mail.
Throughout the month I'm sporadically emailed notifications of people who are leaving the business, it is a hassle to find all these emails at the end of the month and manually enter each employee into a spreadsheet to forward on to bossman.
This procedure, when the rule fires, automatically adds the email subject (including the employee name) to a workbook, and either creates a new sheet for a new month, or adds the strings to the sheet for the existing month.
