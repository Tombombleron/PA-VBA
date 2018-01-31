Attribute VB_Name = "unNoteCardHolders"
Option Compare Database
Option Explicit

Sub sendOutlookEmailunNote(EmailAdd As String, colItem As String, emailOpt As String)

    Dim oApp As Outlook.Application, oMail As MailItem, Attachments() As String, i As Integer, FilePathToAdd As String
    
    ' FilePathToAdd allows multiple attachments to be added to the same email.
    ' each full file path should be separated by commas; the file path should include the file extension, as per below:
    FilePathToAdd = "D:\Users\r.green\Documents\Card Application Letter\PDF Files\Samsung Credit Card Letter 1_0_2.pdf"
    
    Set oApp = CreateObject("Outlook.application")
    Set oMail = oApp.CreateItem(olMailItem)
    
    With oMail
        ' readTXTFile is a function which takes one argument: the full tile path (including file extension) as a string.
        ' it returns the contents of the file as a string.
        ' in the instance below, the txt file is a HTML file, formatted with CSS, and the email will be sent with
        ' HTML structure and formatting.
        .HTMLBody = readTXTFile("D:\Users\r.green\Documents\HTML Emails\YourCardHasArrived.txt")
        .Subject = "[Collection Notification] - Your Citi Corporate " & colItem & " is ready for collection"
        .To = EmailAdd
    End With
    
    ' this will split the FilePathToAdd variable at each comma, and attach each file to the MailItem.
    If FilePathToAdd <> "" Then
        Attachments = Split(FilePathToAdd, ",")
        For i = LBound(Attachments) To UBound(Attachments)
            If Attachments(i) <> "" Then
                oMail.Attachments.Add Trim(Attachments(i))
            End If
        Next i
    End If
    
    ' this handles the option from the first MsgBox/userInput in 'Sub loopThruRecordSet()'
    ' it will display/send the email based on the user's preference
    If emailOpt = "Display" Then
        oMail.Display
    ElseIf emailOpt = "Send" Then
        oMail.Send
    Else
        MsgBox "An invalid command has been entered (somehow). Please press okay to exit the subroutine"
        Exit Sub
    End If
    
    Set oMail = Nothing
    Set oApp = Nothing

End Sub
Sub loopThruRecordSet()

    Dim rs As DAO.Recordset, rsName As String, EmailAdd As String, colItem As String, _
        emailCount As Integer, userInput As Variant, emailOpt As String
    
    rsName = "cr_InventoryUnnotified_q" 'enter desired recordset name
    Set rs = CurrentDb.OpenRecordset(rsName)
    emailCount = 0
    emailOpt = "Send"
    
    userInput = MsgBox("You are about to send " & rs.RecordCount & " emails. Do you wish to continue", vbYesNo, "Inventory UnNotified")
    
    If userInput = vbNo Then
        GoTo endOutput
    End If
    
    ' this input is passed into the send email sub procedure above.
    userInput = MsgBox("Would you like to view the emails before you send them?", vbYesNo, "Inventory Unnotified")
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            EmailAdd = rs!EmailAdd 'this line should reference the field in the query/table which
            'contains the email address you would like to send the email to
            colItem = rs!Item
            'The below code asks the user whether they would like to view the emails before they are sent
            'it then uses that response in the call to the email-send subroutine.
            If userInput = vbYes Then
                emailOpt = "Display"
                Call sendOutlookEmailunNote(EmailAdd, colItem, emailOpt)
            Else
                Call sendOutlookEmailunNote(EmailAdd, colItem, emailOpt)
            End If
            rs.Edit
            ' update the record with current date for future reference
            rs("Date Notified") = Date
            With rs
                .Update
                .MoveNext
            End With
            emailCount = emailCount + 1
        Loop
    Else
        MsgBox "No Records contained in Query", Title:="Inventory UnNotified"
    End If
    
    If emailCount > 0 Then
        MsgBox emailCount & " email(s) have been sent", Title:="Inventory UnNotified"
    End If
    
    Exit Sub
    
endOutput:
    MsgBox "No emails have been sent; the operation was cancelled", Title:="Inventory Unotified"
    Exit Sub

End Sub


