Attribute VB_Name = "sendEmails"
Public Sub sendOutlookMail()

    Dim subjectLine As String, emailAdd As String, filePathToAdd As String, pdfName As String, caller As String: caller = "sendOutlookMail"
    Dim fileSize As Boolean
    Dim OutApp As Object, OutMail As Object
    Dim testCount As Integer, emailCount As Integer, numSending As Integer
    Dim userInput As Variant
    
    ' for convenience of the user, the file path (as copied from file explorer) is pasted in cell K3
    filePathToAdd = Range("K3").Value & "\"
    ' as above, the email address is entered in K6, multiple emails would need to be entered
    ' as they would in outlook: "abc@gmail.com; xyz@gmail.com"
    emailAdd = Range("K6").Value
    testCount = 0
    emailCount = 0
    
    If Range("B6").Value = "" Then
        MsgBox "BAIL!"
        Exit Sub
    End If
    
    numSending = Range(Range("B6"), Range("B6").End(xlDown)).Cells.SpecialCells(xlCellTypeConstants).Count
    ' prompt to ensure the user knows what they're about to do
    userInput = MsgBox("Are you sure you want to generate " & numSending & " emails?", vbYesNo, "Generation Confirmation")
    
    If userInput = vbNo Then
        MsgBox "Operation Cancelled"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set OutApp = CreateObject("Outlook.Application")

    ' cleanup will inform the user that there has been an error via messagebox
    ' and prompt them to contact me.
    ' it will also send me an email directly informing me of the error number, description,
    ' user running the script, and the time the error occurred.
    On Error GoTo cleanup
    
    For Each cell In Range(Range("B6"), Range("B6").End(xlDown))
    
            subjectLine = cell.Offset(0, 5).Value
            pdfName = cell.Value
            
            ' fileSize returns a boolean: True if file<10Mb, False if file>10MB
            fileSize = CheckFileSize(filePathToAdd, pdfName)
            If fileSize = True Then
            
                Set OutMail = OutApp.CreateItem(0)
                With OutMail
                    .To = emailAdd
                    .Subject = subjectLine
                    .Attachments.Add filePathToAdd & pdfName & ".pdf"
                End With
                
                
                OutMail.Display
                
                ' wait for 3 seconds initially to ensure that the file
                ' is attached from the network before the shortcut is sent.
                ' then send "Alt+S" (outlook shortcut for send message)
                ' to mimic the user pressing the send button
                With Application
                    .Wait (Now + TimeValue("00:00:03"))
                    .SendKeys "%s"
                    .Wait (Now + TimeValue("00:00:03"))
                End With

                cell.Offset(0, 7).Value = "Sent"
                Set OutMail = Nothing
            
            Else:
                userInput = MsgBox("File " & pdfName & " is larger than the 10MB limit.", _
                    Title:="10MB limit Exceeded")
                ' highlights the files which are too large to be sent via email
                ' to ensure that the user is able to identify and send each file manually.
                With cell.Offset(0, 6)
                    .Value = "File larger than 10MB!"
                    .Interior.ColorIndex = 3
                End With
                cell.Offset(0, 5).Interior.ColorIndex = 6
            End If
            
        testCount = testCount + 1
        emailCount = emailCount + 1

    Next cell

    If emailCount = 1 Then
        MsgBox "1 email has been generated and sent.", Title:="Emails"
    ElseIf emailCount > 1 Then
        MsgBox emailCount & " emails have been generated and sent.", Title:="Emails"
    End If
    
    ' successemail will notify me by email that the script has run successfully.
    ' partly for peace-of-mind, also for self-satisfaction. ¯\_("_/)_/¯
    SuccessEmail caller, Application.UserName, emailCount
    Set OutApp = Nothing
    Application.ScreenUpdating = True
    Exit Sub

cleanup:
    Set OutApp = Nothing
    MsgBox "Operation has failed" & vbNewLine & "Please contact me@helppls.com for help."
    CustomErrorHandler caller, Application.UserName
    Application.ScreenUpdating = True

End Sub
Public Function CheckFileSize(filePath As String, pdfName As String) As Boolean

    ' this function will return takes two arguments:
        ' filePath: the path to the file (including "\" at the end)
        ' pdfName: the name of the file without the file extension
    ' it returns a boolean "True" if the file is <10MB or "False" if it is >10MB
    
    Dim result As Boolean
    Dim fileLength As Double, fileLim As Double
    Dim fileName As String

    fileLimit = 10000000
    result = False
    fileName = filePath & pdfName & ".pdf"
    fileLength = FileLen(fileName)
    
    If fileLength < fileLimit Then
        CheckFileSize = True
    Else
        CheckFileSize = False
    End If

End Function
Sub CleanSentCol()

    ' this will remove contents and clear formats from columns H and I
    ' and clear formats column g.
    ' if there are no values in any of those columns then an error will
    ' be raised and I will be informed via email that the user has attmpted
    ' to do this.

    Dim userInput As Variant
    Dim caller As String: caller = "CleanSentCol"
    On Error GoTo errorEmail
    
    userInput = MsgBox("Are you sure you want to clear contents from this column?", vbYesNo, "Clear Contents")

    If userInput = vbYes And Range(Range("H6"), Range("i6").End(xlDown)).Cells.SpecialCells(xlCellTypeConstants).Count > 0 Then
        With Range(Range("H6"), Range("I1000"))
            .ClearContents
            .ClearFormats
        End With
        Range(Range("G6"), Range("G6").End(xlDown)).Interior.ColorIndex = 0
    Else
        Err.Raise 3100, Description:="User has tried to clear columns H and I in " & ThisWorkbook.Name & " and the columns are blank." & _
            vbNewLine & vbNewLine & "No action is required on your part if the user does not ask for help."
    End If
    
    Exit Sub

errorEmail:
    CustomErrorHandler caller, Application.UserName
    Exit Sub

End Sub
Sub splitTextToCols()

    Dim splitArray() As String
    Dim rangeToSplit As Range
    
    Application.ScreenUpdating = False
    
    Set rangeToSplit = Range(Range("B6"), Range("B6").End(xlDown))
    
    For Each cell In rangeToSplit
        
        If InStr(1, cell, ",") <= 0 Then
            MsgBox "No Comma found in string in cell " & cell.Address & ". Please ensure all strings are CSV before trying this again."
            Exit Sub
        End If
        
        splitArray = Split(cell, ",")
        For Each itemVal In splitArray
            cell.Value = splitArray(0)
            cell.Offset(0, 1).Value = splitArray(1)
            cell.Offset(0, 2).Value = splitArray(2)
        Next itemVal
    Next cell

    Application.ScreenUpdating = True

End Sub
Sub CustomErrorHandler(caller As String, user As String)

    ' this is called whenever an error occurs in the main sub procedures.
    ' it takes two arguments:
    ' caller: the name of the sub-procedure that was running
    ' user: the user whose environment the script was running in
    ' it will send me an email with details of the error and the time it was raised.

    Dim oApp As Outlook.Application: Set oApp = CreateObject("Outlook.Application")
    Dim oMail As MailItem: Set oMail = CreateItem(olMailItem)
    
    With oMail
        .To = "me@helppls.com"
        .Subject = "VBA Error | Workbook: " & ThisWorkbook.Name & " | Error Number: " & Err.Number
        .Body = "The following error has been raised in sub " & caller & " whilst user " & user & " was running the procedure." & _
            vbNewLine & vbNewLine & _
            "This error was raised at " & Now & "." & _
            vbNewLine & vbNewLine & _
            "----------ERROR DESCRIPTION----------" & _
            vbNewLine & vbNewLine & _
            Err.Description & _
            vbNewLine & vbNewLine & _
            "------------------------------------------------"
            .Importance = olImportanceHigh
            .Send
    End With
    
    Set oMail = Nothing
    Set oApp = Nothing
    
End Sub
Sub SuccessEmail(caller As String, user As String, emailNum As Integer)

    ' this function will always run at the end of the sendOutlookMail sub procedure
    ' it takes three arguments:
    ' caller: name of the subprocedure which called this sub procedure
    ' user: user who is running the script
    ' emailNum: the number of emails sendOutlookMail generated and sent
    ' it will send me an email detailing who ran the script and the number of emails which were sent.

    Dim oApp As Outlook.Application: Set oApp = CreateObject("Outlook.Application")
    Dim oMail As MailItem: Set oMail = CreateItem(olMailItem)
    
    With oMail
        .To = "me@helppls.com"
        .Subject = "VBA Success | Workbook: " & ThisWorkbook.Name
        .Body = "User " & user & " has used " & caller & " sub-procedure in " & ThisWorkbook.Name & _
            " to successfully send " & emailNum & " email(s) at " & Now & "."
        .Send
    End With
    
    Set oMail = Nothing
    Set oApp = Nothing
        
End Sub
