Attribute VB_Name = "sendEmails"
Public Sub sendOutlookMail()

    Dim subjectLine As String, emailAdd As String, filePathToAdd As String, pdfName As String
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

    On Error GoTo cleanup
    
    For Each cell In Range(Range("B6"), Range("B6").End(xlDown))
    
            subjectLine = cell.Offset(0, 5).Value
            pdfName = cell.Value
            
            fileSize = CheckFileSize(filePathToAdd, pdfName)
            If fileSize = True Then
            
                Set OutMail = OutApp.CreateItem(0)
                On Error Resume Next
                With OutMail
                    .To = emailAdd
                    .Subject = subjectLine
                    .Attachments.Add filePathToAdd & pdfName & ".pdf"
                End With
                
                
                OutMail.Display
                
                With Application
                    .Wait (Now + TimeValue("00:00:03"))
                    .SendKeys "%s"
                    .Wait (Now + TimeValue("00:00:03"))
                End With

                cell.Offset(0, 7).Value = "Sent"
                On Error GoTo 0
                Set OutMail = Nothing
            
            Else:
                userInput = MsgBox("File " & pdfName & " is larger than the 10MB limit.", _
                    Title:="10MB limit Exceeded")
                With cell.Offset(0, 6)
                    .Value = "File larger than 10MB!"
                    .Interior.ColorIndex = 3
                End With
            End If
            
        testCount = testCount + 1
        emailCount = emailCount + 1
        ' Application.Wait (Now + TimeValue("0:00:30"))
        ' Wend
    Next cell

    If emailCount = 1 Then
        MsgBox "1 email has been generated and sent.", Title:="Emails"
    ElseIf emailCount > 1 Then
        MsgBox emailCount & " emails have been generated and sent.", Title:="Emails"
    End If

cleanup:
    Set OutApp = Nothing
    Application.ScreenUpdating = True

End Sub
Public Function CheckFileSize(filePath As String, pdfName As String) As Boolean

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

    Dim userInput As Variant
    
    userInput = MsgBox("Are you sure you want to clear contents from this column?", vbYesNo, "Clear Contents")

    If userInput = vbYes Then
        With Range(Range("H6"), Range("I1000"))
            .ClearContents
            .ClearFormats
        End With
    Else
        Exit Sub
    End If

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