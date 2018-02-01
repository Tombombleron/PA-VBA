Attribute VB_Name = "savemail"
Public Sub SaveACopy(Item As Outlook.MailItem)
    Const olMsg As Long = 3

    Dim m As MailItem
    Dim savePath As String, mailSubject As String, mailBody As String
    Dim mailReceived As Date

    If TypeName(Item) <> "MailItem" Then Exit Sub

    Set m = Item

    savePath = "D:\Users\admin\Documents\koala_facts\"  '## Modify as needed
    savePath = savePath & m.Subject & " " & Format(Now(), "yyyy-mm-dd-hhNNss")
    savePath = savePath & ".msg"


    m.SaveAs savePath, olMsg

    mailSubject = m.Subject
    mailReceived = m.SentOn
    mailBody = m.Body
    
    Call addToExcel(mailSubject, mailReceived, mailBody)

End Sub
Public Sub addToExcel(mailSubject As String, mailReceived As Date, mailBody As String)

    Dim xlApp As Object
    Dim xlWB As Excel.Workbook
    Dim xllWS As Excel.Worksheet
    Dim fileStr As String: fileStr = "D:\Users\admin\Documents\koala_facts\koala_recovery_master.xlsx"
    Dim subjectCell As Range
    
    On Error GoTo cleanUp
    
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    With xlApp
        .Visible = False
        .EnableEvents = False
    End With
    
    Set xlWB = Workbooks.Open(fileStr, ReadOnly:=False, Editable:=True)
    
    If xlWB.Worksheets(Sheets.Count).Name = (Month(Date) & "_" & Year(Date)) Then
        Set xlWS = xlWB.Worksheets(Sheets.Count)
    Else
        Set xlWS = xlWB.Sheets.Add(After:=xlWB.Sheets(xlWB.Sheets.Count))
        With xlWS
            .Name = Month(Date) & "_" & Year(Date)
            .Range("A1").Value = "Subject Line"
            .Range("D1").Value = "Body Conent"
            .Range("B1").Value = "Sent on"
            .Range("C1").Value = "Date Entered"
            .Rows(1).Font.Bold = True
        End With
    End If
    
    xlWB.Activate
    
    Set subjectCell = xlWS.Range("A1000").End(xlUp).Offset(1, 0)
    With subjectCell
        .Value = mailSubject
        .Offset(0, 3).Value = mailBody
        .Offset(0, 1).Value = mailReceived
        .Offset(0, 2).Value = Date
    End With
    
    With xlWS
        .Columns("A:C").AutoFit
        .Rows.RowHeight = 14.4
    End With
    
cleanUp:
    With xlWB
        .Save
        .Close
    End With
    xlApp.Quit
    Set xlApp = Nothing

End Sub
