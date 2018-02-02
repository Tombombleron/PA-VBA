Attribute VB_Name = "crLimSave"
Public Sub creditLimitSave(MItem As Outlook.MailItem)

        Dim pathStr As String: pathStr = "N:\koalas\more_koalas\BarryKoala\Eucalyptus\2018\Dingo Limit\"
        Dim mailSubj As String: mailSubj = MItem.Subject
        Dim receivedTime As Date: receivedTime = MItem.SentOn
        Dim folderName As String: folderName = Format(receivedTime, "MM-YY")
        
        Call createFolder(pathStr, folderName)
        
        'On Error GoTo cleanup
        
        mailSubj = Replace(mailSubj, "£", "")
        mailSubj = Replace(mailSubj, "/", "")
        mailSubj = Replace(mailSubj, ":", "")
        mailSubj = Replace(mailSubj, "[", "")
        mailSubj = Replace(mailSubj, "]", "")
        
        MItem.SaveAs pathStr & folderName & "\" & mailSubj & ".msg", olMsg
        
Exit Sub
cleanup:
    MsgBox "Action Failed"

End Sub
Sub createFolder(pathStr As String, folderName As String)

    Dim finDir As String: finDir = pathStr & folderName

    If Len(Dir(finDir, vbDirectory)) = 0 Then
        MkDir finDir
    End If

End Sub
