Attribute VB_Name = "ReadWholeTXT"
Option Compare Database
Option Explicit

Public Function readTXTFile(FileName As String) As String

'    this function takes on argument: the full file path, name, and file extension as a string, e.g:
'    "D:\Users\admin\Documents\koala_facts\koala_cuddles.txt"
'    it will return a string consisting of all of the text within the file

    Dim fileNo As Integer

    fileNo = FreeFile 'Get first free file number
     
    Open FileName For Input As #fileNo
    readTXTFile = Input$(LOF(fileNo), fileNo)
    Close #fileNo

End Function
