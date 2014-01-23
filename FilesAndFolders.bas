Attribute VB_Name = "FilesAndFolders"
Function FolderExists(strPath) As Boolean
    If Len(Dir(strPath, vbDirectory)) = 0 Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

Function FileExists(FileName As String) As Boolean
    FileExists = (Dir(FileName) > "")
End Function


Function ReadTextFile(Fname As String, Length As Integer) As Variant
    
    If FileExists(Fname) Then
        Close #1
        
        Open Fname For Input As #1
        ReadTextFile = Input(Length, 1)
        Close 1
    Else
        ReadTextFile = False
    End If

End Function


Private Function openPipeSeparatedUTF8() As Workbook
'Opens a pipe-separated text file, enforcing UTF8 encoding and US English number separators
'Returns workbook object representing processed pipe-separated file
    Dim fn As String

    On Error Resume Next
    fn = Excel.Application.GetOpenFilename( _
         fileFilter:="Text Files (*.txt), *.txt,All Files (*.*),*.*", _
         title:="Open Pipe-Separated Report...")
    If fn <> "False" Then
        Excel.Workbooks.OpenText fileName:=fn, Origin:=msoEncodingUTF8, _
                                 DataType:=xlDelimited, TextQualifier:=xlTextQualifierNone, _
                                 ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, _
                                 Comma:=False, Space:=False, other:=True, OtherChar:="|", _
                                 DecimalSeparator:=".", ThousandsSeparator:=","
        Set openPipeSeparatedUTF8 = Excel.ActiveWorkbook
    End If
End Function