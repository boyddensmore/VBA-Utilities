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
