Attribute VB_Name = "RegistryAndUserTools"
Option Explicit
 'api call for obtaining the username
Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, _
nSize As Long)

Public Function WindowsUserName() As String
    Dim szBuffer As String * 100
    Dim lBufferLen As Long
     
    lBufferLen = 100
     
    If CBool(GetUserName(szBuffer, lBufferLen)) Then
         
        WindowsUserName = Left$(szBuffer, lBufferLen - 1)
         
    Else
         
        WindowsUserName = CStr(Empty)
         
    End If
     
End Function


Function GetRegValue(Appname As String, Section As String, Keyname As String) As String
    GetRegValue = GetSetting(Appname, Section, Keyname, "")
End Function


Function SetRegValue(Appname As String, Section As String, Keyname As String, Setting As String) As Integer

    On Error Resume Next
    SaveSetting Appname:=Appname, Section:=Section, key:=Keyname, Setting:=Setting
    
    SetRegValue = Err.Number
    
End Function
