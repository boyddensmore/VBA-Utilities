Attribute VB_Name = "Network"
Option Explicit

Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(WSADESCRIPTION_LEN) As Byte
    szSystemStatus(WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal _
   szName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal _
   wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias _
   "RtlMoveMemory" (Destination As Any, Source As Any, ByVal _
   Length As Long)
   
   
Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" _
(ByVal iDestIPAddr As Long, _
ByRef iHopCount As Long, _
ByVal iMaxHops As Long, _
ByRef iRTT As Long) As Long
 
Private Declare Function inet_addr Lib "wsock32.dll" _
(ByVal cp As String) As Long


Public Function LookupIPAddress(ByVal sHostName As String) As String

    Dim wsa As WSADATA
    Dim nRet As Long
    Dim nTemp As Long
    Dim bTemp(0 To 3) As Byte
    Dim sOut As String
    Dim he As HOSTENT
    
    'Initialize WinSock
    WSAStartup &H10, wsa
        
    'Attempt to lookup the host
    nRet = gethostbyname(sHostName)
    
    'If it failed, just return nothing
    If nRet = 0 Then
        sOut = ""
    Else
        'Take a look at the resulting hostent structure
        CopyMemory he, ByVal nRet, Len(he)
        
        'Are there atlest four bytes, then we have
        ' at least one address
        If he.h_length >= 4 Then
            'Copy the address out,
            CopyMemory nTemp, ByVal he.h_addr_list, 4
            CopyMemory bTemp(0), ByVal nTemp, 4
            ' and format it
            sOut = Format(bTemp(0)) & "." & Format(bTemp(1)) & "." _
               & Format(bTemp(2)) & "." & Format(bTemp(3))
        Else
            sOut = ""
        End If
        
    End If
    
    WSACleanup

    LookupIPAddress = sOut

End Function


Public Function Ping(sIPadr As String, iMaxHops As Long) As Boolean
     
     ' Based on an article on CodeGuru by Bill Nolde
     ' Implemented in VBA in Nov 2002 by G. Wirth, Ulm,  Germany
     
    Const SUCCESS   As Long = 1
     
    Dim iIPadr      As Long
    Dim iHopCount   As Long
    Dim iRTT        As Long
     
    iIPadr = inet_addr(sIPadr)
    Ping = (GetRTTAndHopCount(iIPadr, iHopCount, iMaxHops, iRTT) = SUCCESS)
     
    Debug.Print "IP Address ....... " & iIPadr & vbLf _
    & "HopCount ......... " & iHopCount & vbLf _
    & " Round trip, ms ... " & iRTT
End Function



Function IsConnectible(sHost, iPings, iTO)
   ' Returns True or False based on the output from ping.exe
   '
   ' Authors: Alex Angelopoulos/Torgeir Bakken
   ' Modified by: Tom Lavedas
   ' Works an "all" WSH versions
   ' sHost is a hostname or IP
   ' iPings is number of ping attempts
   ' iTO is timeout in milliseconds
   ' if values are set to "", then defaults below used

   Dim nRes
   If iPings = "" Then iPings = 1 ' default number of pings
   If iTO = "" Then iTO = 250     ' default timeout per ping
   With CreateObject("WScript.Shell")
     nRes = .Run("%comspec% /c ping.exe -n " & iPings & " -w " & iTO _
          & " " & sHost & " | find ""TTL="" > nul 2>&1", 0, True)
   End With
   IsConnectible = (nRes = 0)

End Function








