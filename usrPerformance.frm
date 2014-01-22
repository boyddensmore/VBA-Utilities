VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrPerformance 
   Caption         =   "Performance test"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   OleObjectBlob   =   "usrPerformance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public MasterTimer As cTimer

Option Explicit

Public Sub test_the_tester()

Dim x As Integer

Set MasterTimer = New cTimer

    Process_Start
    
    For x = 1 To 10
    
        Chk1_Start
        Chk1_End
        Chk2_Start
        Chk2_End
        Chk3_Start
        Chk3_End
        Chk4_Start
        Chk4_End
        
    Next x
    
    Process_End

Set MasterTimer = Nothing

End Sub

Public Sub Process_Start()

    txtTotStart = Timer


End Sub

Public Sub Process_End()

    txtTotEnd = Timer

End Sub

Public Sub Chk1_Start()

    txtChk1ItStart = 0
   
    MasterTimer.StartCounter
    
End Sub

Public Sub Chk1_End()

    txtChk1ItEnd = MasterTimer.TimeElapsed

    txtChk1ItLength = Round(txtChk1ItEnd - txtChk1ItStart, 5)
    
    If txtChk1ItAvg.Text = "" Then
        txtChk1ItAvg = txtChk1ItLength
    Else
        txtChk1ItAvg = Round((CDbl(txtChk1ItAvg) + CDbl(txtChk1ItLength)) / 2, 6)
    End If
    
    Me.Repaint

End Sub

Public Sub Chk2_Start()
    
    txtChk2ItStart = 0
    
    MasterTimer.StartCounter
    
End Sub

Public Sub Chk2_End()

    txtChk2ItEnd = MasterTimer.TimeElapsed

    txtChk2ItLength = Round(txtChk2ItEnd - txtChk2ItStart, 5)
    
    If txtChk2ItAvg.Text = "" Then
        txtChk2ItAvg = txtChk2ItLength
    Else
        txtChk2ItAvg = Round((CDbl(txtChk2ItAvg) + CDbl(txtChk2ItLength)) / 2, 6)
    End If

    Me.Repaint

End Sub

Public Sub Chk3_Start()
    
    txtChk3ItStart = 0
    
    MasterTimer.StartCounter
    
End Sub

Public Sub Chk3_End()

    txtChk3ItEnd = MasterTimer.TimeElapsed
    
    txtChk3ItLength = Round(txtChk3ItEnd - txtChk3ItStart, 5)
    
    If txtChk3ItAvg.Text = "" Then
        txtChk3ItAvg = txtChk3ItLength
    Else
        txtChk3ItAvg = Round((CDbl(txtChk3ItAvg) + CDbl(txtChk3ItLength)) / 2, 6)
    End If
    
    Me.Repaint

End Sub

Public Sub Chk4_Start()
    
    txtChk4ItStart = 0
    
    MasterTimer.StartCounter
    
End Sub

Public Sub Chk4_End()

    txtChk4ItEnd = MasterTimer.TimeElapsed
    
    txtChk4ItLength = Round(txtChk4ItEnd - txtChk4ItStart, 5)
    
    If txtChk4ItAvg.Text = "" Then
        txtChk4ItAvg = txtChk4ItLength
    Else
        txtChk4ItAvg = Round((CDbl(txtChk4ItAvg) + CDbl(txtChk4ItLength)) / 2, 6)
    End If
    
    Me.Repaint

End Sub

Private Sub UserForm_Initialize()

    Set MasterTimer = New cTimer

End Sub
