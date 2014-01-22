Attribute VB_Name = "Clipboard"
Option Explicit

Function CopyRangeToClipboard(ByVal Target As Range, Optional ColorConst = -1)

    Dim DataObj As New MSForms.DataObject
    Dim S As String
    
    If Not IsError(Target.Value) Then
        
        'If user has selected only one cell
        If Target.Cells.Count = 1 Then
            
            'Copy the string value to the clipboard.
            S = Target.Value
            DataObj.SetText S
            DataObj.PutInClipboard
        
            'Change the color of the cell to show that it's been copied.
            If ColorConst <> -1 Then
            
                Target.Cells.Font.Color = ColorConst
            
            End If
        
        End If
    
    End If
    
    'Clean up
    Set DataObj = Nothing

End Function

'Modularizing, not fully rolled into code yet.
Function CopyStringToClipboard(S As String)

    Dim DataObj As New MSForms.DataObject
    DataObj.SetText S
    DataObj.PutInClipboard

End Function
