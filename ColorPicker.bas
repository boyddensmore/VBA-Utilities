Attribute VB_Name = "ColorPicker"
Public Function PickColor() As Integer
    Dim oDialog As Object
    Dim PickedColor As Integer
    
    Set oDialog = CreateObject("MSComDlg.CommonDialog.1")
        With oDialog
           .ShowColor
           PickedColor = .Color
        End With
    Set oDialog = Nothing
    
    PickColor = PickedColor
End Function
