Attribute VB_Name = "StrToDate"
Function StrToDate(inDate As String) As Date

    StrToDate = Format$(Format$(inDate, "##/##/##"), "mm/dd/yy")

End Function
