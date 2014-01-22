Attribute VB_Name = "Bounds"
Function SheetBoundaries(Sheet As String) As Variant
    Dim ReturnVal(1 To 2) As Variant
    ColArr = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")

    On Error Resume Next

    EndRow = Worksheets(Sheet).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    EndCol = Worksheets(Sheet).Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If EndRow = Empty Then EndRow = "Empty"
    If EndCol = Empty Then EndCol = "Empty"

    ReturnVal(1) = EndRow
    ReturnVal(2) = ColArr(EndCol - 1)

    SheetBoundaries = ReturnVal

End Function


Function LastRow(Sheet As String) As Integer

    On Error Resume Next

    EndRow = Worksheets(Sheet).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    If EndRow = Empty Then EndRow = -1

    Err.Clear

    LastRow = EndRow

End Function


Function LastCol(Sheet As String) As String
    
    ColArr = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")

    On Error Resume Next

    EndCol = Worksheets(Sheet).Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If EndCol = Empty Then
        EndCol = -1
        Err.Clear
        LastCol = EndCol
    Else
        LastCol = ColArr(EndCol - 1)
    End If

End Function
