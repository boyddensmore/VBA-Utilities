Attribute VB_Name = "Sorts"
Function SortStringArray(arrToSort() As String) As String()

    For lLoop = 1 To UBound(arrToSort)
       For lLoop2 = lLoop To UBound(arrToSort)
            If UCase(arrToSort(lLoop2)) < UCase(arrToSort(lLoop)) Then
                str1 = arrToSort(lLoop)
                str2 = arrToSort(lLoop2)
                arrToSort(lLoop) = str2
                arrToSort(lLoop2) = str1
            End If
        Next lLoop2
    Next lLoop

End Function


Function SortCollection(ToSort As Collection)

    'Two loops to bubble sort
    For i = 1 To ToSort.Count - 1
        For j = i + 1 To ToSort.Count
            If ToSort(i) > ToSort(j) Then
                'store the lesser item
                vTemp = ToSort(j)
                'remove the lesser item
                ToSort.Remove j
                're-add the lesser item before the
                'greater Item
                ToSort.Add vTemp, CStr(vTemp), i
            End If
        Next j
    Next i
   
End Function
