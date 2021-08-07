Public Function CountRows(ByRef nmRow As Long)
Dim firstCount As Long, finalCount As Long

        firstCount = Sheets("td").usedRange.Count
                'Counts row if data is present in the td spreadsheet
                If firstCount = 0 Then
                    MsgBox ("Please Enter Data")
                Else
                    finalCount = firstCount
                    finalCount = Sheets("td").Cells(firstCount, 1).End(xlUp).Row
                    nmRow = finalCount
                End If
        End Function