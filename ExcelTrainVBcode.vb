'Print out all members function for _dayone
Sub ButtonPrintall_Click()
Dim R As Range
    For Each R In Sheets("_members").Range("A2:A35")
        If Not IsEmpty(R) Then
            Range("I8") = R
                ActiveSheet.PrintOut
            End If
        Next
End Sub
'Print preview for the seleted sheet for _dayone
Sub ButtonPreview_Click()
        ActiveSheet.PrintOut preview:=True
End Sub
'Print preview for testing sheet
Sub ButtonPrintTesting_Click()
     ActiveSheet.PrintOut preview:=True
End Sub
'Print out full testing sheet
Sub ButtonPrintoutTest_Click()
    ActiveSheet.PrintOut
End Sub
