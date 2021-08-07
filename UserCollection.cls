Option Explicit
Sub CollectionMembers()
Dim Members As New Collection, delMembers As New Collection
Dim destRow As Long, nmRow As Long, dCount As Long, iTwo As Long
Dim outRow As Long, delRow As Long
Dim Member As New MemberDS, delMember As New DMember

'Call to row numbers code in the RowCounter file
Call CountRows(nmRow)

'Turns off the workbooks ability to update the screen till this is changes to true/end of process
Application.ScreenUpdating = False

destRow = 1
outRow = 1
delRow = 2

        'Collecting member phonenumber & description 
        For iTwo = 1 To nmRow
            If InStr(1, LCase(Sheets("td").Range("A" & destRow)), "disabled", vbTextCompare) Or InStr(1, LCase(Sheets("td").Range("A" & destRow)), "old", vbTextCompare) Then

                'Send members to the del speadsheet that are to be deleted
                Set delMember = New DMember
                delMember.PhoneNumber = Sheets("td").Cells(destRow, 1).Value
                delMember.Desc = Sheets("td").Cells(destRow, 2).Value
                delMembers.Add delMember
                Sheets("D").Range("A" & delRow).Rows.Value = delMember.Desc
                Sheets("D").Range("B" & delRow).Rows.Value = Split(delMember.PhoneNumber, "(")
                delRow = delRow + 1
                dCount = dCount + 1
            Else

                'Send members to the Active speadsheet that are to be stored for coaching records
                Set Member = New MemberDS
                Member.PhoneNumber = Sheets("td").Cells(destRow, 1).Value
                Member.Desc = Sheets("td").Cells(destRow, 2).Value
                Members.Add Member
                Sheets("Ori").Range("A" & outRow).Rows.Value = Member.Desc
                Sheets("Ori").Range("B" & outRow).Rows.Value = Split(Member.PhoneNumber, "(")
                outRow = outRow + 1
            End If
            destRow = destRow + 1
        Next
        'displaying Active Member Count & deleting members
        MsgBox "Active Member Count: " & outRow
        MsgBox "Deleting Member Count: " & dCount

        'Turns off the workbooks ability to update the screen & this marks the end of the process
        Application.ScreenUpdating = True
'Call SPFeature, not implemented 
End Sub