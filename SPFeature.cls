Sub SPFeature()
Dim objectRange As Range
Dim wSO As Worksheet
Dim nmRow As Long
Dim MemberObject As ListObject

'destation of the callout location from the logic/main program file
Call CountRows(nmRow)

If nmRow = 0 Then
MsgBox ("Please Enter Data")
Else
Set wSO = Sheets("Ori")
Set objectRange = wSO.Range("A1").CurrentRegion
Set MemberObject = wSO.ListObjects.Add(SourceType:=xlSrcRange, Source:=objectRange)
End If

MsgBox MemberObject.ListRows.Count

'MemberObject.Range.Select
ChDir "C:\Members\Matthew\Documents"
Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Members\Matthew\Documents\" & "List.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

Range("G16").Select

End Sub