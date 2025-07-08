'=====================================================================
' Helper – returns the sheet for a salesperson, creating & adding
' column headers on first use.
'=====================================================================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        'create a clean sheet at the end
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:= _
                                  ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName

        'copy header row from source table
        Dim srcHdr As Range
        Set srcHdr = ThisWorkbook.Worksheets("MainData").ListObjects("SalesData") _
                        .HeaderRowRange
        srcHdr.Copy Destination:=GetOrCreateSheet.Range("A1")
    End If
End Function
