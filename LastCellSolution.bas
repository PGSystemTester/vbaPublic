Function FindLastRowInRange(someColumns As Range) As Long
   Const startFormula = "=IFERROR(MAX(FILTER(ROW(????),NOT(ISBLANK(????)))),0)"

   Dim zWS As Worksheet, zNameOfSheet As String, tangoRange As Range
      Set zWS = someColumns.Worksheet
      zNameOfSheet = "'" & zWS.Name & "'!"
      Set tangoRange = Intersect(someColumns, zWS.UsedRange)
   
   Dim i As Long, cLastROW As Long, aColumn As Range
   For i = 1 To tangoRange.Columns.Count
      Set aColumn = Intersect(tangoRange.Cells(1, i).EntireColumn, _
            zWS.UsedRange) 'narrows search to improve performance
            
      cLastROW = Evaluate(Replace(startFormula, "????", zNameOfSheet & aColumn.Address, 1, -1))
      
      If cLastROW > FindLastRowInRange Then FindLastRowInRange = cLastROW
   Next i

End Function


Function findLastRowInSheet(anywhereInSheet As Range) As Long
      Application.Volatile
      findLastRowInSheet = FindLastRowInRange(anywhereInSheet.Worksheet.UsedRange)
End Function
      
      
Function FindLastColumnInRange(someRows As Range) As Long
   Const startFormula = "=IFERROR(MAX(FILTER(Column(????),NOT(ISBLANK(????)))),0)"

   Dim zWS As Worksheet, zNameOfSheet As String, tangoRange As Range
      Set zWS = someRows.Worksheet
      zNameOfSheet = "'" & zWS.Name & "'!"
      Set tangoRange = Intersect(someRows, zWS.UsedRange)
   
   Dim i As Long, cLastColumn As Long, aRow As Range
   For i = 1 To tangoRange.Rows.Count
      Set aRow = Intersect(tangoRange.Cells(i, 1).EntireRow, _
            zWS.UsedRange) 'narrows search to improve performance
            
      cLastColumn = Evaluate(Replace(startFormula, "????", zNameOfSheet & aRow.Address, 1, -1))
      
      If cLastColumn > FindLastColumnInRange Then FindLastColumnInRange = cLastColumn
   Next i

End Function


Function findLastColumnInSheet(anywhereInSheet As Range) As Long
      Application.Volatile
      findLastColumnInSheet = FindLastColumnInRange(anywhereInSheet.Worksheet.UsedRange)
End Function

