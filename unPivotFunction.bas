Function unPivotData(theDataRange As Range, theColumnRange As Range, theRowRange As Range, Optional includeZeros As Boolean, Optional includeBlanks As Boolean)

  If Intersect(theDataRange.EntireColumn, theColumnRange) Is Nothing Then
     unPivotData = "data/columns not aligned"
     Exit Function
  ElseIf Intersect(theDataRange.EntireRow, theRowRange) Is Nothing Then
     unPivotData = "data/rows not aligned"
     Exit Function
  End If

  'total dimensions
  Dim dimCount As Long
     dimCount = theColumnRange.Rows.Count + theRowRange.Columns.Count

  Dim aCell As Range, addMember As Boolean, i As Long, p As Long
  ReDim newdata(dimCount, i)

  For Each aCell In theDataRange.Cells

     If IsEmpty(aCell) And Not (includeBlanks) Then
        'skip
     ElseIf aCell.Value2 = 0 And Not (includeZeros) Then
        'skip
     Else
        ReDim Preserve newdata(dimCount, i)
        g = 0
        For Each gcell In Intersect(aCell.EntireColumn, theColumnRange).Cells
           newdata(g, i) = gcell.Value
           g = g + 1
        Next gcell

        For Each gcell In Intersect(aCell.EntireRow, theRowRange).Cells
           newdata(g, i) = gcell.Value
           g = g + 1
        Next gcell

        newdata(g, i) = aCell.Value
        i = i + 1
     End If
  Next aCell

  unPivotData = WorksheetFunction.Transpose(newdata)

End Function
