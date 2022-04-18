'This Function will unpivot data into a list of rows for each intersection of table
'Similar to reversing out of a pivot table to a flat file.
Function unPivotData(theDataRange As Range, theColumnRange As Range, theRowRange As Range, Optional skipZerosAsTrue As Boolean, Optional includeBlanksAsTrue As Boolean)
   
'tests Data ranges
'okay to have entire rows/columns for axis members
   If theDataRange.EntireColumn.Address <> Intersect(theDataRange.EntireColumn, theColumnRange).EntireColumn.Address Then
      unPivotData = "datarange missing Column Ranges"

   ElseIf theDataRange.EntireRow.Address <> Intersect(theDataRange.EntireRow, theRowRange).EntireRow.Address Then
      unPivotData = "datarange missing row Ranges"

   ElseIf Not Intersect(theDataRange, theColumnRange) Is Nothing Then
      unPivotData = "datarange may not intersect column range.  " & Intersect(theDataRange, theColumnRange).Address
      
   ElseIf Not Intersect(theDataRange, theRowRange) Is Nothing Then
      unPivotData = "datarange may not intersect row range.  " & Intersect(theDataRange, theRowRange).Address
   
   End If

   'exits if errors were found
   If Len(unPivotData) > 0 Then Exit Function
   
   Dim dimCount As Long
      dimCount = theColumnRange.Rows.Count + theRowRange.Columns.Count
   
   Dim aCell As Range, i As Long, g As Long
   ReDim newdata(dimCount, i)
   
'loops through data ranges
   For Each aCell In theDataRange.Cells
   
      If aCell.Value2 = "" And Not (includeBlanksAsTrue) Then
         'skip
      ElseIf aCell.Value2 = 0 And skipZerosAsTrue Then
         'skip
      Else
         ReDim Preserve newdata(dimCount, i)
         g = 0
         
      'gets DimensionMembers members
         For Each gcell In Union(Intersect(aCell.EntireColumn, theColumnRange), _
            Intersect(aCell.EntireRow, theRowRange)).Cells
               
            newdata(g, i) = IIf(gcell.Value2 = "", "", gcell.Value)
            g = g + 1
         Next gcell
      
         newdata(g, i) = IIf(aCell.Value2 = "", "", aCell.Value)
         i = i + 1
      End If
   Next aCell
            
   unPivotData = WorksheetFunction.Transpose(newdata)

End Function
