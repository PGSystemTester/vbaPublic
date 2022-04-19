'This Function will unpivot data into a list of rows for each intersection of table
'Similar to reversing out of a pivot table to a flat file.
Function unPivotData(theDataRange As Range, theColumnRange As Range, theRowRange As Range, Optional skipZerosAsTrue As Boolean, Optional includeBlanksAsTrue As Boolean)


'Set effecient range
Dim cleanedDataRange As Range
    Set cleanedDataRange = Intersect(theDataRange, theDataRange.Worksheet.UsedRange)
   
'tests Data ranges

    'Use intersect address to account for users selecting full row or column
   If cleanedDataRange.EntireColumn.Address <> Intersect(cleanedDataRange.EntireColumn, theColumnRange).EntireColumn.Address Then
      unPivotData = "datarange missing Column Ranges"

   ElseIf cleanedDataRange.EntireRow.Address <> Intersect(cleanedDataRange.EntireRow, theRowRange).EntireRow.Address Then
      unPivotData = "datarange missing row Ranges"

   ElseIf Not Intersect(cleanedDataRange, theColumnRange) Is Nothing Then
      unPivotData = "datarange may not intersect column range.  " & Intersect(cleanedDataRange, theColumnRange).Address
      
   ElseIf Not Intersect(cleanedDataRange, theRowRange) Is Nothing Then
      unPivotData = "datarange may not intersect row range.  " & Intersect(cleanedDataRange, theRowRange).Address
   
   End If

   'exits if errors were found
   If Len(unPivotData) > 0 Then Exit Function
   
   Dim dimCount As Long
      dimCount = theColumnRange.Rows.Count + theRowRange.Columns.Count
   
   Dim aCell As Range, i As Long, g As Long
   ReDim newdata(dimCount, i)
   
'loops through data ranges
   For Each aCell In cleanedDataRange.Cells
   
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
