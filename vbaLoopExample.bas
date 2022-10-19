Sub LoopAllSheetsExample()
  'this will loop through all sheets
  Dim ws As Worksheet, zAnswer As Long

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
         zAnswer = MsgBox("This is worksheet " & ws.Name & ". It has a usedRange of:" & ws.UsedRange.Address & _
            ". Continue looping through worksheets?", vbYesNo)
    
        If zAnswer <> vbYes Then Exit For    
    Next ws
End Sub
