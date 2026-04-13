# Save Selection as CSV

## About
- Captures current selection on Excel Worksheet and saves as CSV
- Offers prompt with unique date/ID
- Make sure to update the `defaultPath` constant to your 


## Code

```vba

Sub SaveSelectionAsCSV()
    Const defaultPath = "C:\updateThis\" '? adjust to whatever your preferred path is. Must end in a \
    If defaultPath = "C:\updateThis\" Then
        MsgBox "Change VBA parameter of Default path"
        Exit Sub
    End If
    

    Dim rng As Range, cell As Range, rowRange As Range
    Dim filePath As Variant, fileNum As Long
    Dim lineText As String, cellText As String, zNowText As String
    
    Const startFileName = "xxxxx-"
    
    zNowText = startFileName & WorksheetFunction.Text(Now, "yyyyMMdd""_""hhmmss")
    
    Set rng = Selection
    
    ' Prompt for save location with default folder
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultPath & zNowText, _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Save Selection As CSV")
    
    ' User cancelled
    If filePath = False Then Exit Sub
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Loop rows
    For Each rowRange In rng.Rows
        lineText = ""
        
        For Each cell In rowRange.Cells
            
            cellText = cell.Value
            
            ' Escape quotes
            cellText = Replace(cellText, """", """""")
            
            ' Wrap in quotes if needed
            If InStr(cellText, ",") > 0 Or InStr(cellText, """") > 0 Or InStr(cellText, vbLf) > 0 Then
                cellText = """" & cellText & """"
            End If
            
            If lineText = "" Then
                lineText = cellText
            Else
                lineText = lineText & "," & cellText
            End If
            
        Next cell
        
        Print #fileNum, lineText
    Next rowRange
    
    Close #fileNum
    
    MsgBox "CSV file created successfully.", vbInformation

End Sub
````
