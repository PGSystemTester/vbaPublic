'Function will test if Excel instance has spill range capability
Function testForSpill() As Boolean
    On Error GoTo nopE
        testForSpill = IsArray(Application.WorksheetFunction.Unique(Array("a", "b")))
    On Error GoTo 0
    
    Exit Function
nopE:
End Function
