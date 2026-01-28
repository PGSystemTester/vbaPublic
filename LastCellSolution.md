## Find Last Cell
Finding the last row in Excel has been [fiercely debated](https://stackoverflow.com/questions/11169445/find-last-used-cell-in-excel-vba/59081657#59081657) in the Excel Community. However with latest updates in Excel including TRIMRANGE, this should be a settled issue.

## Excel

`=ROWS(TRIMRANGE(A:D,2))`


## VBA
```vb
Function lastRowWithVba(aRange As Range) As Long
    lastRowWithVba = Evaluate("=ROWS(TRIMRANGE('" & aRange.Worksheet.Name & "'!" & aRange.Address & ",2))")
End Function
```
