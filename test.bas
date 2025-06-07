Sub DeleteRowsWithFutureInvoiceDates()
    Dim ws As Worksheet
    Dim cell As Range
    Dim invoiceCol As Long
    Dim lastRow As Long
    Dim maxDate As Date
    Dim searchRange As Range
    Dim i As Long

    ' Get the maxDate value from the named range
    On Error GoTo NoMaxDate
    maxDate = Range("maxDate").Value
    On Error GoTo 0

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) Like "*invoices*" Then
            Set searchRange = ws.UsedRange

            ' Find "Invoice Date" in the used range
            For Each cell In searchRange
                If Trim(LCase(cell.Value)) = "invoice date" Then
                    invoiceCol = cell.Column
                    lastRow = ws.Cells(ws.Rows.Count, invoiceCol).End(xlUp).Row

                    ' Loop from bottom to top to avoid skipping rows after deletion
                    For i = lastRow To cell.Row + 1 Step -1
                        If IsDate(ws.Cells(i, invoiceCol).Value) Then
                            If ws.Cells(i, invoiceCol).Value > maxDate Then
                                ws.Rows(i).Delete
                            End If
                        End If
                    Next i

                    Exit For ' Done once we find the first "Invoice Date"
                End If
            Next cell
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "Rows with future invoice dates removed.", vbInformation
    Exit Sub

NoMaxDate:
    MsgBox "Named range 'maxDate' not found or invalid.", vbCritical
End Sub
