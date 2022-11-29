Attribute VB_Name = "mdlFunctions"
''Option Explicit

Public Function SelFileDialog(ByVal docPath As String) As String
'--------------------------------------------------------------------
' This function Select a file via Dialog box
'--------------------------------------------------------------------
Dim fd As Office.FileDialog

SelFileDialog = ""

Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .Filters.Clear      ' Clear all the filters (if applied before).
    .Title = "Select a Word File"
    .Filters.Add "Word Files", "*.doc?", 1
    .AllowMultiSelect = False
    .Show
    If .SelectedItems.Count = 0 Then
        Exit Function
    Else
        SelFileDialog = .SelectedItems(1)
    End If
End With

End Function

Function GetFolder() As String
'--------------------------------------------------------------------
' This function Select a Folder via Dialog box
'--------------------------------------------------------------------

Dim fldr As FileDialog
Dim sItem As String

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = "" ''Application.DefaultFilePath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With

NextCode:
GetFolder = sItem
Set fldr = Nothing

End Function




Function IsFileOpen(fileName As String)
'--------------------------------------------------------------------
' This function check if a non-Shared file is already open
'--------------------------------------------------------------------
Dim ff As Long, ErrNo As Long
On Error Resume Next
ff = FreeFile()
Open fileName For Input Lock Read As #ff
Close ff
ErrNo = Err
On Error GoTo 0

Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
End Select

End Function

Sub PrintLog(logFile, argument)
    If logFile.Line = 1 Then
        logFile.WriteLine argument
    Else
        logFile.WriteLine logFile.Line - 1 & Chr(9) & Now() & Chr(9) & argument
    End If

End Sub


Function DocGetCoverInfo(oDoc)

'--------------------------------------------------------------------
' This function extracts Cover information: title ans subtitle
'--------------------------------------------------------------------
Dim oHeader, oRange As Range

Dim n, ltitle As Long
Dim LogoHeight, MinLogoHeight As Long
Dim s, minFS, maxFS As Integer
Dim strTitleInfo, strSubTitleInfo, Auxstr As String
Dim GetCoverPagePar

minFS = 200
maxFS = 0
strTitleInfo = ""
strSubTitleInfo = ""
GetCoverPagePar = GetCoverPageRange(oDoc)

 
If GetCoverPagePar(0) = 0 And GetCoverPagePar(1) = 0 Then
    GoTo Exit_DocGetCoverInfo '---> It was not found Cover page
End If

Set oRange = oDoc.Range(oDoc.Paragraphs(GetCoverPagePar(0)).Range.Start, oDoc.Paragraphs(GetCoverPagePar(1)).Range.End)
oRange.Select
'--------------------------------------------------------------------
'--- Get Font Size info ---
For n = 1 To oRange.Paragraphs.Count
    If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size >= maxFS Then
        If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size < 9999999 Then
            maxFS = oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size
        End If
    End If

    If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size <= minFS Then
        If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size < 9999999 Then
            minFS = oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size
        End If
    End If
Next n

'--------------------------------------------------------------------
'--- Get Title info ---
strTitleInfo = ""
For n = 1 To oRange.Paragraphs.Count
    If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size = maxFS And _
        oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Tables.Count = 0 Then
        If Trim(Replace(Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), ""), "", "")) <> "" Then
            If strTitleInfo = "" Then
                strTitleInfo = Replace(Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), ""), "", "")
            Else
                strTitleInfo = strTitleInfo & Chr(13) & Replace(Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), ""), "", "")
            End If
        End If
    ElseIf n = oRange.Paragraphs.Count And strTitleInfo = "" Then
        strTitleInfo = Replace(oDoc.Range(oRange.Paragraphs(1).Range.Start, oRange.Paragraphs(1).Range.End).Text, Chr(13), "")
    End If
Next n

'strTitleInfo = Replace(strTitleInfo, Chr(11), Chr(13))
If Len(strTitleInfo) > 200 Then
    strTitleInfo = Split(strTitleInfo, Chr(13))(0)
End If

If Len(strTitleInfo) > 200 Then
    strTitleInfo = Left(strTitleInfo, 255)
End If

'--------------------------------------------------------------------
'--- Get subtitle info: Method 1 ---
For n = 1 To oRange.Paragraphs.Count
    oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Select
    If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Tables.Count = 0 Then
        If maxFS = minFS Then
            Auxstr = Trim(Replace(Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), ""), "", ""))
        

            If Auxstr <> "" And Auxstr <> strTitleInfo And InStr(1, Replace(strTitleInfo, " ", ""), Replace(Auxstr, " ", "")) = 0 Then
                 If strSubTitleInfo = "" Then
                     strSubTitleInfo = Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), "")
                 Else
                     strSubTitleInfo = strSubTitleInfo & Chr(13) & Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), "")
                 End If
             End If


        Else

oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Select
            If oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size < maxFS And _
               oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size > minFS Then
    
               If Trim(Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), "")) <> "" And _
                  Len(Trim(Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), ""))) > 1 Then
    
                    
                    If strSubTitleInfo = "" Then
                        strSubTitleInfo = Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), "")
                    Else
                        strSubTitleInfo = strSubTitleInfo & Chr(13) & Replace(oDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Text, Chr(13), "")
                    End If
                End If
            End If

        End If
    End If
Next n


'--------------------------------------------------------------------
Exit_DocGetCoverInfo:
If strTitleInfo <> "" Then
    DocGetCoverInfo = Array(strTitleInfo, strSubTitleInfo)
Else
    DocGetCoverInfo = Array("|--> Not Found <--|")
End If

End Function

Function GetCoverPageRange(oDoc)
'--------------------------------------------------------------------
' This function determines Cover page range
'--------------------------------------------------------------------
Dim MinLogoHeight, LogoHeight As Long
Dim lgPage As Long
Dim h, s, startPara, n As Integer
Dim strTitleInfo, strSubTitleInfo As String
Dim oRange As Range
Dim oHeader

'--------------------------------------------------------------------
'--- Defaults ---
MinLogoHeight = 1000000 '--> Minimun first page Header Picture Header

strTitleInfo = ""
strSubTitleInfo = ""



'--------------------------------------------------------------------
'--- Get header logo height ---
For s = 1 To oDoc.Sections.Count
    For h = 1 To oDoc.Sections(s).Headers.Count
        For Shp = 1 To oDoc.Sections(s).Headers(h).Shapes.Count
            If oDoc.Sections(s).Headers(h).Shapes(Shp).Type <> msoAutoShape Then
                If MinLogoHeight > Application.PointsToCentimeters(oDoc.Sections(s).Headers(h).Shapes(Shp).Height) Then
                    MinLogoHeight = oDoc.Sections(s).Headers(h).Shapes(s).Height ''Application.PointsToCentimeters(oDoc.Sections(s).Headers(h).Shapes(s).Height)
                End If
            End If
        Next Shp
    Next h
Next s
If MinLogoHeight = 1000000 Then MinLogoHeight = 5

'--------------------------------------------------------------------
'--- Get if  1st page is Cover by header pic ---
LogoHeight = 0
For h = 1 To oDoc.Sections(1).Headers.Count
    Set oHeader = oDoc.Sections(1).Headers(h)
    With oHeader
        For s = 1 To .Shapes.Count
            If Application.PointsToCentimeters(.Shapes(s).Height) > LogoHeight Then
                LogoHeight = Application.PointsToCentimeters(.Shapes(s).Height)
            End If
        Next s


        If LogoHeight < MinLogoHeight Then
'            GoTo Exit_GetCoverPageRange '---> It was not found Cover page
            h = 1
            Exit For
        Else
            Exit For
        End If
    End With
Next h

'--------------------------------------------------------------------
'--- Select Cover page range: Method 1 ---
startPara = 1
Covertype1 = False

If oDoc.Sections(1).Footers(1).Shapes.Count > 0 And oDoc.Sections(1).Headers(1).Shapes.Count > 0 Then
    ShpF = 1
    Do While ShpF <= oDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Shapes.Count
        If oDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Shapes(ShpF).Type = msoPicture And _
           oDoc.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.Count = 0 Then
            For ShpH = 1 To oDoc.Sections(1).Headers(wdHeaderFooterFirstPage).Shapes.Count
                If oDoc.Sections(1).Headers(wdHeaderFooterFirstPage).Shapes(ShpH).Type = msoPicture Then
                    Covertype1 = True
                    Exit Do
                End If
            Next ShpH
        End If
        ShpF = ShpF + 1
    Loop
End If

If Covertype1 = True Then
    For n = 1 To oDoc.Paragraphs.Count - 1
   
    oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(n).Range.End).Select

        If oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(n).Range.End).Information(wdActiveEndAdjustedPageNumber) > 1 Then
            n = n - 1
            Set oRange = oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(n).Range.End)
            oRange.Select
            Exit For
        ElseIf n = oDoc.Paragraphs.Count - 1 And oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(n).Range.End).Information(wdActiveEndAdjustedPageNumber) = 1 Then
            Set oRange = oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(n).Range.End)
            oRange.Select
        End If
    Next n
ElseIf LogoHeight < MinLogoHeight Then
    For n = 1 To oDoc.Paragraphs.Count - 1
        If Replace(Replace(oDoc.Range(oDoc.Paragraphs(n).Range.Start, oDoc.Paragraphs(n).Range.End).Text, Chr(13), ""), " ", "") <> "" And _
        Replace(Replace(oDoc.Range(oDoc.Paragraphs(n).Range.Start, oDoc.Paragraphs(n).Range.End).Text, Chr(13), ""), " ", "") <> "/" Then
            Set oRange = oDoc.Range(oDoc.Paragraphs(n).Range.Start, oDoc.Paragraphs(n).Range.End)
            oRange.Select
            Exit For
        End If
    Next n

    
Else
    For n = 0 To oDoc.Paragraphs.Count - 1
        Set oRange = oDoc.Paragraphs(1).Next(Count:=n).Range
        lgPage = oRange.Information(wdActiveEndAdjustedPageNumber)
        If lgPage > h Then
            Set oRange = Nothing
            Set oRange = oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(n).Range.End)
            oRange.Select
            Exit For
        End If
    Next n
End If

'--------------------------------------------------------------------

Exit_GetCoverPageRange:
If Not oRange Is Nothing Then
    GetCoverPageRange = Array(1, n)
Else
'    GetCoverPageRange = Array(1, oDoc.Paragraphs.Count)
    GetCoverPageRange = Array(0, 0)
End If

End Function

Function GetContentTablePar(oDoc, AuxDoc, Optional OAPercent)
'--------------------------------------------------------------------
' This function determines Table of content range
'--------------------------------------------------------------------
Dim oRange  As Range
Dim Auxstr As String
Dim p, lgPage, n As Long
Dim startPara, opt As Integer
Dim CTOptionArray

GetContentTablePar = Array(0, 0)

'--------------------------------------------------------------------
'--- Select Cover page range ---
startPara = 1
lgPage = 1
n = 0

Do While n < AuxDoc.Paragraphs.Count

    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(n / AuxDoc.Paragraphs.Count, 4), True, "Preparing Data. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(n / AuxDoc.Paragraphs.Count, 4), True, "Preparing Data. Progress..."
        DoEvents
    End If
    '------------------------------
    
    
    Set oRange = AuxDoc.Paragraphs(startPara).Next(Count:=n).Range
    oRange.Select

    If oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC1) Or oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC2) Or _
       oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC3) Or oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC4) Or _
       oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC5) Or oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC6) Or _
       oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC7) Or oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC8) Or _
       oRange.Paragraphs.Style = AuxDoc.Styles(wdStyleTOC9) Then
        Exit Do
    End If
    
    If oRange.Information(wdActiveEndAdjustedPageNumber) > 10 Then
        GoTo Exit_GetContentTablePar
    End If

    n = n + 1
Loop

If n = AuxDoc.Paragraphs.Count Then
    GoTo Exit_GetContentTablePar
End If

'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 4), True, "Preparing Data. Progress..."
    DoEvents
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 4), True, "Preparing Data. Progress..."
    DoEvents

End If
'------------------------------

startPara = n
n = 1
Do While n < AuxDoc.Paragraphs.Count
    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(n / AuxDoc.Paragraphs.Count, 4), True, "Finding Table of Contents. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(n / AuxDoc.Paragraphs.Count, 4), True, "Finding Table of Contents. Progress..."
        DoEvents
    End If
    '------------------------------
    
    Set oRange = AuxDoc.Paragraphs(startPara).Next(Count:=n).Range
    oRange.Select
    If Replace(oRange.Text, Chr(13), "") <> "" Then
        If oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC1) And oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC2) And _
           oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC3) And oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC4) And _
           oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC5) And oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC6) And _
           oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC7) And oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC8) And _
           oRange.Paragraphs.Style <> AuxDoc.Styles(wdStyleTOC9) Then
                If startPara <= 1 Then
                    GetContentTablePar = Array(startPara, startPara + n)
                Else
                    GetContentTablePar = Array(startPara - 1, startPara + n)
                End If
                GoTo Exit_GetContentTablePar
        End If
    End If
    
    n = n + 1
Loop


    
'--------------------------------------------------------------------
'--- Select TOC manual style ---
CTOptionArray = Array("Table of Content", "Table of Contents", "Inhaltsverzeichnis", "Inhaltsverzeich")
startPara = 1
lgPage = 1
n = 0

Do While n < AuxDoc.Paragraphs.Count
        '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(n / AuxDoc.Paragraphs.Count, 4), True, "Preparing Data. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(n / AuxDoc.Paragraphs.Count, 4), True, "Preparing Data. Progress..."
        DoEvents
    End If
    '------------------------------
    
    Set oRange = AuxDoc.Paragraphs(startPara).Next(Count:=n).Range
    oRange.Select
    If lgPage < oRange.Information(wdActiveEndAdjustedPageNumber) Then
'        lgPage = oRange.Information(wdActiveEndAdjustedPageNumber)
        
        Auxstr = Replace(Replace(Replace(AuxDoc.Range(oRange.Paragraphs(1).Range.Start, oRange.Paragraphs(1).Range.End).Text, "", ""), Chr(13), ""), "", "")

        For opt = 0 To UBound(CTOptionArray)
            If Trim(LCase(Replace(Auxstr, " ", ""))) = Trim(LCase(Replace(CTOptionArray(opt), " ", ""))) Then

                Exit Do

            End If
        Next opt
    End If
    If oRange.Information(wdActiveEndAdjustedPageNumber) > 10 Then
        GoTo Exit_GetContentTablePar
    End If

    n = n + 1
Loop

If n = AuxDoc.Paragraphs.Count Then
    GoTo Exit_GetContentTablePar
End If

    
    '------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 4), True, "Preparing Data. Progress..."
    DoEvents
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 4), True, "Preparing Data. Progress..."
    DoEvents
End If
'------------------------------

startPara = n + 2
n = 1
Do While n < AuxDoc.Paragraphs.Count
    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(n / AuxDoc.Paragraphs.Count, 4), True, "Finding Table of Contents. Progress..."
        DoEvents
    
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(n / AuxDoc.Paragraphs.Count, 4), True, "Finding Table of Contents. Progress..."
        DoEvents
    End If
    '------------------------------
    Set oRange = AuxDoc.Paragraphs(startPara).Next(Count:=n).Range
    oRange.Select
    If Replace(oRange.Text, Chr(13), "") <> "" Then
          If IsNumeric(oRange.Paragraphs(1).Range.Words(1)) = False And _
             IsNumeric(oRange.Paragraphs(1).Range.Words(oRange.Paragraphs(1).Range.Words.Count - 1)) = False Then
                GetContentTablePar = Array(startPara, startPara + n)
                GoTo Exit_GetContentTablePar
          End If
    End If
    
    
    n = n + 1
Loop



'--------------------------------------------------------------------
Exit_GetContentTablePar:
'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 4), True, "Finding Table of Contents. Progress..."
    DoEvents
    Unload ProgressMultipleBarUF

Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 4), True, "Finding Table of Contents. Progress..."
    DoEvents
    Unload ProgressBarUF
End If
'------------------------------




End Function



Function GetHeaderSize(AuxDoc, oRange, Optional OAPercent)

'--------------------------------------------------------------------
' This function determines Header 1 size asociating with maximun font size
'--------------------------------------------------------------------
Dim n As Long
Dim minFS, maxFS As Integer

'--- Find by Font Size ---
minFS = 200
maxFS = 0


'--------------------------------------------------------------------
'--- Get Font Size info ---
For n = 1 To oRange.Paragraphs.Count
    ''heree
    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(n / oRange.Paragraphs.Count, 4), True, "Identifying Headings. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(n / oRange.Paragraphs.Count, 4), True, "Identifying Headings. Progress..."
        DoEvents
   End If
    '------------------------------

    If AuxDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size >= maxFS Then
        maxFS = AuxDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size
    End If

    If AuxDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size <= minFS Then
        minFS = AuxDoc.Range(oRange.Paragraphs(n).Range.Start, oRange.Paragraphs(n).Range.End).Font.Size
    End If
Next n

'If maxFS > minFS Then
    GetHeaderSize = maxFS
'    Exit Function
'End If


    '------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 3), True, "Identifying Headings. Progress..."
    DoEvents
    Unload ProgressMultipleBarUF
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 3), True, "Identifying Headings. Progress..."
    DoEvents
    Unload ProgressBarUF
End If
'------------------------------

End Function




Sub CopyRangeToWord(oDoc, AuxDoc, oAuxRange, ContenType, Optional rgnStyle)
'--------------------------------------------------------------------
' This function copy range to New document
'--------------------------------------------------------------------

Dim oRange1 As Range

If ContenType = "Preface" Then

    parPreface = GetOutputPrefaceRange(oDoc)
    Set oRange1 = oDoc.Range(oDoc.Paragraphs(parPreface(1)).Range.Start, oDoc.Paragraphs(parPreface(1)).Range.End)
    
    oRange1.Select
    Selection.Paragraphs.Add
    
    oRange1.Select
    oAuxRange.Copy
    Selection.Paste
    
    
Else
    Set oRange1 = oDoc.Range(oDoc.Paragraphs(oDoc.Paragraphs.Count).Range.Start, oDoc.Paragraphs(oDoc.Paragraphs.Count).Range.End)
    oRange1.Select
    Selection.Paragraphs.Add
    
    oAuxRange.Copy
    Selection.Paste
    
    Set oRange1 = oDoc.Range(oDoc.Paragraphs(oDoc.Paragraphs.Count - 1).Range.Start, oDoc.Paragraphs(oDoc.Paragraphs.Count - 1).Range.End)


End If

If IsMissing(rgnStyle) = False Then
    oRange1.Style = rgnStyle
    oRange1.Font.Color = oDoc.Styles(rgnStyle).Font.Color
    oRange1.Paragraphs.LineSpacing = oDoc.Styles(rgnStyle).ParagraphFormat.LineSpacing

Else
    oRange1.Font.ColorIndex = wdBlack
End If

If IsMissing(rgnStyle) = False Then
    If rgnStyle = wdStyleHeading1 Then
        oRange1.Font.Bold = False
    End If
ElseIf oAuxRange.Font.Bold = True Then
    oRange1.Font.Bold = True
End If

End Sub


Sub CopyTabletoWord(oDoc, tbl, oRange2, Optional IsCoverTable)
'--------------------------------------------------------------------
'--- This sub Paste and Format selected table to Converted document ---
'--------------------------------------------------------------------
Dim r, c As Long
Dim oCellRange As Range

tbl.Range.HighlightColorIndex = wdNoHighlight

oRange2.Select
For r = 1 To tbl.Rows.Count + 2
    Selection.Paragraphs.Add
    DoEvents
Next r

oRange2.Select
oRange2.FormattedText = tbl.Range.FormattedText


'--- Formating Table ---
If IsMissing(IsCoverTable) = False Then
    If IsCoverTable = False Then
        Set tbl = oRange2.Tables(1)
        tbl.Select
        '--- Font Style ---
        With tbl
            On Error Resume Next
            .Shading.Texture = wdTextureNone
            .Shading.ForegroundPatternColor = wdColorAutomatic
            .Shading.BackgroundPatternColor = wdColorAutomatic

            If .Rows.Count > 1 Then
                Do While Replace(Replace(Replace(.Rows(.Rows.Count).Range.Text, "", ""), Chr(13), ""), " ", "") = ""
                    .Rows(.Rows.Count).Delete
                    
                    If Err.Number <> 0 Then
                        Err.Number = 0
                        Exit Do
                    End If
                    
                    If .Rows.Count = 1 Then Exit Do
                Loop
            End If
            
            If .Rows.Count > 1 Then
                Do While Replace(Replace(.Rows(1).Range.Text, "", ""), Chr(13), "") = ""
                    .Rows(1).Delete
                    
                    If Err.Number <> 0 Then
                        Err.Number = 0
                        Exit Do
                    End If

                    If .Rows.Count = 1 Then Exit Do
                Loop
            End If
        


            For r = 1 To .Rows.Count
                If r = 1 Then
                    Set oRowRange = .Rows(r).Range
                    If .Rows.Count > 1 Then
                        oRowRange.ParagraphFormat.KeepWithNext = wdToggle
                        oRowRange.Font.Color = RGB(188, 67, 40)
                        .Rows(r).HeightRule = wdRowHeightAuto
                    End If
                
                Else
                    For c = 1 To .Rows(r).Cells.Count
                        Set oCellRange = .Rows(r).Cells(c).Range
                        If oCellRange.ListFormat.ListType = wdListBullet Then
'                            oCellRange.Select
                            rgnStyle = "enumeration " & oCellRange.ListFormat.ListLevelNumber
                            oCellRange.Style = rgnStyle
                            oCellRange.Font.Color = oDoc.Styles(rgnStyle).Font.Color
                            oCellRange.Paragraphs.LineSpacing = oDoc.Styles(rgnStyle).ParagraphFormat.LineSpacing
                        Else
                            oCellRange.ParagraphFormat.LeftIndent = 2
                            oCellRange.Font.ColorIndex = wdBlack
                        End If
                    Next c
                End If
                
                oRowRange.Font.Name = "Arial"
                oRowRange.Font.Size = 10
                oRowRange.Paragraphs.LineSpacing = 12
                oRowRange.Paragraphs.LineSpacingRule = wdLineSpaceSingle
                oRowRange.Paragraphs.LineUnitAfter = 0
                oRowRange.Paragraphs.LineUnitBefore = 0
            Next r
           
           
'            '--- Add first Column if needed---
            AddFirstColumn = True
            For r = 1 To tbl.Rows.Count
                If .Rows(r).Range < .Columns.Count Then
                    AddFirstColumn = False
                    Exit For
                End If
            Next r


            If AddFirstColumn = True Then
                For r = 1 To tbl.Rows.Count
                    If Trim(Replace(Replace(tbl.Rows(r).Cells(1).Range.Text, "", ""), Chr(13), "")) <> "" Then
                        Exit For
                    ElseIf r = tbl.Rows.Count Then
                        AddFirstColumn = False
                    End If
                Next r
            End If

            If AddFirstColumn = True Then
                Set newCol = .Columns.Add(BeforeColumn:=.Columns(1))
            End If
            .Rows(2).HeadingFormat = True
            .Columns(1).Width = InchesToPoints(0.19685)
            
'           '--- Add first row if needed ---
            If .Rows.Count = 1 Then
                Set newRow = .Rows.Add(BeforeRow:=tbl.Rows(1))
                .Rows(1).Height = InchesToPoints(0.19685)
            End If
            
            
            '--- Adjust Table width ---
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 95
            .Rows.LeftIndent = CentimetersToPoints(0.1)
            
            Twidth = 0
            MaxTblWidth = 482
            For c = 1 To tbl.Columns.Count
                Twidth = Twidth + tbl.Columns(c).Width
            Next c
            
            If Twidth > MaxTblWidth Then
                tbl.Columns(1).Width = InchesToPoints(0.19685)
                For c = 2 To tbl.Columns.Count
                    tbl.Columns(c).Width = (MaxTblWidth - InchesToPoints(0.19685)) / (tbl.Columns.Count - 1)
                Next c
            End If
            
            
            '--- Borders ---
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            .Borders.Shadow = False
        
            With .Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth025pt
                .Color = RGB(217, 217, 217)
            End With
        
            With .Rows(1).Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth025pt
                .Color = RGB(0, 0, 0)
            End With
        
            With .Columns(1).Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth025pt
                .Color = RGB(0, 0, 0)
            End With
            On Error GoTo 0
        End With

    End If
End If

End Sub




Function GetFooterinfo(oDoc)
'--------------------------------------------------------------------
'--- This function extracts footer information from input file  ---
'--------------------------------------------------------------------

Dim oSection As Section
Dim oFooter As HeaderFooter
Dim strFooterinfo
Dim DocTitle As String
Dim p, d As Integer
Dim strFooter
Dim DocDate
Dim Daysarray
Dim CoverInfo


On Error GoTo GetFooterinfo_Exit:
DocDate = 0
DocTitle = ""




Daysarray = Array("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday", _
                  "january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december", _
                  "montag", "dienstag", "mittwoch", "donnerstag", "freitag", "samstag", "sonntag", _
                  "Jjnuar", "februar", "märz", "april", "mai", "juni", "juli", "august", "september", "oktober", "november", "dezember")

For Each oSection In oDoc.Sections
    For Each oFooter In oSection.Footers
        If oFooter.Exists Then
            Set strFooter = oFooter.Range
            For p = 1 To strFooter.Paragraphs.Count
                strFooterinfo = strFooter.Paragraphs(p)
                
                
                If IsDate(strFooterinfo) = True Then
                    '--- Get Document Date ---
                    If DocDate < strFooterinfo.Text Then
                        DocDate = Replace(strFooterinfo.Text, Chr(13), "")
                    End If

'                ElseIf strFooterinfo.Fields.Count > 0 Then
'                        '--- Get Document pages info ---
'                        For f = 1 To strFooterinfo.Fields.Count
'                            If strFooterinfo.Fields(f).Type = wdFieldPage Or strFooterinfo.Fields(f).Type = wdFieldNumPage Then
'                                DocPagesInfo = strFooterinfo
'                            End If
'
'                        Next f
                Else
                    If Replace(strFooterinfo.Text, Chr(13), "") <> "" Then
                    
                        For d = 0 To UBound(Daysarray)
                            If InStr(LCase(strFooterinfo.Text), LCase(Daysarray(d))) = 1 Then
                                If DocDate < strFooterinfo.Text Then
                                    DocDate = Replace(strFooterinfo.Text, Chr(13), "")
                                    Exit For
                                End If

                            End If

                            If d = UBound(Daysarray) Then
                                '--- Get Document Title ---
                                If DocTitle = "" Then
                                    DocTitle = Replace(strFooterinfo.Text, Chr(13), "")
                                ElseIf InStr(LCase(strFooterinfo.Text), "page") = 0 And InStr(LCase(strFooterinfo.Text), "pag.") = 0 And _
                                        InStr(LCase(strFooterinfo.Text), "seite") = 0 And InStr(LCase(strFooterinfo.Text), "seiten") = 0 Then
                                    DocTitle = DocTitle & ". " & Replace(strFooterinfo.Text, Chr(13), "")
                                End If
                            End If
                        Next d
                    
                    End If
                End If
                
                
                
            Next p
            
        End If
    Next oFooter
Next oSection




'--------------------------------------------------------------------
If DocTitle = "" Then
    '--- Set Footer DocTitle by Cover info ---
    CoverInfo = DocGetCoverInfo(oDoc)
    If CoverInfo(0) <> "" Then
        DocTitle = CoverInfo(0)
    End If
End If

'--------------------------------------------------------------------
If DocDate = 0 Then
    n0 = 0
    lgPage_0 = 1
    Do

        If n0 > oDoc.Paragraphs.Count Or lgPage_0 > 1 Then Exit Do
            Set oRange = oDoc.Paragraphs(1).Next(Count:=n0).Range
            oRange.Select
            For f = 1 To oRange.Fields.Count
                If oRange.Fields(f).Type = wdFieldCreateDate Then
                    DocDate = Format(Replace(oRange.Text, "", ""), "dd.mm.yyyy")
                    Exit Do
                End If
            Next f
        '---
        n0 = n0 + 1
        Set oRange = oDoc.Paragraphs(1).Next(Count:=n0).Range
        lgPage_0 = oRange.Information(wdActiveEndAdjustedPageNumber)
    Loop

End If


If DocDate = 0 Then
    DocDate = ""
End If
'--------------------------------------------------------------------


GetFooterinfo_Exit:

GetFooterinfo = Array(DocTitle, DocDate)
End Function


Function GetOutputPrefaceRange(oDoc)

'--------------------------------------------------------------------
'--- This function find preface paragraphs of output file ---
'--------------------------------------------------------------------
GetOutputPrefaceRange = Array(0, 0)

n0 = 0
Do While n0 <= oDoc.Paragraphs.Count
    Set oRange = oDoc.Paragraphs(1).Next(Count:=n0).Range
    lgPage_0 = oRange.Information(wdActiveEndAdjustedPageNumber)
    If lgPage_0 > 1 Then

        n0 = n0 + 1
        For nf = 0 To oDoc.Paragraphs.Count
            Set oRange = oDoc.Paragraphs(n0).Next(Count:=nf).Range
            oRange.Select
            lgPage_f = oRange.Information(wdActiveEndAdjustedPageNumber)
            If lgPage_f > lgPage_0 Then
                    Set oRange1 = oDoc.Range(oDoc.Paragraphs(n0 + nf - 2).Range.Start, oDoc.Paragraphs(n0 + nf - 2).Range.End)
                    oRange1.Select
                    
                    Exit Do
                Exit For
            End If
        Next nf

    End If
    n0 = n0 + 1
Loop



GetOutputPrefaceRange = Array(n0, n0 + nf - 2)

End Function


Function GetBuiltInQuickGallery(Optional OAPercent)
Dim AuxArray()
Dim Count As Integer
Set oDoc = ActiveDocument

'--------------------------------------------------------------------
'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(0, 3), True, "Copying Cover information. Progress..."
    DoEvents
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(0, 3), True, "Preparing Data. Progress..."
    DoEvents
End If
'--------------------------------------------------------------------

Count = -1
With oDoc
    For Each oSty In .Styles
        If oSty.QuickStyle = True Then
            Count = Count + 1
            ReDim Preserve AuxArray(Count)
            AuxArray(Count) = oSty.NameLocal
        End If
    Next oSty
End With

GetBuiltInQuickGallery = AuxArray

'--------------------------------------------------------------------
'------- update progress ----
'--------------------------------------------------------------------

If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 3), True, "Copying Cover information. Progress..."
    DoEvents
    
    Unload ProgressMultipleBarUF
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 3), True, "Preparing Data. Progress..."
    DoEvents
    Unload ProgressBarUF
End If
End Function



Sub ManageQuickStyleGallery(BuiltInStyleArray, Optional OAPercent)
Dim AuxArray()
Dim QuickGalNumber As Integer
Set oDoc = ActiveDocument

'BuiltInStyleArray = GetBuiltInQuickGallery

'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(0, 3), True, "Finishing. Progress..."
    DoEvents
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(0, 3), True, "Finishing. Progress..."
    DoEvents
End If
'--------------------------------------------------------------------
GalleryOrderArray = Array("Title black", "Title blue", "Title red", wdStyleSubtitle, oDoc.Styles(wdStyleHeading1).NameLocal, oDoc.Styles(wdStyleHeading2).NameLocal, _
                          oDoc.Styles(wdStyleHeading3).NameLocal, oDoc.Styles(wdStyleHeading4).NameLocal, oDoc.Styles(wdStyleHeading5).NameLocal, _
                          oDoc.Styles(wdStyleHeading6).NameLocal, oDoc.Styles(wdStyleHeading7).NameLocal, oDoc.Styles(wdStyleHeading8).NameLocal, _
                          oDoc.Styles(wdStyleHeading9).NameLocal, "Text", "enumeration1", "enumeration2", "enumeration3", "numbering1", "numbering2", _
                          "numbering3", wdStyleEmphasis, "Title footer", wdStyleTocHeading)
With oDoc
    For Each oSty In .Styles
        For s = 0 To UBound(BuiltInStyleArray)
            If oSty.NameLocal = BuiltInStyleArray(s) Then
                oSty.QuickStyle = True
                
                For G = 0 To UBound(GalleryOrderArray)
                   If oSty.NameLocal = GalleryOrderArray(G) Then
                      oSty.Priority = G + 1
                   End If
                Next G
                
                Exit For
            ElseIf s = UBound(BuiltInStyleArray) Then
                If oSty.QuickStyle = True Then
                    oSty.QuickStyle = False
                End If
            End If
        Next s
    Next oSty


End With

'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 3), True, "Finishing. Progress..."
    DoEvents
    
    Unload ProgressMultipleBarUF

Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 3), True, "Finishing. Progress..."
    DoEvents
    
    Unload ProgressBarUF
End If
End Sub

