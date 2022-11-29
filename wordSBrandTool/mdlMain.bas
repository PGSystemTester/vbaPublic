Attribute VB_Name = "mdlMain"
''Option Explicit

Sub ConvertSingleFile()
Dim oDoc
Dim FilePath As String

'--- Select Input File ---
FilePath = SelFileDialog(ActiveDocument)
If FilePath = "" Then
    MsgBox ("Please Select a File")
    Exit Sub
End If

'--------------------------------------------------------------------
Set oDoc = ActiveDocument
boolConvertDocument = ConvertDocument(oDoc, FilePath)


If IsFileOpen(FilePath) = True Then
    Set AuxDoc = GetObject(FilePath)
    AuxDoc.Close (False)
End If

ProgressBarUF.Show
ProgressBarUF.ShowProgress Round(1, 3), True, "Finishing..."
DoEvents

oDoc.Activate
    Unload ProgressBarUF
    DoEvents

If boolConvertDocument = True Then
    MsgBox "Done!"
End If

End Sub

Sub ConvertMultiplesFiles()


Dim projectFolderPath, fileName
Dim FilePath As String
Dim TempFileName As String
Dim FileArray()

projectFolderPath = GetFolder

If projectFolderPath = "" Then Exit Sub
fileName = Dir(projectFolderPath & "\*doc*")
If fileName = "" Then
  MsgBox "No files found"
  Exit Sub
End If

'--- loop over files -----
TemplatePath = ActiveDocument.AttachedTemplate.FullName
fcount = -1
While fileName <> ""
    If fileName <> ActiveDocument.AttachedTemplate Then
        fcount = fcount + 1
        ReDim Preserve FileArray(fcount)
        FileArray(fcount) = projectFolderPath & "\" & fileName
        fileName = Dir
    End If
Wend

fcount = 0
For f = 0 To UBound(FileArray)

    If fcount > 0 Then
        Set oAuxDoc = ActiveDocument
        Set oDoc = Application.Documents.Add(TemplatePath, False)
        oDoc.Activate
        oAuxDoc.Close (True)
    Else
        Set oDoc = ActiveDocument
    End If

    FilePath = FileArray(f) ''projectFolderPath & "\" & fileName
    boolConvertDocument = ConvertDocument(oDoc, FilePath, Round(fcount / (UBound(FileArray) + 1), 3))
    If boolConvertDocument = False Then Exit For
    fcount = fcount + 1

Next f

If boolConvertDocument = True Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress 1, Round(1, 3), True, "Finishing..."
    DoEvents
End If

oDoc.Activate
Unload ProgressMultipleBarUF
DoEvents

If boolConvertDocument = True Then
    MsgBox "Done!"
End If
End Sub


Function ConvertDocument(oDoc, FilePath As String, Optional OAPercent) As Boolean

'Dim oDoc
Dim AuxDoc
Dim AuxLog
Dim parPreface
''Dim FilePath  As String
Dim fileName As String
Dim fileName0 As String
Dim LogArgument As String
Dim OutputFilePath As String
Dim tc As Integer
Dim Count As Integer
Dim p As Integer
Dim lgPage As Long
Dim fso, logFile
Dim logPath As String
Dim oRange As Range
Dim ContentTablePos
Dim oData   As New DataObject 'object to use the clipboard

On Error GoTo ConvertDocument_Exit
ConvertDocument = False
'--------------------------------------------------------------------
'--- Select Input File ---
If FilePath <> "" Then
    If IsFileOpen(FilePath) = True Then
        Set AuxDoc = GetObject(FilePath)
    Else
        Set AuxDoc = Application.Documents.Open(FilePath)
    End If
Else
    MsgBox ("Please Select a File")
    Exit Function
End If

If IsFileOpen(FilePath) = True Then
        Set AuxDoc = GetObject(FilePath)
    Else
        Set AuxDoc = Application.Documents.Open(FilePath)
End If
''--------------------------------------------------------------------
'--- Select Output Folder ---
OutputFilePath = AuxDoc.Path

AuxDoc.ActiveWindow.WindowState = wdWindowStateMinimize
oDoc.Activate
Application.ScreenUpdating = False
Application.DisplayAlerts = False


'--------------------------------------------------------------------
'--- Output file Name ---
fileName0 = ""
strExtension = "." & Split(AuxDoc.Name, ".")(UBound(Split(AuxDoc.Name, ".")))
fileName0 = Replace(AuxDoc.Name, strExtension, "")

If InStr(1, LCase(fileName0), "itelligence") > 0 Then
    fileName0 = Replace(fileName0, "itelligence", "NTT DATA")
    fileName0 = Replace(fileName0, "Itelligence", "NTT DATA")
    fileName0 = Replace(fileName0, "ITELLIGENCE", "NTT DATA")
Else
    fileName0 = Replace(fileName0, "itelli", "NTT DATA")
    fileName0 = Replace(fileName0, "Itelli", "NTT DATA")
    fileName0 = Replace(fileName0, "ITELLI", "NTT DATA")
End If

If fileName0 & ".docx" = AuxDoc.Name Then
    fileName = fileName0 & " - NTT DATA"
Else
    fileName = fileName0
End If

fileName0 = fileName
Do Until Dir(OutputFilePath & "\" & fileName & ".docx") = ""
    Count = Count + 1
    fileName = fileName0 & "_" & Format(Count, "00")
Loop

'--------------------------------------------------------------------
'--- Create Log File ---
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = OutputFilePath & "\" & fileName & " - log.txt"
If Dir(logPath) <> "" Then
    If IsFileOpen(logPath) = True Then
        Set AuxLog = GetObject(logPath)
        AuxLog.Close
    End If
    Kill logPath
End If

Set logFile = fso.CreateTextFile(logPath, False)
LogArgument = "#" & Chr(9) & "Time" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Description"
Call PrintLog(logFile, LogArgument)


Call PrintLog(logFile, "Source file path: " & FilePath)

'--------------------------------------------------------------------
'--- Delete indicative First image ---
If ClearTemplate(oDoc) = False Then
    Call PrintLog(logFile, "ERROR: document template is corrupted")
    MsgBox "Document template is corrupted", vbCritical
    Exit Function
End If

'--------------------------------------------------------------------
'--- Get Quick Style Gallery state at start ---
If IsMissing(OAPercent) = False Then
    BuiltInQuickGalleryArray = GetBuiltInQuickGallery(OAPercent)
Else
    BuiltInQuickGalleryArray = GetBuiltInQuickGallery
End If

'--------------------------------------------------------------------
'--- Formatting Cover page ---
If IsMissing(OAPercent) = False Then
    boolFormatCoverPage = FormatCoverPage(oDoc, AuxDoc, logFile, OAPercent)
Else
    boolFormatCoverPage = FormatCoverPage(oDoc, AuxDoc, logFile)
End If

'--- Delete indicative First image ---
If boolFormatCoverPage = False Then
    Call PrintLog(logFile, "FATAL ERROR: on cover page converting")
    MsgBox "Error on cover page converting", vbCritical
    Exit Function
End If

'--------------------------------------------------------------------
'--- Formatting Preface Section ---
If IsMissing(OAPercent) = False Then
    ContentTablePos = GetContentTablePar(oDoc, AuxDoc, OAPercent)
Else
    ContentTablePos = GetContentTablePar(oDoc, AuxDoc)
End If

If ContentTablePos(0) <> ContentTablePos(1) Then
    Set oRange = AuxDoc.Range(AuxDoc.Paragraphs(ContentTablePos(0)).Range.Start, AuxDoc.Paragraphs(ContentTablePos(0)).Range.End)
    lgPage = oRange.Information(wdActiveEndAdjustedPageNumber)
End If

If lgPage >= 2 Then
    '--- Convert Preface ---
    If IsMissing(OAPercent) = False Then
        boolFormatDocBody = FormatDocBody(oDoc, AuxDoc, logFile, "Preface", ContentTablePos)
    Else
        boolFormatDocBody = FormatDocBody(oDoc, AuxDoc, logFile, "Preface", ContentTablePos, OAPercent)
    End If
    
    If boolFormatCoverPage = False Then
        Call PrintLog(logFile, "FATAL ERROR: on Document preface converting")
        MsgBox "Error on Document  preface converting", vbCritical
        Exit Function
    End If
    
    parPreface = GetOutputPrefaceRange(oDoc)
Else
    '--- Delete Preface section on output file ---
    parPreface = GetOutputPrefaceRange(oDoc)
    Set oRange = oDoc.Range(oDoc.Paragraphs(parPreface(0)).Range.Start, oDoc.Paragraphs(parPreface(1) + 1).Range.End)
    oRange = ""
'    oRange.Delete
End If

'--------------------------------------------------------------------
'--- Formatting Main Section ---
If IsMissing(OAPercent) = False Then
    boolFormatDocBody = FormatDocBody(oDoc, AuxDoc, logFile, "Main", ContentTablePos, OAPercent)
Else
    boolFormatDocBody = FormatDocBody(oDoc, AuxDoc, logFile, "Main", ContentTablePos)
End If

If boolFormatDocBody = False Then
    Call PrintLog(logFile, "FATAL ERROR: on Document main part converting")
    MsgBox "Error on Document main part converting", vbCritical
    Exit Function
End If

'--------------------------------------------------------------------
'--- Formatting Footer ---
If IsMissing(OAPercent) = False Then
    boolFormatFooter = FormatFooter(oDoc, AuxDoc, logFile, OAPercent)
Else
    boolFormatFooter = FormatFooter(oDoc, AuxDoc, logFile)
End If

If boolFormatFooter = False Then
    Call PrintLog(logFile, "FATAL ERROR: on Document footer converting")
    MsgBox "Error on Document footer converting", vbCritical
    Exit Function
End If


'--------------------------------------------------------------------
'--- Create & Update TOC ---
boolCreateTOC = False
For stl = -2 To -10 Step -1
    '---Legend -- wdStyleHeading1=-2, wdStyleHeading2=-3... wdStyleHeading9=-10
    Set oRange = oDoc.Range
    With oRange.Find
       .Style = oDoc.Styles(stl)
       .Wrap = wdFindStop
       While .Execute
          boolCreateTOC = True
          Exit For
       Wend
    End With
Next stl


If boolCreateTOC = True Then
    Set oRange = Nothing
    Set oRange1 = Nothing
    
    If parPreface(0) = parPreface(1) Then
        Set oRange = oDoc.Range(oDoc.Paragraphs(parPreface(0)).Range.Start, oDoc.Paragraphs(parPreface(1)).Range.End)
        oRange = "Table of Contents"
        oRange.Style = wdStyleTocHeading
           
        Set oRange1 = oDoc.Range(oDoc.Paragraphs(parPreface(0) + 1).Range.Start, oDoc.Paragraphs(parPreface(1) + 1).Range.End)
        oDoc.TablesOfContents.Add Range:=oRange1, RightAlignPageNumbers:=True, _
         UseHeadingStyles:=True, IncludePageNumbers:=True, UseHyperlinks:=True, _
         HidePageNumbersInWeb:=True, UseOutlineLevels:=False
    
    Else
        Set oRange = oDoc.Range(oDoc.Paragraphs(parPreface(1)).Range.Start, oDoc.Paragraphs(parPreface(1)).Range.End)
        oRange = "Table of Contents"
        oRange.Style = wdStyleTocHeading
        oRange.Select
        Selection.Paragraphs.Add
        Set oRange1 = oDoc.Range(oDoc.Paragraphs(parPreface(1) + 1).Range.Start, oDoc.Paragraphs(parPreface(1) + 1).Range.End)

        oDoc.TablesOfContents.Add Range:=oRange1, RightAlignPageNumbers:=True, _
         UseHeadingStyles:=True, IncludePageNumbers:=True, UseHyperlinks:=True, _
         HidePageNumbersInWeb:=True, UseOutlineLevels:=False
    
    End If
    oDoc.TablesOfContents(oDoc.TablesOfContents.Count).Range.Font.Bold = False
End If
     
If oDoc.TablesOfContents.Count > 0 Then
    For tc = 1 To oDoc.TablesOfContents.Count
        oDoc.TablesOfContents(tc).Update
    Next tc
    Call PrintLog(logFile, "Table of content updated successfully!")
Else
    Call PrintLog(logFile, "WARNING: It was not found any Table of content")
End If

'--------------------------------------------------------------------
'--- Organize Quck Style Gallery ---
If IsMissing(OAPercent) = False Then
    Call ManageQuickStyleGallery(BuiltInQuickGalleryArray, OAPercent)
Else
    Call ManageQuickStyleGallery(BuiltInQuickGalleryArray)
End If

'--------------------------------------------------------------------
'-- Finishing ---
If IsMissing(OAPercent) = False Then
    boolClearEmptyPages = ClearEmptyPages(oDoc, OAPercent)
Else
    boolClearEmptyPages = ClearEmptyPages(oDoc)
End If

If boolClearEmptyPages = False Then
    Call PrintLog(logFile, "FATAL ERROR clearing output document empty pages")
    MsgBox "Error clearing output document empty pages", vbCritical
    Exit Function
End If



'--- Saving file ---
oDoc.SaveAs2 fileName:=OutputFilePath & "\" & fileName & ".docx", _
    FileFormat:=wdFormatDocumentDefault, _
    LockComments:=False, _
    Password:="", _
    AddToRecentFiles:=True, _
    WritePassword:="", _
    ReadOnlyRecommended:=False, _
    EmbedTrueTypeFonts:=False, _
    SaveNativePictureFormat:=False, _
    SaveFormsData:=False, _
    SaveAsAOCELetter:=False, _
    CompatibilityMode:=15


'--- emptiying clipboard ---
oData.SetText Text:=Empty 'Clear
oData.PutInClipboard 'take in the clipboard to empty it

Application.ScreenUpdating = True
Application.DisplayAlerts = True

AuxDoc.Close (False)

If 1 > 2 Then
ConvertDocument_Exit:
    ConvertDocument = False
Else
    ConvertDocument = True
End If

End Function


Function FormatCoverPage(oDoc, AuxDoc, logFile, Optional OAPercent) As Boolean

Dim rngTableTarget As Range
Dim oAuxFirstPageRange As Range
Dim oRange As Range
Dim oRange2 As Range

Dim n, startPara, t, r As Long
Dim tbl, Shp As Object
Dim lgPage
Dim CoverInfo
Dim GetCoverPagePar, AuxCoverRange

On Error GoTo FormatCoverPage_exit

'--------------------------------------------------------------------
'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(0, 3), True, "Copying Cover information. Progress..."
    DoEvents

Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(0, 3), True, "Copying Cover information. Progress..."
    DoEvents
End If
'--------------------------------------------------------------------
'--- Select Cover page range ---
CoverInfo = DocGetCoverInfo(AuxDoc)
oDoc.Activate

'--------------------------------------------------------------------
'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(0.25, 4), True, "Copying Cover information. Progress..."
    DoEvents
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(0.25, 4), True, "Copying information. Progress..."
    DoEvents
End If

'--- Replace Title ---
If CoverInfo(0) <> "|--> Not Found <--|" Then
    Set oRange = Nothing
    Set oRange = oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(1).Range.End)
    oRange.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .Text = "|-->Title<--|"
        .Replacement.Text = CoverInfo(0)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Call PrintLog(logFile, "Document Title OK!")

    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(0.375, 4), True, "Copying Cover information. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(0.375, 4), True, "Copying information. Progress..."
        DoEvents
    End If
    
    '--- Replace subtitle ---
    Set oRange = Nothing
    Set oRange = oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(2).Range.End)
    oRange.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "|-->Subtitle<--|"
        .Replacement.Text = CoverInfo(1)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
    If CoverInfo(1) <> "" Then
        Call PrintLog(logFile, "Document Subtitle OK!")
    Else
        Call PrintLog(logFile, "WARNING: Document Subtitle was not found")
    End If
    

    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(0.5, 4), True, "Copying Cover information. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(0.5, 4), True, "Copying information. Progress..."
        DoEvents
    End If

    '--- Find Shapes: Tables ---
    GetCoverPagePar = GetCoverPageRange(AuxDoc)
    Set AuxCoverRange = AuxDoc.Range(AuxDoc.Paragraphs(GetCoverPagePar(0)).Range.Start, AuxDoc.Paragraphs(GetCoverPagePar(1)).Range.End)

    n = 3
    If AuxCoverRange.Tables.Count > 0 Then
        For t = 1 To AuxCoverRange.Tables.Count
            Set tbl = AuxCoverRange.Tables(t)
            Set oRange2 = oDoc.Range(oDoc.Paragraphs(n).Range.Start, oDoc.Paragraphs(n).Range.End)
            Call CopyTabletoWord(oDoc, tbl, oRange2, True)
            
            n = n + tbl.Range.Paragraphs.Count + 3
        Next t
        Call PrintLog(logFile, AuxCoverRange.Tables.Count & " Document cover tables OK!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found document any cover page table")
    End If
    
    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(0.75, 4), True, "Copying Cover information. Progress..."
        DoEvents
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(0.75, 4), True, "Copying information. Progress..."
        DoEvents
    End If
    

    
    '--- Find Shapes: inline shapes ---
    If AuxCoverRange.InlineShapes.Count > 0 Then
        For t = 1 To AuxCoverRange.InlineShapes.Count
            Set Shp = AuxCoverRange.InlineShapes(t)
            Set oRange2 = oDoc.Range(oDoc.Paragraphs(n).Range.Start, oDoc.Paragraphs(n).Range.End)
            oRange2.Select

            Selection.Paragraphs.Add
            Selection.Paragraphs.Add

            Set rngTableTarget = oDoc.Range(oDoc.Paragraphs(n).Range.Start, oDoc.Paragraphs(n).Range.End)
            rngTableTarget.Select
        
            rngTableTarget.FormattedText = Shp.Range.FormattedText
            Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            n = n + 1
            
        Next t
        Call PrintLog(logFile, AuxCoverRange.InlineShapes.Count & " Document cover picture OK!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found document any cover page picture")
    End If
Else
    Call PrintLog(logFile, "WARNING: Document Cover page was not found")

End If
    
'--------------------------------------------------------------------
'-- Finishing ---
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 3), True, "Copying Cover information. Progress..."
    DoEvents
    Unload ProgressMultipleBarUF
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 3), True, "Copying information. Progress..."
    DoEvents
    Unload ProgressBarUF
End If

If 1 > 2 Then
FormatCoverPage_exit:
    FormatCoverPage = False
Else
    FormatCoverPage = True
End If

End Function

Function FormatDocBody(oDoc, AuxDoc, logFile, ContenType, Optional ContentTablePos, Optional OAPercent) As Boolean
    
Dim GetCoverPagePar
Dim aux1, aux2
Dim parPreface
Dim oRange, oRange2, AuxRng, MainRange As Range
Dim tbl As Object
Dim HeaderFontSize, SubheaderFontSize, l, lgPage  As Integer
Dim n, p, p0, p1, mrP1, mrP2 As Long
Dim countTable As Long
Dim countPictures  As Long
Dim countHeading1  As Long
Dim countHeading2  As Long
Dim countHeading3  As Long
Dim countHeading4  As Long
Dim countHeading5  As Long
Dim countNormal  As Long
Dim countUndefined  As Long
Dim BoolIscoverTable As Boolean

On Error GoTo FormatDocBody_Exit
FormatDocBody = False
'--------------------------------------------------------------------
'--- Getting Content Table ---
'ContentTablePos = GetContentTablePar(oDoc, AuxDoc)
If ContenType = "Main" Then
    If ContentTablePos(0) <> 0 And ContentTablePos(1) <> 0 Then
        Set MainRange = AuxDoc.Range(AuxDoc.Paragraphs(ContentTablePos(1)).Range.Start, AuxDoc.Paragraphs(AuxDoc.Paragraphs.Count).Range.End)
        
        mrP1 = ContentTablePos(1)
        mrP2 = AuxDoc.Paragraphs.Count
    Else
        GetCoverPagePar = GetCoverPageRange(AuxDoc)
        If GetCoverPagePar(1) + 1 >= AuxDoc.Paragraphs.Count Then
            Call PrintLog(logFile, "WARNING: It was found body section on this document")
            GoTo exit_FormatDocBody
        End If
        Set MainRange = AuxDoc.Range(AuxDoc.Paragraphs(GetCoverPagePar(1)).Range.Start, AuxDoc.Paragraphs(AuxDoc.Paragraphs.Count).Range.End)
        MainRange.Select
        mrP1 = GetCoverPagePar(1)
        mrP2 = AuxDoc.Paragraphs.Count
    End If

ElseIf ContenType = "Preface" Then
    For n = 0 To AuxDoc.Paragraphs.Count
        Set oRange = AuxDoc.Paragraphs(1).Next(Count:=n).Range
        oRange.Select
        lgPage = oRange.Information(wdActiveEndAdjustedPageNumber)
        
        If lgPage > 1 Then
            mrP1 = n
            mrP2 = ContentTablePos(0) - 1
            If mrP2 >= mrP1 Then
                Set MainRange = AuxDoc.Range(AuxDoc.Paragraphs(mrP1).Range.Start, AuxDoc.Paragraphs(mrP2).Range.End)
                Exit For
            Else
                Exit Function
            End If
        End If
    Next n

End If

MainRange.Select


'--------------------------------------------------------------------
'--- Getting Header Font Size ---
'HeaderFontSize = 0
'HeaderFontSize = GetHeaderSize(AuxDoc, MainRange)

If ContenType = "Main" Then
    countTable = 0
    countPictures = 0
    countHeading1 = 0
    countHeading2 = 0
    countHeading3 = 0
    countHeading4 = 0
    countHeading5 = 0
    countNormal = 0
    countUndefined = 0
End If

p = mrP1
Do While p <= mrP2
    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round((p - mrP1) / (mrP2 - mrP1 + 1), 4), True, "Copying information. Progress..."
        DoEvents
        
    
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round((p - mrP1) / (mrP2 - mrP1 + 1), 4), True, "Copying information. Progress..."
        DoEvents
    End If
    '------------------------------

    Set oRange = AuxDoc.Range(AuxDoc.Paragraphs(p).Range.Start, AuxDoc.Paragraphs(p).Range.End)
    oRange.HighlightColorIndex = wdNoHighlight
    
    oRange.Select
    If Replace(oRange.Text, Chr(13), "") <> "" And oRange <> vbFormFeed Then
        If Trim(Replace(oRange.Text, Chr(13), "")) = Chr(12) And p = mrP1 Then GoTo next_par
        
        '--- Set no highlight and set default font ---
        If oRange.Font.Name <> oDoc.Styles("Text").Font.Name Then oRange.Font.Name = oDoc.Styles("Text").Font.Name
        oRange.HighlightColorIndex = wdNoHighlight
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        If oRange.Information(wdWithInTable) Then '--> Table Range identified
            
            If oRange.Tables.Count > 0 Then
                Set tbl = oRange.Tables(1)
                tbl.Select
                If ContenType = "Preface" Then
                    BoolIscoverTable = True
                    parPreface = GetOutputPrefaceRange(oDoc)
                    Set oRange2 = oDoc.Range(oDoc.Paragraphs(parPreface(1)).Range.Start, oDoc.Paragraphs(parPreface(1)).Range.End)
                    Call CopyTabletoWord(oDoc, tbl, oRange2, BoolIscoverTable)
                    
                Else
                    BoolIscoverTable = False
                    Set oRange2 = oDoc.Range(oDoc.Paragraphs(oDoc.Paragraphs.Count).Range.Start, oDoc.Paragraphs(oDoc.Paragraphs.Count).Range.End)
                    Call CopyTabletoWord(oDoc, tbl, oRange2, BoolIscoverTable)
                    Set oRange2 = oDoc.Range(oDoc.Paragraphs(oDoc.Paragraphs.Count).Range.Start, oDoc.Paragraphs(oDoc.Paragraphs.Count).Range.End)
                    oRange2.Select
                    
                    Selection.Paragraphs.Add
    
                
                End If
    
                Do
                    AuxDoc.Range(AuxDoc.Paragraphs(p).Range.Start, AuxDoc.Paragraphs(p).Range.End).Select
                    p = p + 1
                    If AuxDoc.Range(AuxDoc.Paragraphs(p).Range.Start, AuxDoc.Paragraphs(p).Range.End).Tables.Count = 0 Or _
                        p > AuxDoc.Paragraphs.Count Then
                        p = p - 1
                        Exit Do
                    End If
                Loop
                    
    '                p = p + tbl.Range.Paragraphs.Count - 1
            End If
            countTable = countTable + 1

        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.InlineShapes.Count > 0 Then '--> Inlineshape identified
            oRange.Select
            Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, "Text")
            countPictures = countPictures + 1
            
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.ShapeRange.Count > 0 Then '--> ShapeRange Identified
            
            If oRange.ShapeRange(1).Type = msoAutoShape Then
                oRange.Select
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, "Text")
                countPictures = countPictures + 1

            ElseIf oRange.ShapeRange(1).Type <> msoTextBox Then
                sh = 1
                p0 = p
shpagain:


                Do While sh <= oRange.ShapeRange.Count
                    oRange.Select
                    oRange.ShapeRange(sh).ConvertToInlineShape
                    sh = sh + 1
                    GoTo shpagain
                Loop
                
                If p < AuxDoc.Paragraphs.Count - 1 Then
                    p1 = p + 1
                Else
                    p1 = p
                End If
                
'                p0 = p1
                Do While p0 <= mrP2
                   Set AuxRng = AuxDoc.Range(AuxDoc.Paragraphs(p0).Range.Start, AuxDoc.Paragraphs(p0).Range.End)
                   AuxRng.Select
                   If AuxRng.Text = "" Or AuxRng.Text = Chr(13) Then
                        AuxRng.Select
                        Selection.Delete Unit:=wdCharacter, Count:=1
                        mrP2 = AuxDoc.Paragraphs.Count
                   Else
                        mrP2 = AuxDoc.Paragraphs.Count
                        p = p - 1
                        Exit Do
                   End If
                    p0 = p0 + 1
                Loop
            End If
            
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.Style = AuxDoc.Styles(wdStyleHeading1) Then '--> Heading 1 identified WdBuiltinStyle enumeration
            If Replace(oRange.Text, Chr(13), "") <> "" Then
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, wdStyleHeading1)
                countHeading1 = countHeading1 + 1
            End If
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.Style = AuxDoc.Styles(wdStyleHeading2) Then  '--> Heading 2 identified WdBuiltinStyle enumeration
            If Replace(oRange.Text, Chr(13), "") <> "" Then
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, wdStyleHeading2)
                countHeading2 = countHeading2 + 1
            End If
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.Style = AuxDoc.Styles(wdStyleHeading3) Then  '--> Heading 3 identified WdBuiltinStyle enumeration
            If Replace(oRange.Text, Chr(13), "") <> "" Then
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, wdStyleHeading3)
                countHeading3 = countHeading3 + 1
            End If
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.Style = AuxDoc.Styles(wdStyleHeading4) Then  '--> Heading 4 identified WdBuiltinStyle enumeration
            If Replace(oRange.Text, Chr(13), "") <> "" Then
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, wdStyleHeading4)
                countHeading4 = countHeading4 + 1
            End If
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.Style = AuxDoc.Styles(wdStyleHeading5) Then  '--> Heading 5 identified WdBuiltinStyle enumeration
            If Replace(oRange.Text, Chr(13), "") <> "" Then
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, wdStyleHeading5)
                countHeading5 = countHeading5 + 1
            End If
        
        '---------------------------------------------------------------------------------------------------------------------------------------
'        ElseIf oRange.Font.Size = HeaderFontSize Then '--> Check Heading 1 by font Size
'            Call CopyRangeToWord(oDoc, AuxDoc, oRange, wdStyleHeading1)
'            countHeading1 = countHeading1 + 1
            
        '---------------------------------------------------------------------------------------------------------------------------------------
'        ElseIf oRange.Font.Bold = True Then '--> Check Heading 2 by font Bold
'            Call CopyRangeToWord(oDoc, AuxDoc, oRange, wdStyleHeading2)
'            countHeading2 = countHeading2 + 1
            
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.ListFormat.ListType = wdListBullet Or Trim(oRange.Words(1)) = "-" Then '--> Paragraphs bullet identified
            If Replace(oRange.Text, Chr(13), "") <> "" Then
'                '-- Set Bullet ---
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, "enumeration " & oRange.ListFormat.ListLevelNumber)
                countNormal = countNormal + 1
            End If
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf oRange.ListFormat.ListType = wdListSimpleNumbering Then '--> Paragraphs numbering identified
            If Replace(oRange.Text, Chr(13), "") <> "" Then
                        
                If oRange.Font.Name <> oDoc.Styles("Text").Font.Name Then
                    oRange.Font.Name = oDoc.Styles("Text").Font.Name
                End If
                        
                If oRange.Font.Color = oDoc.Styles("Text").Font.Color Then
                    Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, "Text")

                Else

                    If oRange.Font.Color <> oDoc.Styles("Text").Font.Color Then
                        oRange.Font.Color = oDoc.Styles("Text").Font.Color
                    End If

                    If oRange.Font.Size <> oDoc.Styles("Text").Font.Size Then
                        oRange.Font.Size = oDoc.Styles("Text").Font.Size
                    End If

                    Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType)
                End If
            
                Set oRange2 = oDoc.Range(oDoc.Paragraphs(oDoc.Paragraphs.Count - 1).Range.Start, oDoc.Paragraphs(oDoc.Paragraphs.Count - 1).Range.End)
                oRange2.Select
                Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    
                '-- Set nubering ---
                Selection.Range.ListFormat.ApplyListTemplateWithLevel _
                ListTemplate:=oRange.ListFormat.ListTemplate, _
                ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
                DefaultListBehavior:=wdWord10ListBehavior
                           
                Selection.Range.SetListLevel Level:=oRange.ListFormat.ListLevelNumber
                
                If oDoc.Range(oDoc.Paragraphs(oDoc.Paragraphs.Count - 2).Range.Start, oDoc.Paragraphs(oDoc.Paragraphs.Count - 2).Range.End).ListFormat.ListType = wdListBullet Then
                    Selection.Range.ListFormat.ApplyListTemplateWithLevel _
                    ListTemplate:=oRange.ListFormat.ListTemplate, _
                    ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
                    DefaultListBehavior:=wdWord10ListBehavior
                End If
            
                countNormal = countNormal + 1
            End If
            
        '---------------------------------------------------------------------------------------------------------------------------------------
        ElseIf (oRange.Style = AuxDoc.Styles(wdStyleNormal) And oRange.ListFormat.ListType = wdListNoNumbering) And InStr(1, oRange, Chr(9)) = 0 Then '--> Normal range identified
                                
                oRange.Font.Color = oDoc.Styles("Text").Font.Color
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, "Text")
                countNormal = countNormal + 1
        
        ElseIf oRange.Font.Size = AuxDoc.Styles(wdStyleNormal).Font.Size And _
               oRange.Font.Color = AuxDoc.Styles(wdStyleNormal).Font.Color And oRange.ListFormat.ListType = wdListNoNumbering And InStr(1, oRange, Chr(9)) = 0 Then '--> Normal range identified
                                
                oRange.Font.Color = oDoc.Styles("Text").Font.Color
                Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType, "Text")
                countNormal = countNormal + 1

        '---------------------------------------------------------------------------------------------------------------------------------------
        Else '--> unasigned Style
            If oRange.Font.Name <> oDoc.Styles("Text").Font.Name Then
                oRange.Font.Name = oDoc.Styles("Text").Font.Name
            End If
            
            Call CopyRangeToWord(oDoc, AuxDoc, oRange, ContenType)
            countUndefined = countUndefined + 1
        End If
    End If
next_par:
    
    If ContenType = "Main" Then
        mrP2 = AuxDoc.Paragraphs.Count
    Else
        mrP2 = ContentTablePos(0) - 3
    End If
    p = p + 1
    
Loop

'---------------------------------------------------------------------------------------------------------------------------------------
If ContenType = "Main" Then
    
    '--- Write Log File ---
    If countHeading1 > 0 Then
        Call PrintLog(logFile, countHeading1 & "  Heading 1 paragraphs converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any Heading 1 paragraphs")
    End If
    
    '-------------------------
    If countHeading2 > 0 Then
        Call PrintLog(logFile, countHeading2 & "  Heading 2 paragraphs converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any Heading 2 paragraphs")
    End If
    
    '-------------------------
    If countHeading3 > 0 Then
        Call PrintLog(logFile, countHeading3 & "  Heading 3 paragraphs converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any Heading 3 paragraphs")
    End If
    
    '-------------------------
    If countHeading4 > 0 Then
        Call PrintLog(logFile, countHeading4 & "  Heading 4 paragraphs converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any Heading 4 paragraphs")
    End If
    
    '-------------------------
    If countHeading5 > 0 Then
        Call PrintLog(logFile, countHeading5 & "  Heading 5 paragraphs converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any Heading 5 paragraphs")
    End If
    
    '-------------------------
    If countNormal > 0 Then
        Call PrintLog(logFile, countNormal & "  Normal style paragraphs converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any Normal Style paragraphs")
    End If
    
    '-------------------------
    If countTable > 0 Then
        Call PrintLog(logFile, countTable & "  main part table converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any document main part table")
    End If

    '-------------------------
    If countPictures > 0 Then
        Call PrintLog(logFile, countPictures & "  main part Pictures/Figures converted!")
    Else
        Call PrintLog(logFile, "WARNING: It was not found any document main part Pictures/Figures")
    End If
    
    '-------------------------
    If countUndefined > 0 Then
        Call PrintLog(logFile, "WARNING: It was found " & countUndefined & " Unasigned style paragraphs")
    End If
Else
    '---Preface ---
    Call PrintLog(logFile, "Preface section converted!")
End If

If 1 > 2 Then
FormatDocBody_Exit:
    FormatDocBody = False
Else
    FormatDocBody = True
End If
'--------------------------------------------------------------------
'------- update progress ----
exit_FormatDocBody:
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 3), True, "Copying Cover information. Progress..."
    DoEvents
    Unload ProgressMultipleBarUF
Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 3), True, "Copying information. Progress..."
    DoEvents
    
    Unload ProgressBarUF
End If




End Function


Function FormatFooter(oDoc, AuxDoc, logFile, Optional OAPercent) As Boolean
Dim oSection As Section
Dim oFooter As HeaderFooter
Dim strFooter As Range
Dim oRange1
Dim p As Integer
Dim FooterInfo

On Error GoTo FormatFooter_Exit
FormatFooter = False

FooterInfo = GetFooterinfo(AuxDoc)


For Each oSection In oDoc.Sections
    For Each oFooter In oSection.Footers
        If oFooter.Exists Then
            Set strFooter = oFooter.Range
            For p = 1 To 2
                Set oRange1 = strFooter.Paragraphs(p)
                oRange1.Range.Select
                Selection.Paragraphs.Add
                oRange1.Range.Text = FooterInfo(p - 1)
           Next p
        End If
    Next oFooter
Next oSection

oDoc.Activate
If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    ActiveWindow.ActivePane.View.Type = wdPrintView
Else
    ActiveWindow.View.Type = wdPrintView
End If
ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument


If 1 > 2 Then
FormatFooter_Exit:
    FormatFooter = False
Else
    FormatFooter = True
End If


'--- save Log File ---
If FooterInfo(0) <> "" Then
    Call PrintLog(logFile, "Footer document title was successfully converted!")
Else
    Call PrintLog(logFile, "WARNING: It was not found Footer document title")
End If

If FooterInfo(1) <> "" Then
    Call PrintLog(logFile, "Footer document date was successfully converted!")
Else
    Call PrintLog(logFile, "WARNING: It was not found Footer document date")
End If

End Function


Function ClearTemplate(oDoc)

Dim oRange As Range

On Error GoTo ClearTemplate_exit
Set oDoc = ActiveDocument

Set oRange = Nothing
Set oRange = oDoc.Range(oDoc.Paragraphs(1).Range.Start, oDoc.Paragraphs(1).Range.End)
oRange.Select
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting

With Selection.Find
    .Text = "Word Converter Tool"
    .Replacement.Text = "|-->Title<--|"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False
End With
Selection.Find.Execute Replace:=wdReplaceAll


Set oRange = Nothing
Set oRange = oDoc.Range(oDoc.Paragraphs(2).Range.Start, oDoc.Paragraphs(2).Range.End)
oRange.Select
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting

With Selection.Find
    .Text = "Instruction"
    .Replacement.Text = "|-->Subtitle<--|"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = False
    .MatchFuzzy = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

Set oRange = oDoc.Range(oDoc.Paragraphs(4).Range.Start, oDoc.Paragraphs(4).Range.End)
oRange = ""
Set oRange = Nothing

Set oRange = oDoc.Range(oDoc.Paragraphs(8).Range.Start, oDoc.Paragraphs(27).Range.End)
oRange.Select
oRange = ""
Set oRange = Nothing

If 1 > 2 Then
ClearTemplate_exit:
    ClearTemplate = False
Else
    ClearTemplate = True
End If

End Function

Function ClearEmptyPages(oDoc, Optional OAPercent) As Boolean

Dim oRange As Range


On Error GoTo ClearEmptyPages_Exit
ClearEmptyPages = False
Count = 0
CheckEmptyPages:
lgPage = 1
p = 1


Do While p <= oDoc.Paragraphs.Count - 1
    '------- update progress ----
    If IsMissing(OAPercent) = False Then
        ProgressMultipleBarUF.Show
        ProgressMultipleBarUF.ShowProgress OAPercent, Round(p / (oDoc.Paragraphs.Count - 1), 3), True, "Finishing. Progress..."
        DoEvents
    
    Else
        ProgressBarUF.Show
        ProgressBarUF.ShowProgress Round(p / (oDoc.Paragraphs.Count - 1), 3), True, "Finishing. Progress..."
        DoEvents
    End If
        
    p0 = p
    pf = p
    
    lgPage = oDoc.Range(oDoc.Paragraphs(pf).Range.Start, oDoc.Paragraphs(pf).Range.End).Information(wdActiveEndAdjustedPageNumber)
    Do While (oDoc.Range(oDoc.Paragraphs(pf).Range.Start, oDoc.Paragraphs(pf).Range.End).Information(wdActiveEndAdjustedPageNumber) = lgPage And _
              pf <= oDoc.Paragraphs.Count - 1)
        pf = pf + 1
    Loop
    
    Set oRange = oDoc.Range(oDoc.Paragraphs(p0).Range.Start, oDoc.Paragraphs(pf - 1).Range.End)
    If Replace(Replace(Replace(Replace(oRange, "", ""), Chr(13), ""), vbFormFeed, ""), " ", "") = "" Then
        oRange.Select
        oRange.Delete
        Count = Count + 1
        If Count >= oDoc.Paragraphs.Count - 1 Then Exit Do
        GoTo CheckEmptyPages
    End If
    
    p = pf + 1
'    lgPage = lgPage + 1
Loop


'------- update progress ----
If IsMissing(OAPercent) = False Then
    ProgressMultipleBarUF.Show
    ProgressMultipleBarUF.ShowProgress OAPercent, Round(1, 3), True, "Finishing. Progress..."
    DoEvents

Else
    ProgressBarUF.Show
    ProgressBarUF.ShowProgress Round(1, 3), True, "Finishing. Progress..."
    DoEvents
End If

pf = oDoc.Paragraphs.Count
Set oRange = oDoc.Range(oDoc.Paragraphs(pf).Range.Start, oDoc.Paragraphs(pf).Range.End)
If Replace(Replace(Replace(oRange, "", ""), Chr(13), ""), " ", "") = "" Then
    oRange.Delete
End If



If 1 > 2 Then
ClearEmptyPages_Exit:
    If IsMissing(OAPercent) = False Then
        ClearEmptyPages = True
    Else
        ClearEmptyPages = False
    End If
    
Else
    ClearEmptyPages = True
End If

'------- update progress ----
If IsMissing(OAPercent) = False Then
    Unload ProgressMultipleBarUF
    DoEvents

Else
    Unload ProgressBarUF
    DoEvents
End If


End Function
