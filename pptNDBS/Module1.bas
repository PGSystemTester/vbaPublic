Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long

Sub importPPTFunction()
Dim filename1, singlefilename As String
Dim layoutColor, footerColor
Dim srtMoveUp, topLimit
Dim auxSld As Slide

Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")

srtMoveUp = 25
''srtMoveUp = 0
topLimit = -10

selectLayoutUserForm.MoveupTextBox = srtMoveUp
selectLayoutUserForm.ToplimitTextBox = topLimit

selectLayoutUserForm.Show vbModeless 'show UF

Do While selectLayoutUserForm.Visible = True
  DoEvents
Loop

'If selectLayoutUserForm.OptionButton1 = True Then
'   layoutColor = 1 '"Blue"
'Else
'   layoutColor = 2 '"DarkGray"
'End If

layoutColor = selectLayoutUserForm.ColorLayoutComboBox1.ListIndex + 1
footerColor = selectLayoutUserForm.ColorFooterComboBox1.ListIndex + 1


'--- title and paragraph offset ---
srtMoveUp = selectLayoutUserForm.MoveupTextBox
topLimit = selectLayoutUserForm.ToplimitTextBox

'-----------------------

Unload selectLayoutUserForm

filename1 = ShowFileDialog

If filename1 = "" Then Exit Sub

'----- get single file name -------

singlefilename = Replace(Replace(fso.GetFilename(filename1), ".pptx", ""), ".ppt", "")
singlefilename = singlefilename & "-NTTDATA.pptx"
singlefilename = ShowSaveAsDialog(singlefilename)

If singlefilename = "" Then Exit Sub
'singlefilename = Split(filename1, "\")(UBound(Split(filename1, "\")))
'singlefilename = Split(singlefilename, ".")(UBound(Split(singlefilename, ".")) - 1)
'----------------------------------

Call deleteAllSlides

Call saveAsPPT(singlefilename)

Call copySlides(filename1)

Call editShapesInteli

Call changeLayout(layoutColor, footerColor)


Call moveShapeUp(srtMoveUp, topLimit)

Call copyCopyright 'add copyright slide

Call deleteAdditionalPattern
Call deleteAdditionalPattern2

Call bringTextToFront

Call bringTextToFront


Call checkFooters

ActivePresentation.Save


MsgBox "Done"


End Sub


Sub saveAsPPT(filename1 As String)

With Application.ActivePresentation
    '.SaveCopyAs "New Format Copy"
    '.SaveAs ActivePresentation.Path & "\" & filename1 & "-NTTDATA.pptx" ', ppSaveAsPowerPoint4
    .SaveAs filename1
    .BuiltInDocumentProperties.Item("title").Value = .Name
    
End With


End Sub

Sub deleteAllSlides()
Dim Pre As Presentation
Dim x As Long

Set Pre = ActivePresentation

'If Pre.Slides.Count > 1 Then
    For x = Pre.Slides.Count To 1 Step -1
        Pre.Slides(x).Delete
    Next x
'End If

'If Pre.Slides.Count < 1 Then
'   ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutCustom
'
'End If

End Sub

Sub copySlides(strFilename)

Dim objPresentation As Presentation
Dim thisPresentation As Presentation
Dim strProgress
Dim initSlide As Long

Dim i As Integer

Set thisPresentation = ActivePresentation
'open the target presentation

Set objPresentation = Presentations.Open(strFilename)

If objPresentation.Slides.Count < 1 Then Exit Sub 'exit if not enought slides

If LCase(ActivePresentation.Slides.Item(1).CustomLayout.Name) = "title" Then 'set init Slide
  initSlide = 2
Else
  initSlide = 1
End If

 initSlide = 1

For i = initSlide To objPresentation.Slides.Count

    '==== Update progress bar ===
    strProgress = i * 100 / objPresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Importing content, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============
   
    objPresentation.Slides.Item(i).Copy
    
    Call Wait(0.5)
    
    thisPresentation.Slides.Paste
    
    'Copy notes
    'thisPresentation.Slides(thisPresentation.Slides.Count).NotesPage.Shapes(2).TextFrame.TextRange = objPresentation.Slides(i).NotesPage.Shapes(2).TextFrame.TextRange 'copy notes
    
    Presentations.Item(1).Slides.Item(Presentations.Item(1).Slides.Count).Design = _
        objPresentation.Slides.Item(i).Design

    
    ''thisPresentation.Application.CommandBars.ExecuteMso ("PasteSourceFormatting")

    'Presentation.Item(1).Slides.Paste
 

Next i

objPresentation.Close

Unload ProgressBarUserForm

End Sub
Sub changeLayout(layoutType, footerType)
Dim i As Integer, j As Integer, k As Integer
Dim sourceSlideName As String, targetSlideName As String
Dim strProgress
Dim sld As Slide
Dim shapeArray, shapeFeatures
Dim shp As Shape
Dim macroPattern As Long
Dim specTxtArray

If footerType = 1 Then
  macroPattern = 7 'old macro layout
  ''macroPattern = 3
Else
  macroPattern = 6 'old macro layout
  'macroPattern = 5
End If
For i = 1 To ActivePresentation.Slides.Count

    Set sld = ActivePresentation.Slides.Item(i)
    
    'GoTo changeshape

    '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Checking layout, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

    sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
    
    shapeArray = getShapeFeatures(sld)
    
    Select Case sourceSlideName
    
    
    Case "title"
          
          If i > 1 Then
            GoTo anotherTitle
          End If
        
             ' xIndex = getLayoutIndexByName("Title Slide A (White NTT DATA)")
          'sld.CustomLayout = ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)
          ''sld.CustomLayout = ActivePresentation.Designs(7).SlideMaster.CustomLayouts(1)
          
              
          
          sld.CustomLayout = ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)
   
    
    Case "agenda"
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(1) 'Old macro layout
       ''sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2)
       
      
    
    Case "content 1", "content 2", "content 3", "1_content 1", "content & image 1", "content & image 2", "content & image 3"
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2) 'Old macro layout
       ''sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2)
    
    Case "two columns", "1_two columns"
       
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(3) 'Old macro layout
       'sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(3)
        
    Case "sheer", "titel en object"
       
        sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(4) 'Old macro layout
        'sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(1)
    
    Case "sub title", "1_sub title", "title" 'separating slides
anotherTitle:

'       If layoutType = 1 Then
'         'sld.CustomLayout = ActivePresentation.Designs(2).SlideMaster.CustomLayouts(1)
'         sld.CustomLayout = ActivePresentation.Designs(5).SlideMaster.CustomLayouts(2)
'       Else
'         'sld.CustomLayout = ActivePresentation.Designs(2).SlideMaster.CustomLayouts(5)
'         sld.CustomLayout = ActivePresentation.Designs(5).SlideMaster.CustomLayouts(17)
'       End If
       ''sld.CustomLayout = ActivePresentation.Designs(8).SlideMaster.CustomLayouts(layoutType)
       sld.CustomLayout = ActivePresentation.Designs(2).SlideMaster.CustomLayouts(layoutType)

        
    Case "final", "1_final"
       
       'sld.CustomLayout = ActivePresentation.Designs(4).SlideMaster.CustomLayouts(1)
       sld.CustomLayout = ActivePresentation.Designs(4).SlideMaster.CustomLayouts(layoutType)
       
    
    
    Case Else
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2) 'content Slide
    End Select
    
    'MsgBox ActivePresentation.Slides.Item(i).CustomLayout.Name
    
    
changeshape:
    Call getShapePositionBack(sld, sourceSlideName, shapeArray) 'change shape to its original position


    '-----  delete empty boxes ---------
    If sourceSlideName <> "sub title" And sourceSlideName <> "agenda" And sourceSlideName <> "1_sub title" And sourceSlideName <> "title" Then
      Call DeleteShapeWithSpecTxt(sld, "")
    End If
    '-----------------------------------
    
    '----- delete shapes with specific text -----
    specTxtArray = Array("© 2010 itelligence", "© 2011 itelligence", "© 2012 itelligence", "© 2013 itelligence", "© 2014 itelligence", "© 2015 itelligence", "© 2016 itelligence", "© 2017 itelligence", _
    "© 2018 itelligence", "© 2019 itelligence", "© 2020 itelligence", "© 2021 itelligence", "© 2022 itelligence", "© 2023 itelligence")
    For k = 0 To UBound(specTxtArray)
       Call DeleteShapeWithSpecTxt2(sld, CStr(specTxtArray(k)))
    Next k
    Call DeleteShapeWithSpecTxt2(sld, "We Transform. Trust into Value")
    '-------------------------------------------------
    
    '..... add empty boxes .....
'    If sourceSlideName = "final" Or sourceSlideName = "1_final" Then
'       With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=37, Top:=357, Width:=546, Height:=100).TextFrame
'         .AutoSize = ppAutoSizeNone 'do not autosize
'         .TextRange.Text = "Contact & Address"
'         .TextRange.Font.Name = "Arial"
'       End With
'    End If
    '.............................
    
    
    '..... delete image on separate and cover slide if exist ......
    If sourceSlideName = "sub title" Or sourceSlideName = "1_sub title" Or sourceSlideName = "title" Then
      For Each shp In sld.Shapes
         If shp.HasTextFrame = False Then
           shp.Delete
         End If
      Next
    End If
    '.....................................
    
    '..... send image back last sheet .....
    If sourceSlideName = "final" Or sourceSlideName = "1_final" Then
       For Each shp In sld.Shapes
         If shp.HasTextFrame = False Then
            shp.ZOrder msoSendToBack
         End If
       Next
      
    End If
    '...............................
    
'    '.... bring text to front ...
'    For Each shp In sld.Shapes
'         If shp.HasTextFrame = False Then
'           ''If Len(shp.TextFrame.TextRange.Text) > 1 Then
'             shp.ZOrder msoSendBehindText
'           ''End If
'         End If
'    Next
'    '..........................
    
    
   
    
    Call changeFontBulletColor(sld)
    'Call changeBulletFormat(sld)
    
     targetSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
    Call change2FinalLayout(sld, targetSlideName, sourceSlideName, footerType)
    
    Call getShapePositionBack(sld, sourceSlideName, shapeArray) 'change shape to its original position
    
Next i

Unload ProgressBarUserForm




End Sub

Sub change2FinalLayout(sld As Slide, targetSlideName As String, sourceSlideName As String, footerType)
Dim pivotShapeName, pivotShapeNameSub
Dim shp As Shape

On Error Resume Next
If footerType = 1 Then
 
  macroPattern = 3
Else
  
  macroPattern = 5
End If

'----- mapp text Cover -----
If Mid(targetSlideName, 1, 5) = "cover" Then
   pivotShapeName = getLowestCustomName(sld, "Custom Shape Name")
   pivotShapeNameSub = "Custom Shape Name " & pivotShapeName + 1
   pivotShapeName = "Custom Shape Name " & pivotShapeName
   
   
   For Each shp In sld.Shapes
      If shp.HasTextFrame And InStr(shp.Name, "Custom Shape Name") <> 0 And InStr(shp.Name, pivotShapeName) = 0 And InStr(shp.Name, pivotShapeNameSub) = 0 Then
          sld.Shapes(pivotShapeNameSub).TextFrame.TextRange.Text = sld.Shapes(pivotShapeNameSub).TextFrame.TextRange.Text & vbNewLine & vbNewLine & shp.TextFrame.TextRange.Text
          shp.TextFrame.TextRange.Text = ""
      End If
   Next
   
   Call DeleteShapeWithSpecTxt(sld, "")
End If
'--------------------
      
'----- mapp text Divider -----
If Mid(targetSlideName, 1, 7) = "divider" Then
   pivotShapeName = getLowestCustomName(sld, "Custom Shape Name")
   pivotShapeName = "Custom Shape Name " & pivotShapeName
   
   For Each shp In sld.Shapes
      If shp.HasTextFrame And InStr(shp.Name, "Custom Shape Name") <> 0 And InStr(shp.Name, pivotShapeName) = 0 Then
          sld.Shapes(pivotShapeName).TextFrame.TextRange.Text = sld.Shapes(pivotShapeName).TextFrame.TextRange.Text & vbNewLine & vbNewLine & shp.TextFrame.TextRange.Text
          shp.TextFrame.TextRange.Text = ""
      End If
   Next
   
  Call DeleteShapeWithSpecTxt(sld, "")
End If
'--------------------

Select Case targetSlideName


    
    Case "agenda macro"
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2)
       Call changeBullets12Numbers(sld)
    
    Case "content macro"
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2)
       Call DeleteShapeWithSpecTxt(sld, "")
    
    Case "two columns macro"
       
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(3)
       Call DeleteShapeWithSpecTxt(sld, "")
        
    Case "sheer macro"
       
        sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(1)
        Call DeleteShapeWithSpecTxt(sld, "")
    
    End Select

End Sub

Function getLayoutIndexByName(xName As String) As Integer


    ActivePresentation.Designs(1).SlideMaster.CustomLayouts.Item (1)
    
    With ActivePresentation.Designs(1).SlideMaster.CustomLayouts
        For i = 1 To .Count
            If .Item(i).Name = xName Then
            getLayoutIndexByName = i
            Exit Function
            End If
        Next
    End With

    End Function




Sub checkDesignMaster()

    Dim desName As Design



    With ActivePresentation

        For Each desName In .Designs

            MsgBox "The design name is " & .Designs.Item(desName.Index).Name

        Next
        
    End With



End Sub


Function ShowFileDialog()

On Error Resume Next

ShowFileDialog = ""

    Dim dlgOpen As FileDialog

    Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFilePicker)

    With dlgOpen
        .AllowMultiSelect = False
        .Show
        ShowFileDialog = dlgOpen.SelectedItems.Item(1)
    End With
    
    

End Function
Function ShowSaveAsDialog(Optional strFilename)

On Error Resume Next

ShowSaveAsDialog = ""

    Dim dlgSaveAs As FileDialog

    Set dlgSaveAs = Application.FileDialog(Type:=msoFileDialogSaveAs)

    With dlgSaveAs
        ''.AllowMultiSelect = False
        .InitialFileName = strFilename
        .Show
        ''ShowSaveAsDialog = dlgOpen.SelectedItems.Item(1)
        ShowSaveAsDialog = .SelectedItems.Item(1)
    End With
    
 
End Function


Sub editShapesInteli()
Dim shp As Shape
Dim sld As Slide
Dim i As Integer
Dim sourceSlideName As String

'inches x 72 to get points



For i = 1 To ActivePresentation.Slides.Count

    Set sld = ActivePresentation.Slides.Item(i)
    Call deleteShapesByPosition(sld, 762.48, 28.08) 'intelli top right
    Call deleteShapesByPosition(sld, 0, 0) 'intelli top left
    'Call deleteShapesByPosition(sld, 754.56, 276.48) 'date
    Call deleteShapesByPosition(sld, 923.04, 504) 'date
    
    Call deleteShapesByPosition(sld, 921.6, 0) 'page number
    'Call deleteShapesByPosition(sld, 754.56, 277.2) 'brand
    'Call deleteShapesByPosition(sld, 790, 260) 'brand
    
    'Call moveShapeUp(sld, 30) 'move shapes up
    
   
Next i



  

End Sub

Sub deleteShapesByPosition(sld As Slide, strLeft, srtTop)
Dim shp As Shape
Dim strNumber As Double

strNumber = 0.5
'inches x 72 to get points


    For Each shp In sld.Shapes
     
       If shp.Left >= strLeft * 0.95 - strNumber And shp.Left <= strLeft * 1.05 + strNumber And shp.Top >= srtTop * 0.95 - strNumber And shp.Top <= srtTop * 1.05 + strNumber Then
         
         shp.Delete
       
       End If
    Next



End Sub


Public Sub Wait(Seconds As Double)
    Dim endtime As Double
    endtime = DateTime.Timer + Seconds
    Do
        WaitMessage
        DoEvents
    Loop While DateTime.Timer < endtime
End Sub

Sub DeleteShapeWithSpecTxt(oSld As Slide, sSearch As String)
  Dim oShp As Shape
  Dim lShp As Long
  
  
  On Error GoTo errorhandler
  'If sSearch = "" Then sSearch = ActivePresentation.Slides(335).Shapes(4).TextFrame.TextRange.Text

  
  For lShp = oSld.Shapes.Count To 1 Step -1
      With oSld.Shapes(lShp)
        If .HasTextFrame And InStr(oSld.Shapes(lShp).Name, "Custom Shape") = 0 Then
          If StrComp(sSearch, .TextFrame.TextRange.Text) = 0 Then .Delete
        End If
      End With
  Next
  
Exit Sub
errorhandler:
  Debug.Print "Error in DeleteShapeWithSpecTxt : " & Err & ": " & Err.Description
  On Error GoTo 0
End Sub
Sub DeleteShapeWithSpecTxt2(oSld As Slide, sSearch As String)
  Dim oShp As Shape
  Dim lShp As Long
  
  
  On Error GoTo errorhandler
  'If sSearch = "" Then sSearch = ActivePresentation.Slides(335).Shapes(4).TextFrame.TextRange.Text

  
  For lShp = oSld.Shapes.Count To 1 Step -1
      With oSld.Shapes(lShp)
        If .HasTextFrame And InStr(oSld.Shapes(lShp).Name, "Custom Shape") <> 0 Then
          ''If LCase(sSearch) = LCase(Trim(.TextFrame.TextRange.Text)) Then .Delete
          If InStr(.TextFrame.TextRange.Text, sSearch) <> 0 Then .Delete
        End If
      End With
  Next
  
Exit Sub
errorhandler:
  Debug.Print "Error in DeleteShapeWithSpecTxt : " & Err & ": " & Err.Description
  On Error GoTo 0
End Sub

Sub moveShapeUp(srtMoveUp, topLimit)
  Dim oShp As Shape
  Dim oSld As Slide
  Dim lShp As Long
  Dim newTop As Double
  Dim i As Long
  ''Dim topLimit, srtMoveUp As Double
  Dim sourceSlideName As String
  
  
  'On Error Resume Next
  
'  srtMoveUp = 30
'  topLimit = -20

  'On Error GoTo errorhandler
  
For i = 1 To ActivePresentation.Slides.Count


  '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Moving up shapes, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

  Set oSld = ActivePresentation.Slides.Item(i)
  
  sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
  
  If (sourceSlideName = "cover slide 1 macro" Or sourceSlideName = "cover slide 14 macro") And i = 1 Then 'do not affect cover
     GoTo skipShape
  End If
  
  
  
  For lShp = oSld.Shapes.Count To 1 Step -1
  
      If lShp < 1 Then GoTo skipSlide
      
      With oSld.Shapes(lShp)
      
           If .HasTextFrame = True And (.Height >= 530 Or .Top <= topLimit) Then GoTo skipShape  '(.Top <= 30 And .Width >= 900) Or
           
           If .Top < 20 And .Width > 850 Then GoTo skipShape  'do not change  titles
           
           newTop = .Top - srtMoveUp
           
'           If newTop < topLimit And .HasTextFrame = True Then
'             GoTo skipShape
'           End If
           
           
          .Top = newTop
          
            
          '----- set limit ----
          If .Top < topLimit And .HasTextFrame = True Then
             .Top = topLimit
          End If
          '-------------------------------
       
      End With
skipShape:
  Next
skipSlide:
Next i

Unload ProgressBarUserForm

End Sub

Sub bringTextToFront()
  Dim oShp As Shape
  Dim oSld As Slide
  Dim lShp As Long
  Dim newTop As Double
  Dim i As Long
  ''Dim topLimit, srtMoveUp As Double
  Dim sourceSlideName As String
  
  
  'On Error Resume Next
  
'  srtMoveUp = 30
'  topLimit = -20

  'On Error GoTo errorhandler
  
For i = 1 To ActivePresentation.Slides.Count


  '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Bring text to front, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

  Set oSld = ActivePresentation.Slides.Item(i)
  
  sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
  
'  If (sourceSlideName = "cover slide 1 macro" Or sourceSlideName = "cover slide 14 macro") And i = 1 Then 'do not affect cover
'     GoTo skipShape
'  End If
  
  
      '.... bring text to front ...
    For Each oShp In oSld.Shapes
         If oShp.HasTextFrame = True Then
           If Len(oShp.TextFrame.TextRange.Text) > 1 Then
             oShp.ZOrder msoBringToFront
           End If
         End If
    Next
    '..........................
  

Next i

Unload ProgressBarUserForm

End Sub

Function getShapeFeatures(sld As Slide)

Dim shp As Shape
'Dim sld As Slide
Dim shapeArray(), counter As Long
Dim x As Long

'inches x 72 to get points
counter = 0
ReDim shapeArray(counter)
shapeArray(counter) = ""

''Set sld = ActivePresentation.Slides.Item(1)

''Debug.Print ActivePresentation.Slides.Item(1).Shapes("Title 1").Top
For Each shp In sld.Shapes
  
    shp.Name = "Custom Shape Name " & counter
    ReDim Preserve shapeArray(counter)
    shapeArray(counter) = shp.Name & ";" & shp.Top & ";" & shp.Left & ";" & shp.Height & ";" & shp.Width
    'shapeArray(counter) = shp & ";" & shp.Top & ";" & shp.Left & ";" & shp.Height & ";" & shp.Width
    counter = counter + 1
  If shp.Type = msoGroup Then
     
     For x = 1 To shp.GroupItems.Count
     'MsgBox shp.GroupItems.Count
        
        shp.GroupItems(x).Name = "Custom Shape Name " & counter
        ReDim Preserve shapeArray(counter)
        shapeArray(counter) = shp.GroupItems(x).Name & ";" & shp.GroupItems(x).Top & ";" & shp.GroupItems(x).Left & ";" & shp.GroupItems(x).Height & ";" & shp.GroupItems(x).Width
        counter = counter + 1
     Next x
  End If
Next

getShapeFeatures = shapeArray
'Debug.Print shapeArray(0)

End Function

Sub changeFontBulletColor(sld As Slide)
'Dim sld As Slide
Dim i As Long, j As Long, k As Long
Dim x As Long
Dim oTbl
Dim currentSlideName
Dim rustRedRGB

rustRedRGB = RGB(188, 67, 40)


currentSlideName = sld.CustomLayout.Name

 'Set sld = ActivePresentation.Slides.Item(3)
 

For lShp = sld.Shapes.Count To 1 Step -1
      With sld.Shapes(lShp)
        If .HasTextFrame Then
           
           '....... change blue box that was set by mistake .......
           If .Fill.ForeColor.RGB = RGB(0, 128, 177) Then
              .Fill.ForeColor.RGB = RGB(255, 255, 255)
           End If
           '........................................
           
           '............  Change text color and size ........
           If Len(.TextFrame.TextRange.Text) > 1 Then
                .TextFrame.TextRange.Font.Name = "Arial"
                For i = 1 To Len(.TextFrame.TextRange.Text)
                   .TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(.TextFrame.TextRange.Characters(i).Font.Color.RGB)  'Change to Blue
                   
                   '/////   size  /////////
                   If Mid(LCase(currentSlideName), 1, 7) <> "divider" Then
                     If .TextFrame.TextRange.Characters(i).Font.Size > 24 Then
                        '.TextFrame.TextRange.Characters(i).Font.Size = 24
                     End If
                   End If
                   '///////////////////////
                   
                Next
           End If
           '............................................
           
           '......... Change bullet color ..................
           For i = 1 To .TextFrame.TextRange.Paragraphs.Count
             If .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Font.Color = rustRedRGB 'getColorConversion(.TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Font.Color.RGB)  'Change to Blue
             End If
           Next i
           '.......................................................
        
        ElseIf .HasTable Then
          Set oTbl = sld.Shapes(lShp).Table
            For k = 1 To oTbl.Columns.Count
              For j = 1 To oTbl.Rows.Count
                
                Call changeShapeColor(sld.Shapes(lShp).Table.Cell(j, k).Shape) 'Change cell color
                
                With oTbl.Cell(j, k).Shape.TextFrame.TextRange
                  '...... change color text in tables .......
                  '.Size = 12
                  .Font.Name = "Arial"
                  For i = 1 To Len(.Text)
                      .Characters(i).Font.Color = getColorConversion(.Characters(i).Font.Color.RGB) 'Change to Blue
                  Next i
                  '.Bold = True
                  '.....................................
                  
                  
                  '......... Change bullet color ..................
                  For i = 1 To .Paragraphs.Count
                      If .Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                         .Paragraphs(i).ParagraphFormat.Bullet.Font.Color = rustRedRGB ''getColorConversion(.Paragraphs(i).ParagraphFormat.Bullet.Font.Color.RGB)  'Change to Blue
                      End If
                  Next i
                  '.......................................................
                  
                  
                End With
              Next j
            Next k
                   
           
        End If
        
        '---- shape color -----
        Dim oShpNode
        Dim oNode As SmartArtNode
        If .HasSmartArt Then
           For Each oNode In .SmartArt.Nodes
              For Each oShpNode In oNode.Shapes ' As ShapeRange
                 Call changeShapeColor(oShpNode)
              Next
           Next
        End If
        
        
        On Error Resume Next
        If .HasTable = False Then
            If .Type <> msoGroup Then
              Call changeShapeColor(sld.Shapes(lShp))
            Else
    
               'Debug.Print "GROUP"
               For x = 1 To sld.Shapes(lShp).GroupItems.Count
                   Call changeShapeColor(sld.Shapes(lShp).GroupItems(x))
                   
                  
                   
                   '****** check for texts*****************
                   If sld.Shapes(lShp).GroupItems(x).HasTextFrame Then
                      If Len(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Text) > 1 Then
                         For i = 1 To Len(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Text)
                            sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Characters(i).Font.Color.RGB)  'Change to Blue
                         Next i
                      End If
                      
                      
                      '............ Change bullet color ...........................
                      For i = 1 To sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs.Count
                        If sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                           sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Font.Color = rustRedRGB '' getColorConversion(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Font.Color.RGB)  'Change to Blue
                        End If
                      Next i
                      '..........................................................
                   
                   End If
                   '***************************************
                   
                   '******* has chart ************
                   On Error Resume Next
                   If sld.Shapes(lShp).GroupItems(x).HasChart Then
                     If sld.Shapes(lShp).GroupItems(x).Chart.ChartType = 51 Then
                     
                        ''shp.Chart.SeriesCollection(1).DataLabels.Font.Color = RGB(0, 0, 0)
                        sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Interior.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Interior.Color)
                        sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Border.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Border.Color)
                     End If
                   End If
                   On Error GoTo 0
                   
                   '******************************
                   
                   
               Next x
            End If
        End If
        '--------------------------
      
      End With
Next
 
 

End Sub
 


Sub changeShapeColor(oSh)

With oSh
  On Error GoTo skip1
  If .Fill.Visible = msoTrue And CStr(.Fill.ForeColor.RGB) <> 0 Then
     .Fill.ForeColor.RGB = getColorConversion(.Fill.ForeColor.RGB)
     .Fill.BackColor.RGB = getColorConversion(.Fill.BackColor.RGB)
  End If

skip1:
  On Error GoTo skip2
  If .Line.Visible = msoTrue And CStr(.Line.ForeColor.RGB) <> 0 Then
     .Line.ForeColor.RGB = getColorConversion(.Line.ForeColor.RGB)
     '.Line.ForeColor.RGB = .Fill.ForeColor.RGB
  End If
  
skip2:

End With

End Sub


Function getColorConversion(rgbColor)
Dim SourceRGBColor, DestRGBColor
Dim i As Long
Dim intRedRGB, intRedRGB2, intRedRGB3, intBlackRGB
Dim intGrayRGB, intSilverRGB, baseGrayRGB, maroonRGB, greenRGB, coralRGB, turqueisRGB, orangeRGB, blueRGB
Dim altIntRedRGB, altIntBlackRGB, altIntGrayRGB, altIntSilverRGB, altBaseGrayRGB
Dim altMaroonRGB, altGreenRGB, altCoralRGB, altTurqueisRGB, altOrangeRGB, altBlueRGB
Dim lightPinkRGB, altLightPinkRGB


'----- source colors ----------
intRedRGB = RGB(212, 0, 48)
intRedRGB2 = RGB(213, 16, 48)
intRedRGB3 = RGB(192, 0, 0)
intRedRGB4 = RGB(255, 0, 0)
intRedRGB5 = RGB(212, 0, 48)
intBlackRGB = RGB(0, 0, 0)
intGrayRGB = RGB(134, 128, 125)
intSilverRGB = RGB(222, 222, 222)
baseGrayRGB = RGB(236, 236, 237)

maroonRGB = RGB(150, 78, 95)
greenRGB = RGB(117, 182, 95)
coralRGB = RGB(211, 110, 101)
turqueisRGB = RGB(64, 163, 145)
orangeRGB = RGB(244, 148, 63)
blueRGB = RGB(61, 104, 153)
lightPinkRGB = RGB(247, 231, 232)
'-------------------------------

'----- alt colors ----------

altIntRedRGB = RGB(15, 28, 80)
altIntBlackRGB = RGB(0, 0, 0)
altIntGrayRGB = RGB(28, 28, 28)
altIntSilverRGB = RGB(194, 206, 230)
altBaseGrayRGB = RGB(141, 141, 141)

altMaroonRGB = RGB(170, 60, 128)
altGreenRGB = RGB(131, 178, 84)
altCoralRGB = RGB(188, 67, 40)
altTurqueisRGB = RGB(103, 133, 193)
altOrangeRGB = RGB(230, 182, 0)
altBlueRGB = RGB(0, 128, 177)
altLightPinkRGB = RGB(220, 226, 248)

'-----------------------------

SourceRGBColor = Array(intRedRGB, intRedRGB2, intRedRGB3, intRedRGB4, intRedRGB5, intBlackRGB, intGrayRGB, intSilverRGB, baseGrayRGB, _
maroonRGB, greenRGB, coralRGB, turqueisRGB, orangeRGB, blueRGB, lightPinkRGB)

DestRGBColor = Array(altIntRedRGB, altIntRedRGB, altIntRedRGB, altIntRedRGB, altIntRedRGB, altIntBlackRGB, altIntGrayRGB, altIntSilverRGB, altBaseGrayRGB, _
altMaroonRGB, altGreenRGB, altCoralRGB, altTurqueisRGB, altOrangeRGB, altBlueRGB, altLightPinkRGB)

getColorConversion = rgbColor


For i = 0 To UBound(SourceRGBColor)
 If SourceRGBColor(i) = rgbColor Then
     getColorConversion = DestRGBColor(i)
     Exit Function
 End If
Next i

End Function


Sub deleteAdditionalPattern()
Dim layoutNumber As Integer, i As Integer
Dim maxLayoutNumber As Integer

maxLayoutNumber = 3

layoutNumber = ActivePresentation.Designs(1).SlideMaster.CustomLayouts.Count

If layoutNumber > maxLayoutNumber Then
  For i = layoutNumber To maxLayoutNumber + 1 Step -1
     ActivePresentation.Designs(1).SlideMaster.CustomLayouts(i).Delete
  Next i
End If

'MsgBox ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)
'MsgBox ActivePresentation.Designs(1).SlideMaster.CustomLayouts.Count
End Sub

Sub deleteAdditionalPattern2()
Dim layoutNumber As Integer, i As Integer
Dim maxLayoutNumber As Integer

maxLayoutNumber = 5

layoutNumber = ActivePresentation.Designs.Count

If layoutNumber > maxLayoutNumber Then
  For i = layoutNumber To maxLayoutNumber + 1 Step -1
     ActivePresentation.Designs(i).Delete
  Next i
End If

'MsgBox ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)
'MsgBox ActivePresentation.Designs(1).SlideMaster.CustomLayouts.Count
End Sub

Sub checkFooters()
  Dim oShp As Shape
  Dim oSld As Slide
  Dim lShp As Long
  Dim newTop As Double
  Dim i As Long
  ''Dim topLimit, srtMoveUp As Double
  Dim sourceSlideName As String
  
  
  On Error Resume Next
  
'  srtMoveUp = 30
'  topLimit = -20

  'On Error GoTo errorhandler
  
For i = 1 To ActivePresentation.Slides.Count


  '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Check Footers, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

  Set oSld = ActivePresentation.Slides.Item(i)
  
  sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
  
'  If (sourceSlideName = "cover slide 1 macro" Or sourceSlideName = "cover slide 14 macro") And i = 1 Then 'do not affect cover
'     GoTo skipShape
'  End If
  
  
   ' .....FOOTER.......

    With oSld.HeadersFooters
    
        .Footer.Visible = True
    
        '.Footer.Text = "Regional Sales"
    
        .SlideNumber.Visible = True
    
        .DateAndTime.Visible = True
    
        .DateAndTime.UseFormat = True
    
        .DateAndTime.Format = ppDateTimeMdyy
    
    End With
    
    ''............
  

Next i

Unload ProgressBarUserForm

End Sub

Sub getShapePositionBack(sld As Slide, sourceSlideName As String, shapeArray)
Dim shapeFeatures

'------  get the shape back -------
On Error Resume Next

'If sourceSlideName = "title" And i = 1 Then 'skip change shape back
If sourceSlideName = "title" Or sourceSlideName = "sub title" Or sourceSlideName = "1_sub title" Then
   GoTo skipchangeshape
End If

changeshape:
If shapeArray(0) <> "" Then

    For j = 0 To UBound(shapeArray)
      shapeFeatures = Split(shapeArray(j), ";")
      If sld.Shapes(shapeFeatures(0)).Top <= 500 Then   'do not change footers
         
        If sld.Shapes(shapeFeatures(0)).Top < 20 And sld.Shapes(shapeFeatures(0)).Width > 850 Then GoTo nextshape  'do not change  titles
      
        'sld.Shapes(shapeFeatures(0)).TextFrame.AutoSize = ppAutoSizeNone 'do not autosize
        'sld.Shapes(shapeFeatures(0)).TextFrame.AutoSize = ppAutoSizeMixed 'do not autosize
        sld.Shapes(shapeFeatures(0)).TextFrame2.AutoSize = 2 'do not autosize
      
        sld.Shapes(shapeFeatures(0)).Top = shapeFeatures(1)
        sld.Shapes(shapeFeatures(0)).Left = shapeFeatures(2)
        
        On Error GoTo 0
        On Error Resume Next
        If sld.Shapes(shapeFeatures(0)).HasTextFrame = True And Len(sld.Shapes(shapeFeatures(0)).TextFrame.TextRange.Text) > 1 Then
          If Err.Number = 0 Then
            sld.Shapes(shapeFeatures(0)).Height = shapeFeatures(3)
            sld.Shapes(shapeFeatures(0)).Width = shapeFeatures(4)
          End If
        End If
      
      
      Else
      
        'sld.Shapes(shapeFeatures(0)).TextFrame.TextRange.Font.Color = RGB(255, 255, 255) 'set white
      
      End If
nextshape:
    Next j
End If
On Error GoTo 0
'----------------------------------

skipchangeshape:
End Sub


Function getLowestCustomName(sld As Slide, srtCustomName)
Dim shp As Shape
Dim strNumber
Dim auxString

getLowestCustomName = ""
auxString = ""

For Each shp In sld.Shapes
    If shp.HasTextFrame And InStr(shp.Name, srtCustomName) <> 0 Then
    
        strNumber = Trim(Replace(shp.Name, srtCustomName, ""))
        
        If IsNumeric(strNumber) Then
           If auxString = "" Then
             auxString = strNumber
           Else
             If strNumber < auxString Then
               auxString = strNumber
             End If
           End If
        
        End If
    
    End If
  

Next

getLowestCustomName = auxString
End Function

Sub copyCopyright()

Dim copyCopyrightSld As Slide, newSld As Slide
Dim shp As Shape, oSh As Shape
Dim i As Long

'copyright at the end
Set newSld = ActivePresentation.Slides.AddSlide(Index:=ActivePresentation.Slides.Count + 1, pCustomLayout:=ActivePresentation.Designs(3).SlideMaster.CustomLayouts(2))

'Set copyCopyrightSld = ActivePresentation.Designs(8).SlideMaster.CustomLayouts(1)

For Each shp In ActivePresentation.Designs(8).SlideMaster.CustomLayouts(1).Shapes
   
    shp.Copy

    Set oSh = newSld.Shapes.Paste(1)
Next


For i = newSld.Shapes.Count To 1 Step -1
      With newSld.Shapes(i)
        If .HasTextFrame Then
          If .TextFrame.TextRange.Text = "" Then .Delete
        End If
      End With
Next

' ____ shape positioin __---
For i = newSld.Shapes.Count To 1 Step -1
      With newSld.Shapes(i)
         If .Name = "Titel 1" Then
            .Top = 18
            .Left = 30.24
            
         ElseIf .Name = "Textplatzhalter 5" Then
            .Top = 90
            .Left = 30.24
         End If
      End With
Next
'_________________

End Sub

Sub changeBullets12Numbers(sld As Slide)
Dim i As Long

On Error Resume Next


For lShp = sld.Shapes.Count To 1 Step -1

   With sld.Shapes(lShp)

           For i = 1 To .TextFrame.TextRange.Paragraphs.Count
             If .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
               
               With .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet
                  
                  '.Character = 8226
                  .Type = 2
               End With
             End If
           Next i

  End With
Next


End Sub
