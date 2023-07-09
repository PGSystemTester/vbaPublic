Private Sub CommandButton1_Click()

'-----input validation
If IsNumeric(MoveupTextBox) = False Or IsNumeric(ToplimitTextBox) = False Then

   MsgBox "Please enter only numeric values on Titles and Paragraphs Offset"
   Exit Sub

End If
'-------------------------

Me.Hide


End Sub


Private Sub CommandButton2_Click()
End
End Sub



Private Sub UserForm_Initialize()
Dim i As Long

'----- Fill in combobox  -------
For i = 1 To 11

''ColorLayoutComboBox1.AddItem Replace(Replace(ActivePresentation.Designs(8).SlideMaster.CustomLayouts(i).Name, "Divider - ", ""), "Macro", "")
ColorLayoutComboBox1.AddItem Replace(ActivePresentation.Designs(2).SlideMaster.CustomLayouts(i).Name, "Divider - ", "")

Next i
ColorLayoutComboBox1.ListIndex = 3

ColorFooterComboBox1.List = Array("Dark Blue") ', "Human Blue"
ColorFooterComboBox1.ListIndex = 0
'---------------------------------

End Sub
