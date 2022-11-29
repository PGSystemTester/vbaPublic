VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarUF 
   Caption         =   "Please Wait..."
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   OleObjectBlob   =   "ProgressBarUF.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBarUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Sub ShowProgress(Percent As Single, ShowValue As Boolean, Optional status As String)
'    Const PAD = "                         "
     PAD = "                         "
    StatusLabel.Caption = status
    If ShowValue Then
        labPg1v.Caption = PAD & Format(Percent, "0.0%")
        labPg1va.Caption = labPg1v.Caption
        labPg1va.Width = Int((labPg1a.Width - (labPg1v.Left - labPg1a.Left)) * Percent)
        labPg1.Width = Int((labPg1a.Width - 5) * Percent)
    End If
    

'    labPg1.Width = Int(Val(labPg1.Tag) * Percent)
End Sub

Private Sub labPg1va_Click()

End Sub
