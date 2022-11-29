VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressMultipleBarUF 
   Caption         =   "Please Wait..."
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   OleObjectBlob   =   "ProgressMultipleBarUF.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressMultipleBarUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Sub ShowProgress(overallPercent, Percent As Single, ShowValue As Boolean, Optional status As String)
'    Const PAD = "                         "
     PAD = "                         "
    
    
    
    OverallStatusLabel.Caption = "Overall progress..."
    If ShowValue Then
        OverAlllabPg1v.Caption = PAD & Format(overallPercent, "0.0%")
        OverAlllabPg1va.Caption = OverAlllabPg1v.Caption
        OverAlllabPg1va.Width = Int((OveralllabPg1a.Width - (OverAlllabPg1v.Left - OveralllabPg1a.Left)) * overallPercent)
        OverAlllabPg1.Width = Int((OveralllabPg1a.Width - 5) * overallPercent)
    End If
    
    StatusLabel.Caption = status
    If ShowValue Then
        labPg1v.Caption = PAD & Format(Percent, "0.0%")
        labPg1va.Caption = labPg1v.Caption
        labPg1va.Width = Int((labPg1a.Width - (labPg1v.Left - labPg1a.Left)) * Percent)
        labPg1.Width = Int((labPg1a.Width - 5) * Percent)
    End If
    

'    labPg1.Width = Int(Val(labPg1.Tag) * Percent)
End Sub


Private Sub StatusLabel_Click()

End Sub
