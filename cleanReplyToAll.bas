Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Const endingAddress As String = "@nttdata.com"

    Dim theType As String, senderEmail As String, i As Long
        theType = TypeName(Item)
        
        If theType = "MailItem" Or theType = "AppointmentItem" Or theType = "MeetingItem" Then
        
            senderEmail = LCase(Application.Session.Accounts.Item(1).UserName _
                & endingAddress)

            For i = Item.Recipients.Count To 1 Step -1
                With Item.Recipients(i)
                    If senderEmail = LCase(.Address) Then
                            .Delete
                            Exit For
                    End If
                End With
            Next i
        End If
End Sub
