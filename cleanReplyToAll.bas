Sub cleanDyanmic(ByVal Item As Object)
  Const endingAddress As String = "someaddress.com"
    
    Dim theType As String, senderEmail As String, i As Long
        theType = TypeName(Item)
        
        If theType = "MailItem" Or theType = "AppointmentItem" Or theType = "MeetingItem" Then
        
                senderEmail = Item.SendUsingAccount.DisplayName
                senderEmail = LCase(Mid(senderEmail, 1, InStr(1, senderEmail, "@", vbTextCompare)) _
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
