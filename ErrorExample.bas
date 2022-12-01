Sub exampleOfErrorHandling()

    Dim aResponse As String
        aResponse = InputBox("Enter something. Text will trigger an error while a nubmer will be accepted.")
    If aResponse = "" Then Exit Sub

    'programs procedure to jump to section problem with an error
    On Error GoTo ProblemZone
    Dim anyNumber As Integer 'variable will only accept number
        anyNumber = aResponse
        
    On Error GoTo 0 'sets errors to be handled in default method
    
    MsgBox anyNumber & " is a valid number"
    
    'section where other code would typically be inserted
    
    Exit Sub 'where normal code would end
    
ProblemZone:
    'section to handle errors
    Dim tryAgain As Long
    
    tryAgain = MsgBox(aResponse & " is not a number. Try again?", vbYesNo + vbCritical)
    
    If tryAgain = vbYes Then Call exampleOfErrorHandling

End Sub
