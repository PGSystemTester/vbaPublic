Sub mathMultiplicationQuiz()
Dim answer As Long, maxMultiplier As Long, countOfQuestions As Long, _
i As Long, z As Long, y As Long, CorrectAnswers As Long, theTime As Double
    
    countOfQuestions = Application.InputBox("How many questions?", Type:=1)
    If countOfQuestions = 0 Then Exit Sub
    
    maxMultiplier = Application.InputBox("How high of multipliers?", , 12, Type:=1)
    If maxMultiplier = 0 Then Exit Sub

    maxMultiplier = maxMultiplier - 1
        
    theTime = Now
            
    For i = 1 To countOfQuestions
        z = Int(maxMultiplier * Rnd) + 2
        y = Int(maxMultiplier * Rnd) + 2
        
        answer = Application.InputBox(prompt:=z & " x " & y, Type:=1)
        
        CorrectAnswers = CorrectAnswers + -(answer = (z * y))
    
    Next i

    theTime = Round((Now() - theTime) * 24 * 3600, 1)

    MsgBox "You answered " & CorrectAnswers & " out of " & countOfQuestions & _
            " questions correctly (" & Round((CorrectAnswers / countOfQuestions) * 100, 0) _
            & "%) and finished in " & theTime & " seconds."

End Sub


Sub mathAddingQuiz()
Dim answer As Long, maxMultiplier As Long, countOfQuestions As Long, _
i As Long, z As Long, y As Long, CorrectAnswers As Long
Dim theTime As Double

    
    countOfQuestions = Application.InputBox("How many questions?", Type:=1)
    If countOfQuestions = 0 Then Exit Sub
    
    maxMultiplier = Application.InputBox("How many digits?", , 1, Type:=1)
    If maxMultiplier = 0 Then Exit Sub
    
        
    theTime = Now
            
    For i = 1 To countOfQuestions
        z = WorksheetFunction.RandBetween(1, 10 ^ maxMultiplier - 1)
        y = WorksheetFunction.RandBetween(1, 10 ^ maxMultiplier - 1)

        
        
        answer = Application.InputBox(prompt:=z & " + " & y, Type:=1)
        
        CorrectAnswers = CorrectAnswers + -(answer = (z + y))
    
    Next i

    theTime = Round((Now() - theTime) * 24 * 3600, 1)

    MsgBox "You answered " & CorrectAnswers & " out of " & countOfQuestions & _
            " questions correctly (" & Round((CorrectAnswers / countOfQuestions) * 100, 0) _
            & "%) and finished in " & theTime & " seconds."

End Sub
