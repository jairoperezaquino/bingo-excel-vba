Private Sub btnChocolatear_Click()
    Dim answer As Integer
    
    ReDim RandomNumbers(1) As Integer
    Dim Numbers_Bot As Integer
    Dim Numbers_Top As Integer

    Numbers_Bot = 1
    Numbers_Top = 75
    
    If Range("F2") <> "" Then
        answer = MsgBox("Juego en curso, deseas reiniciarlo?", vbQuestion + vbYesNo)
        
        If answer = vbYes Then
            '******************CLEAN FORMAT******************'
            Call GenericFormatCleanUp
            Sheets("DB").Range("A2:D76") = ""
            
            '******************GENERATE NUMBERS AND BALLS******************'
            RandomNumbers = GenerateRandomNumbers(Numbers_Bot, Numbers_Top)
            
            'Fill DB with IdGame, Number Ball, and PlayedFlag
            For i = 1 To Numbers_Top
                Sheets("DB").Cells(i + 1, 1) = 1
                Sheets("DB").Cells(i + 1, 2) = RandomNumbers(i)
                Sheets("DB").Cells(i + 1, 3) = GenerateBall(RandomNumbers(i))
                Sheets("DB").Cells(i + 1, 4) = 0
            Next i
        End If
    Else
        '******************CLEAN FORMAT******************'
        Call GenericFormatCleanUp
        Sheets("DB").Range("A2:D76") = ""
          
        '******************GENERATE NUMBERS AND BALLS******************'
        RandomNumbers = GenerateRandomNumbers(Numbers_Bot, Numbers_Top)
            
        'Fill DB with IdGame, Number Ball, and PlayedFlag
        For i = 1 To Numbers_Top
            Sheets("DB").Cells(i + 1, 1) = 1
            Sheets("DB").Cells(i + 1, 2) = RandomNumbers(i)
            Sheets("DB").Cells(i + 1, 3) = GenerateBall(RandomNumbers(i))
            Sheets("DB").Cells(i + 1, 4) = 0
        Next i
    End If

    
End Sub
Private Sub btnSacar_Click()
    'Settings
    Dim AutoShowBall As String
    Dim AutoShowBallDelay As String
    Dim AutoMarkCard As String
    Dim ReadBall As String
    Dim TimeToWait As String
    
    AutoShowBall = Sheets("Settings").Range("B2")
    If AutoShowBall = "SI" Then
        AutoShowBallDelay = Sheets("Settings").Range("C2")
        TimeToWait = "0:00:" & Right("00" & CStr(AutoShowBallDelay), 2)
    End If
    AutoMarkCard = Sheets("Settings").Range("B3")
    ReadBall = Sheets("Settings").Range("B4")
    
    If Not ((AutoShowBall = "NO" Or AutoShowBall = "SI") And (AutoMarkCard = "NO" Or AutoMarkCard = "SI") And (ReadBall = "NO" Or ReadBall = "SI")) Then
        MsgBox "Error de configuración"
        Exit Sub
    End If
    
    '******************CLEAN FORMAT******************'
    If AutoShowBall = "SI" Then Call GenericFormatCleanUp
        
    '******************SHOW NUMBERS AND BALLS******************'
    Dim FirstGeneratedNumberRow As Integer
    Dim LastGeneratedNumberRow As Integer
    LastGeneratedNumberRow = Sheets("DB").Range("A1").End(xlDown).Row
    
    For i = 2 To LastGeneratedNumberRow
        If Sheets("DB").Range("D" & i) = 0 Then
            FirstGeneratedNumberRow = i
            Exit For
        End If
    Next i
    
    'Check if all generated numbers were played
    If FirstGeneratedNumberRow = 0 Then
        MsgBox "Juego terminado. Debes chocolatear de nuevo"
        Exit Sub
    End If
    
    'GeneratedCard data
    Dim CardFirstRow As Integer
    Dim CardLastRow As Integer
    CardFirstRow = Sheets("DB").Range("F1").End(xlDown).Offset(-23, 0).Row
    CardLastRow = Sheets("DB").Range("F1").End(xlDown).Row
        
    Dim ListInitialRow As Integer
    Dim ListInitialColumn As Integer
    Dim Step As Integer
    Dim ListTargetRow As Integer
    Dim ListTargetColumn As Integer
       
    ListInitialRow = 2
    ListInitialColumn = 18 'Column R

    Dim GeneratedNumber As Integer
    Dim GeneratedBall As String
    Dim GeneratedBall_Minus1 As String
    Dim GeneratedBall_Minus2 As String
    Dim GeneratedBall_Minus3 As String
    
    'Check AutoShowBall
    If AutoShowBall = "SI" Then
       For i = FirstGeneratedNumberRow To LastGeneratedNumberRow
        GeneratedNumber = Sheets("DB").Range("B" & i)
        GeneratedBall = Sheets("DB").Range("C" & i)
        
        'Speak
        If ReadBall = "SI" Then
            Application.Speech.Speak (GeneratedBall), SpeakAsync:=True
        End If
        
        'Show Ball
        Range("F2") = GeneratedBall
        
        'Show Previous Balls
        If i >= 5 Then
            GeneratedBall_Minus1 = Sheets("DB").Range("C" & i - 1)
            GeneratedBall_Minus2 = Sheets("DB").Range("C" & i - 2)
            GeneratedBall_Minus3 = Sheets("DB").Range("C" & i - 3)
            Range("B13") = GeneratedBall_Minus1
            Range("G13") = GeneratedBall_Minus2
            Range("L13") = GeneratedBall_Minus3
        ElseIf i >= 4 Then
            GeneratedBall_Minus1 = Sheets("DB").Range("C" & i - 1)
            GeneratedBall_Minus2 = Sheets("DB").Range("C" & i - 2)
            Range("B13") = GeneratedBall_Minus1
            Range("G13") = GeneratedBall_Minus2
        ElseIf i >= 3 Then
            GeneratedBall_Minus1 = Sheets("DB").Range("C" & i - 1)
            Range("B13") = GeneratedBall_Minus1
        End If
        
        'Paint List
        Step = Application.WorksheetFunction.RoundUp(GeneratedNumber / 5, 0) - 1
        ListTargetRow = ListInitialRow + Step
        ListTargetColumn = ListInitialColumn + GeneratedNumber - (5 * Step) - 1
        
        Cells(ListTargetRow, ListTargetColumn).Interior.Color = RGB(229, 186, 181)
        
        'Mark Card
        If AutoMarkCard = "SI" Then
            Call MarkCard(CardFirstRow, CardLastRow, GeneratedNumber)
        End If
        
        'Mark as Played in DB
        Sheets("DB").Range("D" & i) = 1
        
        'Wait
        Application.ScreenUpdating = True
        Application.Wait (Now + TimeValue(TimeToWait))
        Next i
        
        MsgBox "Se acabó el bingo y nadie gritó bingo, gilipollas!"
    Else
        GeneratedNumber = Sheets("DB").Range("B" & FirstGeneratedNumberRow)
        GeneratedBall = Sheets("DB").Range("C" & FirstGeneratedNumberRow)

        'Speak
        If ReadBall = "SI" Then
            Application.Speech.Speak (GeneratedBall), SpeakAsync:=True
        End If
        
        'Show Ball
        Range("F2") = GeneratedBall

        'Show Previous Balls
        If FirstGeneratedNumberRow >= 5 Then
            GeneratedBall_Minus1 = Sheets("DB").Range("C" & FirstGeneratedNumberRow - 1)
            GeneratedBall_Minus2 = Sheets("DB").Range("C" & FirstGeneratedNumberRow - 2)
            GeneratedBall_Minus3 = Sheets("DB").Range("C" & FirstGeneratedNumberRow - 3)
            Range("B13") = GeneratedBall_Minus1
            Range("G13") = GeneratedBall_Minus2
            Range("L13") = GeneratedBall_Minus3
        ElseIf FirstGeneratedNumberRow >= 4 Then
            GeneratedBall_Minus1 = Sheets("DB").Range("C" & FirstGeneratedNumberRow - 1)
            GeneratedBall_Minus2 = Sheets("DB").Range("C" & FirstGeneratedNumberRow - 2)
            Range("B13") = GeneratedBall_Minus1
            Range("G13") = GeneratedBall_Minus2
        ElseIf FirstGeneratedNumberRow >= 3 Then
            GeneratedBall_Minus1 = Sheets("DB").Range("C" & FirstGeneratedNumberRow - 1)
            Range("B13") = GeneratedBall_Minus1
        End If
        
        'Paint List
        Step = Application.WorksheetFunction.RoundUp(GeneratedNumber / 5, 0) - 1
        ListTargetRow = ListInitialRow + Step
        ListTargetColumn = ListInitialColumn + GeneratedNumber - (5 * Step) - 1
        
        Cells(ListTargetRow, ListTargetColumn).Interior.Color = RGB(229, 186, 181)
        
        'Mark Card
        If AutoMarkCard = "SI" Then
            Call MarkCard(CardFirstRow, CardLastRow, GeneratedNumber)
        End If
        
        'Mark as Played in DB
        Sheets("DB").Range("D" & FirstGeneratedNumberRow) = 1
    End If
    
End Sub
Private Sub btnGenerarCartilla_Click()
    '******************CLEAN FORMAT******************'
    Call GenericFormatCleanUp
    Range("Z6:AN20") = ""
    Sheets("DB").Range("F2:I25") = ""
    
    '******************GENERATE NUMBERS AND BALLS******************'
    'Define the range of numbers
    ReDim RandomNumbers(1) As Integer
    Dim Numbers_Bot As Integer
    Dim Numbers_Top As Integer
    
    Numbers_Bot = 1
    Numbers_Top = 75
        
    RandomNumbers = GenerateRandomNumbers(Numbers_Bot, Numbers_Top)
        
    '******************SHOW NUMBERS******************'
    Dim CardInitialRow As Integer
    Dim CardInitialColumn As Integer
    Dim Step As Integer
    Dim CardTargetRow As Integer
    Dim CardTargetColumn As Integer
    
    Dim Number As Integer
    Dim Numbers_B As Integer
    Dim Numbers_I As Integer
    Dim Numbers_N As Integer
    Dim Numbers_G As Integer
    Dim Numbers_O As Integer
    
    Dim DBRowIndex As Integer
    
    CardInitialRow = 6
    CardInitialColumn = 26 'Column Z
    
    For i = 1 To Numbers_Top
        Number = RandomNumbers(i)
        Step = ((Application.WorksheetFunction.RoundUp(Number / 15, 0) - 1) * 5)
        
        If Numbers_B < 5 Or Numbers_I < 5 Or Numbers_N Or Numbers_G < 5 Or Numbers_O < 5 Then
            If Number >= 1 And Number <= 15 Then
                If Numbers_B < 5 Then
                    CardTargetRow = CardInitialRow + (Numbers_B * 3)
                    CardTargetColumn = CardInitialColumn
                    
                    Cells(CardTargetRow, CardTargetColumn) = Number
                    
                    'Fill DB
                    DBRowIndex = Step + Numbers_B + 1
                    Sheets("DB").Cells(DBRowIndex + 1, 6) = 1
                    Sheets("DB").Cells(DBRowIndex + 1, 7) = Number
                    Sheets("DB").Cells(DBRowIndex + 1, 8) = CardTargetRow
                    Sheets("DB").Cells(DBRowIndex + 1, 9) = CardTargetColumn
                    DBRowIndex = DBRowIndex + 1
                                   
                    Numbers_B = Numbers_B + 1
                End If
            ElseIf Number > 15 And Number <= 30 Then
                If Numbers_I < 5 Then
                    CardTargetRow = CardInitialRow + (Numbers_I * 3)
                    CardTargetColumn = CardInitialColumn + 3
                        
                    Cells(CardTargetRow, CardTargetColumn) = Number
                    
                    'Fill DB
                    DBRowIndex = Step + Numbers_I + 1
                    Sheets("DB").Cells(DBRowIndex + 1, 6) = 1
                    Sheets("DB").Cells(DBRowIndex + 1, 7) = Number
                    Sheets("DB").Cells(DBRowIndex + 1, 8) = CardTargetRow
                    Sheets("DB").Cells(DBRowIndex + 1, 9) = CardTargetColumn
                    DBRowIndex = DBRowIndex + 1
                                       
                    Numbers_I = Numbers_I + 1
                End If
            ElseIf Number > 30 And Number <= 45 Then
                If Numbers_N < 5 Then
                    CardTargetRow = CardInitialRow + (Numbers_N * 3)
                    CardTargetColumn = CardInitialColumn + 6
                        
                    If Numbers_N <> 2 Then
                        Cells(CardTargetRow, CardTargetColumn) = Number
                        
                        'Fill DB
                        If Numbers_N <= 2 Then
                            DBRowIndex = Step + Numbers_N + 1
                        Else
                            DBRowIndex = Step + Numbers_N
                        End If
                        
                        Sheets("DB").Cells(DBRowIndex + 1, 6) = 1
                        Sheets("DB").Cells(DBRowIndex + 1, 7) = Number
                        Sheets("DB").Cells(DBRowIndex + 1, 8) = CardTargetRow
                        Sheets("DB").Cells(DBRowIndex + 1, 9) = CardTargetColumn
                        DBRowIndex = DBRowIndex + 1
                    End If
                                       
                    Numbers_N = Numbers_N + 1
                End If
            ElseIf Number > 45 And Number <= 60 Then
                If Numbers_G < 5 Then
                    CardTargetRow = CardInitialRow + (Numbers_G * 3)
                    CardTargetColumn = CardInitialColumn + 9
                        
                    Cells(CardTargetRow, CardTargetColumn) = Number
                    
                    'Fill DB
                    DBRowIndex = Step + Numbers_G
                    Sheets("DB").Cells(DBRowIndex + 1, 6) = 1
                    Sheets("DB").Cells(DBRowIndex + 1, 7) = Number
                    Sheets("DB").Cells(DBRowIndex + 1, 8) = CardTargetRow
                    Sheets("DB").Cells(DBRowIndex + 1, 9) = CardTargetColumn
                    DBRowIndex = DBRowIndex + 1
                                       
                    Numbers_G = Numbers_G + 1
                End If
            ElseIf Number > 60 And Number <= 75 Then
                If Numbers_O < 5 Then
                    CardTargetRow = CardInitialRow + (Numbers_O * 3)
                    CardTargetColumn = CardInitialColumn + 12
                        
                    Cells(CardTargetRow, CardTargetColumn) = Number
                    
                    'Fill DB
                    DBRowIndex = Step + Numbers_O
                    Sheets("DB").Cells(DBRowIndex + 1, 6) = 1
                    Sheets("DB").Cells(DBRowIndex + 1, 7) = Number
                    Sheets("DB").Cells(DBRowIndex + 1, 8) = CardTargetRow
                    Sheets("DB").Cells(DBRowIndex + 1, 9) = CardTargetColumn
                    DBRowIndex = DBRowIndex + 1
                                       
                    Numbers_O = Numbers_O + 1
                End If
            End If
        Else: Exit For
        End If
    Next i
    
End Sub

Private Sub btnScreenshot_Click()
    Range("X1:AP22").Copy
End Sub