VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMastermind 
   Caption         =   "UserForm1"
   ClientHeight    =   8160
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5925
   OleObjectBlob   =   "frmMastermind.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMastermind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const CODE_LENGTH As Long = 4
Private Const MIN_NUMBER As Long = 0
Private Const MAX_NUMBER As Long = 9
Private Const MAX_TURNS As Long = 10
Private Const ALLOW_DUPLICATES As Boolean = True

Private SecretCode(1 To CODE_LENGTH) As Long
Private CurrentTurn As Long
Private gameOver As Boolean

Private Sub fraBoard_Click()

End Sub

Private Sub UserForm_Initialize()
    SetupForm
    NewGame
End Sub

Private Sub SetupForm()
    Me.caption = "Mastermind"
    Me.width = 280
    Me.Height = 430
    
    txtGuess1.MaxLength = 1
    txtGuess2.MaxLength = 1
    txtGuess3.MaxLength = 1
    txtGuess4.MaxLength = 1
    
    txtGuess1.TextAlign = fmTextAlignCenter
    txtGuess2.TextAlign = fmTextAlignCenter
    txtGuess3.TextAlign = fmTextAlignCenter
    txtGuess4.TextAlign = fmTextAlignCenter
    
    txtGuess1.Font.Size = 12
    txtGuess2.Font.Size = 12
    txtGuess3.Font.Size = 12
    txtGuess4.Font.Size = 12
    
    cmdOK.caption = "OK"
    cmdReset.caption = "Reset"
    cmdHelp.caption = "Help"
    
    fraBoard.caption = "Guesses"
    
    lblStatus.caption = vbNullString
    lblStatus.WordWrap = True
End Sub

Private Sub NewGame()
    Dim i As Long
    
    Randomize
    
    CurrentTurn = 1
    gameOver = False
    
    GenerateSecretCode
    ClearGuessBoxes
    ClearBoard
    
    lblStatus.caption = "Enter four numbers from " & MIN_NUMBER & " to " & MAX_NUMBER & "."
    
    For i = 1 To CODE_LENGTH
        Me.Controls("txtGuess" & i).Enabled = True
    Next i
    
    cmdOK.Enabled = True
    txtGuess1.SetFocus
End Sub

Private Sub GenerateSecretCode()
    Dim i As Long
    Dim n As Long
    Dim usedNumbers As Object
    
    Set usedNumbers = CreateObject("Scripting.Dictionary")
    
    For i = 1 To CODE_LENGTH
        
        Do
            n = WorksheetFunction.RandBetween(MIN_NUMBER, MAX_NUMBER)
        Loop While ALLOW_DUPLICATES = False And usedNumbers.Exists(CStr(n))
        
        SecretCode(i) = n
        usedNumbers(CStr(n)) = True
        
    Next i
End Sub

Private Sub ClearGuessBoxes()
    Dim i As Long
    
    For i = 1 To CODE_LENGTH
        Me.Controls("txtGuess" & i).value = vbNullString
    Next i
End Sub

Private Sub ClearBoard()
    Dim i As Long
    
    On Error Resume Next
    
    For i = fraBoard.Controls.count - 1 To 0 Step -1
        fraBoard.Controls.Remove fraBoard.Controls(i).Name
    Next i
    
    On Error GoTo 0
End Sub

Private Sub cmdOK_Click()
    Dim Guess(1 To CODE_LENGTH) As Long
    Dim PositionResult(1 To CODE_LENGTH) As String
    Dim ResultText As String
    Dim TickCount As Long
    
    If gameOver Then Exit Sub
    
    If Not ReadGuess(Guess) Then Exit Sub
    
    ScoreGuess Guess, PositionResult, TickCount
    ResultText = BuildResultText(PositionResult)
    
    AddTurnToBoard Guess, ResultText
    
    If TickCount = CODE_LENGTH Then
        lblStatus.caption = "Congratulations! You cracked the code in " & CurrentTurn & " turn" & IIf(CurrentTurn = 1, "", "s") & "."
        EndGame
        Exit Sub
    End If
    
    If CurrentTurn >= MAX_TURNS Then
        lblStatus.caption = "Game over. The code was " & SecretCodeText() & "."
        EndGame
        Exit Sub
    End If
    
    CurrentTurn = CurrentTurn + 1
    ClearGuessBoxes
    lblStatus.caption = "Try again. Turn " & CurrentTurn & " of " & MAX_TURNS & "."
    txtGuess1.SetFocus
End Sub

Private Function ReadGuess(ByRef Guess() As Long) As Boolean
    Dim i As Long
    Dim v As String
    
    For i = 1 To CODE_LENGTH
        v = Trim$(Me.Controls("txtGuess" & i).value)
        
        If Len(v) = 0 Then
            MsgBox "Please enter all four numbers.", vbExclamation, "Incomplete Guess"
            Me.Controls("txtGuess" & i).SetFocus
            ReadGuess = False
            Exit Function
        End If
        
        If Not IsNumeric(v) Then
            MsgBox "Please enter numbers only.", vbExclamation, "Invalid Guess"
            Me.Controls("txtGuess" & i).SetFocus
            ReadGuess = False
            Exit Function
        End If
        
        Guess(i) = CLng(v)
        
        If Guess(i) < MIN_NUMBER Or Guess(i) > MAX_NUMBER Then
            MsgBox "Each number must be between " & MIN_NUMBER & " and " & MAX_NUMBER & ".", vbExclamation, "Invalid Number"
            Me.Controls("txtGuess" & i).SetFocus
            ReadGuess = False
            Exit Function
        End If
    Next i
    
    ReadGuess = True
End Function

Private Sub ScoreGuess(ByRef Guess() As Long, ByRef PositionResult() As String, ByRef TickCount As Long)
    Dim i As Long
    Dim n As Long
    Dim SecretUsed(1 To CODE_LENGTH) As Boolean
    Dim GuessUsed(1 To CODE_LENGTH) As Boolean
    
    TickCount = 0
    
    ' Default every position to a cross.
    ' This means "wrong number" unless later changed to tick or dot.
    For i = 1 To CODE_LENGTH
        PositionResult(i) = ChrW(&H2715)   ' ?
    Next i
    
    ' First pass: exact matches.
    ' These get a tick in the same position as the number entered.
    For i = 1 To CODE_LENGTH
        If Guess(i) = SecretCode(i) Then
            PositionResult(i) = ChrW(&H2713)   ' ?
            TickCount = TickCount + 1
            SecretUsed(i) = True
            GuessUsed(i) = True
        End If
    Next i
    
    ' Second pass: correct number, wrong position.
    ' These get a dot in the same position as the number entered.
    ' This also handles duplicate numbers correctly.
    For i = 1 To CODE_LENGTH
        If Not GuessUsed(i) Then
            For n = 1 To CODE_LENGTH
                If Not SecretUsed(n) Then
                    If Guess(i) = SecretCode(n) Then
                        PositionResult(i) = ChrW(&H2022)   ' •
                        SecretUsed(n) = True
                        GuessUsed(i) = True
                        Exit For
                    End If
                End If
            Next n
        End If
    Next i
End Sub

Private Function BuildResultText(ByRef PositionResult() As String) As String
    Dim i As Long
    Dim s As String
    
    For i = 1 To CODE_LENGTH
        s = s & PositionResult(i)
        If i < CODE_LENGTH Then s = s & " "
    Next i
    
    BuildResultText = s
End Function

Private Sub AddTurnToBoard(ByRef Guess() As Long, ByVal ResultText As String)
    Dim topPos As Long
    Dim lblGuess As MSForms.Label
    Dim lblResult As MSForms.Label
    
    topPos = 18 + ((CurrentTurn - 1) * 28)
    
    Set lblGuess = fraBoard.Controls.Add("Forms.Label.1", "lblGuess" & CurrentTurn, True)
    
    With lblGuess
        .Left = 12
        .Top = topPos
        .width = 90
        .Height = 20
        .caption = GuessText(Guess)
        .Font.Size = 11
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .BackColor = RGB(245, 245, 235)
    End With
    
    Set lblResult = fraBoard.Controls.Add("Forms.Label.1", "lblResult" & CurrentTurn, True)
    
    With lblResult
        .Left = 120
        .Top = topPos
        .width = 110
        .Height = 20
        .caption = ResultText
        .Font.Size = 12
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .BackColor = RGB(245, 245, 235)
    End With
End Sub

Private Function GuessText(ByRef Guess() As Long) As String
    Dim i As Long
    Dim s As String
    
    For i = 1 To CODE_LENGTH
        s = s & Guess(i)
        If i < CODE_LENGTH Then s = s & " "
    Next i
    
    GuessText = s
End Function

Private Function SecretCodeText() As String
    Dim i As Long
    Dim s As String
    
    For i = 1 To CODE_LENGTH
        s = s & SecretCode(i)
        If i < CODE_LENGTH Then s = s & " "
    Next i
    
    SecretCodeText = s
End Function

Private Sub EndGame()
    Dim i As Long
    
    gameOver = True
    cmdOK.Enabled = False
    
    For i = 1 To CODE_LENGTH
        Me.Controls("txtGuess" & i).Enabled = False
    Next i
End Sub

Private Sub cmdReset_Click()
    NewGame
End Sub

Private Sub cmdHelp_Click()
    MsgBox _
        "Mastermind rules:" & vbCrLf & vbCrLf & _
        "Guess the hidden four-number code." & vbCrLf & _
        "Each number must be between " & MIN_NUMBER & " and " & MAX_NUMBER & "." & vbCrLf & vbCrLf & _
        "Tick = correct number in the correct position" & vbCrLf & _
        "Dot = correct number in the wrong position" & vbCrLf & _
        "Cross = wrong number" & vbCrLf & vbCrLf & _
        "Feedback is shown in the same order as the numbers you entered." & vbCrLf & vbCrLf & _
        "You have " & MAX_TURNS & " turns.", _
        vbInformation, _
        "Mastermind Help"
End Sub

Private Sub txtGuess1_Change()
    MoveNextIfComplete txtGuess1, txtGuess2
End Sub

Private Sub txtGuess2_Change()
    MoveNextIfComplete txtGuess2, txtGuess3
End Sub

Private Sub txtGuess3_Change()
    MoveNextIfComplete txtGuess3, txtGuess4
End Sub

Private Sub txtGuess4_Change()
    If Len(txtGuess4.value) = 1 Then cmdOK.SetFocus
End Sub

Private Sub MoveNextIfComplete(ByVal CurrentBox As MSForms.TextBox, ByVal NextBox As MSForms.TextBox)
    If Len(CurrentBox.value) = 1 Then
        NextBox.SetFocus
    End If
End Sub

Private Sub txtGuess1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToValidNumbers KeyAscii
End Sub

Private Sub txtGuess2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToValidNumbers KeyAscii
End Sub

Private Sub txtGuess3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToValidNumbers KeyAscii
End Sub

Private Sub txtGuess4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToValidNumbers KeyAscii
End Sub

Private Sub RestrictToValidNumbers(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim ch As String
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    ch = Chr$(KeyAscii)
    
    If Not IsNumeric(ch) Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If CLng(ch) < MIN_NUMBER Or CLng(ch) > MAX_NUMBER Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtGuess4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

