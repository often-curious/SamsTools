VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFlappy 
   Caption         =   "Flappy Spaceship"
   ClientHeight    =   5832
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8985.001
   OleObjectBlob   =   "frmFlappy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFlappy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gameRunning As Boolean
Private gameStarted As Boolean
Private gameOver As Boolean
Private isEndingGame As Boolean

Private playerVelocity As Single
Private gravity As Single
Private flapStrength As Single

Private obstacleSpeed As Single
Private gapSize As Single
Private obstacleWidth As Single

Private score As Long
Private currentLevel As Long
Private highScore As Long
Private obstacleCount As Long

Private Const STAR_COUNT As Long = 35

Private Sub UserForm_Initialize()

    Me.width = 460
    Me.Height = 320
    Me.caption = "Flappy Spaceship"
    Me.BackColor = RGB(0, 0, 0)

    gravity = 0.9
    flapStrength = -8.5
    obstacleSpeed = 4
    gapSize = 150
    obstacleWidth = 36

    highScore = 0

    SetupPlayer
    SetupLabels
    CreateStarfield
    ShowStartScreen

End Sub

Private Sub SetupPlayer()

    With lblPlayer
        .caption = ChrW(&H27A4) ' ?
        .Font.Size = 18
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .width = 28
        .Height = 24
        .Left = 70
        .Top = 135
        .ForeColor = RGB(0, 255, 120)
        .BackStyle = fmBackStyleTransparent
    End With

End Sub

Private Sub SetupLabels()

    With lblScore
        .caption = "Press SPACE to Start"
        .Left = 10
        .Top = 8
        .width = 320
        .Height = 20
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackStyle = fmBackStyleTransparent
    End With

    With lblMessage
        .caption = ""
        .Left = 55
        .Top = 95
        .width = 340
        .Height = 110
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .ForeColor = RGB(0, 255, 120)
        .BackStyle = fmBackStyleTransparent
    End With

End Sub

Private Sub ShowStartScreen()

    gameRunning = False
    gameStarted = False
    gameOver = False
    isEndingGame = False

    StopFlappyTimer

    lblPlayer.Left = 70
    lblPlayer.Top = 135
    playerVelocity = 0

    lblScore.caption = "High Score: " & highScore

    lblMessage.caption = _
        "FLAPPY SPACESHIP" & vbCrLf & vbCrLf & _
        "Press SPACE to Start" & vbCrLf & _
        "Press ESC to Exit"

    RemoveObstacles

End Sub

Private Sub StartGame()

    On Error GoTo ErrHandler

    gameRunning = True
    gameStarted = True
    gameOver = False
    isEndingGame = False

    score = 0
    obstacleCount = 0
    currentLevel = 1

    playerVelocity = 0

    obstacleSpeed = 4
    gapSize = 150

    lblPlayer.Left = 70
    lblPlayer.Top = 135

    lblScore.caption = "Score: 0   Level: 1   High Score: " & highScore

    lblMessage.caption = ""

    RemoveObstacles

    CreateObstacle Me.InsideWidth + 30
    CreateObstacle Me.InsideWidth + 230

    DoEvents

    StartFlappyTimer

    Exit Sub

ErrHandler:

    MsgBox "StartGame Error: " & Err.Number & _
           vbCrLf & Err.Description

End Sub

Public Sub Tick()

    If Not gameRunning Then Exit Sub
    If gameOver Then Exit Sub
    If isEndingGame Then Exit Sub

    MoveStarfield
    MovePlayer
    MoveObstacles

    If Not gameRunning Then Exit Sub
    If gameOver Then Exit Sub
    If isEndingGame Then Exit Sub

    CheckCollisions

End Sub

Public Sub HandleSpace()

    Static LastPress As Double

    If Timer - LastPress < 0.2 Then Exit Sub

    LastPress = Timer

    If Not gameStarted Or gameOver Then

        isEndingGame = False
        StartGame

    ElseIf gameRunning Then

        playerVelocity = flapStrength

    End If

End Sub

Public Sub HandleEscape()

    gameRunning = False
    StopFlappyGame

End Sub

Private Sub MovePlayer()

    playerVelocity = playerVelocity + gravity
    lblPlayer.Top = lblPlayer.Top + playerVelocity

End Sub

Private Sub CreateStarfield()

    Dim i As Long
    Dim star As MSForms.Label
    Dim starSymbol As String

    Randomize

    For i = 1 To STAR_COUNT

        Set star = Me.Controls.Add("Forms.Label.1", "star_" & i)

        Select Case Int(Rnd() * 4)
            Case 0
                starSymbol = "."
            Case 1
                starSymbol = "*"
            Case 2
                starSymbol = ChrW(&HB7) ' ·
            Case Else
                starSymbol = ChrW(&H2726) ' ?
        End Select

        With star
            .caption = starSymbol
            .Left = Rnd() * Me.InsideWidth
            .Top = 25 + Rnd() * (Me.InsideHeight - 35)
            .width = 14
            .Height = 14
            .Font.Size = 7 + Int(Rnd() * 5)
            .ForeColor = RGB(120 + Int(Rnd() * 135), 120 + Int(Rnd() * 135), 120 + Int(Rnd() * 135))
            .BackStyle = fmBackStyleTransparent
            .Tag = 1 + Int(Rnd() * 3)
        End With

        On Error Resume Next
        star.ZOrder 1
        On Error GoTo 0

    Next i

    BringMainLabelsToFront

End Sub

Private Sub MoveStarfield()

    Dim ctrl As control
    Dim newTop As Single

    For Each ctrl In Me.Controls

        If Left(ctrl.Name, 5) = "star_" Then

            ctrl.Left = ctrl.Left - CSng(ctrl.Tag)

            If ctrl.Left + ctrl.width < 0 Then
                ctrl.Left = Me.InsideWidth + Rnd() * 40
                newTop = 25 + Rnd() * (Me.InsideHeight - 35)
                ctrl.Top = newTop
            End If

        End If

    Next ctrl

End Sub

Private Sub MoveObstacles()

    Dim ctrl As control
    Dim rightMost As Single
    Dim level As Long

    rightMost = 0

    For Each ctrl In Me.Controls

        If IsObstacle(ctrl.Name) Then

            ctrl.Left = ctrl.Left - obstacleSpeed

            If ctrl.Left > rightMost Then
                rightMost = ctrl.Left
            End If

            If Left(ctrl.Name, 7) = "obsBot_" Then

                If ctrl.Tag <> "Passed" Then

                    If ctrl.Left + ctrl.width < lblPlayer.Left Then
                
                        ctrl.Tag = "Passed"
                        score = score + 1
                
                        ' Adjust difficulty based on score
                        Select Case score
                
                            Case 0 To 5
                                gapSize = 150
                                currentLevel = 1
                
                            Case 6 To 15
                                gapSize = 130
                                currentLevel = 2
                
                            Case 16 To 25
                                gapSize = 110
                                currentLevel = 3
                
                            Case Else
                                gapSize = 90
                                currentLevel = 4
                
                        End Select
                
                        If score > highScore Then
                            highScore = score
                        End If
                
                        lblScore.caption = "Score: " & score & _
                                           "   Level: " & currentLevel & _
                                           "   High Score: " & highScore
                
                        If score Mod 5 = 0 Then
                            obstacleSpeed = obstacleSpeed + 0.4
                        End If
                
                    End If
                
                End If

            End If

        End If

    Next ctrl

    RemoveOldObstacles

    If rightMost < Me.InsideWidth - 170 Then
        CreateObstacle Me.InsideWidth + 30
    End If

End Sub

Private Sub CreateObstacle(ByVal startLeft As Single)

    Dim topObstacle As MSForms.Label
    Dim bottomObstacle As MSForms.Label

    Dim gapTop As Single
    Dim topHeight As Single
    Dim bottomTop As Single
    Dim bottomHeight As Single

    Dim minObstacleHeight As Single
    Dim minGapTop As Single
    Dim maxGapTop As Single

    obstacleCount = obstacleCount + 1

    Randomize

    minObstacleHeight = 25
    minGapTop = 35

    ' Make sure the gap can fit while leaving room for top and bottom obstacles
    maxGapTop = Me.InsideHeight - gapSize - minObstacleHeight - 8

    ' Safety fallback if the gap is too large for the form
    If maxGapTop < minGapTop Then
        maxGapTop = minGapTop
        gapSize = Me.InsideHeight - minGapTop - minObstacleHeight - 8
    End If

    ' Pick a random gap position within safe limits
    gapTop = minGapTop + Rnd() * (maxGapTop - minGapTop)

    topHeight = gapTop
    bottomTop = gapTop + gapSize
    bottomHeight = Me.InsideHeight - bottomTop - 8

    ' Final safety checks
    If topHeight < minObstacleHeight Then topHeight = minObstacleHeight
    If bottomHeight < minObstacleHeight Then bottomHeight = minObstacleHeight

    Set topObstacle = Me.Controls.Add("Forms.Label.1", "obsTop_" & obstacleCount)

    With topObstacle
        .caption = "¦¦¦" & vbCrLf & "¦¦¦" & vbCrLf & "¦¦¦" & vbCrLf & "¦¦¦"
        .Left = startLeft
        .Top = 0
        .width = obstacleWidth
        .Height = topHeight
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = RGB(0, 255, 90)
        .BackColor = RGB(0, 80, 25)
        .TextAlign = fmTextAlignCenter
        .Tag = ""
    End With

    Set bottomObstacle = Me.Controls.Add("Forms.Label.1", "obsBot_" & obstacleCount)

    With bottomObstacle
        .caption = "¦¦¦" & vbCrLf & "¦¦¦" & vbCrLf & "¦¦¦" & vbCrLf & "¦¦¦"
        .Left = startLeft
        .Top = bottomTop
        .width = obstacleWidth
        .Height = bottomHeight
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = RGB(0, 255, 90)
        .BackColor = RGB(0, 80, 25)
        .TextAlign = fmTextAlignCenter
        .Tag = ""
    End With

    BringMainLabelsToFront

End Sub

Private Sub RemoveOldObstacles()

    Dim i As Long
    Dim ctrlName As String

    For i = Me.Controls.count - 1 To 0 Step -1

        ctrlName = Me.Controls(i).Name

        If IsObstacle(ctrlName) Then
            If Me.Controls(i).Left + Me.Controls(i).width < 0 Then
                Me.Controls.Remove ctrlName
            End If
        End If

    Next i

End Sub

Private Sub RemoveObstacles()

    Dim i As Long

    For i = Me.Controls.count - 1 To 0 Step -1

        If Left(Me.Controls(i).Name, 7) = "obsTop_" _
        Or Left(Me.Controls(i).Name, 7) = "obsBot_" Then

            Me.Controls.Remove Me.Controls(i).Name

        End If

    Next i

End Sub

Private Function IsObstacle(ByVal ctrlName As String) As Boolean

    IsObstacle = _
        Left(ctrlName, 7) = "obsTop_" Or _
        Left(ctrlName, 7) = "obsBot_"

End Function

Private Sub CheckCollisions()

    Dim ctrl As control

    If Not gameRunning Then Exit Sub
    If gameOver Then Exit Sub
    If isEndingGame Then Exit Sub

    If lblPlayer.Top < 0 Then
        EndGame
        Exit Sub
    End If

    If lblPlayer.Top + lblPlayer.Height > Me.InsideHeight Then
        EndGame
        Exit Sub
    End If

    For Each ctrl In Me.Controls

        If Not gameRunning Then Exit Sub
        If gameOver Then Exit Sub
        If isEndingGame Then Exit Sub

        If IsObstacle(ctrl.Name) Then
            If IsCollision(lblPlayer, ctrl) Then
                EndGame
                Exit Sub
            End If
        End If

    Next ctrl

End Sub

Private Function IsCollision(ByVal a As MSForms.Label, ByVal b As MSForms.Label) As Boolean

    IsCollision = _
        a.Left < b.Left + b.width And _
        a.Left + a.width > b.Left And _
        a.Top < b.Top + b.Height And _
        a.Top + a.Height > b.Top

End Function

Private Sub EndGame()

    If isEndingGame Then Exit Sub

    isEndingGame = True
    gameRunning = False
    gameOver = True

    StopFlappyTimer

    If score > highScore Then
        highScore = score
    End If

    lblScore.caption = "Score: " & score & _
                       "   Level: " & currentLevel & _
                       "   High Score: " & highScore

    lblMessage.caption = _
        "GAME OVER" & vbCrLf & vbCrLf & _
        "Final Score: " & score & vbCrLf & _
        "High Score: " & highScore & vbCrLf & vbCrLf & _
        "Press SPACE to Play Again"

    BringMainLabelsToFront

End Sub

Private Sub BringMainLabelsToFront()

    On Error Resume Next

    lblPlayer.ZOrder 0
    lblScore.ZOrder 0
    lblMessage.ZOrder 0

    On Error GoTo 0

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    gameRunning = False

    StopFlappyTimer
    UnregisterFlappyKeys

    Set gFlappyGame = Nothing

End Sub

Private Sub UserForm_Activate()
    ' Intentionally empty: leave focus on the form so UserForm_KeyDown gets the keys
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Select Case KeyCode
        Case vbKeySpace
            HandleSpace
            KeyCode = 0
        Case vbKeyEscape
            HandleEscape
            KeyCode = 0
    End Select

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    Me.SetFocus
End Sub
