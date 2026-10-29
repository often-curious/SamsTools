Attribute VB_Name = "m_Flappy"
Option Explicit

Public gFlappyGame As frmFlappy

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr) As Long

    Private gTimerId As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long

    Private gTimerId As Long
#End If

Public Sub ShowFlappyGame(control As IRibbonControl)

    On Error GoTo ErrHandler

    Set gFlappyGame = New frmFlappy

    RegisterFlappyKeys

    gFlappyGame.Show vbModeless
    'ThisWorkbook.Activate
    'AppActivate Application.caption

    Exit Sub

ErrHandler:
    MsgBox "Could not start Flappy Game: " & Err.Description, vbExclamation

End Sub

Public Sub RegisterFlappyKeys()

    ' Spacebar key is a literal space character
    Application.OnKey " ", "FlappySpace"
    Application.OnKey "{ESC}", "FlappyEscape"

End Sub

Public Sub UnregisterFlappyKeys()

    Application.OnKey " "
    Application.OnKey "{ESC}"

End Sub

Public Sub FlappySpace()

    On Error Resume Next

    If Not gFlappyGame Is Nothing Then
        gFlappyGame.HandleSpace
    End If

    On Error GoTo 0

End Sub

Public Sub FlappyEscape()

    On Error Resume Next

    If Not gFlappyGame Is Nothing Then
        gFlappyGame.HandleEscape
    Else
        StopFlappyTimer
        UnregisterFlappyKeys
    End If

    On Error GoTo 0

End Sub

Public Sub StartFlappyTimer()

    StopFlappyTimer

    #If VBA7 Then
        gTimerId = SetTimer(0, 0, 30, AddressOf FlappyTimerProc)
    #Else
        gTimerId = SetTimer(0, 0, 30, AddressOf FlappyTimerProc)
    #End If

End Sub

Public Sub StopFlappyTimer()

    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
    End If

End Sub

#If VBA7 Then
Public Sub FlappyTimerProc( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long)
#Else
Public Sub FlappyTimerProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal idEvent As Long, _
    ByVal dwTime As Long)
#End If

    On Error Resume Next

    If Not gFlappyGame Is Nothing Then
        gFlappyGame.Tick
    End If

    On Error GoTo 0

End Sub

Public Sub StopFlappyGame()

    On Error Resume Next

    StopFlappyTimer
    UnregisterFlappyKeys

    If Not gFlappyGame Is Nothing Then
        Unload gFlappyGame
        Set gFlappyGame = Nothing
    End If

    On Error GoTo 0

End Sub

