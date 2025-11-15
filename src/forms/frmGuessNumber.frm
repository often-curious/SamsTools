VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGuessNumber 
   Caption         =   "Guess The Number"
   ClientHeight    =   5208
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5430
   OleObjectBlob   =   "frmGuessNumber.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmGuessNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'================= Constants & State =================
Private Const MIN_NUM As Long = 1
Private Const MAX_NUM As Long = 50
Private bits As Long
Private bitIndex As Long
Private sumGuess As Long
Private bitOrder() As Long   ' holds randomized order of bit indices

'================= Lifecycle =================
Private Sub UserForm_Initialize()
    Me.caption = "Guess The Number"
    On Error Resume Next
    Me.txtCard.Font.Name = "Consolas"
    Me.txtCard.Font.SIZE = 12
    On Error GoTo 0
    ResetForm
End Sub

'================= Buttons =================
Private Sub cmdStart_Click()
    bits = CeilLog2(MAX_NUM)
    bitIndex = 0
    sumGuess = 0
    
    '--- Create a randomized order of bit indices ---
    ReDim bitOrder(0 To bits - 1)
    Dim i As Long, j As Long, temp As Long
    For i = 0 To bits - 1
        bitOrder(i) = i
    Next i
    
    Randomize
    For i = bits - 1 To 1 Step -1
        j = Int((i + 1) * Rnd)
        temp = bitOrder(i)
        bitOrder(i) = bitOrder(j)
        bitOrder(j) = temp
    Next i
    
    TogglePlayUI True
    ShowCurrentCard
End Sub

Private Sub cmdYes_Click()
    sumGuess = sumGuess + (2 ^ bitOrder(bitIndex))
    NextCardOrFinish
End Sub

Private Sub cmdNo_Click()
    NextCardOrFinish
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdAgain_Click()
    ResetForm
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' Optional keyboard shortcuts
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyY: If Me.cmdYes.Visible Then cmdYes_Click
        Case vbKeyN: If Me.cmdNo.Visible Then cmdNo_Click
        Case vbKeyEscape: If Me.cmdCancel.Visible Then cmdCancel_Click
    End Select
End Sub

'================= Flow =================
Private Sub ShowCurrentCard()
    If bitIndex >= bits Then
        RevealAnswer
        Exit Sub
    End If
    
    Dim b As Long
    b = bitOrder(bitIndex)
    Dim bitVal As Long: bitVal = (2 ^ b)
    
    Me.lblQNum.caption = "Card " & (bitIndex + 1) & " of " & bits
    Me.lblInstruction.caption = "Is your number shown on this card?"
    
    Me.txtCard.text = BuildCardText(b, MIN_NUM, MAX_NUM)
End Sub

Private Sub NextCardOrFinish()
    bitIndex = bitIndex + 1
    If bitIndex >= bits Then
        RevealAnswer
    Else
        ShowCurrentCard
    End If
End Sub

Private Sub RevealAnswer()
    TogglePlayUI False
    
    Dim msg As String, n As Long
    n = sumGuess
    
    If n < MIN_NUM Or n > MAX_NUM Then
        msg = "Hmm… your answers don’t map to 1–50." & vbCrLf & "Please try again."
    Else
        msg = "Your number is: " & n
    End If
    
    Me.lblResult.caption = msg
    Me.lblResult.Visible = True
    Me.cmdAgain.Visible = True
    Me.cmdClose.Visible = True
    Me.cmdStart.Visible = False
    
End Sub

Private Sub ResetForm()
    Me.lblIntro.caption = "Think of a number between 1 and 50!"
    Me.lblQNum.caption = ""
    Me.txtCard.text = ""
    Me.lblInstruction.caption = "I'm going to try guess your number based on numbers visualised across 6 cards." & vbCrLf & vbCrLf & "Click Start to see the first card."
    
    Me.lblResult.Visible = False
    Me.cmdAgain.Visible = False
    Me.cmdClose.Visible = False
    
    TogglePlayUI False
End Sub

Private Sub TogglePlayUI(ByVal playing As Boolean)
    Me.cmdStart.Visible = Not playing
    Me.cmdYes.Visible = playing
    Me.cmdNo.Visible = playing
    Me.cmdCancel.Visible = playing
    Me.lblQNum.Visible = playing
    Me.txtCard.Visible = playing
    
End Sub

'================= Card Builder =================
Private Function BuildCardText(ByVal bitIdx As Long, ByVal lo As Long, ByVal hi As Long) As String
    Dim arr() As Long, i As Long, count As Long
    ReDim arr(1 To (hi - lo + 1))
    
    For i = lo To hi
        If (i And (2 ^ bitIdx)) <> 0 Then
            count = count + 1
            arr(count) = i
        End If
    Next i
    If count = 0 Then
        BuildCardText = "(No numbers on this card.)"
        Exit Function
    End If
    
    Dim cols As Long, colWidth As Long
    cols = 5
    colWidth = 4
    
    Dim r As Long, c As Long, idx As Long, rows As Long
    rows = (count + cols - 1) \ cols
    
    Dim sb As String
    idx = 1
    For r = 1 To rows
        Dim line As String: line = ""
        For c = 1 To cols
            If idx <= count Then
                line = line & RPad(CStr(arr(idx)), colWidth)
                idx = idx + 1
            End If
        Next c
        sb = sb & line & vbCrLf
    Next r
    
    BuildCardText = Trim$(sb)
End Function

Private Function RPad(ByVal s As String, ByVal width As Long) As String
    If Len(s) >= width Then
        RPad = s & " "
    Else
        RPad = s & Space$(width - Len(s))
    End If
End Function

'================= Helpers =================
Private Function CeilLog2(ByVal n As Long) As Long
    If n <= 1 Then
        CeilLog2 = 0
    Else
        CeilLog2 = Fix(Log(n - 1) / Log(2)) + 1
    End If
End Function

