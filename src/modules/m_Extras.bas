Attribute VB_Name = "m_Extras"
Option Explicit

Sub ShowAbout(control As IRibbonControl)
    With frmAbout
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show vbModeless
    End With
End Sub


Sub ShowMotivationalMessage(control As IRibbonControl)
    Dim messagesPart1 As Variant
    Dim messagesPart2 As Variant
    Dim messagesPart3 As Variant
    Dim messages(1 To 50) As String
    Dim i As Long, idx As Long
    Dim randomIndex As Long

    ' First 20 messages
    messagesPart1 = Array( _
        "You’ve got this! Keep pushing forward.", _
        "Small steps every day lead to big results.", _
        "Believe in yourself and all that you are.", _
        "Stay positive, work hard, make it happen.", _
        "You are capable of amazing things.", _
        "Progress, not perfection.", _
        "Your only limit is your mind.", _
        "Push yourself, because no one else will do it for you.", _
        "Success is no accident. It’s hard work and determination.", _
        "Every day is a new opportunity to grow.", _
        "Dream it. Wish it. Do it.", _
        "Don’t watch the clock; do what it does. Keep going.", _
        "Start where you are. Use what you have. Do what you can.", _
        "Don’t limit your challenges. Challenge your limits.", _
        "Great things never come from comfort zones.", _
        "Doubt kills more dreams than failure ever will.", _
        "Hardships often prepare ordinary people for an extraordinary destiny.", _
        "Your future is created by what you do today, not tomorrow.", _
        "Wake up with determination. Go to bed with satisfaction.", _
        "The harder you work for something, the greater you’ll feel when you achieve it." _
    )

    ' Next 15 messages
    messagesPart2 = Array( _
        "Push harder than yesterday if you want a different tomorrow.", _
        "Don’t stop until you’re proud.", _
        "Difficult roads often lead to beautiful destinations.", _
        "Success doesn’t just find you. You have to go out and get it.", _
        "Believe you can and you're halfway there.", _
        "Work in silence. Let success make the noise.", _
        "Action is the foundational key to all success.", _
        "Opportunities don’t happen. You create them.", _
        "It always seems impossible until it’s done.", _
        "Stay humble. Work hard. Be kind.", _
        "You are stronger than you think.", _
        "The secret of getting ahead is getting started.", _
        "If you’re tired, learn to rest, not quit.", _
        "Make today so awesome yesterday gets jealous.", _
        "You miss 100% of the shots you don’t take." _
    )

    ' Last 15 messages
    messagesPart3 = Array( _
        "One day or day one. You decide.", _
        "Discipline is the bridge between goals and accomplishment.", _
        "Everything you need is already inside you.", _
        "Keep going. Everything you need will come to you.", _
        "Success is the sum of small efforts repeated daily.", _
        "Nothing will work unless you do.", _
        "The way to get started is to quit talking and begin doing.", _
        "Be fearless in the pursuit of what sets your soul on fire.", _
        "Don’t wait for opportunity. Create it.", _
        "Act as if what you do makes a difference. It does.", _
        "Strive for progress, not perfection.", _
        "Keep your face always toward the sunshine—and shadows will fall behind you.", _
        "Sometimes later becomes never. Do it now.", _
        "Don’t be pushed by your problems. Be led by your dreams.", _
        "Little by little, a little becomes a lot." _
    )

    ' Combine arrays into messages()
    idx = 1
    For i = LBound(messagesPart1) To UBound(messagesPart1)
        messages(idx) = messagesPart1(i)
        idx = idx + 1
    Next i
    For i = LBound(messagesPart2) To UBound(messagesPart2)
        messages(idx) = messagesPart2(i)
        idx = idx + 1
    Next i
    For i = LBound(messagesPart3) To UBound(messagesPart3)
        messages(idx) = messagesPart3(i)
        idx = idx + 1
    Next i

    ' Show random message
    Randomize
    randomIndex = Int((50) * Rnd + 1)
    MsgBox messages(randomIndex), , "Motivation Boost!"
End Sub

    
Sub TriggerEasterEgg()
    Dim visibleRange As Range
    Dim flashCells As Collection
    Dim originalColors As Collection
    Dim i As Long, cycle As Long
    Dim r As Range
    Dim colorOptions(1 To 3) As Long

    On Error GoTo Cleanup
    
    frmParty.btnEndParty.Enabled = False

    ' Define 3 alternating colors
    colorOptions(1) = RGB(250, 250, 0) ' Yellow
    colorOptions(2) = RGB(41, 209, 213) ' Teal
    colorOptions(3) = RGB(240, 14, 132) ' Pink

    Set flashCells = New Collection
    Set originalColors = New Collection

    Set visibleRange = ActiveWindow.visibleRange
    Randomize

    frmParty.Show vbModeless
    
    ' Select 30 random cells from visible range
    For i = 1 To 30
        Dim randRow As Long, randCol As Long
        randRow = Int(Rnd() * visibleRange.rows.count) + 1
        randCol = Int(Rnd() * visibleRange.Columns.count) + 1

        Set r = visibleRange.Cells(randRow, randCol)
        If Not r Is Nothing Then
            flashCells.Add r
            originalColors.Add r.Interior.color
        End If
    Next i

    ' Flash for 7 cycles
    For cycle = 1 To 7
        For i = 1 To flashCells.count
            Set r = flashCells(i)
            Dim ColorIndex As Long
            ColorIndex = Int((Rnd() * 3) + 1)
            r.Interior.color = colorOptions(ColorIndex)
        Next i
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next cycle

Cleanup:
    ' Restore original colors
    For i = 1 To flashCells.count
        flashCells(i).Interior.color = originalColors(i)
    Next i
    
    frmParty.btnEndParty.Enabled = True
End Sub

Sub Speak(control As IRibbonControl)
    Dim speechRange As Range
    Set speechRange = Selection
    
    Dim txt As String
    Dim cell As Range
    
    txt = ""
    For Each cell In Selection
        If Not IsEmpty(cell.value) Then
            txt = txt & CStr(cell.value) & " "
        End If
    Next cell
    
    txt = Trim(txt)
    
    If Len(txt) = 0 Then Exit Sub
    
    ' Show the modeless form
    With frmSpeakStatus
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show vbModeless
    End With
    
    ' Start speaking asynchronously
    Application.Speech.Speak txt, SpeakAsync:=True
    
End Sub

Public Sub PlayGuessTheNumber(control As IRibbonControl)
    VBA.UserForms.Add("frmGuessNumber").Show
End Sub
