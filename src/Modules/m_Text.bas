Attribute VB_Name = "m_Text"
Sub Off()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Debug.Print Now
End Sub
Sub Onn()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Debug.Print Now
End Sub

Sub ShowTextToolsForm(control As IRibbonControl)
    With frmTextTools
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub

Sub ConvertTextToUpper()
    Dim c As Range
    Dim shp As Shape
    Dim i As Long
    Dim sel As Object
    
    Call Off

    ' Check if selection is cells
    If TypeName(Selection) = "Range" Then
        For Each c In Selection.Cells
            If Not c.HasFormula Then
                c.value = StrConv(c.value, vbUpperCase)
            End If
        Next c
        Exit Sub
    End If

    ' If shapes are selected
    If TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "ShapeRange" Then
        For Each sel In Selection.ShapeRange
            If sel.Type = msoTextBox Or sel.Type = msoAutoShape Then
                If sel.TextFrame.HasText Then
                    sel.TextFrame.Characters.text = StrConv(sel.TextFrame.Characters.text, vbUpperCase)
                End If
            End If
        Next sel
        Exit Sub
    End If

    Call Onn

    ' Unsupported selection
    MsgBox "The selected object is not supported. Please select cells or text-based shapes.", vbExclamation, "Invalid Selection"
End Sub

Sub ConvertTextToLower()
    Dim c As Range
    Dim shp As Shape
    Dim i As Long
    Dim sel As Object

    Call Off
    
    ' Check if selection is cells
    If TypeName(Selection) = "Range" Then
        For Each c In Selection.Cells
            If Not c.HasFormula Then
                c.value = StrConv(c.value, vbLowerCase)
            End If
        Next c
        Call Onn
        Exit Sub
    End If

    ' If shapes are selected
    If TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "ShapeRange" Then
        For Each sel In Selection.ShapeRange
            If sel.Type = msoTextBox Or sel.Type = msoAutoShape Then
                If sel.TextFrame.HasText Then
                    sel.TextFrame.Characters.text = StrConv(sel.TextFrame.Characters.text, vbLowerCase)
                End If
            End If
        Next sel
        Call Onn
        Exit Sub
    End If

    Call Onn
    
    ' Unsupported selection
    MsgBox "The selected object is not supported. Please select cells or text-based shapes.", vbExclamation, "Invalid Selection"
End Sub

Sub ConvertTextToProper()
    Dim c As Range
    Dim shp As Shape
    Dim sel As Object

    Call Off
    
    ' Check if selection is a cell range
    If TypeName(Selection) = "Range" Then
        For Each c In Selection.Cells
            If Not c.HasFormula Then
                c.value = StrConv(c.value, vbProperCase)
            End If
        Next c
        Call Onn
        Exit Sub
    End If

    ' Check if selection includes shapes
    If TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "ShapeRange" Then
        For Each sel In Selection.ShapeRange
            If sel.Type = msoTextBox Or sel.Type = msoAutoShape Then
                If sel.TextFrame.HasText Then
                    sel.TextFrame.Characters.text = StrConv(sel.TextFrame.Characters.text, vbProperCase)
                End If
            End If
        Next sel
        Call Onn
        Exit Sub
    End If
    
    Call Onn
    
    ' Unsupported selection
    MsgBox "The selected object is not supported. Please select cells or text-based shapes.", vbExclamation, "Invalid Selection"
End Sub

Sub ConvertTextToSentenceCase()
    Dim c As Range
    Dim shp As Shape
    Dim sel As Object
    
    Call Off

    ' If Range selected
    If TypeName(Selection) = "Range" Then
        For Each c In Selection.Cells
            If Not c.HasFormula Then
                c.value = ApplySentenceCase(CStr(c.value))
            End If
        Next c
        Call Onn
        Exit Sub
    End If

    ' If ShapeRange selected
    If TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "ShapeRange" Then
        For Each sel In Selection.ShapeRange
            If sel.Type = msoTextBox Or sel.Type = msoAutoShape Then
                If sel.TextFrame.HasText Then
                    sel.TextFrame.Characters.text = ApplySentenceCase(sel.TextFrame.Characters.text)
                End If
            End If
        Next sel
        Call Onn
        Exit Sub
    End If
    
    Call Onn

    MsgBox "The selected object is not supported. Please select cells or text-based shapes.", vbExclamation, "Invalid Selection"
End Sub

Function ApplySentenceCase(text As String) As String
    Dim result As String
    Dim i As Long, ch As String, nextUpper As Boolean

    Call Off
    
    text = LCase(text)
    result = ""
    nextUpper = True ' Capitalize first letter

    For i = 1 To Len(text)
        ch = Mid(text, i, 1)

        If nextUpper And ch Like "[a-z]" Then
            result = result & UCase(ch)
            nextUpper = False
        Else
            result = result & ch
        End If

        If ch = "." Or ch = "!" Or ch = "?" Then
            nextUpper = True
        ElseIf ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
            ' Don't reset on spaces, only reset if non-whitespace
            nextUpper = nextUpper
        End If
    Next i

    ApplySentenceCase = result
    
    Call Onn
    
End Function

Sub ToggleBullets(control As IRibbonControl)
    Dim c As Range
    Dim lines As Variant
    Dim i As Long
    Dim newText As String
    Dim bulletState As String
    Dim line As String
    Dim firstLine As String

    ' Confirm it's a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cells only.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call Off

    For Each c In Selection.Cells
        If Not c.HasFormula Then
            If Len(c.value) > 0 Then
                lines = Split(c.value, vbLf)
                firstLine = Trim(lines(0))

                ' Detect current bullet state
                If Left(firstLine, 2) = "- " Then
                    bulletState = "dash"
                ElseIf Left(firstLine, 2) = "• " Then
                    bulletState = "dot"
                Else
                    bulletState = "none"
                End If

                newText = ""
                For i = LBound(lines) To UBound(lines)
                    line = Trim(lines(i))

                    Select Case bulletState
                        Case "none"
                            newText = newText & "- " & line
                        Case "dash"
                            If Left(line, 2) = "- " Then line = Mid(line, 3)
                            newText = newText & "• " & line
                        Case "dot"
                            If Left(line, 2) = "• " Then line = Mid(line, 3)
                            newText = newText & line
                    End Select

                    If i < UBound(lines) Then newText = newText & vbLf
                Next i

                c.value = newText
            End If
        End If
    Next c
    
    Call Onn
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub ToggleLetterBullets(control As IRibbonControl)
    Dim c As Range
    Dim allLines As Collection
    Dim lineTexts() As String
    Dim i As Long, idx As Long
    Dim bulletState As String
    Dim line As String
    Dim firstLine As String
    Dim totalLines As Long
    
    ' Validate selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cells only.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    Call Off
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set allLines = New Collection
    
    ' Detect bullet state by inspecting first line of first non-empty cell
    For Each c In Selection.Cells
        If Not c.HasFormula And Len(c.value) > 0 Then
            lineTexts = Split(c.value, vbLf)
            firstLine = Trim(lineTexts(0))
            
            If firstLine Like "#.*" And Mid(firstLine, 2, 1) = "." Then
                bulletState = "number"
            ElseIf firstLine Like "[a-zA-Z].*" And Mid(firstLine, 2, 1) = "." Then
                bulletState = "letter"
            Else
                bulletState = "none"
            End If
            Exit For
        End If
    Next c
    
    ' Gather all lines from all cells in order
    For Each c In Selection.Cells
        If Not c.HasFormula And Len(c.value) > 0 Then
            lineTexts = Split(c.value, vbLf)
            For i = LBound(lineTexts) To UBound(lineTexts)
                allLines.Add Trim(lineTexts(i))
            Next i
        Else
            ' Add blank line if empty or formula (to keep count consistent)
            allLines.Add ""
        End If
    Next c
    
    totalLines = allLines.count
    idx = 1 ' global line counter for bullets
    
    ' Now write back to cells with new bullets based on bulletState
    For Each c In Selection.Cells
        If Not c.HasFormula And Len(c.value) > 0 Then
            lineTexts = Split(c.value, vbLf)
            Dim newCellText As String
            newCellText = ""
            
            For i = LBound(lineTexts) To UBound(lineTexts)
                line = Trim(lineTexts(i))
                
                Select Case bulletState
                    Case "none"
                        newCellText = newCellText & idx & ". " & line
                    Case "number"
                        ' Remove existing number bullet if present
                        If Mid(line, 2, 1) = "." Then
                            line = Trim(Mid(line, InStr(line, ".") + 1))
                        End If
                        newCellText = newCellText & Chr(96 + idx) & ". " & line ' a., b., c.
                    Case "letter"
                        ' Remove existing letter bullet if present
                        If Mid(line, 2, 1) = "." Then
                            line = Trim(Mid(line, InStr(line, ".") + 1))
                        End If
                        newCellText = newCellText & line ' Plain text (remove bullets)
                End Select
                
                If i < UBound(lineTexts) Then newCellText = newCellText & vbLf
                
                idx = idx + 1
            Next i
            
            c.value = newCellText
        ElseIf Not c.HasFormula And Len(c.value) = 0 Then
            ' Clear empty cells so they don't break count
            c.value = ""
        End If
    Next c
    
    Call Onn
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub DeleteFirstChars(n As Integer)
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula And Len(c.value) > 0 Then
            c.value = Mid(c.value, n + 1)
        End If
    Next c
    Call Onn
End Sub

Sub DeleteLastChars(n As Integer)
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula And Len(c.value) > n Then
            c.value = Left(c.value, Len(c.value) - n)
        End If
    Next c
    Call Onn
End Sub

Sub DeleteAtPosition(startPos As Integer, count As Integer)
    Dim c As Range
    Dim txt As String
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            txt = c.value
            If startPos > 0 And startPos <= Len(txt) Then
                c.value = Left(txt, startPos - 1) & Mid(txt, startPos + count)
            End If
        End If
    Next c
    Call Onn
End Sub

Sub DeleteExtraSpaces()
    Dim c As Range
    Dim txt As String
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            txt = Application.WorksheetFunction.Trim(c.value) ' Removes extra spaces
            c.value = txt
        End If
    Next c
    Call Onn
End Sub

Sub DeleteNonPrintable()
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            c.value = Application.WorksheetFunction.Clean(c.value)
        End If
    Next c
    Call Onn
End Sub

Sub DeleteInitialApostrophes()
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            If Left(c.value, 1) = "'" Then
                c.value = Mid(c.value, 2)
            End If
        End If
    Next c
    Call Onn
End Sub


Sub DeleteAllExceptNumbers()
    Dim c As Range, i As Long, ch As String, result As String
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            result = ""
            For i = 1 To Len(c.value)
                ch = Mid(c.value, i, 1)
                If ch Like "#" Then result = result & ch
            Next i
            c.value = result
        End If
    Next c
    Call Onn
End Sub

Sub DeleteAllExceptLettersAndSpaces()
    Dim c As Range, i As Long, ch As String, result As String
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            result = ""
            For i = 1 To Len(c.value)
                ch = Mid(c.value, i, 1)
                If ch Like "[A-Za-z ]" Then result = result & ch
            Next i
            c.value = result
        End If
    Next c
    Call Onn
End Sub

Sub DeleteBeforeText(targetText As String)
    Dim c As Range, Pos As Long
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            Pos = InStr(1, c.value, targetText, vbTextCompare)
            If Pos > 0 Then
                c.value = Mid(c.value, Pos)
            End If
        End If
    Next c
    Call Onn
End Sub

Sub DeleteAfterText(targetText As String)
    Dim c As Range, Pos As Long
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            Pos = InStr(1, c.value, targetText, vbTextCompare)
            If Pos > 0 Then
                c.value = Left(c.value, Pos - 1)
            End If
        End If
    Next c
    Call Onn
End Sub

Sub DeleteLineBreaks()
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            c.value = Replace(c.value, vbLf, " ")
        End If
    Next c
    Call Onn
End Sub


Sub InsertTextBefore(textToInsert As String)
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            c.value = textToInsert & c.value
        End If
    Next c
    Call Onn
End Sub

Sub InsertTextAfter(textToInsert As String)
    Dim c As Range
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            c.value = c.value & textToInsert
        End If
    Next c
    Call Onn
End Sub

Sub InsertTextAtPosition(textToInsert As String, position As Long)
    Dim c As Range, original As String
    Call Off
    For Each c In Selection
        If Not c.HasFormula Then
            original = c.value
            If position <= 0 Then position = 1
            If position > Len(original) Then
                c.value = original & textToInsert
            Else
                c.value = Left(original, position - 1) & textToInsert & Mid(original, position)
            End If
        End If
    Next c
    Call Onn
End Sub

