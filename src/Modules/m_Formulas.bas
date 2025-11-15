Attribute VB_Name = "m_Formulas"
Function BeautifyString(InputString As String) As String
    Dim i As Long
    Dim StringLength As Long
    Dim InQuotes As Boolean
    Dim IndentLevel As Long
    Dim MaxIndent As Long: MaxIndent = 10
    Dim NewLineIndented(0 To 20) As String
    Dim OutputParts() As String
    Dim Pos As Long
    Dim InputPart As String
    Dim InlineBuffer As String
    Dim j As Long, Depth As Long
    Dim commaCount As Long, nestedParens As Long

    StringLength = Len(InputString)
    ReDim OutputParts(1 To StringLength * 2)

    ' Precompute indent strings
    For i = 0 To MaxIndent
        NewLineIndented(i) = vbLf & Space$(i * 4)
    Next i

    InQuotes = False
    IndentLevel = 0
    Pos = 1
    i = 1

    Do While i <= StringLength
        InputPart = Mid(InputString, i, 1)

        If InputPart = """" Then
            InQuotes = Not InQuotes
            OutputParts(Pos) = InputPart
            Pos = Pos + 1
            i = i + 1
            GoTo NextChar
        End If

        If InQuotes Then
            OutputParts(Pos) = InputPart
            Pos = Pos + 1
            i = i + 1
            GoTo NextChar
        End If

        If InputPart = "(" Then
            InlineBuffer = ""
            Depth = 1
            commaCount = 0
            nestedParens = 0

            For j = i + 1 To StringLength
                Dim ch As String: ch = Mid(InputString, j, 1)
                If ch = """" Then Exit For
                If ch = "(" Then
                    Depth = Depth + 1
                    nestedParens = nestedParens + 1
                End If
                If ch = ")" Then Depth = Depth - 1
                If ch = "," And Depth = 1 Then commaCount = commaCount + 1
                If Depth = 0 Then Exit For
                InlineBuffer = InlineBuffer & ch
            Next j

            ' Allow inline if it's small, no commas, and shallow
            If Depth = 0 And Len(InlineBuffer) <= 30 And commaCount = 0 And nestedParens = 0 Then
                OutputParts(Pos) = "(" & InlineBuffer & ")"
                Pos = Pos + 1
                i = j + 1
                GoTo NextChar
            Else
                IndentLevel = Application.WorksheetFunction.Min(IndentLevel + 1, MaxIndent)
                OutputParts(Pos) = "(" & NewLineIndented(IndentLevel)
                Pos = Pos + 1
                i = i + 1
                GoTo NextChar
            End If

        ElseIf InputPart = ")" Then
            IndentLevel = Application.WorksheetFunction.Max(IndentLevel - 1, 0)
            OutputParts(Pos) = NewLineIndented(IndentLevel) & ")"
            Pos = Pos + 1

        ElseIf InputPart = "," Then
            OutputParts(Pos) = "," & NewLineIndented(IndentLevel)
            Pos = Pos + 1

        ElseIf InputPart = "+" Or InputPart = "-" Then
            OutputParts(Pos) = NewLineIndented(IndentLevel) & InputPart & " "
            Pos = Pos + 1

        ElseIf InputPart = "*" Or InputPart = "/" Then
            OutputParts(Pos) = " " & InputPart & " "
            Pos = Pos + 1

        ElseIf InputPart = "^" Then
            OutputParts(Pos) = " ^ "
            Pos = Pos + 1

        Else
            OutputParts(Pos) = InputPart
            Pos = Pos + 1
        End If

        i = i + 1
NextChar:
    Loop

    BeautifyString = Join(OutputParts, "")
End Function

Sub BeautifyFormula(control As IRibbonControl)
    Dim cell As Range
    Dim originalFormula As String
    Dim beautified As String
    Dim parenCount As Long
    Dim minParenToBeautify As Long: minParenToBeautify = 2
    Dim minLengthToBeautify As Long: minLengthToBeautify = 20

    For Each cell In Selection
        If cell.HasFormula Then
            originalFormula = cell.formula

            ' Count number of parentheses
            parenCount = Len(originalFormula) - Len(Replace(originalFormula, "(", ""))

            ' Skip short/simple formulas
            If Len(originalFormula) < minLengthToBeautify And parenCount < minParenToBeautify Then
                GoTo SkipCell
            End If

            beautified = BeautifyString(originalFormula)
            cell.formula = beautified
        End If
SkipCell:
    Next cell
End Sub


Function MinifyString(InputString As String) As String
    Dim i As Long
    Dim Char As String
    Dim InQuotes As Boolean
    Dim CleanParts() As String
    Dim Pos As Long
    ReDim CleanParts(1 To Len(InputString) * 2)

    InQuotes = False
    Pos = 1

    For i = 1 To Len(InputString)
        Char = Mid(InputString, i, 1)

        If Char = """" Then
            InQuotes = Not InQuotes
            CleanParts(Pos) = Char
            Pos = Pos + 1
        ElseIf InQuotes Then
            CleanParts(Pos) = Char
            Pos = Pos + 1
        Else
            Select Case Char
                Case vbCr, vbLf, vbTab
                    ' Skip line breaks and tabs
                Case " "
                    ' Collapse multiple spaces outside strings
                    If Pos > 1 And CleanParts(Pos - 1) <> " " Then
                        CleanParts(Pos) = " "
                        Pos = Pos + 1
                    End If
                Case Else
                    CleanParts(Pos) = Char
                    Pos = Pos + 1
            End Select
        End If
    Next i

    MinifyString = Trim(Join(CleanParts, ""))
End Function

Sub MinifyFormula(control As IRibbonControl)
    Dim cell As Range
    Dim originalFormula As String
    Dim minified As String

    For Each cell In Selection
        If cell.HasFormula Then
            originalFormula = cell.formula
            minified = MinifyString(originalFormula)
            cell.formula = minified
        End If
    Next cell
End Sub

Sub EncapsulateIFERROR_ZERO(control As IRibbonControl)

    Dim c As Range

    For Each c In Selection.Cells
        Select Case Left(c.formula, 1)
            Case "="
                c.formula = "=IFERROR(" & Right(c.formula, Len(c.formula) - 1) & ",0)"
            Case "+"
                c.formula = "=IFERROR(" & Right(c.formula, Len(c.formula) - 1) & ",0)"
            Case Else
                c.formula = "=IFERROR(" & c.formula & ",0)"
        End Select
    Next c
End Sub

Sub EncapsulateIFERROR_BLANK(control As IRibbonControl)

    Dim c As Range

    For Each c In Selection.Cells
        Select Case Left(c.formula, 1)
            Case "="
                c.formula = "=IFERROR(" & Right(c.formula, Len(c.formula) - 1) & ","""")"
            Case "+"
                c.formula = "=IFERROR(" & Right(c.formula, Len(c.formula) - 1) & ","""")"
            Case Else
                c.formula = "=IFERROR(" & c.formula & ","""")"
        End Select
    Next c
End Sub

Sub RemoveSheetNameReferences(control As IRibbonControl)
    Dim ws As Worksheet
    Dim cell As Range
    Dim f As String
    Dim re As Object
    Dim sheetName As String, sheetQuoted As String, bookName As String, pattern As String

    ' Identify active sheet and workbook
    sheetName = ActiveSheet.Name
    bookName = ThisWorkbook.Name

    ' Escape regex special characters in sheet name
    sheetQuoted = Replace(sheetName, "'", "''") ' doubled single-quotes for escaped sheet names

    ' Regex: matches ONLY current sheet (with or without quotes or workbook prefix)
    ' e.g. Sheet1!A1 or 'Sheet 1'!B2 or [Book1.xlsx]Sheet1!C3
    pattern = "(\[(" & Replace(bookName, ".", "\.") & ")\])?('?(" & sheetQuoted & ")'?)!"

    Set re = CreateObject("VBScript.RegExp")
    With re
        .Global = True
        .IgnoreCase = False
        .pattern = pattern
    End With

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error Resume Next

    ' Work on selection or used range if none selected
    If TypeName(Selection) = "Range" Then
        For Each cell In Selection
            If cell.HasFormula Then
                f = cell.formula
                If re.Test(f) Then cell.formula = re.Replace(f, "")
            End If
        Next cell
    Else
        Set ws = ActiveSheet
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                f = cell.formula
                If re.Test(f) Then cell.formula = re.Replace(f, "")
            End If
        Next cell
    End If

    On Error GoTo 0

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    'MsgBox "Removed references to the active sheet only (" & sheetName & ").", vbInformation
End Sub

