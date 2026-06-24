Attribute VB_Name = "m_ExplainFormula"
Option Explicit

Public Sub FormulaTranslator()

    Dim rng As Range
    Dim f As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell with a formula.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection.Cells(1, 1)

    If Not rng.HasFormula Then
        MsgBox "The selected cell does not contain a formula.", vbExclamation
        Exit Sub
    End If

    f = rng.formula

    MsgBox TranslateFormula(f), vbInformation, "Formula Explainer"

End Sub

Private Function TranslateFormula(ByVal formulaText As String) As String

    Dim f As String
    f = Trim$(formulaText)

    If Left$(f, 1) = "=" Then f = Mid$(f, 2)

    Select Case True

        ' Basic aggregation
        Case StartsWithFunction(f, "SUM")
            TranslateFormula = "Adds together the values in " & ArgsText(f) & "."

        Case StartsWithFunction(f, "AVERAGE")
            TranslateFormula = "Calculates the average of " & ArgsText(f) & "."

        Case StartsWithFunction(f, "COUNT")
            TranslateFormula = "Counts the numeric values in " & ArgsText(f) & "."

        Case StartsWithFunction(f, "COUNTA")
            TranslateFormula = "Counts the non-blank values in " & ArgsText(f) & "."

        Case StartsWithFunction(f, "MAX")
            TranslateFormula = "Returns the highest value from " & ArgsText(f) & "."

        Case StartsWithFunction(f, "MIN")
            TranslateFormula = "Returns the lowest value from " & ArgsText(f) & "."

        ' Conditional aggregation
        Case StartsWithFunction(f, "SUMIF")
            TranslateFormula = TranslateSUMIF(f)

        Case StartsWithFunction(f, "SUMIFS")
            TranslateFormula = TranslateSUMIFS(f)

        Case StartsWithFunction(f, "COUNTIF")
            TranslateFormula = TranslateCOUNTIF(f)

        Case StartsWithFunction(f, "COUNTIFS")
            TranslateFormula = TranslateCOUNTIFS(f)

        ' Lookup
        Case StartsWithFunction(f, "XLOOKUP")
            TranslateFormula = TranslateXLOOKUP(f)

        Case StartsWithFunction(f, "VLOOKUP")
            TranslateFormula = TranslateVLOOKUP(f)

        Case StartsWithFunction(f, "INDEX")
            TranslateFormula = TranslateINDEX(f)

        Case StartsWithFunction(f, "MATCH")
            TranslateFormula = TranslateMATCH(f)

        Case StartsWithFunction(f, "XMATCH")
            TranslateFormula = TranslateXMATCH(f)

        ' Logical
        Case StartsWithFunction(f, "IFERROR")
            TranslateFormula = TranslateIFERROR(f)

        Case StartsWithFunction(f, "IF")
            TranslateFormula = TranslateIF(f)

        Case StartsWithFunction(f, "IFS")
            TranslateFormula = TranslateIFS(f)

        Case StartsWithFunction(f, "AND")
            TranslateFormula = "Returns TRUE only if all of these conditions are met: " & ArgsText(f) & "."

        Case StartsWithFunction(f, "OR")
            TranslateFormula = "Returns TRUE if any of these conditions are met: " & ArgsText(f) & "."

        ' Dates
        Case StartsWithFunction(f, "TODAY")
            TranslateFormula = "Returns today's date."

        Case StartsWithFunction(f, "YEAR")
            TranslateFormula = "Extracts the year from " & ArgsText(f) & "."

        Case StartsWithFunction(f, "MONTH")
            TranslateFormula = "Extracts the month number from " & ArgsText(f) & "."

        Case StartsWithFunction(f, "DAY")
            TranslateFormula = "Extracts the day number from " & ArgsText(f) & "."

        Case StartsWithFunction(f, "EOMONTH")
            TranslateFormula = TranslateEOMONTH(f)

        Case StartsWithFunction(f, "DATEDIF")
            TranslateFormula = TranslateDATEDIF(f)

        ' Text
        Case StartsWithFunction(f, "LEFT")
            TranslateFormula = TranslateLEFT(f)

        Case StartsWithFunction(f, "RIGHT")
            TranslateFormula = TranslateRIGHT(f)

        Case StartsWithFunction(f, "MID")
            TranslateFormula = TranslateMID(f)

        Case StartsWithFunction(f, "LEN")
            TranslateFormula = "Returns the number of characters in " & ArgsText(f) & "."

        Case StartsWithFunction(f, "SUBSTITUTE")
            TranslateFormula = TranslateSUBSTITUTE(f)

        Case StartsWithFunction(f, "TEXTJOIN")
            TranslateFormula = TranslateTEXTJOIN(f)

        ' Dynamic arrays
        Case StartsWithFunction(f, "FILTER")
            TranslateFormula = TranslateFILTER(f)

        Case StartsWithFunction(f, "UNIQUE")
            TranslateFormula = "Returns the unique values from " & ArgsText(f) & "."

        Case StartsWithFunction(f, "SORTBY")
            TranslateFormula = TranslateSORTBY(f)

        Case StartsWithFunction(f, "SORT")
            TranslateFormula = "Returns the values from " & ArgsText(f) & " sorted in order."

        ' Financial
        Case StartsWithFunction(f, "NPV")
            TranslateFormula = TranslateNPV(f)

        Case StartsWithFunction(f, "IRR")
            TranslateFormula = "Calculates the internal rate of return for the cash flows in " & ArgsText(f) & "."

        ' Maths
        Case StartsWithFunction(f, "ROUND")
            TranslateFormula = TranslateROUND(f, "rounds")

        Case StartsWithFunction(f, "ROUNDUP")
            TranslateFormula = TranslateROUND(f, "rounds up")

        Case StartsWithFunction(f, "ROUNDDOWN")
            TranslateFormula = TranslateROUND(f, "rounds down")

        Case StartsWithFunction(f, "SUMPRODUCT")
            TranslateFormula = "Multiplies corresponding values in the selected ranges, then adds the results together."

        ' References
        Case StartsWithFunction(f, "OFFSET")
            TranslateFormula = TranslateOFFSET(f)

        Case StartsWithFunction(f, "INDIRECT")
            TranslateFormula = "Converts the text in " & ArgsText(f) & " into a worksheet reference."

        Case StartsWithFunction(f, "CELL")
            TranslateFormula = "Returns information about a cell, workbook, or worksheet."

        Case Else
            TranslateFormula = "This formula calculates:" & vbCrLf & vbCrLf & formulaText & vbCrLf & vbCrLf & _
                               "This formula type is not yet included in the translator."

    End Select

End Function

Private Function TranslateIF(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count < 3 Then
        TranslateIF = "Checks a condition and returns one result if TRUE and another if FALSE."
    Else
        TranslateIF = "If " & CleanExpression(a(1)) & ", then return " & CleanExpression(a(2)) & _
                      ". Otherwise, return " & CleanExpression(a(3)) & "."
    End If
End Function

Private Function TranslateIFERROR(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count < 2 Then
        TranslateIFERROR = "Returns a fallback value if the formula results in an error."
    Else
        TranslateIFERROR = "Calculates " & CleanExpression(a(1)) & _
                           ". If that results in an error, return " & CleanExpression(a(2)) & "."
    End If
End Function

Private Function TranslateIFS(ByVal f As String) As String
    Dim a As Collection, i As Long, txt As String
    Set a = SplitFunctionArguments(ArgsText(f))

    txt = "Checks each condition in order and returns the first matching result:" & vbCrLf

    For i = 1 To a.count Step 2
        If i + 1 <= a.count Then
            txt = txt & vbCrLf & "• If " & CleanExpression(a(i)) & ", return " & CleanExpression(a(i + 1))
        End If
    Next i

    TranslateIFS = txt
End Function

Private Function TranslateSUMIF(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count = 2 Then
        TranslateSUMIF = "Adds values in " & a(1) & " where the value matches " & CleanExpression(a(2)) & "."
    ElseIf a.count >= 3 Then
        TranslateSUMIF = "Adds values in " & a(3) & " where " & a(1) & " matches " & CleanExpression(a(2)) & "."
    Else
        TranslateSUMIF = "Adds values that meet a single condition."
    End If
End Function

Private Function TranslateSUMIFS(ByVal f As String) As String
    Dim a As Collection, i As Long, txt As String
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count < 3 Then
        TranslateSUMIFS = "Adds values that meet multiple conditions."
        Exit Function
    End If

    txt = "Adds values in " & a(1) & " where "

    For i = 2 To a.count Step 2
        If i + 1 <= a.count Then
            txt = txt & a(i) & " matches " & CleanExpression(a(i + 1))
            If i + 2 <= a.count Then txt = txt & ", and "
        End If
    Next i

    TranslateSUMIFS = txt & "."
End Function

Private Function TranslateCOUNTIF(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateCOUNTIF = "Counts cells in " & a(1) & " where the value matches " & CleanExpression(a(2)) & "."
    Else
        TranslateCOUNTIF = "Counts values that meet a single condition."
    End If
End Function

Private Function TranslateCOUNTIFS(ByVal f As String) As String
    Dim a As Collection, i As Long, txt As String
    Set a = SplitFunctionArguments(ArgsText(f))

    txt = "Counts rows where "

    For i = 1 To a.count Step 2
        If i + 1 <= a.count Then
            txt = txt & a(i) & " matches " & CleanExpression(a(i + 1))
            If i + 2 <= a.count Then txt = txt & ", and "
        End If
    Next i

    TranslateCOUNTIFS = txt & "."
End Function

Private Function TranslateXLOOKUP(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateXLOOKUP = "Looks for " & CleanExpression(a(1)) & " in " & a(2) & _
                           ", then returns the matching value from " & a(3) & "."
    Else
        TranslateXLOOKUP = "Looks up a value and returns a matching result."
    End If
End Function

Private Function TranslateVLOOKUP(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateVLOOKUP = "Looks for " & CleanExpression(a(1)) & " in the first column of " & a(2) & _
                           ", then returns the matching value from column " & a(3) & "."
    Else
        TranslateVLOOKUP = "Looks up a value in the first column of a table."
    End If
End Function

Private Function TranslateINDEX(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateINDEX = "Returns the value from " & a(1) & " at row position " & CleanExpression(a(2)) & "."
    Else
        TranslateINDEX = "Returns a value from a specific position in a range."
    End If
End Function

Private Function TranslateMATCH(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateMATCH = "Finds the position of " & CleanExpression(a(1)) & " within " & a(2) & "."
    Else
        TranslateMATCH = "Finds the position of a value within a range."
    End If
End Function

Private Function TranslateXMATCH(ByVal f As String) As String
    TranslateXMATCH = Replace(TranslateMATCH(f), "Finds", "Returns")
End Function

Private Function TranslateEOMONTH(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateEOMONTH = "Returns the last day of the month that is " & CleanExpression(a(2)) & _
                           " month(s) from " & CleanExpression(a(1)) & "."
    Else
        TranslateEOMONTH = "Returns the last day of a month."
    End If
End Function

Private Function TranslateDATEDIF(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateDATEDIF = "Calculates the difference between " & a(1) & " and " & a(2) & _
                           " using the unit " & CleanExpression(a(3)) & "."
    Else
        TranslateDATEDIF = "Calculates the difference between two dates."
    End If
End Function

Private Function TranslateLEFT(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateLEFT = "Returns the first " & a(2) & " character(s) from " & a(1) & "."
    Else
        TranslateLEFT = "Returns characters from the left side of text."
    End If
End Function

Private Function TranslateRIGHT(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateRIGHT = "Returns the last " & a(2) & " character(s) from " & a(1) & "."
    Else
        TranslateRIGHT = "Returns characters from the right side of text."
    End If
End Function

Private Function TranslateMID(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateMID = "Returns " & a(3) & " character(s) from " & a(1) & _
                       ", starting at position " & a(2) & "."
    Else
        TranslateMID = "Returns characters from the middle of text."
    End If
End Function

Private Function TranslateSUBSTITUTE(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateSUBSTITUTE = "Replaces " & CleanExpression(a(2)) & " with " & _
                              CleanExpression(a(3)) & " in " & a(1) & "."
    Else
        TranslateSUBSTITUTE = "Replaces existing text with new text."
    End If
End Function

Private Function TranslateTEXTJOIN(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateTEXTJOIN = "Combines values from " & a(3) & " onward, separated by " & _
                            CleanExpression(a(1)) & ". Blank cells are " & _
                            IIf(UCase$(a(2)) = "TRUE", "ignored", "included") & "."
    Else
        TranslateTEXTJOIN = "Combines multiple values into one text string."
    End If
End Function

Private Function TranslateFILTER(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateFILTER = "Returns values from " & a(1) & " where " & CleanExpression(a(2)) & "."
    Else
        TranslateFILTER = "Filters a range based on a condition."
    End If
End Function

Private Function TranslateSORTBY(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateSORTBY = "Returns " & a(1) & " sorted by " & a(2) & "."
    Else
        TranslateSORTBY = "Sorts one range by another range."
    End If
End Function

Private Function TranslateNPV(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateNPV = "Calculates the present value of future cash flows in " & a(2) & _
                       " using a discount rate of " & a(1) & "."
    Else
        TranslateNPV = "Calculates the net present value of future cash flows."
    End If
End Function

Private Function TranslateROUND(ByVal f As String, ByVal actionText As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 2 Then
        TranslateROUND = actionText & " " & CleanExpression(a(1)) & " to " & a(2) & " decimal place(s)."
    Else
        TranslateROUND = actionText & " a number."
    End If
End Function

Private Function TranslateOFFSET(ByVal f As String) As String
    Dim a As Collection
    Set a = SplitFunctionArguments(ArgsText(f))

    If a.count >= 3 Then
        TranslateOFFSET = "Returns a reference starting from " & a(1) & ", moved " & _
                          a(2) & " row(s) and " & a(3) & " column(s)."
    Else
        TranslateOFFSET = "Returns a reference offset from a starting cell."
    End If
End Function

Private Function StartsWithFunction(ByVal f As String, ByVal functionName As String) As Boolean
    f = Trim$(f)
    StartsWithFunction = UCase$(Left$(f, Len(functionName) + 1)) = UCase$(functionName & "(")
End Function

Private Function ArgsText(ByVal f As String) As String
    Dim firstParen As Long
    Dim lastParen As Long

    firstParen = InStr(1, f, "(")
    lastParen = InStrRev(f, ")")

    If firstParen = 0 Or lastParen = 0 Or lastParen <= firstParen Then
        ArgsText = f
    Else
        ArgsText = Mid$(f, firstParen + 1, lastParen - firstParen - 1)
    End If
End Function

Private Function SplitFunctionArguments(ByVal argText As String) As Collection

    Dim args As New Collection
    Dim i As Long
    Dim ch As String
    Dim currentArg As String
    Dim depth As Long
    Dim inQuotes As Boolean

    For i = 1 To Len(argText)

        ch = Mid$(argText, i, 1)

        If ch = """" Then
            inQuotes = Not inQuotes
            currentArg = currentArg & ch

        ElseIf ch = "(" And Not inQuotes Then
            depth = depth + 1
            currentArg = currentArg & ch

        ElseIf ch = ")" And Not inQuotes Then
            depth = depth - 1
            currentArg = currentArg & ch

        ElseIf ch = "," And depth = 0 And Not inQuotes Then
            args.Add Trim$(currentArg)
            currentArg = ""

        Else
            currentArg = currentArg & ch
        End If

    Next i

    If Len(Trim$(currentArg)) > 0 Then args.Add Trim$(currentArg)

    Set SplitFunctionArguments = args

End Function

Private Function CleanExpression(ByVal txt As String) As String

    txt = Trim$(txt)
    txt = Replace(txt, """", "")

    txt = Replace(txt, ">=", " is greater than or equal to ")
    txt = Replace(txt, "<=", " is less than or equal to ")
    txt = Replace(txt, "<>", " does not equal ")
    txt = Replace(txt, ">", " is greater than ")
    txt = Replace(txt, "<", " is less than ")
    txt = Replace(txt, "=", " equals ")

    CleanExpression = Trim$(txt)

End Function

