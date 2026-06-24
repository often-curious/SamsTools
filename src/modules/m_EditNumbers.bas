Attribute VB_Name = "m_EditNumbers"
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

Sub MultiplyNumbers(control As IRibbonControl)

    Dim rng As Range
    Dim data As Variant
    Dim r As Long, c As Long
    Dim multInput As String
    Dim mult As Double
    Dim f As Variant

    ' Turn off Excel's overhead
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    

    ' Ask for multiplier
GetInput:
    multInput = InputBox("Enter number to multiply by", "Input Required")

    ' Handle cancel (empty string AND user cancelled)
    If multInput = "" Then GoTo CleanUp

    ' Validate: must be numeric only
    If Not IsNumeric(multInput) Then
        MsgBox "Error: Please enter a valid numeric value (e.g. 2, -5, 3.5).", vbExclamation
        GoTo GetInput
    End If
    
    ShowLoading "Updating numbers..."
    
    mult = CDbl(multInput)

    Set rng = Selection
    data = rng.value

    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)

            If Not IsEmpty(data(r, c)) And IsNumeric(data(r, c)) Then

                f = rng.Cells(r, c).formula

                If f <> "" Then
                    If Left$(f, 1) = "=" Then f = Mid$(f, 2)
                    data(r, c) = "=" & "(" & f & ")*" & mult
                Else
                    data(r, c) = "=" & data(r, c) & "*" & mult
                End If

            End If

        Next c
    Next r

    rng.value = data

CleanUp:
    HideLoading
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub




Sub DivideNumbers(control As IRibbonControl)
    Dim rng As Range
    Dim i As Variant
    Dim originalFormula As String

    Call Off

    i = InputBox("Enter number to divide by", "Input Required")

    ' Validate input
    If Not IsNumeric(i) Or i = 0 Then
        MsgBox "Please enter a valid non-zero number.", vbExclamation
        Exit Sub
    End If

    ShowLoading "Updating numbers..."
    
    For Each rng In Selection
        ' Skip empty or non-numeric cells
        If Not IsEmpty(rng.value) And IsNumeric(rng.value) Then
            If rng.HasFormula Then
                ' Append division to existing formula
                originalFormula = rng.formula
                rng.formula = "=" & "(" & Mid(originalFormula, 2) & ") / " & i
            Else
                ' Create new formula from cell value
                rng.formula = "=" & rng.value & " / " & i
            End If
        End If
    Next rng
    
    HideLoading
    Call Onn
    
End Sub

Sub ConvertFormulasToValuesInSelection(control As IRibbonControl)

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    Dim rng As Range
    
    ShowLoading "Converting to values..."

    For Each rng In Selection

        If rng.HasFormula Then

            rng.formula = rng.value

        End If

    Next rng
    
    HideLoading
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub

Sub TogglePercentNumber(control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection ' Change this if you want a specific range
    
    Dim cell As Range
    Dim fmt As String
    Dim val As Variant
    
    Call Off
    ShowLoading "Updating numbers..."
    
    For Each cell In rng
        If Not IsEmpty(cell) And IsNumeric(cell.value) Then
            fmt = cell.numberFormat
            
            ' Check if cell is formatted as percentage
            If InStr(fmt, "%") > 0 Then
                ' Convert percent to number: multiply by 100 and remove % format
                val = cell.value * 100
                cell.value = val
                cell.numberFormat = "0.00" ' or "General" for no decimals
            Else
                ' Convert number to percent: divide by 100 and set % format
                val = cell.value / 100
                cell.value = val
                cell.numberFormat = "0.00%" ' 2 decimal places percent
            End If
        End If
    Next cell
    
    HideLoading
    Call Onn
    
End Sub

Sub ToggleSign(control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection ' Or specify a range like Range("A1:A10")
    
    Call Off
    ShowLoading "Changing sign..."
    
    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell.value) And Not IsEmpty(cell) Then
            cell.value = cell.value * -1
        End If
    Next cell
    
    HideLoading
    Call Onn
    
End Sub
