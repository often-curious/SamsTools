Attribute VB_Name = "m_EditNumbers"
Sub MultiplyNumbers(control As IRibbonControl)
    Dim rng As Range
    Dim i As Variant
    Dim originalFormula As String

    i = InputBox("Enter number to multiply by", "Input Required")

    ' Validate input
    If Not IsNumeric(i) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If

    For Each rng In Selection
        ' Skip empty cells or cells with non-numeric values
        If Not IsEmpty(rng.value) And IsNumeric(rng.value) Then
            If rng.HasFormula Then
                ' Append multiplication to existing formula
                originalFormula = rng.formula
                rng.formula = "=" & "(" & Mid(originalFormula, 2) & ") * " & i
            Else
                ' Create new formula from cell value
                rng.formula = "=" & rng.value & " * " & i
            End If
        End If
    Next rng
End Sub


Sub DivideNumbers(control As IRibbonControl)
    Dim rng As Range
    Dim i As Variant
    Dim originalFormula As String

    i = InputBox("Enter number to divide by", "Input Required")

    ' Validate input
    If Not IsNumeric(i) Or i = 0 Then
        MsgBox "Please enter a valid non-zero number.", vbExclamation
        Exit Sub
    End If

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
End Sub

Sub ConvertFormulasToValuesInSelection(control As IRibbonControl)

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    Dim rng As Range

    For Each rng In Selection

        If rng.HasFormula Then

            rng.formula = rng.value

        End If

    Next rng
    
        
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub

Sub TogglePercentNumber(control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection ' Change this if you want a specific range
    
    Dim cell As Range
    Dim fmt As String
    Dim val As Variant
    
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
End Sub

Sub ToggleSign(control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection ' Or specify a range like Range("A1:A10")
    
    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell.value) And Not IsEmpty(cell) Then
            cell.value = cell.value * -1
        End If
    Next cell
End Sub
