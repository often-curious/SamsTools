Attribute VB_Name = "m_ReverseRange"
Option Explicit

' Reverses the order of the longer dimension of each selected area.
Public Sub ReverseSelectedRange(control As IRibbonControl)
    Dim ar As Range
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim errNum As Long
    Dim errMsg As String
    Dim merged As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range of cells first.", vbExclamation
        Exit Sub
    End If

    ' Pre-checks: merged cells and protected sheets
    For Each ar In Selection.Areas
        If IsNull(ar.MergeCells) Then
            merged = True
        ElseIf ar.MergeCells Then
            merged = True
        End If
        If ar.Worksheet.ProtectContents Then
            MsgBox "The sheet is protected. Unprotect it and try again.", vbExclamation
            Exit Sub
        End If
    Next ar
    If merged Then
        MsgBox "The selection contains merged cells. Unmerge them and try again.", vbExclamation
        Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For Each ar In Selection.Areas
        ReverseArea ar
    Next ar

CleanExit:
    On Error Resume Next            ' single restore block, nothing else runs here
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0
    If errNum <> 0 Then MsgBox "Reverse failed: " & errMsg, vbCritical
    Exit Sub

CleanFail:
    errNum = Err.Number
    errMsg = Err.Description
    Resume CleanExit
End Sub

Private Sub ReverseArea(ByVal ar As Range)
    Dim src As Variant, dst() As Variant
    Dim r As Long, c As Long, nR As Long, nC As Long
    Dim flipRows As Boolean

    If ar.Cells.count = 1 Then Exit Sub

    src = ar.formula                ' 2D array, 1-based; keeps formulas as text
    nR = UBound(src, 1)
    nC = UBound(src, 2)
    flipRows = (nR >= nC)           ' tie or square: flip top to bottom
    ReDim dst(1 To nR, 1 To nC)

    For r = 1 To nR
        For c = 1 To nC
            If flipRows Then
                dst(nR - r + 1, c) = src(r, c)
            Else
                dst(r, nC - c + 1) = src(r, c)
            End If
        Next c
    Next r

    ar.formula = dst
End Sub
