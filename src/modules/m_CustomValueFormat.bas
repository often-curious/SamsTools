Attribute VB_Name = "m_CustomValueFormat"
Sub ConfigureCustomNumberFormatToggle(control As IRibbonControl)

    frmNumberFormats.Show

End Sub

Sub ToggleCustomNumberFormat(control As IRibbonControl)

    Dim formats() As String
    Dim FormatCount As Long
    Dim i As Long
    Dim Fmt As String
    Dim TestCell As Range
    
    Dim rng As Range
    Dim currentFormat As String
    Dim NextIndex As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection

    Set TestCell = ThisWorkbook.Worksheets("Hidden").Range("A1")

    FormatCount = 0

    '----------------------------------------
    ' Build valid format list
    '----------------------------------------
    For i = 1 To 6

        Fmt = Trim(GetSetting("SamsTools", "FormatToggle", "Format" & i, ""))

        If Fmt <> "" Then

            On Error Resume Next
            TestCell.numberFormat = Fmt

            If Err.Number = 0 Then
                FormatCount = FormatCount + 1
                ReDim Preserve formats(1 To FormatCount)
                formats(FormatCount) = Fmt
            End If

            Err.Clear
            On Error GoTo 0

        End If

    Next i

    If FormatCount = 0 Then
        MsgBox "No valid number formats have been configured.", vbExclamation
        Exit Sub
    End If

    '----------------------------------------
    ' Determine current format (based on first cell)
    '----------------------------------------
    currentFormat = rng.Cells(1, 1).numberFormat

    NextIndex = 1

    For i = 1 To FormatCount

        If currentFormat = formats(i) Then

            If i = FormatCount Then
                NextIndex = 1
            Else
                NextIndex = i + 1
            End If

            Exit For

        End If

    Next i

    '----------------------------------------
    ' Apply to ALL selected cells
    '----------------------------------------
    rng.numberFormat = formats(NextIndex)


End Sub
