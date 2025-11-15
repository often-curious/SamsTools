Attribute VB_Name = "m_Pivots"
Sub PivotTable_ClassicSettings(control As IRibbonControl)

    Const DATA_NUM_FMT As String = "#,##0;[Red](#,##0)"

    Dim pt As PivotTable
    Dim pf As PivotField
    Dim i As Long

    On Error Resume Next
    ' Safely get the current PivotTable from the active cell
    Set pt = ActiveCell.PivotTable
    If pt Is Nothing Then
        Set pt = ActiveSheet.PivotTables(ActiveCell.PivotTable.Name)
    End If
    On Error GoTo 0

    If pt Is Nothing Then
        MsgBox "No Pivot Table selected." & vbCrLf & _
               "Please click a cell inside a PivotTable and try again.", _
               vbExclamation, "Pivot Table Selection"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    pt.ManualUpdate = True

    ' --- Classic / Tabular look and feel ---
    With pt
        On Error Resume Next
        .RowAxisLayout xlTabularRow
        .DisplayContextTooltips = False
        .ShowDrillIndicators = False
        .RepeatAllLabels xlRepeatLabels
        On Error GoTo 0
    End With

    ' --- Field-level settings (all fields) ---
    On Error Resume Next
    For Each pf In pt.PivotFields
        pf.AutoSort xlAscending, pf.Name
        For i = 1 To 12
            pf.Subtotals(i) = False
        Next i
    Next pf
    On Error GoTo 0

    ' --- Data fields: only affect SUM fields ---
    On Error Resume Next
    For Each pf In pt.DataFields
        ' Only change caption + format if it's a SUM field
        If pf.Function = xlSum Then
            pf.caption = pf.SourceName & " "  ' remove "Sum of" look
            pf.numberFormat = DATA_NUM_FMT
        End If
    Next pf
    On Error GoTo 0

    ' --- Housekeeping ---
    On Error Resume Next
    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
    pt.ManualUpdate = False
    pt.RefreshTable
    pt.HasAutoFormat = False
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Cleanup
    Set pf = Nothing
    Set pt = Nothing
End Sub


