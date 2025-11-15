Attribute VB_Name = "m_Charts"
Sub FormatChart(control As IRibbonControl)
    Dim cht As Chart
    Dim ser As Series

    'Check if a chart is selected
    On Error Resume Next
    Set cht = ActiveChart
    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "Please select a chart before running this macro.", vbExclamation
        Exit Sub
    End If

    'Remove axes (primary and secondary)
    On Error Resume Next
    With cht.Axes(xlValue, xlPrimary)
        .HasTitle = False
        .HasMajorGridlines = False
        .TickLabels.numberFormat = ";;;" ' Hides the tick labels
        .Format.line.Visible = msoFalse
    End With
    
    With cht.Axes(xlValue, xlSecondary)
        .HasTitle = False
        .HasMajorGridlines = False
        .TickLabels.numberFormat = ";;;"
        .Format.line.Visible = msoFalse
    End With
    On Error GoTo 0

    'Remove gridlines
    On Error Resume Next
    cht.Axes(xlCategory).MajorGridlines.Delete
    cht.Axes(xlCategory).MinorGridlines.Delete
    cht.Axes(xlValue).MajorGridlines.Delete
    cht.Axes(xlValue).MinorGridlines.Delete
    On Error GoTo 0

    'Set X-axis label position to low
    On Error Resume Next
    With cht.Axes(xlCategory)
        .TickLabelPosition = xlLow
    End With
    On Error GoTo 0

    'Add value data labels to each series
    For Each ser In cht.SeriesCollection
        With ser
            .HasDataLabels = True
            With .DataLabels
                .ShowValue = True
                .ShowSeriesName = False
                .ShowCategoryName = False
            End With
        End With
    Next ser
End Sub


Sub AdjustVerticalAxis(control As IRibbonControl)
    'PURPOSE: Adjust Y-Axis according to Min/Max of Chart Data

    Dim cht As Chart
    Dim srs As Series
    Dim FirstTime As Boolean
    Dim MaxNumber As Double
    Dim MinNumber As Double
    Dim MaxChartNumber As Double
    Dim MinChartNumber As Double
    Dim Padding As Double

    'Input Padding on Top of Min/Max Numbers (Percentage)
    Padding = 0.1  'Number between 0-1

    'Check if a chart is selected
    On Error Resume Next
    Set cht = ActiveChart
    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "Please select a chart before running this macro.", vbExclamation
        Exit Sub
    End If

    'Optimize Code
    Application.ScreenUpdating = False

    'First Time Looking at This Chart?
    FirstTime = True

    'Determine Chart's Overall Max/Min From Connected Data Source
    For Each srs In cht.SeriesCollection
        'Determine Maximum value in Series
        MaxNumber = Application.WorksheetFunction.Max(srs.Values)

        'Store value if currently the overall Maximum Value
        If FirstTime Then
            MaxChartNumber = MaxNumber
        ElseIf MaxNumber > MaxChartNumber Then
            MaxChartNumber = MaxNumber
        End If

        'Determine Minimum value in Series
        MinNumber = Application.WorksheetFunction.Min(srs.Values)

        'Store value if currently the overall Minimum Value
        If FirstTime Then
            MinChartNumber = MinNumber
        ElseIf MinNumber < MinChartNumber Or MinChartNumber = 0 Then
            MinChartNumber = MinNumber
        End If

        FirstTime = False
    Next srs

    'Rescale Y-Axis (Value Axis)
    On Error Resume Next ' In case axis is missing or not value axis
    With cht.Axes(xlValue)
        .MinimumScale = MinChartNumber * (1 - Padding)
        .MaximumScale = MaxChartNumber * (1 + Padding)
    End With
    On Error GoTo 0

    Application.ScreenUpdating = True
    'MsgBox "Y-axis scaled based on data range.", vbInformation
End Sub

Sub InsertHorizontalWaterfall(control As IRibbonControl)
    Dim tableTopLeft As Range
    
    Application.ScreenUpdating = False
    
    Set tableTopLeft = InsertWaterfallTableTemplate
    If tableTopLeft Is Nothing Then Exit Sub

    Call InsertWaterfallChart(xlColumnStacked, tableTopLeft)

    Application.ScreenUpdating = True
    MsgBox "Waterfall chart & input table created. You can now enter or paste your data in the yellow columns.", vbInformation
End Sub

Sub InsertVerticalWaterfall(control As IRibbonControl)
    Dim tableTopLeft As Range
    
    Application.ScreenUpdating = False
    Set tableTopLeft = InsertWaterfallTableTemplate
    If tableTopLeft Is Nothing Then Exit Sub

    Call InsertWaterfallChart(xlBarStacked, tableTopLeft)

    Application.ScreenUpdating = True
    MsgBox "Waterfall chart & input table created. You can now enter or paste your data in the yellow columns.", vbInformation
End Sub

Function InsertWaterfallTableTemplate() As Range
    Dim targetCell As Range
    Dim ws As Worksheet
    Dim tableTopLeft As Range
    Dim headers As Variant
    Dim i As Integer
    Dim headerRange1 As Range, headerRange2 As Range
    Dim fullRange As Range, inputRange As Range, dataValidationRange As Range

    ' Ask user to select top-left cell
    On Error Resume Next
    Set targetCell = Application.InputBox( _
        Prompt:="Select the top-left cell where the Waterfall Chart data should start:", _
        Title:="Insert Waterfall Table", _
        Type:=8)
    On Error GoTo 0

    If targetCell Is Nothing Then
        MsgBox "No cell selected. Macro cancelled.", vbExclamation
        Exit Function
    End If

    Set ws = targetCell.Worksheet
    Set tableTopLeft = targetCell

    ' Define headers
    headers = Array( _
        "Field Name", "Value", "Is Total?", _
        "Cumulative Total", "Totals", "Blank", "Up > 0", "Up < 0", "Down > 0", "Down < 0")



    ' --- Add Column Headers ---
    For i = 0 To UBound(headers)
        With tableTopLeft.Offset(1, i)
            .value = headers(i)
            .Font.Bold = True
            .Interior.color = RGB(0, 0, 0)
            .Font.color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.color = RGB(255, 255, 255)
        End With
    Next i

    ' --- Set Yellow Fill for First 3 Columns, Rows 1 to 3 ---
    Set inputRange = ws.Range(tableTopLeft.Offset(2, 0), tableTopLeft.Offset(20, 2))
    inputRange.Interior.color = RGB(255, 255, 200)

    ' --- Clear Fill for Remaining Columns, Rows 1 to 3 ---
    With ws.Range(tableTopLeft.Offset(2, 3), tableTopLeft.Offset(20, 9))
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' --- Apply Light Grey Borders to All Initial Table Area (Rows 1 to 21) ---
    Set fullRange = ws.Range(tableTopLeft.Offset(2, 0), tableTopLeft.Offset(20, 9))
    With fullRange.Borders
        .LineStyle = xlContinuous
        .color = RGB(200, 200, 200)
        .Weight = xlThin
    End With

    ' --- Add Data Validation to 'Total?' Column (Column C) ---
    Set dataValidationRange = ws.Range(tableTopLeft.Offset(2, 2), tableTopLeft.Offset(20, 2))
    With dataValidationRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Start,Y,N"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' --- Add Formulas to Calculation Columns (D to I) ---
        Dim r As Long
        Dim startRow As Long: startRow = tableTopLeft.row + 2 ' Row where formulas begin
        Dim endRow As Long: endRow = startRow + 18           ' Fill 19 rows
        Dim topRow As Long: topRow = tableTopLeft.row + 2 ' Row where data starts
    
        Dim baseCol As Long: baseCol = tableTopLeft.Column   ' Col A = 1
    
        For r = startRow To endRow
            With ws
                ' Cumulative Total (D)
                .Cells(r, baseCol + 3).formula = _
                    "=SUMIFS(" & .Cells(startRow, baseCol + 1).Address(False, False) & ":" & .Cells(r, baseCol + 1).Address(False, False) & "," & _
                                .Cells(startRow, baseCol + 2).Address(False, False) & ":" & .Cells(r, baseCol + 2).Address(False, False) & ",""<>" & "Y" & """)"
        
                ' Totals (E)
                .Cells(r, baseCol + 4).formula = "=IF(OR(" & .Cells(r, baseCol + 2).Address(False, False) & "=""Y""," & _
                                                        .Cells(r, baseCol + 2).Address(False, False) & "=""Start"")," & _
                                                        .Cells(r, baseCol + 3).Address(False, False) & ","""")"
        
                If r = startRow Then
                    ' Top row – fill others with 0
                    .Cells(r, baseCol + 5).value = 0 ' Blank
                    .Cells(r, baseCol + 6).value = 0 ' Up > 0
                    .Cells(r, baseCol + 7).value = 0 ' Up < 0
                    .Cells(r, baseCol + 8).value = 0 ' Down > 0
                    .Cells(r, baseCol + 9).value = 0 ' Down < 0
                Else
                    ' Full formulas below top row
                    .Cells(r, baseCol + 5).formula = "=IFERROR(IF(" & .Cells(r, baseCol + 2).Address(False, False) & "=""" & "Y" & """,""""," & _
                                                      "IF(" & .Cells(r - 1, baseCol + 3).Address(False, False) & "<0,MAX(" & .Cells(r - 1, baseCol + 3).Address(False, False) & "," & _
                                                      .Cells(r - 1, baseCol + 3).Address(False, False) & "-" & .Cells(r, baseCol + 7).Address(False, False) & ")," & _
                                                      "MIN(" & .Cells(r - 1, baseCol + 3).Address(False, False) & "," & .Cells(r - 1, baseCol + 3).Address(False, False) & "-" & _
                                                      .Cells(r, baseCol + 8).Address(False, False) & "))),0)"
        
                    .Cells(r, baseCol + 6).formula = "=IF(" & .Cells(r, baseCol + 2).Address(False, False) & "=""" & "Y" & """,0,MAX(0,MIN(" & _
                                                      .Cells(r, baseCol + 3).Address(False, False) & "," & .Cells(r, baseCol + 1).Address(False, False) & ")))"
        
                    .Cells(r, baseCol + 7).formula = "=IF(" & .Cells(r, baseCol + 2).Address(False, False) & "=""" & "Y" & """,0,-MAX(0," & _
                                                      .Cells(r, baseCol + 1).Address(False, False) & "-" & .Cells(r, baseCol + 6).Address(False, False) & "))"
        
                    .Cells(r, baseCol + 8).formula = "=IFERROR(IF(" & .Cells(r, baseCol + 2).Address(False, False) & "=""" & "Y" & """,0,MAX(0," & _
                                                      .Cells(r, baseCol + 9).Address(False, False) & "-" & .Cells(r, baseCol + 1).Address(False, False) & ")),0)"
        
                    .Cells(r, baseCol + 9).formula = "=IFERROR(IF(" & .Cells(r, baseCol + 2).Address(False, False) & "=""" & "Y" & """,0,MIN(0,MAX(" & _
                                                      .Cells(r - 1, baseCol + 3).Address(False, False) & "+" & .Cells(r, baseCol + 1).Address(False, False) & "," & _
                                                      .Cells(r, baseCol + 1).Address(False, False) & "))),0)"
                End If
            End With
        Next r


    ' Add marker at bottom of table
        Dim markerRow As Long
        markerRow = tableTopLeft.row + 2 + 18 + 1 ' adjust if your table has a different size
        
        With ws.Cells(markerRow, tableTopLeft.Column)
            .value = "<-- Delete unneeded cells above this -->"
            .Font.Italic = True
            .Font.color = RGB(128, 128, 128)
        End With



    ' --- Add Top Group Headers ---
    Set headerRange1 = ws.Range(tableTopLeft, tableTopLeft.Offset(0, 2)) ' First 3 cols
    Set headerRange2 = ws.Range(tableTopLeft.Offset(0, 3), tableTopLeft.Offset(0, 9)) ' Remaining 7 cols

    With headerRange1
        .Merge
        .value = "Enter Chart Data Below"
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.color = RGB(0, 0, 0)
        .Font.color = RGB(255, 255, 255)
        .Borders.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
    
    With headerRange1
        .MergeCells = False
        .HorizontalAlignment = xlCenterAcrossSelection
    End With

    With headerRange2
        .Merge
        .value = "Chart Data Table"
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.color = RGB(0, 0, 0)
        .Font.color = RGB(255, 255, 255)
        .Borders.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .MergeCells = False
    End With
    
    ws.Names.Add Name:="WaterfallInputData", RefersTo:=inputRange
    
    With headerRange2
        .MergeCells = False
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    
    Set InsertWaterfallTableTemplate = tableTopLeft
    
End Function

Private Sub InsertWaterfallChart(chartType As XlChartType, tableTopLeft As Range)
    Dim ws As Worksheet
    Dim baseCol As Long, startRow As Long, endRow As Long
    Dim chartObj As ChartObject
    Dim i As Integer
    Dim chartLeftPos As Double
    Dim seriesNames As Variant
    Dim seriesColors As Variant

    Set ws = tableTopLeft.Worksheet
    baseCol = tableTopLeft.Column
    startRow = tableTopLeft.row + 2
    endRow = startRow + 18

    chartLeftPos = ws.Cells(startRow, baseCol + 11).Left

    ' Add chart object
    Set chartObj = ws.ChartObjects.Add( _
        Left:=chartLeftPos, _
        Top:=ws.Cells(startRow, baseCol).Top, _
        width:=500, Height:=300)

    ' Define series headers and their colors
    seriesNames = Array("Totals", "Blank", "Up > 0", "Up < 0", "Down > 0", "Down < 0")
    seriesColors = Array(RGB(0, 0, 128), -1, RGB(0, 176, 80), RGB(0, 176, 80), RGB(255, 0, 0), RGB(255, 0, 0))

    With chartObj.Chart
        .chartType = chartType
        .HasTitle = True
        .ChartTitle.text = "Waterfall Chart Breakdown"
        .HasLegend = False

        ' Clear any auto series added
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add each series explicitly from columns E to J (baseCol + 4 to baseCol + 9)
        For i = 0 To 5
            With .SeriesCollection.NewSeries
                .Name = seriesNames(i)
                .Values = ws.Range(ws.Cells(startRow, baseCol + 4 + i), ws.Cells(endRow, baseCol + 4 + i))
                .XValues = ws.Range(ws.Cells(startRow, baseCol), ws.Cells(endRow, baseCol)) ' Category labels (Column A)
                
                ' Format series
                If seriesColors(i) = -1 Then
                    ' Blank series = No fill
                    .Format.Fill.Visible = msoFalse
                    .Format.line.Visible = msoFalse
                Else
                    .Format.Fill.ForeColor.RGB = seriesColors(i)
                End If
                
            End With
        Next i
        
        
        ' Format chart
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        .Axes(xlValue).Delete
        .ChartGroups(1).GapWidth = 50
        On Error GoTo 0
        .Axes(xlCategory).TickLabelPosition = xlLow
        If chartType = xlBarStacked Then
            .Axes(xlCategory).ReversePlotOrder = True
        End If
        
    End With
End Sub



