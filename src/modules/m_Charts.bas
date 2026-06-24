Attribute VB_Name = "m_Charts"

Option Explicit

Public gSourceChart As Chart

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
    
    ShowLoading "Formatting chart..."

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
    
    HideLoading
    
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
    ShowLoading "Creating chart..."
    Set tableTopLeft = InsertWaterfallTableTemplate
    If tableTopLeft Is Nothing Then Exit Sub
    
    Call InsertWaterfallChart(xlColumnStacked, tableTopLeft)

    HideLoading
    
    Application.ScreenUpdating = True
    MsgBox "Waterfall chart & input table created. You can now enter or paste your data in the yellow columns.", vbInformation
End Sub

Sub InsertVerticalWaterfall(control As IRibbonControl)
    Dim tableTopLeft As Range
    
    Application.ScreenUpdating = False
    ShowLoading "Creating chart..."
    Set tableTopLeft = InsertWaterfallTableTemplate
    If tableTopLeft Is Nothing Then Exit Sub
    
    Call InsertWaterfallChart(xlBarStacked, tableTopLeft)

    HideLoading
    
    Application.ScreenUpdating = True
    MsgBox "Waterfall chart & input table created. You can now enter or paste your data in the yellow columns.", vbInformation
End Sub

Function InsertWaterfallTableTemplate() As Range
    Dim TargetCell As Range
    Dim ws As Worksheet
    Dim tableTopLeft As Range
    Dim headers As Variant
    Dim i As Integer
    Dim headerRange1 As Range, headerRange2 As Range
    Dim fullRange As Range, inputRange As Range, dataValidationRange As Range

    ' Ask user to select top-left cell
    On Error Resume Next
    Set TargetCell = Application.InputBox( _
        Prompt:="Select the top-left cell where the Waterfall Chart data should start:", _
        Title:="Insert Waterfall Table", _
        Type:=8)
    On Error GoTo 0

    If TargetCell Is Nothing Then
        MsgBox "No cell selected. Macro cancelled.", vbExclamation
        Exit Function
    End If

    Set ws = TargetCell.Worksheet
    Set tableTopLeft = TargetCell

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
        .Interior.colorIndex = xlColorIndexNone
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
        .chartTitle.text = "Waterfall Chart Breakdown"
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


Sub CreateForecastChart(control As IRibbonControl)

    Dim ws As Worksheet
    Dim startCell As Range
    Dim colourCell As Range
    Dim baseColor As Long
    Dim ch As ChartObject
    Dim dataRange As Range
    Dim markerColor As Long
    
    Application.ScreenUpdating = False
    
    ShowLoading "Creating Chart..."

    '========================
    ' PICK BASE COLOUR
    '========================
    
    baseColor = RGB(0, 0, 0)
    markerColor = LighterColor(baseColor, 0.5)

    '========================
    ' PICK INSERT LOCATION
    '========================
    Set startCell = Application.InputBox( _
        Prompt:="Select the top-left cell where the chart data should be inserted.", _
        Title:="Select Insert Location", _
        Type:=8)

    On Error GoTo 0

    If startCell Is Nothing Then Exit Sub

    Set ws = startCell.Worksheet

    With ws

        '========================
        ' BUILD DATA TABLE
        '========================
        startCell.value = "[Chart Name Here]"
        With startCell
            .Interior.color = RGB(255, 247, 209)
            With .Borders
                .LineStyle = xlContinuous
                .color = RGB(217, 217, 217)
                .Weight = xlThin
            End With
        End With

        startCell.Offset(1, 0).Resize(4, 1).value = _
            Application.Transpose(Array("Prior Year", "Plan", "Actual", "Forecast"))

        startCell.Offset(0, 1).value = "Version"

        startCell.Offset(1, 1).Resize(4, 1).value = _
            Application.Transpose(Array("PY", "PL", "AC", "FC"))

        startCell.Offset(0, 2).Resize(1, 12).formula = _
            "=""Period ""&COLUMN()-" & startCell.Offset(0, 1).Column

        startCell.Resize(1, 14).Font.Bold = True

        ' Sample data
        startCell.Offset(1, 2).Resize(4, 12).value = "=RANDBETWEEN(10,25)"
        With startCell.Offset(1, 2).Resize(4, 12)
            .Interior.color = RGB(255, 247, 209)
            With .Borders
                .LineStyle = xlContinuous
                .color = RGB(217, 217, 217)
                .Weight = xlThin
            End With
        End With
        

        ' Clear future periods
        startCell.Offset(4, 2).Resize(1, 6).ClearContents
        startCell.Offset(3, 8).Resize(1, 6).ClearContents

        Set dataRange = startCell.Offset(0, 1).Resize(5, 13)

        '========================
        ' CREATE CHART
        '========================
        Set ch = .ChartObjects.Add( _
            Left:=startCell.Left, _
            Top:=startCell.Offset(7, 0).Top, _
            width:=620, _
            Height:=300)

    End With

    With ch.Chart

        .SetSourceData Source:=dataRange
        .chartType = xlColumnClustered

        .HasTitle = True
        .chartTitle.text = startCell.value
        
        ' Remove Y axis line
        With .Axes(xlValue)
            .Format.line.Visible = msoFalse
        End With
    
        ' Light grey gridlines
        With .Axes(xlValue).MajorGridlines.Format.line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(217, 217, 217)
            .Weight = 0.75
        End With

        '========================
        ' SERIES 1 - MARKERS
        '========================
        With .FullSeriesCollection(1)

            .chartType = xlLineMarkers
            .Format.line.Visible = msoFalse

            .markerStyle = xlMarkerStyleCircle
            .MarkerSize = 10

            .MarkerBackgroundColor = markerColor
            .MarkerForegroundColor = RGB(255, 255, 255)
            

        End With

        '========================
        ' SERIES 2
        '========================
        With .FullSeriesCollection(2)

            .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)

            .Format.line.ForeColor.RGB = baseColor
            .Format.line.Weight = 0.75

        End With

        '========================
        ' SERIES 3
        '========================
        With .FullSeriesCollection(3)

            .Format.Fill.ForeColor.RGB = baseColor

            .Format.line.ForeColor.RGB = baseColor
            .Format.line.Weight = 0.75

        End With

        '========================
        ' SERIES 4
        '========================
        With .FullSeriesCollection(4)
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.RGB = baseColor
            .Format.Fill.BackColor.RGB = RGB(255, 255, 255)
            .Format.Fill.Patterned msoPatternLightUpwardDiagonal

            .Format.line.ForeColor.RGB = baseColor
            .Format.line.Weight = 0.75

        End With

        .ChartGroups(1).GapWidth = 100
        .ChartGroups(1).Overlap = 80

    End With

HideLoading

Application.ScreenUpdating = True
MsgBox "Chart & input table created. You can now enter or paste your data in the yellow cells."

End Sub

Function LighterColor(ByVal baseColor As Long, ByVal pct As Double) As Long

    Dim r As Long, g As Long, b As Long

    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256

    r = r + (255 - r) * pct
    g = g + (255 - g) * pct
    b = b + (255 - b) * pct

    LighterColor = RGB(r, g, b)

End Function

Sub ToggleSmoothLines(control As IRibbonControl)

    Dim cht As Chart
    Dim s As Series
    Dim makeSmooth As Boolean
    Dim hasLineSeries As Boolean

    On Error Resume Next

    If TypeName(Selection) = "ChartArea" Or TypeName(Selection) = "PlotArea" Then
        Set cht = ActiveChart
    ElseIf Not ActiveChart Is Nothing Then
        Set cht = ActiveChart
    End If

    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    ' Check if chart contains line-type series
    hasLineSeries = False

    For Each s In cht.SeriesCollection
        Select Case s.chartType
            Case xlLine, xlLineMarkers, xlLineMarkersStacked, _
                 xlLineMarkersStacked100, xlLineStacked, _
                 xlLineStacked100, xlXYScatterLines, _
                 xlXYScatterLinesNoMarkers, xlXYScatterSmooth, _
                 xlXYScatterSmoothNoMarkers

                hasLineSeries = True
                Exit For
        End Select
    Next s

    If Not hasLineSeries Then
        MsgBox "This chart type does not support smoothed lines.", vbExclamation
        Exit Sub
    End If

    ' Determine current state from first valid series
    On Error Resume Next
    makeSmooth = Not cht.SeriesCollection(1).Smooth
    On Error GoTo 0

    ' Apply toggle only to supported series
    For Each s In cht.SeriesCollection

        Select Case s.chartType
            Case xlLine, xlLineMarkers, xlLineMarkersStacked, _
                 xlLineMarkersStacked100, xlLineStacked, _
                 xlLineStacked100, xlXYScatterLines, _
                 xlXYScatterLinesNoMarkers, xlXYScatterSmooth, _
                 xlXYScatterSmoothNoMarkers

                On Error Resume Next
                s.Smooth = makeSmooth
                On Error GoTo 0

        End Select

    Next s

    If makeSmooth Then
        MsgBox "Smoothed lines turned ON.", vbInformation
    Else
        MsgBox "Smoothed lines turned OFF.", vbInformation
    End If

End Sub

Sub AddSeriesNameToDataLabels(control As IRibbonControl)

    Dim cht As Chart
    Dim s As Series
    Dim lastPoint As Long

    On Error Resume Next
    Set cht = ActiveChart
    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    If cht.SeriesCollection.count = 0 Then
        MsgBox "The selected chart has no series.", vbExclamation
        Exit Sub
    End If

    For Each s In cht.SeriesCollection

        ' Only apply to line or column series, including combo charts
        If IsLineChartSeries(s.chartType) Or IsColumnChartSeries(s.chartType) Then

            lastPoint = s.Points.count

            If lastPoint > 0 Then

                ' Add a data label to the last point only if needed
                s.Points(lastPoint).HasDataLabel = True

                With s.Points(lastPoint).DataLabel
                    .ShowSeriesName = True
                    .ShowValue = False
                    .ShowCategoryName = False
                    .ShowLegendKey = False

                    On Error Resume Next

                    If IsLineChartSeries(s.chartType) Then
                            .position = xlLabelPositionRight
                        Else
                            .position = xlLabelPositionAbove
                    End If
                    
                    On Error GoTo 0
                End With

            End If

        End If

    Next s

End Sub


Private Function IsLineChartSeries(ByVal chartType As XlChartType) As Boolean

    Select Case chartType
        Case xlLine, _
             xlLineMarkers, _
             xlLineMarkersStacked, _
             xlLineMarkersStacked100, _
             xlLineStacked, _
             xlLineStacked100

            IsLineChartSeries = True

        Case Else
            IsLineChartSeries = False
    End Select

End Function


Private Function IsColumnChartSeries(ByVal chartType As XlChartType) As Boolean

    Select Case chartType
        Case xlColumnClustered, _
             xlColumnStacked, _
             xlColumnStacked100

            IsColumnChartSeries = True

        Case Else
            IsColumnChartSeries = False
    End Select

End Function


'========================
' COPY CHART STYLES
'========================

Public Sub CopyChartFormatting(control As IRibbonControl)

    If ActiveChart Is Nothing Then
        MsgBox "Please select the SOURCE chart first.", vbExclamation
        Exit Sub
    End If

    Set gSourceChart = ActiveChart
    
    If IsSpecialChartType(ActiveChart.chartType) Then
        MsgBox "Source chart type is not supported by the Copy Chart Formatting tool.", vbExclamation
        Exit Sub
    End If
    
    frmChartPicker.Show
    
    Set gSourceChart = Nothing

End Sub

Public Sub ApplyFormattingToTargetChart(ByVal wsName As String, ByVal chartName As String)

Dim tgtChart As Chart

Set tgtChart = ActiveWorkbook.Worksheets(wsName).ChartObjects(chartName).Chart

If gSourceChart Is Nothing Then
    MsgBox "Source chart was lost. Please try again.", vbExclamation
    Exit Sub
End If

CopyChartFormatFull gSourceChart, tgtChart

'MsgBox "Chart formatting copied to " & chartName & ".", vbInformation

End Sub

Private Sub CopyChartFormatFull(ByVal src As Chart, ByVal tgt As Chart)

Dim i As Long
Dim sSrc As Series
Dim sTgt As Series

    If IsSpecialChartType(ActiveChart.chartType) Or IsSpecialChartType(tgt.chartType) Then
        MsgBox "Target chart type is not supported by the Copy Chart Formatting tool.", vbExclamation
        Exit Sub
    End If

Application.ScreenUpdating = False

' Basic chart settings
On Error Resume Next
tgt.chartType = src.chartType
tgt.HasLegend = src.HasLegend
tgt.HasTitle = src.HasTitle
On Error GoTo 0

' Chart title
tgt.HasTitle = src.HasTitle

If src.HasTitle And tgt.HasTitle Then
    CopyChartTitleFormat src, tgt
End If


' Chart and plot area
tgt.ChartArea.Format.Fill.Visible = msoFalse
tgt.ChartArea.Format.line.Visible = msoFalse

tgt.PlotArea.Format.Fill.Visible = msoFalse
tgt.PlotArea.Format.line.Visible = msoFalse

' Axes
CopyAxis src, tgt, xlCategory, xlPrimary
CopyAxis src, tgt, xlValue, xlPrimary
CopyAxis src, tgt, xlCategory, xlSecondary
CopyAxis src, tgt, xlValue, xlSecondary

' Series
For i = 1 To WorksheetFunction.Min(src.SeriesCollection.count, tgt.SeriesCollection.count)

    Set sSrc = src.SeriesCollection(i)
    Set sTgt = tgt.SeriesCollection(i)

    On Error Resume Next
    sTgt.chartType = sSrc.chartType
    sTgt.axisGroup = sSrc.axisGroup
    
    CopySeriesLineFormat sSrc, sTgt

    CopyFillFormat sSrc.Format.Fill, sTgt.Format.Fill
    
    'Loop through series in case they have different styles
    Dim p As Long
    For p = 1 To WorksheetFunction.Min(sSrc.Points.count, sTgt.Points.count)
        CopyFillFormat sSrc.Points(p).Format.Fill, sTgt.Points(p).Format.Fill
    Next p

    CopyMarkerFormat sSrc, sTgt
    
    sTgt.Smooth = sSrc.Smooth
    On Error GoTo 0

    CopySeriesDataLabels sSrc, sTgt

' Legend
If src.HasLegend Then
    tgt.HasLegend = True
    tgt.Legend.position = src.Legend.position

    ' Legend fill/line
    tgt.Legend.Format.Fill.Visible = msoFalse
    tgt.Legend.Format.line.Visible = msoFalse

Else
    tgt.HasLegend = False
End If

Next i

CopyLegendFont src, tgt

Application.ScreenUpdating = True

End Sub

Private Sub CopyAxis(ByVal src As Chart, ByVal tgt As Chart, _
                     ByVal axisType As XlAxisType, _
                     ByVal axisGroup As XlAxisGroup)
Dim srcAxis As Axis
Dim tgtAxis As Axis

On Error Resume Next

tgt.HasAxis(axisType, axisGroup) = src.HasAxis(axisType, axisGroup)

If src.HasAxis(axisType, axisGroup) = False Then Exit Sub

Set srcAxis = src.Axes(axisType, axisGroup)
Set tgtAxis = tgt.Axes(axisType, axisGroup)

tgtAxis.TickLabelPosition = srcAxis.TickLabelPosition
tgtAxis.MajorTickMark = srcAxis.MajorTickMark
tgtAxis.MinorTickMark = srcAxis.MinorTickMark
tgtAxis.HasTitle = srcAxis.HasTitle
tgtAxis.TickLabels.numberFormat = srcAxis.TickLabels.numberFormat

CopyAxisLineFormat srcAxis, tgtAxis

tgtAxis.HasMajorGridlines = srcAxis.HasMajorGridlines
If srcAxis.HasMajorGridlines Then
    tgtAxis.MajorGridlines.Format.line.Visible = srcAxis.MajorGridlines.Format.line.Visible
    tgtAxis.MajorGridlines.Format.line.ForeColor.RGB = srcAxis.MajorGridlines.Format.line.ForeColor.RGB
    tgtAxis.MajorGridlines.Format.line.Weight = srcAxis.MajorGridlines.Format.line.Weight
    tgtAxis.MajorGridlines.Format.line.DashStyle = srcAxis.MajorGridlines.Format.line.DashStyle
End If

If srcAxis.HasTitle Then
    tgtAxis.AxisTitle.text = srcAxis.AxisTitle.text
End If

On Error GoTo 0

End Sub

Private Sub CopySeriesDataLabels(ByVal sSrc As Series, ByVal sTgt As Series)

Dim i As Long
Dim maxPoints As Long

On Error Resume Next

sTgt.HasDataLabels = sSrc.HasDataLabels

maxPoints = WorksheetFunction.Min(sSrc.Points.count, sTgt.Points.count)

For i = 1 To maxPoints

    sTgt.Points(i).HasDataLabel = sSrc.Points(i).HasDataLabel

    If sSrc.Points(i).HasDataLabel Then

        With sTgt.Points(i).DataLabel

            ' Label content/options
            .ShowValue = sSrc.Points(i).DataLabel.ShowValue
            .ShowSeriesName = sSrc.Points(i).DataLabel.ShowSeriesName
            .ShowCategoryName = sSrc.Points(i).DataLabel.ShowCategoryName
            .ShowPercentage = sSrc.Points(i).DataLabel.ShowPercentage
            .ShowLegendKey = sSrc.Points(i).DataLabel.ShowLegendKey
            .Separator = sSrc.Points(i).DataLabel.Separator
            .position = sSrc.Points(i).DataLabel.position
            .numberFormat = sSrc.Points(i).DataLabel.numberFormat

            ' Text colour/font
            .Font.Name = sSrc.Points(i).DataLabel.Font.Name
            .Font.SIZE = sSrc.Points(i).DataLabel.Font.SIZE
            .Font.Bold = sSrc.Points(i).DataLabel.Font.Bold
            .Font.Italic = sSrc.Points(i).DataLabel.Font.Italic
            .Font.color = sSrc.Points(i).DataLabel.Font.color

            ' Background fill
            If sSrc.Points(i).DataLabel.Format.Fill.Visible = msoFalse Then
                .Format.Fill.Visible = msoFalse
            Else
                .Format.Fill.Visible = msoTrue
                .Format.Fill.Solid
                .Format.Fill.ForeColor.RGB = _
                    sSrc.Points(i).DataLabel.Format.Fill.ForeColor.RGB
                .Format.Fill.Transparency = _
                    sSrc.Points(i).DataLabel.Format.Fill.Transparency
            End If

            ' Label border
            If sSrc.Points(i).DataLabel.Format.line.Visible = msoFalse Then
                .Format.line.Visible = msoFalse
            Else
                .Format.line.Visible = msoTrue
                .Format.line.ForeColor.RGB = _
                    sSrc.Points(i).DataLabel.Format.line.ForeColor.RGB
                .Format.line.Weight = _
                    sSrc.Points(i).DataLabel.Format.line.Weight
                .Format.line.DashStyle = _
                    sSrc.Points(i).DataLabel.Format.line.DashStyle
            End If

        End With

    End If

Next i

On Error GoTo 0

End Sub

Private Sub CopyLineAndFill(ByVal srcFmt As ChartFormat, ByVal tgtFmt As ChartFormat)

On Error Resume Next

' Fill
If srcFmt.Fill.Visible = msoFalse Then
    tgtFmt.Fill.Visible = msoFalse
Else
    tgtFmt.Fill.Visible = msoTrue
    tgtFmt.Fill.Solid
    tgtFmt.Fill.ForeColor.RGB = srcFmt.Fill.ForeColor.RGB
    tgtFmt.Fill.Transparency = srcFmt.Fill.Transparency
End If

' Line
If srcFmt.line.Visible = msoFalse Then
    tgtFmt.line.Visible = msoFalse
Else
    tgtFmt.line.Visible = msoTrue
    tgtFmt.line.ForeColor.RGB = srcFmt.line.ForeColor.RGB
    tgtFmt.line.Transparency = srcFmt.line.Transparency
    tgtFmt.line.Weight = srcFmt.line.Weight
    tgtFmt.line.DashStyle = srcFmt.line.DashStyle
End If

On Error GoTo 0

End Sub

Private Sub CopyFontStandard(ByVal srcFont As Object, ByVal tgtFont As Object)

On Error Resume Next

tgtFont.Name = srcFont.Name
tgtFont.SIZE = srcFont.SIZE
tgtFont.Bold = srcFont.Bold
tgtFont.Italic = srcFont.Italic
tgtFont.Fill.ForeColor.RGB = srcFont.Fill.ForeColor.RGB

On Error GoTo 0

End Sub

Private Sub CopyFillFormat(ByVal srcFill As FillFormat, ByVal tgtFill As FillFormat)

On Error Resume Next

If srcFill.Visible = msoFalse Then
    tgtFill.Visible = msoFalse
Else
    tgtFill.Visible = msoTrue

    ' Try to copy pattern fill first
    Err.Clear
    tgtFill.Patterned srcFill.pattern

    If Err.Number = 0 Then
        tgtFill.ForeColor.RGB = srcFill.ForeColor.RGB
        tgtFill.BackColor.RGB = srcFill.BackColor.RGB
    Else
        Err.Clear
        tgtFill.Solid
        tgtFill.ForeColor.RGB = srcFill.ForeColor.RGB
        tgtFill.Transparency = srcFill.Transparency
    End If
End If

On Error GoTo 0

End Sub

Private Sub CopySeriesLineFormat(ByVal sSrc As Series, ByVal sTgt As Series)

On Error Resume Next

' Handle default/automatic borders
If sSrc.Format.line.Visible = msoFalse Then

    sTgt.Format.line.Visible = msoFalse
    Exit Sub

End If

' Excel often reports Automatic as black
If sSrc.Format.line.ForeColor.RGB = RGB(0, 0, 0) _
   And sSrc.Format.line.Weight <= 1 Then

    sTgt.Format.line.Visible = msoFalse
    Exit Sub

End If

' Genuine custom border
With sTgt.Format.line
    .Visible = msoTrue
    .ForeColor.RGB = sSrc.Format.line.ForeColor.RGB
    .Transparency = sSrc.Format.line.Transparency
    .Weight = sSrc.Format.line.Weight
    .DashStyle = sSrc.Format.line.DashStyle
End With

On Error GoTo 0

End Sub

Private Sub CopyLegendFont(ByVal src As Chart, ByVal tgt As Chart)

On Error Resume Next

If src.HasLegend = False Then
    tgt.HasLegend = False
    Exit Sub
End If

tgt.HasLegend = True
tgt.Legend.position = src.Legend.position

' Force legacy legend font
With tgt.Legend.Font
    .Name = src.Legend.Font.Name
    .SIZE = src.Legend.Font.SIZE
    .Bold = src.Legend.Font.Bold
    .Italic = src.Legend.Font.Italic
    .color = src.Legend.Font.color
End With

' Force each legend entry font individually
Dim i As Long
For i = 1 To WorksheetFunction.Min(src.Legend.LegendEntries.count, tgt.Legend.LegendEntries.count)
    With tgt.Legend.LegendEntries(i).Font
        .Name = src.Legend.LegendEntries(i).Font.Name
        .SIZE = src.Legend.LegendEntries(i).Font.SIZE
        .Bold = src.Legend.LegendEntries(i).Font.Bold
        .Italic = src.Legend.LegendEntries(i).Font.Italic
        .color = src.Legend.LegendEntries(i).Font.color
    End With
Next i

On Error GoTo 0

End Sub

Private Sub CopyMarkerFormat(ByVal sSrc As Series, ByVal sTgt As Series)

On Error Resume Next

' If source has no markers, remove markers from target
If sSrc.markerStyle = xlMarkerStyleNone Then
    sTgt.markerStyle = xlMarkerStyleNone
    Exit Sub
End If

' Copy marker type and size
sTgt.markerStyle = sSrc.markerStyle
sTgt.MarkerSize = sSrc.MarkerSize

' Copy marker colours
sTgt.MarkerForegroundColor = sSrc.MarkerForegroundColor
sTgt.MarkerBackgroundColor = sSrc.MarkerBackgroundColor

' Copy newer marker fill/line formatting where available
sTgt.Format.Fill.Visible = sSrc.Format.Fill.Visible
sTgt.Format.line.Visible = sSrc.Format.line.Visible

On Error GoTo 0

End Sub

Private Sub CopyChartTitleFormat(ByVal src As Chart, ByVal tgt As Chart)

On Error Resume Next

' Font
With tgt.chartTitle.Font
    .Name = src.chartTitle.Font.Name
    .SIZE = src.chartTitle.Font.SIZE
    .Bold = src.chartTitle.Font.Bold
    .Italic = src.chartTitle.Font.Italic
    .color = src.chartTitle.Font.color
End With

' Fill and border
tgt.chartTitle.Format.Fill.Visible = msoFalse
tgt.chartTitle.Format.line.Visible = msoFalse

If src.chartTitle.Format.line.Visible Then
    tgt.chartTitle.Format.line.ForeColor.RGB = src.chartTitle.Format.line.ForeColor.RGB
    tgt.chartTitle.Format.line.Weight = src.chartTitle.Format.line.Weight
    tgt.chartTitle.Format.line.DashStyle = src.chartTitle.Format.line.DashStyle
End If

On Error GoTo 0

End Sub

Private Sub CopyAxisLineFormat(ByVal srcAxis As Axis, ByVal tgtAxis As Axis)

On Error Resume Next

' If source axis line is hidden, keep target hidden
If srcAxis.Format.line.Visible = msoFalse Then
    tgtAxis.Format.line.Visible = msoFalse
    Exit Sub
End If

' Treat default/automatic black axis line as no line
If srcAxis.Format.line.ForeColor.RGB = RGB(0, 0, 0) _
   And srcAxis.Format.line.Weight <= 1 Then

    tgtAxis.Format.line.Visible = msoFalse
    Exit Sub

End If

' Copy genuine custom axis line
With tgtAxis.Format.line
    .Visible = msoTrue
    .ForeColor.RGB = srcAxis.Format.line.ForeColor.RGB
    .Transparency = srcAxis.Format.line.Transparency
    .Weight = srcAxis.Format.line.Weight
    .DashStyle = srcAxis.Format.line.DashStyle
End With

On Error GoTo 0

End Sub

Private Function IsSpecialChartType(ByVal chartType As XlChartType) As Boolean
    Select Case chartType
        Case xlWaterfall, xlHistogram, xlPareto, xlBoxwhisker, _
             xlTreemap, xlSunburst, xlFunnel, xlRegionMap
            IsSpecialChartType = True
    End Select
End Function

'========================
' EMBED DATA IN CHART
'========================

Public Sub EmbedChartDataInChart(control As IRibbonControl)

    Dim cht As Chart
    Dim s As Series

    On Error Resume Next
    Set cht = ActiveChart
    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If
    
    Select Case cht.chartType
        Case xlWaterfall, _
             xlTreemap, _
             xlSunburst, _
             xlHistogram, _
             xlPareto, _
             xlBoxwhisker, _
             xlFunnel
    
            MsgBox "This chart type is not supported.", vbExclamation
            Exit Sub
    End Select

    Application.ScreenUpdating = False

    EmbedChartTitle cht

    For Each s In cht.SeriesCollection
        s.Name = EmbeddedText(s.Name)

        On Error Resume Next
        s.XValues = ArrayFromValues(s.XValues)
        s.Values = ArrayFromValues(s.Values)
        On Error GoTo 0
    Next s

    Application.ScreenUpdating = True

    MsgBox "Chart data and title have been embedded.", vbInformation

End Sub

Private Sub EmbedChartTitle(ByVal cht As Chart)

    Dim titleText As String

    If cht.HasTitle Then
        On Error Resume Next
        titleText = cht.chartTitle.text
        On Error GoTo 0

        If Len(titleText) > 0 Then
            cht.chartTitle.text = titleText
        End If
    End If

End Sub


Private Function EmbeddedText(ByVal txt As String) As String

    On Error GoTo SafeExit

    If Left$(txt, 1) = "=" Then
        EmbeddedText = CStr(Application.Evaluate(txt))
    Else
        EmbeddedText = txt
    End If

    Exit Function

SafeExit:
    EmbeddedText = txt

End Function


Private Function ArrayFromValues(ByVal vals As Variant) As Variant

    Dim arr() As Variant
    Dim i As Long

    If IsArray(vals) Then
        ReDim arr(LBound(vals) To UBound(vals))

        For i = LBound(vals) To UBound(vals)
            arr(i) = vals(i)
        Next i

        ArrayFromValues = arr
    Else
        ArrayFromValues = vals
    End If

End Function
