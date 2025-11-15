Attribute VB_Name = "m_Formatting"
Sub CentreContent(control As IRibbonControl)
    Dim rng As Range
    Dim vAlign As XlVAlign
    Dim hAlign As XlHAlign
    
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Please select one or more cells first.", vbExclamation
        Exit Sub
    End If
    
    ' Read current alignment from the first cell in the selection
    With rng.Cells(1)
        vAlign = .VerticalAlignment
        hAlign = .HorizontalAlignment
    End With
    
    ' Toggle logic
    If vAlign = xlCenter And hAlign <> xlCenter Then
        ' Currently vertically centered only -> set both vertically and horizontally centered
        rng.VerticalAlignment = xlCenter
        rng.HorizontalAlignment = xlCenter
    ElseIf vAlign = xlCenter And hAlign = xlCenter Then
        ' Currently both centered -> set horizontally centered only
        rng.VerticalAlignment = xlBottom  ' Or xlTop or xlGeneral if you want default
        rng.HorizontalAlignment = xlCenter
    ElseIf vAlign <> xlCenter And hAlign = xlCenter Then
        ' Currently horizontally centered only -> set vertically centered only
        rng.VerticalAlignment = xlCenter
        rng.HorizontalAlignment = xlLeft  ' Or xlGeneral if you prefer default left alignment
    Else
        ' Anything else (e.g. default alignment) -> set vertically centered only (start state)
        rng.VerticalAlignment = xlCenter
        rng.HorizontalAlignment = xlLeft
    End If
End Sub


Sub ClearSelectedCellFormats(control As IRibbonControl)
    Dim cellRange As Range
    On Error Resume Next
    Set cellRange = Intersect(Selection, Selection.Worksheet.UsedRange)
    On Error GoTo 0

    If TypeName(Selection) = "Range" Then
        Selection.ClearFormats
    Else
        MsgBox "Only cell selections are supported. Objects like shapes or charts will be ignored.", vbInformation, "Invalid Selection"
    End If
End Sub

Sub ToggleTotalFormatting(control As IRibbonControl)
    Dim cell As Range
    Dim topBorder As Border
    Dim bottomBorder As Border

    For Each cell In Selection
        Set topBorder = cell.Borders(xlEdgeTop)
        Set bottomBorder = cell.Borders(xlEdgeBottom)

        If Not cell.Font.Bold And Not topBorder.color = RGB(0, 0, 0) Then
            cell.Font.Bold = True
            With topBorder
                .LineStyle = xlContinuous
                .Weight = xlThin
                .color = RGB(0, 0, 0)
            End With
            'bottomBorder.LineStyle = xlNone
            
        ElseIf Not cell.Font.Bold And topBorder.color = RGB(0, 0, 0) Then
            cell.Font.Bold = True
            With topBorder
                .LineStyle = xlContinuous
                .Weight = xlThin
                .color = RGB(0, 0, 0)
            End With
            'bottomBorder.LineStyle = xlNone

        ElseIf topBorder.LineStyle = xlContinuous And topBorder.Weight = xlThin And _
               bottomBorder.LineStyle = xlNone Then
            With bottomBorder
                .LineStyle = xlContinuous
                .Weight = xlThin
                .color = RGB(0, 0, 0)
            End With

        ElseIf bottomBorder.LineStyle = xlContinuous And bottomBorder.Weight = xlThin Then
            With bottomBorder
                .LineStyle = xlContinuous
                .Weight = xlThick
                .color = RGB(0, 0, 0)
            End With

        ElseIf bottomBorder.LineStyle = xlContinuous And bottomBorder.Weight = xlThick Then
            With bottomBorder
                .LineStyle = xlDouble
                .Weight = xlThick
                .color = RGB(0, 0, 0)
            End With

        Else
            cell.Font.Bold = False
            topBorder.LineStyle = xlNone
            bottomBorder.LineStyle = xlNone
        End If
    Next cell
End Sub

Sub ConvertMergedCellsToCenterAcross(control As IRibbonControl)

Dim c As Range
Dim mergedRange As Range

'Loop through all cells in Used range
For Each c In ActiveSheet.UsedRange

    'If merged and single row
    If c.MergeCells = True And c.MergeArea.rows.count = 1 Then

        'Set variable for the merged range
        Set mergedRange = c.MergeArea

        'Unmerge the cell and apply Centre Across Selection
        mergedRange.UnMerge
        mergedRange.HorizontalAlignment = xlCenterAcrossSelection

    End If

Next

End Sub

Sub GetColorCodeFromCellFill(control As IRibbonControl)

'Create variables hold the color data
Dim fillColor As Long
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim Hex As String

'Get the fill color
fillColor = ActiveCell.Interior.color

'Convert fill color to RGB
r = (fillColor Mod 256)
g = (fillColor \ 256) Mod 256
b = (fillColor \ 65536) Mod 256

'Convert fill color to Hex
Hex = "#" & Application.WorksheetFunction.Dec2Hex(fillColor)

'Display fill color codes
MsgBox "Color codes for active cell" & vbNewLine & _
    "R:" & r & ", G:" & g & ", B:" & b & vbNewLine & _
    "Hex: " & Hex, Title:="Color Codes"

End Sub


Sub FormatMetres(control As IRibbonControl)
    Dim cellranges As Range
    Dim formats As Variant
    Dim i As Integer, foundIndex As Integer
    Dim currentFormat As String, nextFormat As String

    Set cellranges = Application.Selection
    currentFormat = Application.ActiveCell.numberFormat

    formats = Array( _
        "#,##0.0""m³"";[Red](#,##0.0""m³"");-", _
        "[Red]#,##0.0""m³"";(#,##0.0""m³"");-", _
        "+#,##0.0""m³"";[Red]-#,##0.0""m³"";-", _
        "[Red]+#,##0.0""m³"";-#,##0.0""m³"";-", _
        "#,##0.0""m²"";[Red](#,##0.0""m²"");-", _
        "[Red]#,##0.0""m²"";(#,##0.0""m²"");-", _
        "+#,##0.0""m²"";[Red]-#,##0.0""m²"";-", _
        "[Red]+#,##0.0""m²"";-#,##0.0""m²"";-" _
    )

    foundIndex = -1
    For i = LBound(formats) To UBound(formats)
        If currentFormat = formats(i) Then
            foundIndex = i
            Exit For
        End If
    Next i

    If foundIndex >= 0 Then
        nextFormat = formats((foundIndex + 1) Mod (UBound(formats) + 1))
    Else
        nextFormat = formats(0)
    End If

    On Error Resume Next
    cellranges.numberFormat = nextFormat
    On Error GoTo 0
End Sub

Sub FormatThousands(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "#,##0.0,;[Red](#,##0.0,);-" Then
cellranges.numberFormat = "[Red]#,##0.0,;(#,##0.0,);-"
ElseIf cellranges.numberFormat = "[Red]#,##0.0,;(#,##0.0,);-" Then
cellranges.numberFormat = "+#,##0.0,;[Red]-#,##0.0,;-"
ElseIf cellranges.numberFormat = "+#,##0.0,;[Red]-#,##0.0,;-" Then
cellranges.numberFormat = "[Red]+#,##0.0,;-#,##0.0,;-"

ElseIf cellranges.numberFormat = "[Red]+#,##0.0,;-#,##0.0,;-" Then
cellranges.numberFormat = "#,##0.0,""k"";[Red](#,##0.0,""k"");-"
ElseIf cellranges.numberFormat = "#,##0.0,""k"";[Red](#,##0.0,""k"");-" Then
cellranges.numberFormat = "[Red]#,##0.0,""k"";(#,##0.0,""k"");-"
ElseIf cellranges.numberFormat = "[Red]#,##0.0,""k"";(#,##0.0,""k"");-" Then
cellranges.numberFormat = "+#,##0.0,""k"";[Red]-#,##0.0,""k"";-"
ElseIf cellranges.numberFormat = "+#,##0.0,""k"";[Red]-#,##0.0,""k"";-" Then
cellranges.numberFormat = "[Red]+#,##0.0,""k"";-#,##0.0,""k"";-"
Else
cellranges.numberFormat = "#,##0.0,;[Red](#,##0.0,);-"
End If
End Sub

Sub FormatThousandsDollars(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "$#,##0.0,;[Red]($#,##0.0,);-" Then
cellranges.numberFormat = "[Red]$#,##0.0,;($#,##0.0,);-"
ElseIf cellranges.numberFormat = "[Red]$#,##0.0,;($#,##0.0,);-" Then
cellranges.numberFormat = "+$#,##0.0,;[Red]-$#,##0.0,;-"
ElseIf cellranges.numberFormat = "+$#,##0.0,;[Red]-$#,##0.0,;-" Then
cellranges.numberFormat = "[Red]+$#,##0.0,;-$#,##0.0,;-"

ElseIf cellranges.numberFormat = "[Red]+$#,##0.0,;-$#,##0.0,;-" Then
cellranges.numberFormat = "$#,##0.0,""k"";[Red]($#,##0.0,""k"");-"
ElseIf cellranges.numberFormat = "$#,##0.0,""k"";[Red]($#,##0.0,""k"");-" Then
cellranges.numberFormat = "[Red]$#,##0.0,""k"";($#,##0.0,""k"");-"
ElseIf cellranges.numberFormat = "[Red]$#,##0.0,""k"";($#,##0.0,""k"");-" Then
cellranges.numberFormat = "+$#,##0.0,""k"";[Red]-$#,##0.0,""k"";-"
ElseIf cellranges.numberFormat = "+$#,##0.0,""k"";[Red]-$#,##0.0,""k"";-" Then
cellranges.numberFormat = "[Red]+$#,##0.0,""k"";-$#,##0.0,""k"";-"
Else
cellranges.numberFormat = "$#,##0.0,;[Red]($#,##0.0,);-"
End If
End Sub

Sub FormatMillions(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "#,##0.0,,;[Red](#,##0.0,,);-" Then
cellranges.numberFormat = "[Red]#,##0.0,,;(#,##0.0,,);-"
ElseIf cellranges.numberFormat = "[Red]#,##0.0,,;(#,##0.0,,);-" Then
cellranges.numberFormat = "+#,##0.0,,;[Red]-#,##0.0,,;-"
ElseIf cellranges.numberFormat = "+#,##0.0,,;[Red]-#,##0.0,,;-" Then
cellranges.numberFormat = "[Red]+#,##0.0,,;-#,##0.0,,;-"

ElseIf cellranges.numberFormat = "[Red]+#,##0.0,,;-#,##0.0,,;-" Then
cellranges.numberFormat = "#,##0.0,,""m"";[Red](#,##0.0,,""m"");-"
ElseIf cellranges.numberFormat = "#,##0.0,,""m"";[Red](#,##0.0,,""m"");-" Then
cellranges.numberFormat = "[Red]#,##0.0,,""m"";(#,##0.0,,""m"");-"
ElseIf cellranges.numberFormat = "[Red]#,##0.0,,""m"";(#,##0.0,,""m"");-" Then
cellranges.numberFormat = "+#,##0.0,,""m"";[Red]-#,##0.0,,""m"";-"
ElseIf cellranges.numberFormat = "+#,##0.0,,""m"";[Red]-#,##0.0,,""m"";-" Then
cellranges.numberFormat = "[Red]+#,##0.0,,""m"";-#,##0.0,,""m"";-"
Else
cellranges.numberFormat = "#,##0.0,,;[Red](#,##0.0,,);-"
End If
End Sub

Sub FormatMillionsDollars(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "$#,##0.0,,;[Red]($#,##0.0,,);-" Then
cellranges.numberFormat = "[Red]$#,##0.0,,;($#,##0.0,,);-"
ElseIf cellranges.numberFormat = "[Red]$#,##0.0,,;($#,##0.0,,);-" Then
cellranges.numberFormat = "+$#,##0.0,,;[Red]-$#,##0.0,,;-"
ElseIf cellranges.numberFormat = "+$#,##0.0,,;[Red]-$#,##0.0,,;-" Then
cellranges.numberFormat = "[Red]+$#,##0.0,,;-$#,##0.0,,;-"

ElseIf cellranges.numberFormat = "[Red]+$#,##0.0,,;-$#,##0.0,,;-" Then
cellranges.numberFormat = "$#,##0.0,,""m"";[Red]($#,##0.0,,""m"");-"
ElseIf cellranges.numberFormat = "$#,##0.0,,""m"";[Red]($#,##0.0,,""m"");-" Then
cellranges.numberFormat = "[Red]$#,##0.0,,""m"";($#,##0.0,,""m"");-"
ElseIf cellranges.numberFormat = "[Red]$#,##0.0,,""m"";($#,##0.0,,""m"");-" Then
cellranges.numberFormat = "+$#,##0.0,,""m"";[Red]-$#,##0.0,,""m"";-"
ElseIf cellranges.numberFormat = "+$#,##0.0,,""m"";[Red]-$#,##0.0,,""m"";-" Then
cellranges.numberFormat = "[Red]+$#,##0.0,,""m"";-$#,##0.0,,""m"";-"
Else
cellranges.numberFormat = "$#,##0.0,,;[Red]($#,##0.0,,);-"
End If
End Sub

Sub FormatAccounting(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "#,##0;[Red](#,##0);-" Then
cellranges.numberFormat = "[Red]#,##0;(#,##0);-"
ElseIf cellranges.numberFormat = "[Red]#,##0;(#,##0);-" Then
cellranges.numberFormat = "+#,##0;[Red]-#,##0;-"
ElseIf cellranges.numberFormat = "+#,##0;[Red]-#,##0;-" Then
cellranges.numberFormat = "[Red]+#,##0;-#,##0;-"
Else
cellranges.numberFormat = "#,##0;[Red](#,##0);-"
End If
End Sub

Sub FormatAccountingDollars(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "$#,##0;[Red]($#,##0);-" Then
cellranges.numberFormat = "[Red]$#,##0;($#,##0);-"
ElseIf cellranges.numberFormat = "[Red]$#,##0;($#,##0);-" Then
cellranges.numberFormat = "+$#,##0;[Red]-$#,##0;-"
ElseIf cellranges.numberFormat = "+$#,##0;[Red]-$#,##0;-" Then
cellranges.numberFormat = "[Red]+$#,##0;-$#,##0;-"
Else
cellranges.numberFormat = "$#,##0;[Red]($#,##0);-"
End If
End Sub

Sub FormatBPS(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "+#,##0""bps"";[Red]-#,##0""bps""" Then
cellranges.numberFormat = "[Red]+#,##0""bps"";-#,##0""bps"""
ElseIf cellranges.numberFormat = "[Red]+#,##0""bps"";-#,##0""bps""" Then
cellranges.numberFormat = "#,##0""bps"";[Red](#,##0)""bps"""
ElseIf cellranges.numberFormat = "#,##0""bps"";[Red](#,##0)""bps""" Then
cellranges.numberFormat = "[Red]#,##0""bps"";(#,##0)""bps"""
Else
cellranges.numberFormat = "+#,##0""bps"";[Red]-#,##0""bps"""
End If
End Sub

Sub FormatPlusMinusPercent(control As IRibbonControl)
Set cellranges = Application.Selection
If cellranges.numberFormat = "+#,##0.0%;[Red]-#,##0.0%;"" - """ Then
cellranges.numberFormat = "[Red]+#,##0.0%;-#,##0.0%;"" - """
ElseIf cellranges.numberFormat = "[Red]+#,##0.0%;-#,##0.0%;"" - """ Then
cellranges.numberFormat = "#,##0.0%;[Red](#,##0.0%);"" - """
ElseIf cellranges.numberFormat = "#,##0.0%;[Red](#,##0.0%);"" - """ Then
cellranges.numberFormat = "[Red]#,##0.0%;(#,##0.0%);"" - """
Else
cellranges.numberFormat = "+#,##0.0%;[Red]-#,##0.0%;"" - """
End If
End Sub


' Ribbon callback keeps its signature
Public Sub Cell_Model_Formatting(control As IRibbonControl)
    Do_Cell_Model_Formatting
End Sub

' OnKey wrapper with no parameters
Public Sub Run_Cell_Model_Formatting()
    Do_Cell_Model_Formatting
End Sub

Public Sub Do_Cell_Model_Formatting()
    Dim cellranges As Range
    Dim cellselected As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    On Error Resume Next ' Start error bypassing

    ' Apply number format and border to each cell
    For Each cellselected In cellranges
        'cellselected.NumberFormat = "#,##0.0;(#,##0.0);"" - """
        cellselected.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin, _
            ColorIndex:=15
    Next cellselected

    ' Conditional formatting & styling
    If cellranges.Font.ColorIndex = 5 And cellranges.Interior.ColorIndex = 19 Then
        cellranges.Interior.ColorIndex = 0
        cellranges.Font.ColorIndex = 45
        ActiveWorkbook.Styles.Add Name:="Linked Cell", BasedOn:=ActiveCell
        cellranges.Style = "Linked Cell"
    
    ElseIf cellranges.Font.ColorIndex = 45 Then
        cellranges.Interior.ColorIndex = 15
        cellranges.Font.ColorIndex = 10
        For Each cellselected In cellranges
            'cellselected.NumberFormat = "#,##0.0;(#,##0.0);"" - """
            cellselected.BorderAround _
                LineStyle:=xlDouble, _
                Weight:=xlThin, _
                ColorIndex:=16
        Next cellselected
        cellranges.numberFormat = "[Color10][=0]""Ok"";[Red]""Error"""
        ActiveWorkbook.Styles.Add Name:="Check Cell", BasedOn:=ActiveCell
        cellranges.Style = "Check Cell"
    
    ElseIf cellranges.Interior.ColorIndex = 15 Then
        cellranges.Interior.ColorIndex = 0
        cellranges.Font.ColorIndex = 1
        ActiveWorkbook.Styles.Add Name:="Calculation", BasedOn:=ActiveCell
        cellranges.Style = "Calculation"
    
    Else
        cellranges.Interior.ColorIndex = 19
        cellranges.Font.ColorIndex = 5
        ActiveWorkbook.Styles.Add Name:="Input", BasedOn:=ActiveCell
        cellranges.Style = "Input"
    End If

    On Error GoTo 0 ' Turn off error bypassing

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_Input_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        'cellselected.NumberFormat = "#,##0.0;(#,##0.0);"" - """
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=15
    Next cellselected

    cellranges.Interior.ColorIndex = 19
    cellranges.Font.ColorIndex = 5
    ActiveWorkbook.Styles.Add Name:="Input", BasedOn:=ActiveCell
    cellranges.Style = "Input"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_Linked_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection
    
    For Each cellselected In cellranges
        'cellselected.NumberFormat = "#,##0.0;(#,##0.0);"" - """
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=15
    Next cellselected

    'cellranges.NumberFormat = "#,##0.0;(#,##0.0);"" - """
    cellranges.Interior.ColorIndex = 0
    cellranges.Font.ColorIndex = 45
    ActiveWorkbook.Styles.Add Name:="Linked Cell", BasedOn:=ActiveCell
    cellranges.Style = "Linked Cell"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_Check_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        'cellselected.NumberFormat = "#,##0.0;(#,##0.0);"" - """
        cellselected.BorderAround LineStyle:=xlDouble, Weight:=xlThin, ColorIndex:=16
    Next cellselected

    cellranges.Interior.ColorIndex = 15
    cellranges.Font.ColorIndex = 10
    cellranges.numberFormat = "[Color10][=0]""Ok"";[Red]""Error"""
    ActiveWorkbook.Styles.Add Name:="Check Cell", BasedOn:=ActiveCell
    cellranges.Style = "Check Cell"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_Calculation_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        'cellselected.NumberFormat = "#,##0.0;(#,##0.0);"" - """
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=15
    Next cellselected
    
    cellranges.Interior.ColorIndex = 0
    cellranges.Font.ColorIndex = 1
    ActiveWorkbook.Styles.Add Name:="Calculation", BasedOn:=ActiveCell
    cellranges.Style = "Calculation"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_NotUsed_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        With cellselected.Interior
            .pattern = xlLightUp
            .PatternColor = RGB(193, 193, 193)
            .color = RGB(255, 255, 255)
        End With
    Next cellselected

    cellranges.Font.ColorIndex = 1 ' Black font

    ActiveWorkbook.Styles.Add Name:="Not Used", BasedOn:=ActiveCell
    cellranges.Style = "Not Used"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_Header_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        With cellselected.Interior
            .color = RGB(1, 57, 118)
        End With
    Next cellselected

    cellranges.numberFormat = "#,##0.0;(#,##0.0);"" - """
    cellranges.Font.Name = "Arial"
    cellranges.Font.Bold = True
    cellranges.Font.ColorIndex = 2 ' White font

    ActiveWorkbook.Styles.Add Name:="Header", BasedOn:=ActiveCell
    cellranges.Style = "Header"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_SubHeader_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        With cellselected.Interior
            .color = RGB(232, 232, 232)
        End With
    Next cellselected

    cellranges.numberFormat = "#,##0.0;(#,##0.0);"" - """
    cellranges.Font.Name = "Arial"
    cellranges.Font.ColorIndex = 1 ' Black font

    ActiveWorkbook.Styles.Add Name:="Subheader", BasedOn:=ActiveCell
    cellranges.Style = "Subheader"

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_DotBorder_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        With cellselected.Borders
            .LineStyle = xlDot
            .color = RGB(150, 150, 150) ' Light grey
            .Weight = xlHairline
        End With
    Next cellselected

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_DashBorder_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        With cellselected.Borders
            .LineStyle = xlDot
            .color = RGB(150, 150, 150) ' Light grey
            .Weight = xlThin
        End With
    Next cellselected

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub RebuildDefaultStyles(control As IRibbonControl)

'The purpose of this macro is to remove all styles in the active
'workbook and rebuild the default styles.
'It rebuilds the default styles by merging them from a new workbook.

'Dimension variables.
   Dim MyBook As Workbook
   Dim tempBook As Workbook
   Dim CurStyle As Style
   
    Application.ScreenUpdating = False
    Application.EnableEvents = False

   'Set MyBook to the active workbook.
   Set MyBook = ActiveWorkbook
   On Error Resume Next
   'Delete all the styles in the workbook.
   For Each CurStyle In MyBook.Styles
      'If CurStyle.Name <> "Normal" Then CurStyle.Delete
      Select Case CurStyle.Name
         Case "20% - Accent1", "20% - Accent2", _
               "20% - Accent3", "20% - Accent4", "20% - Accent5", "20% - Accent6", _
               "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4", _
               "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", _
               "60% - Accent3", "60% - Accent4", "60% - Accent5", "60% - Accent6", _
               "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", _
               "Bad", "Calculation", "Check Cell", "Comma", "Comma [0]", "Currency", _
               "Currency [0]", "Explanatory Text", "Good", "Heading 1", "Heading 2", _
               "Heading 3", "Heading 4", "Input", "Linked Cell", "Neutral", "Normal", _
               "Note", "Output", "Percent", "Title", "Total", "Warning Text"
            'Do nothing, these are the default styles
         Case Else
            CurStyle.Delete
      End Select

   Next CurStyle

   'Open a new workbook.
   Set tempBook = Workbooks.Add

   'Disable alerts so you may merge changes to the Normal style
   'from the new workbook.
   Application.DisplayAlerts = False

   'Merge styles from the new workbook into the existing workbook.
   MyBook.Styles.Merge Workbook:=tempBook

   'Close the new workbook.
   tempBook.Close
   
    'Enable alerts.
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


Sub Apply_ShortScale_NumberFormat(control As IRibbonControl)
    Const NUMFMT As String = "[<1000]##,##0;[<1000000]#,###,""k"";#,###,,""m"""
    
    Dim appliedTo As Long
    Dim selType As String
    
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    selType = TypeName(Selection)
    
    ' 1) If cells are selected, apply directly
    If selType = "Range" Then
        Selection.numberFormat = NUMFMT
        appliedTo = appliedTo + Selection.Cells.count
    End If
    
    ' 2) If a chart or a chart part is selected, format axes + data labels
    If Not ActiveChart Is Nothing Then
        appliedTo = appliedTo + FormatChartNumbers(ActiveChart, NUMFMT)
    End If
    
    ' 3) If shapes are selected, iterate them (charts inside shapes, etc.)
    If selType = "DrawingObjects" Or selType = "Picture" Or selType = "TextBox" _
       Or selType = "GroupObject" Or selType = "ShapeRange" Then
       
        Dim sr As ShapeRange
        Dim Sh As Shape
        On Error Resume Next
        Set sr = Selection.ShapeRange
        On Error GoTo 0
        
        If Not sr Is Nothing Then
            For Each Sh In sr
                appliedTo = appliedTo + HandleShapeNumberFormat(Sh, NUMFMT)
            Next Sh
        End If
    End If
    
    ' 4) If a single object like Axis/DataLabels is selected, try direct set
    '    (safe no-op if property doesn’t exist)
    On Error Resume Next
    Selection.numberFormat = NUMFMT
    If Err.Number = 0 Then appliedTo = appliedTo + 1
    Err.Clear
    
    ' Axis specifically:
    Selection.TickLabels.numberFormat = NUMFMT
    If Err.Number = 0 Then appliedTo = appliedTo + 1
    Err.Clear
    
    ' DataLabels specifically:
    Selection.DataLabels.numberFormat = NUMFMT
    If Err.Number = 0 Then appliedTo = appliedTo + 1
    Err.Clear
    On Error GoTo 0
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If appliedTo > 0 Then
        'MsgBox "Applied number format to " & appliedTo & " target(s).", vbInformation
    Else
        MsgBox "No applicable targets found in the selection.", vbExclamation
    End If
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Oops—couldn't complete. Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

'==================== Helpers ====================

Private Function FormatChartNumbers(ch As Chart, ByVal NUMFMT As String) As Long
    Dim cnt As Long
    Dim ax As Axis
    Dim s As Series
    
    On Error Resume Next
    
    ' Axes tick labels
    For Each ax In ch.Axes
        ax.TickLabels.numberFormat = NUMFMT
        If Err.Number = 0 Then cnt = cnt + 1 Else Err.Clear
    Next ax
    
    ' Data labels for all series (works for normal & pivot charts)
    For Each s In ch.FullSeriesCollection
        If s.HasDataLabels Then
            s.DataLabels.numberFormat = NUMFMT
            If Err.Number = 0 Then cnt = cnt + 1 Else Err.Clear
        End If
    Next s
    
    ' Chart-level selection fallback (rarely needed, safe no-op if unsupported)
    ch.ChartArea.Format.TextFrame2.TextRange.ParagraphFormat
    ' (no number format at chart area level)
    
    On Error GoTo 0
    FormatChartNumbers = cnt
End Function

Private Function HandleShapeNumberFormat(Sh As Shape, ByVal NUMFMT As String) As Long
    Dim cnt As Long
    
    On Error Resume Next
    
    ' If the shape hosts a chart
    If Sh.Type = msoChart Then
        cnt = cnt + FormatChartNumbers(Sh.Chart, NUMFMT)
    End If
    
    ' If the shape is a Form Control with a linked cell, format that cell
    ' (Some shapes expose ControlFormat.LinkedCell)
    Dim linked As String
    linked = ""
    linked = Sh.ControlFormat.LinkedCell
    If Err.Number = 0 Then
        If Len(linked) > 0 Then
            Range(linked).numberFormat = NUMFMT
            cnt = cnt + 1
        End If
    Else
        Err.Clear
    End If
    
    ' For plain text boxes: number formats don’t apply to text directly.
    ' If you’re using a text box linked via formula, Excel doesn’t expose
    ' a NumberFormat on the shape—format the source cell instead.
    
    On Error GoTo 0
    HandleShapeNumberFormat = cnt
End Function


