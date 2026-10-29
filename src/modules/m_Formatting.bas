Attribute VB_Name = "m_Formatting"
Option Explicit

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

    Clear_Selected_Cell_Formats

End Sub

Sub Clear_Selected_Cell_Formats()
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

    Toggle_Total_Formatting

End Sub

Sub Toggle_Total_Formatting()
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

Sub ToggleCentreAcrossSelection(control As IRibbonControl)

    Toggle_Centre_Across_Selection

End Sub

Sub Toggle_Centre_Across_Selection()

    With Selection
        If .HorizontalAlignment = xlCenterAcrossSelection Then
            .HorizontalAlignment = xlGeneral
        Else
            .HorizontalAlignment = xlCenterAcrossSelection
        End If
    End With

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
Dim fillColor As Long
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim Hex As String
Dim outputText As String
Dim clipboard As Object

'Get the fill color
fillColor = ActiveCell.Interior.color

'Convert fill color to RGB
r = (fillColor Mod 256)
g = (fillColor \ 256) Mod 256
b = (fillColor \ 65536) Mod 256

'Convert fill color to Hex
Hex = "#" & Application.WorksheetFunction.Dec2Hex(fillColor)

'Build the output text
outputText = "R:" & r & ", G:" & g & ", B:" & b & vbNewLine & "Hex: " & Hex

'Copy to clipboard using MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText outputText
clipboard.PutInClipboard

'Display fill color codes and confirm copy
MsgBox "Color codes for active cell (copied to clipboard)" & vbNewLine & _
    outputText, Title:="Color Codes"
End Sub

Private Sub ToggleNumberFormat(ByVal formats As Variant)

    On Error GoTo ErrorHandler

    Dim cellranges As Range
    Dim currentFormat As String
    Dim i As Long
    
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Please select cells before applying a format.", vbInformation
        Exit Sub
    End If
    
    Set cellranges = Application.Selection
    
    If IsNull(cellranges.numberFormat) Then
        currentFormat = ""
    Else
        currentFormat = cellranges.numberFormat
    End If
    
    For i = LBound(formats) To UBound(formats)
        
        If currentFormat = formats(i) Then
            
            If i = UBound(formats) Then
                cellranges.numberFormat = formats(LBound(formats))
            Else
                cellranges.numberFormat = formats(i + 1)
            End If
            
            Exit Sub
            
        End If
        
    Next i
    
    cellranges.numberFormat = formats(LBound(formats))

    Exit Sub

ErrorHandler:

    MsgBox "Unable to apply number format." & vbCrLf & _
           Err.Description, vbExclamation, "Sam's Tools"

End Sub

Sub FormatThousands(control As IRibbonControl)

    ToggleNumberFormat Array( _
        "#,##0.0,;[Red](#,##0.0,);-", _
        "[Red]#,##0.0,;(#,##0.0,);-", _
        "+#,##0.0,;[Red]-#,##0.0,;-", _
        "[Red]+#,##0.0,;-#,##0.0,;-", _
        "#,##0.0,""k"";[Red](#,##0.0,""k"");-", _
        "[Red]#,##0.0,""k"";(#,##0.0,""k"");-", _
        "+#,##0.0,""k"";[Red]-#,##0.0,""k"";-", _
        "[Red]+#,##0.0,""k"";-#,##0.0,""k"";-", _
        ChrW(9650) & "#,##0.0,""k"";[Red]" & ChrW(9660) & "#,##0.0,""k"";"" - """, _
        "[Red]" & ChrW(9650) & "#,##0.0,""k"";" & ChrW(9660) & "#,##0.0,""k"";"" - """ _
    )

End Sub

Sub FormatThousandsDollars(control As IRibbonControl)

    ToggleNumberFormat Array( _
        "$#,##0.0,;[Red]($#,##0.0,);-", _
        "[Red]$#,##0.0,;($#,##0.0,);-", _
        "+$#,##0.0,;[Red]-$#,##0.0,;-", _
        "[Red]+$#,##0.0,;-$#,##0.0,;-", _
        "$#,##0.0,""k"";[Red]($#,##0.0,""k"");-", _
        "[Red]$#,##0.0,""k"";($#,##0.0,""k"");-", _
        "+$#,##0.0,""k"";[Red]-$#,##0.0,""k"";-", _
        "[Red]+$#,##0.0,""k"";-$#,##0.0,""k"";-", _
        ChrW(9650) & "$#,##0.0,""k"";[Red]" & ChrW(9660) & "$#,##0.0,""k"";"" - """, _
        "[Red]" & ChrW(9650) & "$#,##0.0,""k"";" & ChrW(9660) & "$#,##0.0,""k"";"" - """ _
    )

End Sub

Sub FormatMillions(control As IRibbonControl)

    ToggleNumberFormat Array( _
        "#,##0.0,,;[Red](#,##0.0,,);-", _
        "[Red]#,##0.0,,;(#,##0.0,,);-", _
        "+#,##0.0,,;[Red]-#,##0.0,,;-", _
        "[Red]+#,##0.0,,;-#,##0.0,,;-", _
        "#,##0.0,,""m"";[Red](#,##0.0,,""m"");-", _
        "[Red]#,##0.0,,""m"";(#,##0.0,,""m"");-", _
        "+#,##0.0,,""m"";[Red]-#,##0.0,,""m"";-", _
        "[Red]+#,##0.0,,""m"";-#,##0.0,,""m"";-", _
        ChrW(9650) & "#,##0.0,,""m"";[Red]" & ChrW(9660) & "#,##0.0,,""m"";"" - """, _
        "[Red]" & ChrW(9650) & "#,##0.0,,""m"";" & ChrW(9660) & "#,##0.0,,""m"";"" - """ _
    )

End Sub

Sub FormatMillionsDollars(control As IRibbonControl)

    ToggleNumberFormat Array( _
        "$#,##0.0,,;[Red]($#,##0.0,,);-", _
        "[Red]$#,##0.0,,;($#,##0.0,,);-", _
        "+$#,##0.0,,;[Red]-$#,##0.0,,;-", _
        "[Red]+$#,##0.0,,;-$#,##0.0,,;-", _
        "$#,##0.0,,""m"";[Red]($#,##0.0,,""m"");-", _
        "[Red]$#,##0.0,,""m"";($#,##0.0,,""m"");-", _
        "+$#,##0.0,,""m"";[Red]-$#,##0.0,,""m"";-", _
        "[Red]+$#,##0.0,,""m"";-$#,##0.0,,""m"";-", _
        ChrW(9650) & "$#,##0.0,,""m"";[Red]" & ChrW(9660) & "$#,##0.0,,""m"";"" - """, _
        "[Red]" & ChrW(9650) & "$#,##0.0,,""m"";" & ChrW(9660) & "$#,##0.0,,""m"";"" - """ _
    )

End Sub

Sub FormatAccounting(control As IRibbonControl)

Format_Accounting

End Sub

Sub Format_Accounting()

    ToggleNumberFormat Array( _
        "#,##0;[Red](#,##0);-", _
        "[Red]#,##0;(#,##0);-", _
        "+#,##0;[Red]-#,##0;-", _
        "[Red]+#,##0;-#,##0;-", _
        ChrW(9650) & "#,##0;[Red]" & ChrW(9660) & "#,##0;-", _
        "[Red]" & ChrW(9650) & "#,##0;" & ChrW(9660) & "#,##0;-" _
    )

End Sub

Sub FormatAccountingDollars(control As IRibbonControl)

Format_Accounting_Dollars

End Sub

Sub Format_Accounting_Dollars()


    ToggleNumberFormat Array( _
        "$#,##0;[Red]($#,##0);-", _
        "[Red]$#,##0;($#,##0);-", _
        "+$#,##0;[Red]-$#,##0;-", _
        "[Red]+$#,##0;-$#,##0;-", _
        ChrW(9650) & "$#,##0;[Red]" & ChrW(9660) & "$#,##0;-", _
        "[Red]" & ChrW(9650) & "$#,##0;" & ChrW(9660) & "$#,##0;-" _
    )
    
End Sub


Sub FormatBPS(control As IRibbonControl)
    
    ToggleNumberFormat Array( _
        "+#,##0""bps"";[Red]-#,##0""bps""", _
        "[Red]+#,##0""bps"";-#,##0""bps""", _
        "#,##0""bps"";[Red](#,##0)""bps""", _
        "[Red]#,##0""bps"";(#,##0)""bps""", _
        ChrW(9650) & "#,##0""bps"";[Red]" & ChrW(9660) & "#,##0""bps"";"" - """, _
        "[Red]" & ChrW(9650) & "#,##0""bps"";" & ChrW(9660) & "#,##0""bps"";"" - """ _
    )
    
End Sub
    

Sub FormatPlusMinusPercent(control As IRibbonControl)

    ToggleNumberFormat Array( _
        "+#,##0.0%;[Red]-#,##0.0%;"" - """, _
        "[Red]+#,##0.0%;-#,##0.0%;"" - """, _
        "#,##0.0%;[Red](#,##0.0%);"" - """, _
        "[Red]#,##0.0%;(#,##0.0%);"" - """, _
        ChrW(9650) & "#,##0.0%;[Red]" & ChrW(9660) & "#,##0.0%;"" - """, _
        "[Red]" & ChrW(9650) & "#,##0.0%;" & ChrW(9660) & "#,##0.0%;"" - """ _
    )

End Sub


' Ribbon callback keeps its signature
Public Sub Cell_Model_Formatting(control As IRibbonControl)
    Cycle_Cell_Style
End Sub

' OnKey wrapper with no parameters
Public Sub Run_Cycle_Cell_Style()
    Cycle_Cell_Style
End Sub

Public Sub Cycle_Cell_Style()
    Dim rng As Range
    Dim c As Range
    Dim CurrentStyle As Long
    Dim NextStyle As Long

    On Error GoTo CleanExit
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set rng = Selection

    'Determine style from first cell in selection
    CurrentStyle = GetStyleIndex(rng.Cells(1))

    'Move to next style
    NextStyle = CurrentStyle + 1
    If NextStyle > 6 Then NextStyle = 1

    ApplyStyle rng, NextStyle

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Function GetStyleIndex(c As Range) As Long

    With c

        'Not Used
        If .Interior.pattern = xlLightUp _
           And .Interior.PatternColor = RGB(193, 193, 193) Then
            GetStyleIndex = 6
            Exit Function
        End If

        'Input
        If .Interior.color = RGB(255, 255, 204) _
           And .Font.color = RGB(0, 0, 255) Then
            GetStyleIndex = 1
            Exit Function
        End If

        'Calculation
        'Important: also checks for the grey border so blank/default cells are not treated as calculation cells
        If .Interior.color = RGB(255, 255, 255) _
           And .Font.color = RGB(0, 0, 0) _
           And HasModelBorder(c) Then
            GetStyleIndex = 2
            Exit Function
        End If

        'Linked Calculation
        If .Interior.color = RGB(221, 235, 247) _
           And .Font.color = RGB(0, 80, 215) Then
            GetStyleIndex = 3
            Exit Function
        End If

        'Unique Calculation
        If .Interior.color = RGB(252, 228, 214) _
           And .Font.color = RGB(204, 0, 0) Then
            GetStyleIndex = 4
            Exit Function
        End If

        'External Calculation
        If .Interior.color = RGB(255, 242, 204) _
           And .Font.color = RGB(204, 0, 0) Then
            GetStyleIndex = 5
            Exit Function
        End If

    End With

    'Unknown or blank/default cell
    'This means the next style applied will be Input
    GetStyleIndex = 0

End Function

Private Function HasModelBorder(c As Range) As Boolean

    On Error GoTo SafeExit

    With c.Borders(xlEdgeLeft)
        If .LineStyle = xlContinuous _
           And .Weight = xlThin _
           And .color = RGB(192, 192, 192) Then
            HasModelBorder = True
            Exit Function
        End If
    End With

SafeExit:

End Function

Private Sub ApplyStyle(rng As Range, StyleIndex As Long)

    Dim c As Range

    Select Case StyleIndex

        Case 1      'Input

            For Each c In rng
                c.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, _
                               color:=RGB(192, 192, 192)
            Next c

            rng.Interior.color = RGB(255, 255, 204)
            rng.Font.color = RGB(0, 0, 255)
            rng.Interior.pattern = xlSolid

        Case 2      'Calculation

            For Each c In rng
                c.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, _
                               color:=RGB(192, 192, 192)
            Next c

            rng.Interior.color = RGB(255, 255, 255)
            rng.Font.color = RGB(0, 0, 0)
            rng.Interior.pattern = xlSolid

        Case 3      'Linked Calculation

            For Each c In rng
                c.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, _
                               color:=RGB(192, 192, 192)
            Next c

            rng.Interior.color = RGB(221, 235, 247)
            rng.Font.color = RGB(0, 80, 215)
            rng.Interior.pattern = xlSolid

        Case 4      'Unique Calculation

            For Each c In rng
                c.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, _
                               color:=RGB(192, 192, 192)
            Next c

            rng.Interior.color = RGB(252, 228, 214)
            rng.Font.color = RGB(204, 0, 0)
            rng.Interior.pattern = xlSolid

        Case 5      'External Calculation

            For Each c In rng
                c.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, _
                               color:=RGB(192, 192, 192)
            Next c

            rng.Interior.color = RGB(255, 242, 204)
            rng.Font.color = RGB(204, 0, 0)
            rng.Interior.pattern = xlSolid

        Case 6      'Not Used

            For Each c In rng
                With c.Interior
                    .pattern = xlLightUp
                    .PatternColor = RGB(193, 193, 193)
                    .color = RGB(255, 255, 255)
                End With
            Next c

            rng.Font.color = RGB(0, 0, 0)

    End Select

End Sub


Sub Apply_Input_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection
    
    For Each cellselected In cellranges
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, color:=RGB(192, 192, 192)
    Next cellselected

    cellranges.Interior.color = RGB(255, 255, 204)
    cellranges.Font.color = RGB(0, 0, 255)


    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


Sub Apply_Linked_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection
    
    For Each cellselected In cellranges
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, color:=RGB(192, 192, 192)
    Next cellselected

    cellranges.Interior.color = RGB(255, 242, 204)
    cellranges.Font.color = RGB(204, 0, 0)


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
        cellselected.BorderAround LineStyle:=xlDouble, Weight:=xlThin, color:=RGB(192, 192, 192)
    Next cellselected

    cellranges.Interior.colorIndex = 15
    cellranges.Font.colorIndex = 10
    cellranges.numberFormat = "[Color10][=0]""Ok"";[Red]""Error"""

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_Calculation_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, color:=RGB(192, 192, 192)
    Next cellselected
    
    cellranges.Interior.color = RGB(255, 255, 255)
    cellranges.Font.color = RGB(0, 0, 0)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_Linked_Calculation_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, color:=RGB(192, 192, 192)
    Next cellselected
    
    cellranges.Interior.color = RGB(221, 235, 247)
    cellranges.Font.color = RGB(0, 80, 215)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_Unique_Calculation_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

    For Each cellselected In cellranges
        cellselected.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, color:=RGB(192, 192, 192)
    Next cellselected
    
    cellranges.Interior.color = RGB(252, 228, 214)
    cellranges.Font.color = RGB(204, 0, 0)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Sub Apply_NotUsed_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

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

    cellranges.Font.colorIndex = 1 ' Black font

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_Header_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

        With cellranges.Interior
            .color = RGB(1, 57, 118)
        End With

    cellranges.Font.Bold = True
    cellranges.Font.colorIndex = 2 ' White font

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_SubHeader_Style(control As IRibbonControl)
    Dim cellranges As Range

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set cellranges = Application.Selection

        With cellranges.Interior
            .color = RGB(232, 232, 232)
        End With

    cellranges.Font.Bold = True
    cellranges.Font.colorIndex = 1 ' Black font

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub Apply_DotBorder_Style(control As IRibbonControl)
    Dim cellranges As Range
    Dim cellselected As Range

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
    Dim cellselected As Range

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
    Const numFmt As String = "[<1000]##,##0;[<1000000]#,###,""k"";#,###,,""m"""
    
    Dim appliedTo As Long
    Dim selType As String
    
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    selType = TypeName(Selection)
    
    ' 1) If cells are selected, apply directly
    If selType = "Range" Then
        Selection.numberFormat = numFmt
        appliedTo = appliedTo + Selection.Cells.count
    End If
    
    ' 2) If a chart or a chart part is selected, format axes + data labels
    If Not ActiveChart Is Nothing Then
        appliedTo = appliedTo + FormatChartNumbers(ActiveChart, numFmt)
    End If
    
    ' 3) If shapes are selected, iterate them (charts inside shapes, etc.)
    If selType = "DrawingObjects" Or selType = "Picture" Or selType = "TextBox" _
       Or selType = "GroupObject" Or selType = "ShapeRange" Then
       
        Dim sr As ShapeRange
        Dim sh As Shape
        On Error Resume Next
        Set sr = Selection.ShapeRange
        On Error GoTo 0
        
        If Not sr Is Nothing Then
            For Each sh In sr
                appliedTo = appliedTo + HandleShapeNumberFormat(sh, numFmt)
            Next sh
        End If
    End If
    
    ' 4) If a single object like Axis/DataLabels is selected, try direct set
    '    (safe no-op if property doesn’t exist)
    On Error Resume Next
    Selection.numberFormat = numFmt
    If Err.Number = 0 Then appliedTo = appliedTo + 1
    Err.Clear
    
    ' Axis specifically:
    Selection.TickLabels.numberFormat = numFmt
    If Err.Number = 0 Then appliedTo = appliedTo + 1
    Err.Clear
    
    ' DataLabels specifically:
    Selection.DataLabels.numberFormat = numFmt
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

Private Function FormatChartNumbers(ch As Chart, ByVal numFmt As String) As Long
    Dim cnt As Long
    Dim ax As Axis
    Dim s As Series
    
    On Error Resume Next
    
    ' Axes tick labels
    For Each ax In ch.Axes
        ax.TickLabels.numberFormat = numFmt
        If Err.Number = 0 Then cnt = cnt + 1 Else Err.Clear
    Next ax
    
    ' Data labels for all series (works for normal & pivot charts)
    For Each s In ch.FullSeriesCollection
        If s.HasDataLabels Then
            s.DataLabels.numberFormat = numFmt
            If Err.Number = 0 Then cnt = cnt + 1 Else Err.Clear
        End If
    Next s
    
    On Error GoTo 0
    FormatChartNumbers = cnt
End Function

Private Function HandleShapeNumberFormat(sh As Shape, ByVal numFmt As String) As Long
    Dim cnt As Long
    
    On Error Resume Next
    
    ' If the shape hosts a chart
    If sh.Type = msoChart Then
        cnt = cnt + FormatChartNumbers(sh.Chart, numFmt)
    End If
    
    ' If the shape is a Form Control with a linked cell, format that cell
    ' (Some shapes expose ControlFormat.LinkedCell)
    Dim linked As String
    linked = ""
    linked = sh.ControlFormat.LinkedCell
    If Err.Number = 0 Then
        If Len(linked) > 0 Then
            Range(linked).numberFormat = numFmt
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

Sub ApplyOkabeItoTheme(control As IRibbonControl)

    Dim theme As Office.ThemeColorScheme
    
    Set theme = ActiveWorkbook.theme.ThemeColorScheme

    With theme
    
        ' Dark / Light
        .colors(msoThemeDark1).RGB = RGB(0, 0, 0)
        .colors(msoThemeLight1).RGB = RGB(255, 255, 255)
        
        .colors(msoThemeDark2).RGB = RGB(60, 60, 60)
        .colors(msoThemeLight2).RGB = RGB(242, 242, 242)

        ' Accent colours
        .colors(msoThemeAccent1).RGB = RGB(0, 114, 178)     ' Blue
        .colors(msoThemeAccent2).RGB = RGB(230, 159, 0)     ' Orange
        .colors(msoThemeAccent3).RGB = RGB(0, 158, 115)     ' Bluish Green
        .colors(msoThemeAccent4).RGB = RGB(204, 121, 167)   ' Reddish Purple
        .colors(msoThemeAccent5).RGB = RGB(86, 180, 233)    ' Sky Blue
        .colors(msoThemeAccent6).RGB = RGB(213, 94, 0)      ' Vermillion

        ' Hyperlinks
        .colors(msoThemeHyperlink).RGB = RGB(0, 114, 178)
        .colors(msoThemeFollowedHyperlink).RGB = RGB(204, 121, 167)

    End With

    MsgBox "Okabe-Ito colour theme applied.", vbInformation

End Sub


Sub ReplaceFormattingByColourAndFormat(control As IRibbonControl)

    Dim ws As Worksheet
    Dim RefCell As Range
    Dim SourceCell As Range
    Dim cell As Range
    Dim MatchRange As Range

    Dim RefColour As Long
    Dim RefNumberFormat As String
    Dim UserChoice As VbMsgBoxResult

    ' Ensure exactly one reference cell selected
    If Selection.Cells.count <> 1 Then
        MsgBox "Please select a single reference cell first.", vbExclamation
        Exit Sub
    End If

    Set RefCell = Selection
    Set ws = RefCell.Worksheet

    RefColour = RefCell.Interior.color
    RefNumberFormat = RefCell.numberFormat

    ' Find matching cells
    For Each cell In ws.UsedRange

        If cell.Interior.color = RefColour _
           And cell.numberFormat = RefNumberFormat Then

            If MatchRange Is Nothing Then
                Set MatchRange = cell
            Else
                Set MatchRange = Union(MatchRange, cell)
            End If

        End If

    Next cell

    If MatchRange Is Nothing Then
        MsgBox "No matching cells found.", vbInformation
        Exit Sub
    End If

    MsgBox MatchRange.Cells.count & _
           " matching cells found." & vbCrLf & vbCrLf & _
           "Select the source cell containing the desired formatting.", _
           vbInformation

    On Error Resume Next
    Set SourceCell = Application.InputBox( _
        Prompt:="Select the source cell", _
        Title:="Select Format Source", _
        Type:=8)
    On Error GoTo 0

    If SourceCell Is Nothing Then Exit Sub

    ' Ask user what to copy
    UserChoice = MsgBox( _
        "Yes = Copy ALL formatting" & vbCrLf & _
        "No = Copy VALUE formatting only (Number Format)" & vbCrLf & _
        "Cancel = Exit", _
        vbYesNoCancel + vbQuestion, _
        "Formatting Options")

    If UserChoice = vbCancel Then Exit Sub

    Application.ScreenUpdating = False

    If UserChoice = vbYes Then

        ' Copy all formatting
        SourceCell.Copy
        MatchRange.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

    ElseIf UserChoice = vbNo Then

        ' Copy Number Format only
        MatchRange.numberFormat = SourceCell.numberFormat

    End If

    Application.ScreenUpdating = True

    MsgBox "Updated " & MatchRange.Cells.count & " cells.", vbInformation

End Sub


Sub CreateWorkbookFormattingLegend(control As IRibbonControl)

    Dim ws As Worksheet
    Dim rngStart As Range
    Dim rngLegend As Range
    Dim sampleCol As Range
    Dim descCol As Range
    
    'Ask user where to place the legend
    On Error Resume Next
    Set rngStart = Application.InputBox( _
        Prompt:="Select the top-left cell where you would like to place the formatting legend.", _
        Title:="Place Formatting Legend", _
        Type:=8)
    On Error GoTo 0
    
    'Exit if user cancels
    If rngStart Is Nothing Then Exit Sub
    
    Set ws = rngStart.Worksheet
    
    Application.ScreenUpdating = False
    
    'Legend uses 2 columns x 9 rows
    Set rngLegend = rngStart.Resize(9, 2)
    
    'Clear existing content and formatting in the legend area
    rngLegend.Clear
    
    'Set column widths
    rngStart.EntireColumn.ColumnWidth = 12
    rngStart.Offset(0, 1).EntireColumn.ColumnWidth = 85
    
    'General font
    With rngLegend
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    
    '========================
    ' Header
    '========================
    With rngStart.Resize(1, 2)
        .Interior.color = RGB(0, 70, 140)
        .Font.color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    rngStart.value = "1.0"
    rngStart.numberFormat = "0.0"
    rngStart.Offset(0, 1).value = "Header"
    
    '========================
    ' Subheader
    '========================
    With rngStart.Offset(1, 0).Resize(1, 2)
        .Interior.color = RGB(217, 217, 217)
        .Font.Bold = True
    End With
    
    rngStart.Offset(1, 0).value = "1.1"
    rngStart.Offset(1, 0).numberFormat = "0.0"
    rngStart.Offset(1, 1).value = "Subheader"
    
    'Blank spacer row
    rngStart.Offset(2, 0).Resize(1, 2).RowHeight = 18
    
    'Set references for legend body
    Set sampleCol = rngStart.Offset(3, 0).Resize(6, 1)
    Set descCol = rngStart.Offset(3, 1).Resize(6, 1)
    
    'Descriptions
    descCol.Cells(1, 1).value = "Inputs or hard coded data"
    descCol.Cells(2, 1).value = "Calculations / references to cell in same sheet"
    descCol.Cells(3, 1).value = "Calculations / references to cells in other sheets"
    descCol.Cells(4, 1).value = "Unique calculations relative to others in same group / section / row / column"
    descCol.Cells(5, 1).value = "Calculations / references to cells in external links & files"
    descCol.Cells(6, 1).value = "Not used cell when inline with calculation cells in a range"
    
    'Sample values
    sampleCol.Cells(1, 1).value = 1000
    sampleCol.Cells(2, 1).value = 1000
    sampleCol.Cells(3, 1).value = 1000
    sampleCol.Cells(4, 1).value = 1000
    sampleCol.Cells(5, 1).value = 1000
    
    'Number formatting
    sampleCol.numberFormat = "#,##0"
    sampleCol.HorizontalAlignment = xlRight
    
    '========================
    ' Formatting key styles
    '========================
    
    'Inputs / hard coded data
    With sampleCol.Cells(1, 1)
        .Font.color = RGB(0, 0, 255)        'Blue font
        .Interior.color = RGB(255, 255, 204) 'Very light yellow
    End With
    
    'Calculations / references to cell in same sheet
    With sampleCol.Cells(2, 1)
        .Font.color = RGB(0, 0, 0)          'Black font
        .Interior.color = RGB(255, 255, 255)
    End With
    
    'Calculations / references to cells in other sheets
    With sampleCol.Cells(3, 1)
        .Font.color = RGB(0, 0, 255)        'Blue font
        .Interior.color = RGB(221, 235, 247) 'Light blue fill
    End With
    
    'Unique calculations relative to others in same group / section / row / column
    With sampleCol.Cells(4, 1)
        .Font.color = RGB(255, 0, 0)        'Red font
        .Interior.color = RGB(252, 228, 214) 'Light orange fill
    End With
    
    'Calculations / references to cells in external links & files
    With sampleCol.Cells(5, 1)
        .Font.color = RGB(255, 0, 0)        'Red font
        .Interior.color = RGB(255, 242, 204)
    End With
    
    'Not used cell when inline with calculation cells in a range
    With sampleCol.Cells(6, 1)
        .value = vbNullString
        .Interior.pattern = xlPatternLightUp
        .Interior.PatternColor = RGB(191, 191, 191)
        .Interior.color = RGB(255, 255, 255)
    End With
    
    'Description formatting
    With descCol
        .Font.color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    
    'Row heights
    rngLegend.RowHeight = 15
    rngStart.Resize(2, 2).RowHeight = 18
    
    'Borders around example cells only
    With sampleCol.Borders
        .LineStyle = xlContinuous
        .color = RGB(191, 191, 191)
        .Weight = xlThin
    End With
    
    'Remove borders from description cells
    descCol.Borders.LineStyle = xlNone
    
    'Tidy up
    rngLegend.Columns.AutoFit
    rngStart.EntireColumn.ColumnWidth = 12
    rngStart.Offset(0, 1).EntireColumn.ColumnWidth = 85
    
    Application.ScreenUpdating = True
    
    MsgBox "Formatting legend created successfully.", vbInformation, "Legend Created"

End Sub
