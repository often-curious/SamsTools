Attribute VB_Name = "m_HideUnhideSheets"
Sub HideAllSelectedSheets(control As IRibbonControl)

'Create variable to hold worksheets
Dim ws As Worksheet

'Ignore error if trying to hide the last worksheet
On Error Resume Next

'Loop through each worksheet in the active workbook
For Each ws In ActiveWindow.selectedSheets

    'Hide each sheet
    ws.Visible = xlSheetVeryHidden

Next ws

'Allow errors to appear
On Error GoTo 0

End Sub

Sub UnhideAllWorksheets(control As IRibbonControl)

'Create variable to hold worksheets
Dim ws As Worksheet

'Loop through each worksheet in the active workbook
For Each ws In ActiveWorkbook.Worksheets

    'Unhide each sheet
    ws.Visible = xlSheetVisible

Next ws

End Sub

Sub HideAllOtherWorksheets(control As IRibbonControl)

Dim ws As Worksheet

'Loop through the worksheets
For Each ws In ActiveWorkbook.Worksheets

'Hide the sheet if it's not the active sheet
If ws.Name <> ActiveWorkbook.ActiveSheet.Name Then
    ws.Visible = xlSheetHidden
End If

Next ws

End Sub
