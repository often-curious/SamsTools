Attribute VB_Name = "m_ProtectSheet"
Sub ProtectWorkbookAndAllSheets(control As IRibbonControl)
    Dim ws As Worksheet
    Dim myPassword As String
    
    myPassword = "#K1w1Bird"
    
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Protect
    For Each ws In Worksheets
        ws.Protect _
        Password:=(myPassword), _
        AllowFormattingRows:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingCells:=True, _
        AllowFiltering:=True
    Next
    Application.ScreenUpdating = True
End Sub
 
Sub UnProtectWorkbookAndAllSheets(control As IRibbonControl)
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    ActiveWorkbook.Unprotect
    For Each ws In Worksheets
        ws.Unprotect "#K1w1Bird"
    Next
    Application.ScreenUpdating = True
End Sub


Sub LockOrUnlockCurrentSheet(control As IRibbonControl)
    Dim myPassword As String
    Dim sheet As Worksheet
    Dim selectedSheets As Sheets
    Dim originalSheet As Worksheet

    myPassword = "#K1w1Bird"
    Set selectedSheets = ActiveWindow.selectedSheets
    Set originalSheet = ActiveSheet

    Application.ScreenUpdating = False

    ' Exit group mode by selecting the first sheet alone
    selectedSheets(1).Select

    ' Loop through each selected sheet
    For Each sheet In selectedSheets
        If TypeOf sheet Is Worksheet Then
            On Error Resume Next
            sheet.Select

            If sheet.ProtectContents = True Then
                sheet.Unprotect myPassword
            Else
                sheet.Protect _
                    Password:=myPassword, _
                    AllowFormattingRows:=True, _
                    AllowFormattingColumns:=True, _
                    AllowFormattingCells:=True, _
                    AllowFiltering:=True
            End If

            On Error GoTo 0
        End If
    Next sheet

    ' Reselect original active sheet
    originalSheet.Select
    Application.ScreenUpdating = True
End Sub





Sub UserProtectWorkbook(control As IRibbonControl)
    Dim myPassword As String
    Dim savePath As Variant
    Dim Wb As Workbook

    ' Prompt for password
    myPassword = InputBox("Enter password for the workbook:")

    If myPassword = "" Then
        MsgBox "Password is required. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Prompt for save location and file name
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ActiveWorkbook.Name, _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="Save Workbook As")

    ' User cancelled the dialog
    If savePath = False Then
        MsgBox "Save cancelled.", vbInformation
        Exit Sub
    End If

    ' Save the workbook with password protection
    ActiveWorkbook.SaveAs _
        fileName:=savePath, _
        FileFormat:=xlOpenXMLWorkbook, _
        Password:=myPassword

    MsgBox "New workbook saved with password.", vbInformation
End Sub


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



