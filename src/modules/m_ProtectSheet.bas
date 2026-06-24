Attribute VB_Name = "m_ProtectSheet"
Private Const DEFAULT_PASSWORD_NAME As String = "_DefaultProtectPassword"
Private Const FALLBACK_PASSWORD As String = "#K1w1Bird"

Private Function GetDefaultPassword() As String
    On Error GoTo UseFallback

    GetDefaultPassword = Replace(ActiveWorkbook.Names(DEFAULT_PASSWORD_NAME).RefersTo, "=", "")
    GetDefaultPassword = Replace(GetDefaultPassword, """", "")

    If Len(GetDefaultPassword) = 0 Then GoTo UseFallback
    Exit Function

UseFallback:
    GetDefaultPassword = FALLBACK_PASSWORD
End Function

Private Sub SaveDefaultPassword(ByVal newPassword As String)
    On Error Resume Next
    ActiveWorkbook.Names(DEFAULT_PASSWORD_NAME).Delete
    On Error GoTo 0

    ActiveWorkbook.Names.Add _
        Name:=DEFAULT_PASSWORD_NAME, _
        RefersTo:="=""" & Replace(newPassword, """", """""") & """"

    ActiveWorkbook.Names(DEFAULT_PASSWORD_NAME).Visible = False
End Sub

Public Sub UpdateDefaultPassword(control As IRibbonControl)
    Dim oldPassword As String
    Dim newPassword As String
    Dim confirmPassword As String

    oldPassword = InputBox("Enter current default password:", "Current Password")
    If oldPassword = vbNullString Then Exit Sub

    If oldPassword <> GetDefaultPassword() Then
        MsgBox "Incorrect current password.", vbExclamation
        Exit Sub
    End If

    newPassword = InputBox("Enter new default password:", "New Password")
    If newPassword = vbNullString Then Exit Sub

    confirmPassword = InputBox("Confirm new default password:", "Confirm Password")
    If newPassword <> confirmPassword Then
        MsgBox "Passwords do not match.", vbExclamation
        Exit Sub
    End If

    SaveDefaultPassword newPassword

    MsgBox "Default password updated.", vbInformation
End Sub

Public Sub ShowSavedPassword(control As IRibbonControl)

    Dim savedPassword As String
    Dim response As VbMsgBoxResult

    savedPassword = GetDefaultPassword()

    MsgBox "Current saved password:" & vbCrLf & vbCrLf & savedPassword, _
           vbInformation, _
           "Saved Password"

End Sub
Sub ProtectWorkbookAndAllSheets(control As IRibbonControl)
    Dim ws As Worksheet
    Dim myPassword As String
    
    myPassword = GetDefaultPassword()
    
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
        ws.Unprotect GetDefaultPassword()
    Next
    Application.ScreenUpdating = True
End Sub


Sub LockOrUnlockCurrentSheet(control As IRibbonControl)
    Dim myPassword As String
    Dim sheet As Worksheet
    Dim selectedSheets As Sheets
    Dim originalSheet As Worksheet

    myPassword = GetDefaultPassword()
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
    Dim wb As Workbook

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






