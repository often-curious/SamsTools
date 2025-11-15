Attribute VB_Name = "m_Mail"
Sub Mail_ActiveSheet(control As IRibbonControl)
'Working in Excel 2000-2016

    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim sourceWb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim TempName As Variant

    On Error GoTo ErrorHandler
    
    'Ask for temp file name
    TempName = InputBox("What do you want to call this file?")
    'cancel pressed
    If StrPtr(TempName) = 0 Then Exit Sub
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set sourceWb = ActiveWorkbook

    'Copy the ActiveSheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            Select Case sourceWb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    '    'Change all cells in the worksheet to values if you want
    '    With Destwb.Sheets(1).UsedRange
    '        .Cells.Copy
    '        .Cells.PasteSpecial xlPasteValues
    '        .Cells(1).Select
    '    End With
    '    Application.CutCopyMode = False



    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    'TempFileName = Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
    TempFileName = TempName & " " & Format(Now, "dd-mmm-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = ""
            .cc = ""
            .BCC = ""
            .Subject = TempName
            .Body = "Hi there,"
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Display   'or use .Send
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred - please try again."
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
Exit Sub

End Sub

Sub Mail_ActiveSheetVALUES(control As IRibbonControl)
'Working in Excel 2000-2016

    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim sourceWb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim TempName As Variant

    On Error GoTo ErrorHandler
    
    'Ask for temp file name
    TempName = InputBox("What do you want to call this file?")
    'cancel pressed
    If StrPtr(TempName) = 0 Then Exit Sub
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set sourceWb = ActiveWorkbook

    'Copy the ActiveSheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            Select Case sourceWb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

        'Change all cells in the worksheet to values if you want
        With Destwb.Sheets(1).UsedRange
            .Cells.Copy
            .Cells.PasteSpecial xlPasteValues
            .Cells(1).Select
        End With
        Application.CutCopyMode = False



    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    'TempFileName = Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
    TempFileName = TempName & " " & Format(Now, "dd-mmm-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = ""
            .cc = ""
            .BCC = ""
            .Subject = TempName
            .Body = "Hi there,"
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Display   'or use .Send
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred - please try again."
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
Exit Sub

End Sub

Sub Mail_ActiveWorkbook(control As IRibbonControl)
'Working in Excel 2000-2016

    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim sourceWb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim TempName As Variant

    On Error GoTo ErrorHandler
    
    'Ask for temp file name
    TempName = InputBox("What do you want to call this file?")
    'cancel pressed
    If StrPtr(TempName) = 0 Then Exit Sub
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set sourceWb = ActiveWorkbook


    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    'TempFileName = Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
    TempFileName = TempName & " " & Format(Now, "dd-mmm-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With sourceWb
        .SaveAs TempFilePath & TempFileName & ".xlsx"
        On Error Resume Next
        With OutMail
            .To = ""
            .cc = ""
            .BCC = ""
            .Subject = TempName
            .Body = "Hi there,"
            .Attachments.Add sourceWb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Display   'or use .Send
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    'Delete the file you have send
    'Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred - please try again."
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
Exit Sub


End Sub

Sub Mail_Selected_Sheets_VALUES(control As IRibbonControl)
'Working in Excel 2000-2016
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim sourceWb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim Sh As Worksheet
    Dim TheActiveWindow As Window
    Dim TempWindow As Window
    Dim TempName As Variant

    On Error GoTo ErrorHandler

    'Ask for temp file name
    TempName = InputBox("What do you want to call this file?")
    'cancel pressed
    If StrPtr(TempName) = 0 Then Exit Sub

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set sourceWb = ActiveWorkbook

    'Copy the sheets to a new workbook
    'We add a temporary Window to avoid the Copy problem
    'if there is a List or Table in one of the sheets and
    'if the sheets are grouped
    With sourceWb
        Set TheActiveWindow = ActiveWindow
        Set TempWindow = .NewWindow
        '.Sheets(Array("Sheet1", "Sheet3")).Copy
        TheActiveWindow.selectedSheets.Copy
    End With

    'Close temporary Window
    TempWindow.Close

    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            Select Case sourceWb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

        'Change all cells in the worksheets to values if you want
        For Each Sh In Destwb.Worksheets
            Sh.Select
            With Sh.UsedRange
                .Cells.Copy
                .Cells.PasteSpecial xlPasteValues
                .Cells(1).Select
            End With
            Application.CutCopyMode = False
            Destwb.Worksheets(1).Select
        Next Sh

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    'TempFileName = "Part of " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")
    TempFileName = TempName & " " & Format(Now, "dd-mmm-yy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = ""
            .cc = ""
            .BCC = ""
            .Subject = TempName
            .Body = "Hi there"
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Display   'or use .Send
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred - please try again."
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
Exit Sub
    
End Sub

