Attribute VB_Name = "m_WorkbookSetup"


Sub TableOfContents_Create(control As IRibbonControl)
'PURPOSE: Add a Table of Contents worksheet to easily navigate to visible sheets

If ActiveWorkbook.path = "" Then
    MsgBox "Please save the workbook before generating the Table of Contents.", vbExclamation
    Exit Sub
End If

Dim sht As Worksheet
Dim Content_sht As Worksheet
Dim myArray() As Variant
Dim x As Long, y As Long
Dim shtName1 As String
Dim ContentName As String
Dim myAnswer As VbMsgBoxResult
Dim sortAlpha As Boolean

'Ask for sorting preference
myAnswer = MsgBox("Would you like to sort the Table of Contents alphabetically?" & vbCrLf & vbCrLf & _
                  "Click Yes for Current Tab Order, No for Alphabetical Order.", vbYesNo + vbQuestion, "Sort Order")

If myAnswer = vbCancel Then Exit Sub
sortAlpha = (myAnswer = vbNo)

'Get TOC name
ContentName = InputBox("Enter name for Contents Page", "Input Required")
If ContentName = "" Then GoTo ErrorHandler

'Optimize
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Check if Contents already exists
On Error Resume Next
Worksheets(ContentName).Activate
On Error GoTo 0

If ActiveSheet.Name = ContentName Then
    myAnswer = MsgBox("A worksheet named [" & ContentName & "] already exists. Replace it?", vbYesNo)
    If myAnswer <> vbYes Then GoTo ExitSub
    Worksheets(ContentName).Delete
End If

'Create new TOC sheet
Worksheets.Add before:=Worksheets(1)
Set Content_sht = ActiveSheet
Content_sht.Name = ContentName
Content_sht.Range("B2") = "Table of Contents"
Content_sht.Range("B2").Font.Bold = True

'Build array of VISIBLE sheet names (excluding the TOC itself)
x = 0
ReDim myArray(1 To Worksheets.count) 'over-allocate initially
For Each sht In ActiveWorkbook.Worksheets
    If sht.Name <> ContentName And sht.Visible = xlSheetVisible Then
        x = x + 1
        myArray(x) = sht.Name
    End If
Next sht
If x = 0 Then GoTo ErrorHandler 'no visible sheets to link
ReDim Preserve myArray(1 To x)

'Sort alphabetically if selected
If sortAlpha Then
    For x = LBound(myArray) To UBound(myArray)
        For y = x + 1 To UBound(myArray)
            If UCase(myArray(y)) < UCase(myArray(x)) Then
                shtName1 = myArray(x)
                myArray(x) = myArray(y)
                myArray(y) = shtName1
            End If
        Next y
    Next x
End If

'Add hyperlinks
For x = LBound(myArray) To UBound(myArray)
    Set sht = Worksheets(myArray(x))
    With Content_sht
        .Hyperlinks.Add .Cells(x + 3, 3), "", "'" & sht.Name & "'!A1", TextToDisplay:=sht.Name
        .Cells(x + 3, 2).value = x
    End With
Next x

'Format sheet
Content_sht.Activate
Content_sht.Columns(3).EntireColumn.AutoFit
Columns("A:A").ColumnWidth = 2.14
Columns("B:B").ColumnWidth = 3.86
Range("B2").Font.SIZE = 16
Range("B2:F2").Borders(xlEdgeBottom).Weight = xlThin

With Range("B4:B" & x + 2)
    .Borders(xlInsideHorizontal).color = RGB(255, 255, 255)
    .Borders(xlInsideHorizontal).Weight = xlMedium
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.color = RGB(255, 255, 255)
    .Interior.color = RGB(21, 96, 130)
End With

ActiveWindow.DisplayGridlines = False
ActiveWindow.Zoom = 90

ExitSub:
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
MsgBox "An error occurred - please try again."
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub Contents_Hyperlinks(control As IRibbonControl)
'PURPOSE: Add hyperlinked buttons back to Table of Contents worksheet tab

Dim sht As Worksheet
Dim shp As Shape
Dim ContentName As String
Dim ButtonID As String
Dim tocSheetExists As Boolean
Dim ws As Worksheet

'Inputs:
ContentName = InputBox("Enter name for Contents Page. This must exactly match the sheet name!", "Input Required")
If ContentName = "" Then Exit Sub 'User cancelled

ButtonID = "_ContentsButton"
tocSheetExists = False

'Check if TOC sheet exists
For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = ContentName Then
        tocSheetExists = True
        Exit For
    End If
Next ws

If Not tocSheetExists Then
    MsgBox "The Table of Contents sheet '" & ContentName & "' was not found. Please check the name and try again.", vbExclamation
    Exit Sub
End If

Application.ScreenUpdating = False

'Loop through each sheet to add/update hyperlink button
For Each sht In ActiveWorkbook.Worksheets
    If sht.Name <> ContentName Then
        
        'Delete old button if it exists
        For Each shp In sht.Shapes
            If Right(shp.Name, Len(ButtonID)) = ButtonID Then
                shp.Delete
                Exit For
            End If
        Next shp
        
        'Set first row height
        sht.rows(1).RowHeight = 30
        
        'Add home symbol to text
        Dim buttonText As String
        buttonText = "<  " & ContentName
        
        'Estimate width based on character count (approx. 6.5 pixels per char at size 10 font)
        Dim btnWidth As Single
        btnWidth = Len(buttonText) * 6.5 + 12 ' padding
        
        'Add new button with dynamic width
        Set shp = sht.Shapes.AddShape(msoShapeRoundedRectangle, 4, 4, btnWidth, 20)
        
        With shp
            .Fill.ForeColor.RGB = RGB(21, 96, 130)
            .line.Visible = msoFalse
            .TextFrame2.TextRange.Font.SIZE = 10
            .TextFrame2.TextRange.text = buttonText
            .TextFrame2.TextRange.Font.Name = "Lucida Sans Unicode"
            .TextFrame2.TextRange.Font.Bold = True
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Name = .Name & ButtonID
        
            ' Center text
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With

        
        'Assign hyperlink
        sht.Hyperlinks.Add shp, "", "'" & ContentName & "'!A1"
        
    End If
Next sht

Application.ScreenUpdating = True
End Sub

Sub RemoveContentsButtons(control As IRibbonControl)
    Dim sht As Worksheet
    Dim shp As Shape
    Dim ButtonID As String
    Dim deletedCount As Long

    ButtonID = "_ContentsButton"
    deletedCount = 0

    Application.ScreenUpdating = False

    For Each sht In ActiveWorkbook.Worksheets
        ' Check all shapes on the sheet
        For Each shp In sht.Shapes
            If Right(shp.Name, Len(ButtonID)) = ButtonID Then
                shp.Delete
                deletedCount = deletedCount + 1
                ' No Exit For — in case multiple exist
            End If
        Next shp
    Next sht

    Application.ScreenUpdating = True
    MsgBox deletedCount & " contents button(s) removed.", vbInformation, "Cleanup Complete"
End Sub


Sub Workbook_sheet_setup(control As IRibbonControl)
'PURPOSE:Set sheet to my preferences
    Dim tb As ListObject
    Dim boolTable As Boolean
'Prevent screen flicker
Application.ScreenUpdating = False

ActiveWindow.Zoom = 90

ActiveWorkbook.Windows(1).DisplayGridlines = False

With ActiveSheet
    .Cells.EntireColumn.AutoFit
    .Cells.EntireRow.AutoFit
    .Range("B2").Font.Bold = True
    .Range("B2").Font.SIZE = 16
End With

Columns("A:A").ColumnWidth = 2.14

ActiveSheet.Range("B2").Select


Application.ScreenUpdating = True
End Sub



Sub ListAllLinks(control As IRibbonControl)

    Dim Wks             As Worksheet
    Dim rFormulas       As Range
    Dim rCell           As Range
    Dim aLinks()        As String
    Dim cnt             As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    
    cnt = 0
    For Each Wks In Worksheets
        On Error Resume Next
        Set rFormulas = Wks.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not rFormulas Is Nothing Then
            For Each rCell In rFormulas
                If InStr(1, rCell.formula, "[") > 0 Then
                    cnt = cnt + 1
                    ReDim Preserve aLinks(1 To 2, 1 To cnt)
                    aLinks(1, cnt) = rCell.Address(, , , True)
                    aLinks(2, cnt) = "'" & rCell.formula
                End If
            Next rCell
        End If
    Next Wks
    
    If cnt > 0 Then
        Worksheets.Add before:=Worksheets(1)
        Range("A1").Resize(, 2).value = Array("Location", "Reference")
        Range("A2").Resize(UBound(aLinks, 2), UBound(aLinks, 1)).value = Application.Transpose(aLinks)
        Columns("A:B").AutoFit
    Else
        MsgBox "No links were found within the active workbook.", vbInformation
    End If
    
End Sub

Sub FileBackUp(control As IRibbonControl)
    Dim FPath As String
    FPath = Application.ActiveWorkbook.path

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrorHandler
    
    If Not ActiveWorkbook.Saved Then
        MsgBox "You must save this file first before running!"
        Exit Sub
    End If
    
    MsgBox "Please note that all the backup file will be saved in the same folder as this file."
    
    With ActiveWorkbook
    .SaveCopyAs fileName:=FPath & _
    "\" & Format(Date, "yy-mm-dd") & " " & _
    ActiveWorkbook.Name
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
        MsgBox "An error occurred - please try again."
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    Exit Sub
    
End Sub


Sub SaveWorksheetsAsPDF(control As IRibbonControl)
    Dim ws As Worksheet
    Dim folderName As String
    Dim FldrPicker As FileDialog
    Dim Wb As Workbook
    Dim selectedSheets As Sheets
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo ErrorHandler

    Set Wb = ActiveWorkbook
    
    ' Check if workbook is saved
    If Not Wb.Saved Then
        MsgBox "Please save the workbook before running this macro.", vbExclamation
        Exit Sub
    End If

    ' Ensure user has selected at least one sheet
    If ActiveWindow.selectedSheets.count = 0 Then
        MsgBox "Please select at least one worksheet to export.", vbExclamation
        Exit Sub
    End If

    ' Prompt user to select folder
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FldrPicker
        .Title = "Select where to save PDFs."
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub ' User cancelled
        folderName = .SelectedItems(1) & "\"
    End With

    ' Export selected sheets as PDFs
    Set selectedSheets = ActiveWindow.selectedSheets
    For Each ws In selectedSheets
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            fileName:=folderName & ws.Name & ".pdf"
    Next ws

    MsgBox "Selected sheets saved as PDFs in: " & folderName, vbInformation

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Sub Issue_log(control As IRibbonControl)

    ' Creates the Issue log or adds to existing issue log

    If ActiveWorkbook.path = "" Then
        MsgBox "Please save the workbook before generating the issue log.", vbExclamation
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Or Selection.Cells.count = 0 Then
        MsgBox "Please select a cell or range before running the issue log.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wsTest As Worksheet
    Dim issueWS As Worksheet
    Dim targetCell As Range
    Dim lastRow As Long
    Const strSheetName As String = "Issues"

    ' Check if Issues sheet exists
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(strSheetName)
    On Error GoTo 0

    ' Get the top-left cell of the selection or merged area
    If Selection.MergeCells Then
        Set targetCell = Selection.MergeArea.Cells(1, 1)
    Else
        Set targetCell = Selection.Cells(1, 1)
    End If

    ' Create sheet if it doesn't exist
    If wsTest Is Nothing Then
        Set issueWS = Worksheets.Add(before:=Worksheets(1))
        issueWS.Name = strSheetName
        With issueWS
            .Range("A1:H1").value = Array("Sheet Name", "Range Ref", "Date/Time", "Colour", "Formula", "Value", "Description", "Action")
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("A1:H1").Borders(xlEdgeBottom).Weight = xlMedium
            .Columns("A:F").ColumnWidth = 15
            .Columns("G:G").ColumnWidth = 50
            .Columns("H:H").ColumnWidth = 30
            .Range(.Cells(1, 9), .Cells(1, Columns.count)).EntireColumn.Hidden = True
            
            .Activate
            .Range("A2").Select
            ActiveWindow.FreezePanes = True
        End With
    Else
        Set issueWS = wsTest
    End If



    ' Add new issue row
    lastRow = issueWS.Cells(issueWS.rows.count, "A").End(xlUp).row + 1

    With issueWS
        ' Sheet name
        .Cells(lastRow, "A").value = targetCell.Worksheet.Name

        ' Hyperlinked reference to top-left cell
        Dim addr As String, linkFormula As String
        Dim labelAddr As String
        addr = targetCell.Address(False, False) ' Link target
        labelAddr = Selection.Address(False, False) ' Display text (full range)
        linkFormula = "=HYPERLINK(""#'" & targetCell.Worksheet.Name & "'!" & addr & """,""" & labelAddr & """)"
        .Cells(lastRow, "B").formula = linkFormula

        ' Timestamp
        .Cells(lastRow, "C").value = Now

        ' Cell formatting
        targetCell.Copy
        .Cells(lastRow, "D").PasteSpecial Paste:=xlPasteFormats

        ' Formula (as text)
        If targetCell.HasFormula Then
            .Cells(lastRow, "E").formula = "'" & targetCell.formula
        Else
            .Cells(lastRow, "E").value = ""
        End If

        ' Value
        .Cells(lastRow, "F").value = targetCell.value
    End With

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub

Sub CreateCalendarv2(control As IRibbonControl)
    Dim lMonth As Long
    Dim strMonth As String
    Dim rStart As Range
    Dim strAddress As String
    Dim rCell As Range
    Dim lDays As Long
    Dim dDate As Date
    Dim dayofweek As Long
    Dim dayofweekref As Range

    Application.ScreenUpdating = False
    
    
    
    Dim s As String: s = "Calendar"
    If Not Evaluate("isref('" & s & "'!A1)") Then

    
    'Add new sheet and format
    Worksheets.Add.Name = "Calendar"
    ActiveWindow.DisplayGridlines = False
        With Cells
            .ColumnWidth = 3.5
            .Font.SIZE = 8
            .HorizontalAlignment = xlCenter
        End With

    'Create Master header day of week name
    For lMonth = 1 To 15
        strAddress = Choose(lMonth, "A1:G1", "H1:N1", "O1:U1", _
                                    "A3:G3", "H3:N3", "O3:U3", _
                                    "A11:G11", "H11:N11", "O11:U11", _
                                    "A19:G19", "H19:N19", "O19:U19", _
                                    "A27:G27", "H27:N27", "O27:U27")
        lDays = 1
        Range(strAddress).BorderAround LineStyle:=xlContinuous
        'Add day of week to month range and format
        For Each rCell In Range(strAddress)
            lDays = lDays + 1
            dDate = lDays
                If Month(dDate) >= 0 Then ' It's a valid date
                    With rCell
                        .value = dDate
                        .numberFormat = "ddd"
                    End With
                End If
        Next rCell
    Next lMonth
    

    'Create the Month headings
    For lMonth = 1 To 4
            Select Case lMonth
                    Case 1
                        strMonth = "January"
                        Set rStart = Range("A2")
                    Case 2
                        strMonth = "April"
                        Set rStart = Range("A10")
                    Case 3
                        strMonth = "July"
                        Set rStart = Range("A18")
                    Case 4
                        strMonth = "October"
                        Set rStart = Range("A26")
            End Select
          
            'Merge, AutoFill and align months
            With rStart
                .value = strMonth
                .HorizontalAlignment = xlCenter
                .Interior.ColorIndex = 15
                .Font.Bold = True
                    With .Range("A1:G1")
                        .Merge
                        .BorderAround LineStyle:=xlContinuous
                    End With
                .Range("A1:G1").AutoFill Destination:=.Range("A1:U1")
            End With
    Next lMonth
    

     'Pass ranges for months
     For lMonth = 1 To 12
        strAddress = Choose(lMonth, "A4:G9", "H4:N9", "O4:U9", _
                            "A12:G17", "H12:N17", "O12:U17", _
                            "A20:G25", "H20:N25", "O20:U25", _
                            "A28:G34", "H28:N34", "O28:U34")
        lDays = 1
        Range(strAddress).BorderAround LineStyle:=xlContinuous
        'Add dates to month range and format
        For Each rCell In Range(strAddress)
            dDate = DateSerial(Year(Date), lMonth, lDays)
            dayofweek = Range(rCell.Address).Column
                If Month(dDate) = lMonth And Weekday(dDate) = Weekday(Cells(1, dayofweek)) Then ' It's a valid date
                    With rCell
                        .value = dDate
                        .numberFormat = "dd"
                    End With
                    lDays = lDays + 1
                End If
                
        Next rCell
    Next lMonth
        
    'add con formatting
     With Range("A1:U35")
           .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TODAY()"
           .FormatConditions(1).Font.ColorIndex = 2
           .FormatConditions(1).Interior.ColorIndex = 1
    End With
        
    Range("A1").EntireRow.Insert
        
    With Range("A1:U1")
            .HorizontalAlignment = xlCenter
            .Interior.ColorIndex = 46
            .Font.Bold = True
            .numberFormat = "yyyy"
                With .Range("A1:U1")
                    .Merge
                    .BorderAround LineStyle:=xlContinuous
                    .numberFormat = "yyyy"
                End With
        End With
        
        Range("A1").value = Range("A5")
            
    Range("A2").EntireRow.Delete
    Range("A1").EntireRow.Insert
    Range("A1").EntireColumn.Insert
    
    Else
    MsgBox "Sheet called 'Calendar' already exists."
    
    End If
    
    Application.ScreenUpdating = True
    
    
End Sub

Sub CreateCalendar(control As IRibbonControl)
    'If ActiveWorkbook.Path = "" Then
    '    MsgBox "Please save the workbook before generating the calendar.", vbExclamation
    '    Exit Sub
    'End If
    
    
    Dim lMonth As Long
    Dim strMonth As Long
    Dim rStart As Range
    Dim strAddress As String
    Dim rCell As Range
    Dim lDays As Long
    Dim dDate As Date
    Dim dayofweek As Long
    Dim dayofweekref As Range

    Application.ScreenUpdating = False
    
    
    Dim s As String: s = "Calendar"
    If Not Evaluate("isref('" & s & "'!A1)") Then

    
    'Add new sheet and format
    Worksheets.Add.Name = "Calendar"
    ActiveWindow.DisplayGridlines = False
        With Cells
            .ColumnWidth = 3.5
            .Font.SIZE = 8
            .HorizontalAlignment = xlCenter
        End With

    'Create Master header day of week name
    For lMonth = 1 To 15
        strAddress = Choose(lMonth, "A1:G1", "H1:N1", "O1:U1", _
                                    "A3:G3", "H3:N3", "O3:U3", _
                                    "A11:G11", "H11:N11", "O11:U11", _
                                    "A19:G19", "H19:N19", "O19:U19", _
                                    "A27:G27", "H27:N27", "O27:U27")
        lDays = 1
        Range(strAddress).BorderAround LineStyle:=xlContinuous
        'Add day of week to month range and format
        For Each rCell In Range(strAddress)
            lDays = lDays + 1
            dDate = lDays
                If Month(dDate) >= 0 Then ' It's a valid date
                    With rCell
                        .value = dDate
                        .numberFormat = "ddd"
                    End With
                End If
        Next rCell
    Next lMonth
    

    'Create the Month headings
    For lMonth = 1 To 12
            Select Case lMonth
                    Case 1
                        strMonth = 1
                        Set rStart = Range("A2")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """January"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("A4").Formula2 = "=IF(MONTH(DATE($A$1,$A$2,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$2,1),2))=$A$2,DATE($A$1,$A$2,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$2,1),2),"""")"
                    Case 2
                        strMonth = 2
                        Set rStart = Range("H2")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """Feburary"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("H4").Formula2 = "=IF(MONTH(DATE($A$1,$H$2,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$2,1),2))=$H$2,DATE($A$1,$H$2,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$2,1),2),"""")"
                    Case 3
                        strMonth = 3
                        Set rStart = Range("O2")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """March"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("O4").Formula2 = "=IF(MONTH(DATE($A$1,$O$2,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$2,1),2))=$O$2,DATE($A$1,$O$2,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$2,1),2),"""")"
                    Case 4
                        strMonth = 4
                        Set rStart = Range("A10")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """April"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("A12").Formula2 = "=IF(MONTH(DATE($A$1,$A$10,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$10,1),2))=$A$10,DATE($A$1,$A$10,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$10,1),2),"""")"
                    Case 5
                        strMonth = 5
                        Set rStart = Range("H10")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """May"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("H12").Formula2 = "=IF(MONTH(DATE($A$1,$H$10,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$10,1),2))=$H$10,DATE($A$1,$H$10,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$10,1),2),"""")"
                    Case 6
                        strMonth = 6
                        Set rStart = Range("O10")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """June"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("O12").Formula2 = "=IF(MONTH(DATE($A$1,$O$10,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$10,1),2))=$O$10,DATE($A$1,$O$10,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$10,1),2),"""")"
                    Case 7
                        strMonth = 7
                        Set rStart = Range("A18")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """July"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("A20").Formula2 = "=IF(MONTH(DATE($A$1,$A$18,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$18,1),2))=$A$18,DATE($A$1,$A$18,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$18,1),2),"""")"
                    Case 8
                        strMonth = 8
                        Set rStart = Range("H18")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """August"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("H20").Formula2 = "=IF(MONTH(DATE($A$1,$H$18,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$18,1),2))=$H$18,DATE($A$1,$H$18,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$18,1),2),"""")"
                    Case 9
                        strMonth = 9
                        Set rStart = Range("O18")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """September"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("O20").Formula2 = "=IF(MONTH(DATE($A$1,$O$18,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$18,1),2))=$O$18,DATE($A$1,$O$18,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$18,1),2),"""")"
                    Case 10
                        strMonth = 10
                        Set rStart = Range("A26")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """October"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("A28").Formula2 = "=IF(MONTH(DATE($A$1,$A$26,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$26,1),2))=$A$26,DATE($A$1,$A$26,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$A$26,1),2),"""")"
                    Case 11
                        strMonth = 11
                        Set rStart = Range("H26")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """November"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("H28").Formula2 = "=IF(MONTH(DATE($A$1,$H$26,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$26,1),2))=$H$26,DATE($A$1,$H$26,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$H$26,1),2),"""")"
                    Case 12
                        strMonth = 12
                        Set rStart = Range("O26")
                        'Merge, format and align month
                        With rStart
                            .value = strMonth
                            .numberFormat = """December"""
                            .HorizontalAlignment = xlCenter
                                With .Range("A1:G1")
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .BorderAround LineStyle:=xlContinuous
                                    .Interior.ColorIndex = 15
                                    .Font.Bold = True
                                End With
                        End With
                        Range("O28").Formula2 = "=IF(MONTH(DATE($A$1,$O$26,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$26,1),2))=$O$26,DATE($A$1,$O$26,1)+SEQUENCE(6,7)-WEEKDAY(DATE($A$1,$O$26,1),2),"""")"
            End Select
          
    Next lMonth
    

     'Pass ranges for months
     For lMonth = 1 To 12
        strAddress = Choose(lMonth, "A4:G9", "H4:N9", "O4:U9", _
                            "A12:G17", "H12:N17", "O12:U17", _
                            "A20:G25", "H20:N25", "O20:U25", _
                            "A28:G34", "H28:N34", "O28:U34")
        lDays = 1
        Range(strAddress).BorderAround LineStyle:=xlContinuous
        Range(strAddress).numberFormat = "dd"
    Next lMonth
        
    'add con formatting
     With Range("A1:U35")
           .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TODAY()"
           .FormatConditions(1).Font.ColorIndex = 2
           .FormatConditions(1).Interior.ColorIndex = 1
    End With
        
    'Range("A1").EntireRow.Insert
        
    With Range("A1:U1")
            .ClearContents
            .HorizontalAlignment = xlCenter
            .Interior.ColorIndex = 46
            .Font.Bold = True
            .numberFormat = "0000"
                With .Range("A1:U1")
                    .Merge
                    .BorderAround LineStyle:=xlContinuous
                    .numberFormat = "0000"
                End With
        End With
        
        Range("A1").value = Year(CLng(Date))
            
    'Range("A2").EntireRow.Delete
    Range("A1").EntireRow.Insert
    Range("A1").EntireColumn.Insert
    
    Range("W2").value = "<<< You can edit the year here to change the calendar."
    Range("W2").HorizontalAlignment = xlLeft
    
    Else
    MsgBox "Sheet called 'Calendar' already exists."
    
    End If
    
    Application.ScreenUpdating = True
    
    
End Sub

Sub CreateAgenda(control As IRibbonControl)
    Dim ws As Worksheet
    Dim agendaName As String: agendaName = "Agenda"
    Dim Wb As Workbook
    Dim row As Long
    Dim inputRange As Range
    Dim cfRange As Range
    
    Set Wb = ActiveWorkbook
    
    Application.ScreenUpdating = False
    
    ' Delete existing Agenda sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    Wb.Sheets(agendaName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Add new Agenda sheet at the end
    Set ws = Wb.Sheets.Add(After:=Wb.Sheets(Wb.Sheets.count))
    ws.Name = agendaName
    
    ' Make column A narrow as margin
    ws.Columns("A").ColumnWidth = 2
    
    ' Agenda header in row 2
    ws.Range("B2").value = "Agenda"
    With ws.Range("B2")
        .Font.Bold = True
        .Font.SIZE = 16
    End With
    
    ' Help text in row 3
    ws.Range("B3").value = "Fill in the left table (B7:E...) with agenda items. The '#' column in the right table lists order numbers 1 to 15, so you can change these around as needed whilst building the agenda."
    ws.Range("B3").Font.Italic = True
    
    ' Meeting Start Time label and input cell (B4 and C4)
    ws.Range("B4").value = "Meeting Start Time"
    ws.Range("C4").numberFormat = "h:mm AM/PM"
    ws.Range("C4").value = Time ' default current time
    
    ' Format meeting start time input cell
    With ws.Range("C4")
        .Interior.color = RGB(255, 255, 204) ' light yellow fill
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.color = RGB(191, 191, 191) ' light grey border
    End With
    
    ' Left table headers and user input area start at row 6 now
    ws.Range("B6:E6").value = Array("Topic", "Attendees Required", "Minutes", "Order #")
    With ws.Range("B6:E6")
        .Font.Bold = True
        .Interior.color = RGB(0, 0, 0)    ' black fill
        .Font.color = RGB(255, 255, 255)  ' white font
    End With
    
    ' Right table headers starting at row 6 too - rename Start Time and End Time to Start and End
    ws.Range("G6:N6").value = Array("#", "Start", "End", "Topic", "Attendees Required", "Topic Length (min)", "Gap (min)", "Highlight Breaks")
    With ws.Range("G6:N6")
        .Font.Bold = True
        .Interior.color = RGB(0, 0, 0)    ' black fill
        .Font.color = RGB(255, 255, 255)  ' white font
    End With
    
    ' Set column widths
    ws.Columns("B:E").ColumnWidth = 18
    ws.Columns("G:G").ColumnWidth = 4    ' Smaller for "#"
    ws.Columns("H:I").ColumnWidth = 12   ' Shrunk Start and End
    ws.Columns("J:K").ColumnWidth = 18
    ws.Columns("L:N").ColumnWidth = 12
    
    ' Format time columns in right table
    ws.Columns("H:H").numberFormat = "h:mm AM/PM"
    ws.Columns("I:I").numberFormat = "h:mm AM/PM"
    ws.Columns("L:L").numberFormat = "0" ' no decimals for Topic Length
    ws.Columns("M:M").numberFormat = "0" ' Gap as integer
    
    ' Clear any existing content in tables (optional)
    ws.Range("B7:E100").ClearContents
    ws.Range("G7:N100").ClearContents
    
    ' Pre-fill Order Number (#) column with numbers 1 to 15 in output table (rows 7 to 21)
    For row = 7 To 21
        ws.Cells(row, "G").value = row - 6
    Next row
    
    ' Insert formulas and formatting
    For row = 7 To 21
        ' Start - blank if no topic
        If row = 7 Then
            ws.Cells(row, "H").formula = "=IFERROR(IF(INDEX($B$7:$B$100,MATCH(G" & row & ",$E$7:$E$100,0))="""","""",$C$4),"""")"
        Else
            ' Start time = Previous End + Gap from previous row, blank if no topic
            ws.Cells(row, "H").formula = "=IFERROR(IF(INDEX($B$7:$B$100,MATCH(G" & row & ",$E$7:$E$100,0))="""","""",I" & (row - 1) & "+IFERROR(M" & (row - 1) & "/(24*60),0)),"""")"
        End If
        
        ' End - blank if no topic
        ws.Cells(row, "I").formula = "=IFERROR(IF(INDEX($B$7:$B$100,MATCH(G" & row & ",$E$7:$E$100,0))="""","""",H" & row & "+TIME(0,INDEX($D$7:$D$100,MATCH(G" & row & ",$E$7:$E$100,0)),0)),"""")"
        
        ' Topic
        ws.Cells(row, "J").formula = "=IFERROR(INDEX($B$7:$B$100,MATCH(G" & row & ",$E$7:$E$100,0)), """")"
        
        ' Attendees Required - show blank if 0 or no match
        ws.Cells(row, "K").formula = "=IFERROR(IF(INDEX($C$7:$C$100,MATCH(G" & row & ",$E$7:$E$100,0))=0,"""",INDEX($C$7:$C$100,MATCH(G" & row & ",$E$7:$E$100,0))),"""")"
        
        ' Topic Length (Minutes)
        ws.Cells(row, "L").formula = "=IFERROR(INDEX($D$7:$D$100,MATCH(G" & row & ",$E$7:$E$100,0)), """")"
        
        ' Gap (minutes) - User input, so clear any formula
        ws.Cells(row, "M").value = ""
        
        ' Break - will be dropdown defaulting to "No"
        ws.Cells(row, "N").value = "No"
    Next row
    
    ' Format all user input cells with yellow fill and grey border:
    ' Meeting Start Time input cell: C4 (already formatted above)
    
    ' Left table user input columns: B7:E21
    Set inputRange = ws.Range("B7:E21")
    FormatInputRange inputRange
    
    ' Right table Gap column (M7:M21) and Break column (N7:N21)
    Set inputRange = ws.Range("M7:N21")
    FormatInputRange inputRange
    
    ' Add data validation dropdown to Break column (N7:N21)
    With ws.Range("N7:N21").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Light borders on tables
    With ws.Range("B6:E21").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 15
    End With
    With ws.Range("G6:N21").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 15
    End With
    
    ' Conditional formatting: highlight entire output row if Break = "Yes"
    Set cfRange = ws.Range("G7:N21")
    With cfRange.FormatConditions
        .Delete
        .Add Type:=xlExpression, Formula1:="=$N7=""Yes"""
        With .Item(1)
            .Interior.color = RGB(217, 217, 217) ' light grey
        End With
    End With
    
    ' Group columns L to N for easy hide/unhide
    ws.Columns("L:N").Group
    
    Application.ScreenUpdating = True
    
    MsgBox "Agenda sheet created in " & Wb.Name & ". Fill the left table and enter gap times and breaks in the right table.", vbInformation
End Sub

' Helper sub to format user input ranges
Sub FormatInputRange(rng As Range)
    With rng
        .Interior.color = RGB(255, 255, 204) ' Light yellow fill
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.color = RGB(191, 191, 191) ' Light grey border
    End With
End Sub


Sub SplitEachWorksheet(control As IRibbonControl)
    Dim ws As Worksheet
    Dim folderPath As String
    Dim selectedSheets As Sheets
    Dim wbNew As Workbook

    On Error GoTo ErrorHandler

    ' Check if workbook is saved
    If Not ActiveWorkbook.Saved Then
        MsgBox "Please save this workbook before running the macro.", vbExclamation
        Exit Sub
    End If

    ' Check that at least one worksheet is selected
    If TypeName(Selection) <> "Nothing" Then
        If ActiveWindow.selectedSheets.count = 0 Then
            MsgBox "Please select at least one worksheet.", vbExclamation
            Exit Sub
        End If
    End If

    ' Ask user for destination folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to save split worksheets"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub ' User cancelled
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through only selected sheets
    Set selectedSheets = ActiveWindow.selectedSheets
    For Each ws In selectedSheets
        ws.Copy
        Set wbNew = ActiveWorkbook
        wbNew.SaveAs fileName:=folderPath & "\" & ws.Name & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next ws

    MsgBox "Selected worksheets have been saved to: " & folderPath, vbInformation

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Sub Split_Selected_Sheets_VALUES(control As IRibbonControl)
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim sourceWb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim Sh As Worksheet
    Dim TheActiveWindow As Window
    Dim TempWindow As Window
    Dim TempName As Variant
    Dim keepFormulas As VbMsgBoxResult

    On Error GoTo ErrorHandler

    Set sourceWb = ActiveWorkbook

    ' Check if saved
    If sourceWb.path = "" Then
        MsgBox "The active workbook has not been saved yet. Please save it before running this macro.", vbExclamation
        Exit Sub
    End If

    'Ask for temp file name
    TempName = InputBox("What do you want to call this file?")
    If StrPtr(TempName) = 0 Then Exit Sub ' Cancel pressed

    ' Ask whether to retain formulas
    keepFormulas = MsgBox("Do you want to retain formulas and links in the copied sheets?" & vbCrLf & vbCrLf & _
                          "Click 'Yes' to retain formulas." & vbCrLf & "Click 'No' to convert all to values.", _
                          vbYesNoCancel + vbQuestion, "Retain Formulas?")
    
    If keepFormulas = vbCancel Then Exit Sub

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    'Copy the sheets to a new workbook
    With sourceWb
        Set TheActiveWindow = ActiveWindow
        Set TempWindow = .NewWindow
        TheActiveWindow.selectedSheets.Copy
    End With

    TempWindow.Close
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If val(Application.Version) < 12 Then
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
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

    ' Optional: Convert cells to values
    If keepFormulas = vbNo Then
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
    End If

    ' Save the workbook
    TempFilePath = sourceWb.path & "\"
    TempFileName = TempName & " " & Format(Now, "dd-mmm-yy")
    Destwb.SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

    MsgBox "Sheets successfully exported to:" & vbCrLf & TempFilePath, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred - please try again."
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub



