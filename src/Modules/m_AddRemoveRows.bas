Attribute VB_Name = "m_AddRemoveRows"
Sub DeleteEveryXthRow(control As IRibbonControl)
    Dim rng As Range
    Dim x As Long, y As Long
    Dim i As Long
    Dim deleteRow As Long
    Dim deletedCount As Long
    Dim startRow As Long
    Dim xRowCount As Long
    Dim userInput As String

    On Error GoTo ErrorHandler

    ' Validate selection
    If Selection Is Nothing Then
        MsgBox "Please select a range before running the macro.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    startRow = rng.rows(1).row

    ' Ask user for interval (X)
    userInput = InputBox("Enter the interval for deleting rows (e.g. 3 will delete every 3rd row)", _
                         "Delete Every Xth Row")
    If Not IsNumeric(userInput) Or val(userInput) < 1 Then
        MsgBox "Please enter a valid positive number for the interval.", vbExclamation
        Exit Sub
    End If
    x = CLng(userInput)

    xRowCount = rng.rows.count \ x

    ' Ask user how many rows to delete each time (Y)
    userInput = InputBox("You've selected " & rng.rows.count & " rows." & vbNewLine & _
                         "That means " & xRowCount & " delete points (every " & x & "th row)." & vbNewLine & _
                         "How many rows would you like to delete at each point?", _
                         "Rows to Delete", 1)
    If userInput = "" Then Exit Sub ' User cancelled
    If Not IsNumeric(userInput) Or val(userInput) < 1 Then
        MsgBox "Please enter a valid positive number for rows to delete.", vbExclamation
        Exit Sub
    End If
    y = CLng(userInput)

    deletedCount = 0

    ' Delete rows from top to bottom, adjusting for already deleted rows
    For i = 1 To rng.rows.count
        If i Mod x = 0 Then
            deleteRow = startRow + i - 1 - deletedCount
            rows(deleteRow & ":" & deleteRow + y - 1).Delete Shift:=xlUp
            deletedCount = deletedCount + y
        End If
    Next i

    MsgBox "Deleted " & y & " row(s) every " & x & "th row (" & xRowCount & " times).", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub


Sub InsertRowsEveryXth(control As IRibbonControl)
    Dim rng As Range
    Dim x As Long, y As Long
    Dim i As Long
    Dim insertRow As Long
    Dim insertedCount As Long
    Dim startRow As Long
    Dim xRowCount As Long
    Dim userInput As String

    On Error GoTo ErrorHandler

    ' Validate selection
    If Selection Is Nothing Then
        MsgBox "Please select a range before running the macro.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    startRow = rng.rows(1).row

    ' Prompt for interval X
    userInput = InputBox("Enter the interval for inserting rows (e.g. 3 will insert rows after every 3rd row)", _
                         "Insert Every Xth Row")

    If Not IsNumeric(userInput) Or val(userInput) < 1 Then
        MsgBox "Please enter a valid positive number for the interval.", vbExclamation
        Exit Sub
    End If

    x = CLng(userInput)
    xRowCount = rng.rows.count \ x

    ' Prompt for how many rows to insert each time (Y)
    userInput = InputBox("You've selected " & rng.rows.count & " rows." & vbNewLine & _
                         "That means " & xRowCount & " insert points (every " & x & "th row)." & vbNewLine & _
                         "How many rows would you like to insert each time?", _
                         "Rows to Insert", 1)

    If userInput = "" Then Exit Sub ' User cancelled
    If Not IsNumeric(userInput) Or val(userInput) < 1 Then
        MsgBox "Please enter a valid positive number for rows to insert.", vbExclamation
        Exit Sub
    End If

    y = CLng(userInput)
    insertedCount = 0

    ' Loop top to bottom, adjust for inserted rows
    For i = 1 To rng.rows.count
        If i Mod x = 0 Then
            insertRow = startRow + i + insertedCount
            rows(insertRow & ":" & insertRow + y - 1).Insert Shift:=xlDown
            insertedCount = insertedCount + y
        End If
    Next i

    MsgBox "Inserted " & y & " row(s) after every " & x & "th row (" & xRowCount & " times).", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub


