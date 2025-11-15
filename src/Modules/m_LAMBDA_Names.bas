Attribute VB_Name = "m_LAMBDA_Names"
Sub LoadLambdas(control As IRibbonControl)
'PURPOSE: Load your favorite LAMBDA functions into the ActiveWorkbook

Dim LambdaName As String
Dim LambdaFormula As String
Dim LambdaComments As String
Dim LambdaList As String

'******************************************************************************

'NULL() Function
  'Lambda Info
    LambdaName = "NULL"
    LambdaFormula = "="""""
    LambdaComments = "Return a blank value"
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersToR1C1:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'SLOTS() Function
  'Lambda Info
    LambdaName = "SLOTS"
    LambdaFormula = "=CHOOSE(RANDBETWEEN(1,5), UNICHAR(127826), UNICHAR(127819), UNICHAR(128276), UNICHAR(127808), UNICHAR(128142))"
    LambdaComments = "Randomly generates an emoji similar to a slot machine - enter 3 of them side by side and hit calculate to test your luck."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersToR1C1:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'FINWEEK() Function
  'Lambda Info
    LambdaName = "FINWEEK"
    LambdaFormula = "=LAMBDA(refDate,fyYear,fyMonth,fyDay, LET( rawStart, DATE(fyYear, fyMonth, fyDay), adjustedFYStart, IF(refDate < rawStart, DATE(fyYear - 1, fyMonth, fyDay), rawStart), firstMonday, adjustedFYStart + MOD(8 - WEEKDAY(adjustedFYStart, 2), 7), weekNum, INT((refDate - firstMonday) / 7) + 1, IF(weekNum < 1, 1, weekNum) ) )"
    LambdaComments = "Return the financial week number based on the reference date (refDate) and the start of the Financial Year (fyYear, fyMonth, fyDay). refDate should be a real date (e.g. Date(2025,04,05) or Today()). Financial Year should reflect the start day of the fin year."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersToR1C1:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'PROGRESS() Function
  'Lambda Info
    LambdaName = "PROGRESS"
    LambdaFormula = "=LAMBDA(current,total, LET( filled, REPT(UNICHAR(9608), INT(10 *@ current /@ total)), empty, REPT(UNICHAR(9617), 10 - INT(10 *@ current /@ total)), filled & empty ) )"
    LambdaComments = "Creates a progress bar in the cell based on the Current and Total variables."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersToR1C1:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'CALENDAR() Function
  'Lambda Info
    LambdaName = "CALENDAR"
    LambdaFormula = "=LAMBDA(Year,Month,Day,LET(INPUT,DATE(Year,Month,Day), A, EXPAND(TEXT(SEQUENCE(7),""ddd""),6+WEEKDAY(INPUT,1),,""""), B, DAY(SEQUENCE(EOMONTH(INPUT,0)-INPUT+1,,INPUT)), C, EXPAND(UPPER(TEXT(INPUT,""MMM"")),7,,""""), D, WRAPROWS(VSTACK(C,A,B),7,""""),D))"
    LambdaComments = "Generate a monthly calendar that starts from a given date. Usage: CALENDAR(Year,Month,Day)"
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'BOMONTH() Function
  'Lambda Info
    LambdaName = "BOMONTH"
    LambdaFormula = "=LAMBDA(Start_Date,Months,EOMONTH(Start_Date,Months-1)+1)"
    LambdaComments = "Behaves like the native EOMONTH function, but returns the beginning of month instead of end of month."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'FIRSTWORD() Function
  'Lambda Info
    LambdaName = "FIRSTWORD"
    LambdaFormula = "=LAMBDA(Text_String,IFERROR(LEFT(Text_String,FIND("" "",Text_String)-1),Text_String))"
    LambdaComments = "Extracts the first word from a text string."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'CAGR() Function
  'Lambda Info
    LambdaName = "CAGR"
    LambdaFormula = "=LAMBDA(Beginning_Value,Ending_Value,IF(ROW(Beginning_Value) = ROW(Ending_Value), RRI(COLUMN(Ending_Value) - COLUMN(Beginning_Value), Beginning_Value, Ending_Value), RRI(ROW(Ending_Value) - ROW(Beginning_Value), Beginning_Value, Ending_Value)))"
    LambdaComments = "Calculates the compounded annual growth rate (CAGR) between two values and assumes the number of columns/rows reflects the number of periods; works both horizontally and vertically."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'TIMESTAMP() Function
  'Lambda Info
    LambdaName = "TIMESTAMP"
    LambdaFormula = "=LAMBDA([Include_Time?],IF(OR(ISOMITTED(Include_Time?)=TRUE),""Last Saved: ""&TEXT(NOW(),""m/d/yyyy""),""Last Saved: ""&TEXT(NOW(),""m/d/yyyy, h:mm am/pm"")))"
    LambdaComments = "Returns the current date when file is saved; optional argument is to add the time."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'BPS() Function
  'Lambda Info
    LambdaName = "BPS"
    LambdaFormula = "=LAMBDA(Percentage_Start,Percentage_End,IFERROR((Percentage_End-Percentage_Start)*10000,0))"
    LambdaComments = "Calculates Basis Point movement between two percentages."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments
    
'******************************************************************************

'QTRFY() Function
  'Lambda Info
    LambdaName = "QTRFY"
    LambdaFormula = "=LAMBDA(mths, IF(COLUMNS(mths)<>12, ""#INVALID - Range must be 12 cells"", HSTACK( BYCOL( WRAPCOLS(mths, 3), LAMBDA(mths,SUM(mths))), SUM(mths))))"
    LambdaComments = "Input range of 12 cells (months) and will output corresponding quarter and full year totals."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'DATAARRAY() Function
  'Lambda Info
    LambdaName = "DATAARRAY"
    LambdaFormula = "=LAMBDA(DataArray, LET(data, DataArray, combos, REDUCE("""", SEQUENCE(COLUMNS(data)), LAMBDA(a, v, TRIM(TOCOL(a & TOROW("" "" & INDEX(data, , v)))))), length, LEN(combos) - LEN(SUBSTITUTE(combos, "" "", """")) + 1, FILTER(combos, COLUMNS(data) = length) ))"
    LambdaComments = "Creates an array from data in multiple columns to reflect all the permutations when each column is combined together. For example, you could have 6 data points in Col A, 4 data points in Col B and 2 data points in Col C - this would combine each permutation across the cols to create the new list."
    LambdaList = LambdaList & LambdaName & "( )" & vbNewLine
    
  'Create Named Range + Formula
    ActiveWorkbook.Names.Add Name:=LambdaName, RefersTo:=LambdaFormula
    
  'Add Comments for the LAMBDA function
    ActiveWorkbook.Names(LambdaName).Comment = LambdaComments

'******************************************************************************

'Completion Message
  MsgBox "Your LAMBDA functions have been loaded into this workbook: " _
  & vbNewLine & vbNewLine & LambdaList

End Sub

Sub ListAllDefinedNames()
    Dim nm As Name
    Dim wsReport As Worksheet
    Dim rowNum As Long

    ' Make sure there's an active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    ' Add a new worksheet for the report
    Set wsReport = ActiveWorkbook.Worksheets.Add
    wsReport.Name = "Defined Names Report"
    
    ' Add headers
    With wsReport
        .Cells(1, 1).value = "Name"
        .Cells(1, 2).value = "RefersTo"
        .Cells(1, 3).value = "Scope"
        .Cells(1, 4).value = "Visible"
        .rows(1).Font.Bold = True
    End With
    
    rowNum = 2

    ' Loop through all names in the active workbook
    For Each nm In ActiveWorkbook.Names
        With wsReport
            .Cells(rowNum, 1).value = nm.Name
            .Cells(rowNum, 2).value = "'" & nm.RefersTo ' Keep formula readable
            On Error Resume Next
            .Cells(rowNum, 3).value = IIf(nm.Parent.Name = ActiveWorkbook.Name, "Workbook", nm.Parent.Name)
            .Cells(rowNum, 4).value = nm.Visible
            On Error GoTo 0
        End With
        rowNum = rowNum + 1
    Next nm

    ' Autofit columns
    wsReport.Columns("A:D").AutoFit

    MsgBox "Report created on sheet: '" & wsReport.Name & "'", vbInformation
End Sub

Sub UnhideAllHiddenNames(control As IRibbonControl)
    Dim nmArray() As Name
    Dim i As Long
    Dim countUnhidden As Long
    Dim countSkipped As Long
    Dim failedNames As String
    Dim nameCount As Long

    ' Ensure there's an active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    ' Copy names to an array to avoid dynamic collection access
    nameCount = ActiveWorkbook.Names.count
    If nameCount = 0 Then
        MsgBox "No defined names found in the active workbook.", vbInformation
        Exit Sub
    End If

    ReDim nmArray(1 To nameCount)
    For i = 1 To nameCount
        Set nmArray(i) = ActiveWorkbook.Names(i)
    Next i

    ' Loop through the static array
    For i = 1 To nameCount
        On Error Resume Next
        If Not nmArray(i).Visible Then
            nmArray(i).Visible = True
            If Err.Number = 0 Then
                countUnhidden = countUnhidden + 1
            Else
                failedNames = failedNames & nmArray(i).Name & " - " & Err.Description & vbNewLine
                countSkipped = countSkipped + 1
                Err.Clear
            End If
        End If
        On Error GoTo 0
    Next i

    ' Show summary
    Dim msg As String
    msg = countUnhidden & " hidden name(s) were unhidden."
    If countSkipped > 0 Then
        msg = msg & vbNewLine & countSkipped & " name(s) could not be modified:" & vbNewLine & failedNames
    End If

    MsgBox msg, vbInformation, "Unhide Names Summary"
End Sub

Sub AnalyzeDefinedNames()
    Dim nmArray() As Name
    Dim i As Long
    Dim totalCount As Long
    Dim visibleCount As Long
    Dim hiddenCount As Long
    Dim refErrorCount As Long
    Dim refErrorNames As String
    Dim nameCount As Long
    Dim tempRef As String

    ' Check for active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    nameCount = ActiveWorkbook.Names.count
    If nameCount = 0 Then
        MsgBox "No defined names in the active workbook.", vbInformation
        Exit Sub
    End If

    ' Cache names to an array to avoid repeated access
    ReDim nmArray(1 To nameCount)
    For i = 1 To nameCount
        Set nmArray(i) = ActiveWorkbook.Names(i)
    Next i

    ' Analyze each name
    For i = 1 To nameCount
        totalCount = totalCount + 1

        On Error Resume Next
        If nmArray(i).Visible Then
            visibleCount = visibleCount + 1
        Else
            hiddenCount = hiddenCount + 1
        End If

        tempRef = nmArray(i).RefersTo
        If InStr(1, tempRef, "#REF!", vbTextCompare) > 0 Then
            refErrorCount = refErrorCount + 1
            refErrorNames = refErrorNames & nmArray(i).Name & " - RefersTo: " & tempRef & vbNewLine
        End If
        On Error GoTo 0
    Next i

    ' Build summary
    Dim msg As String
    msg = "Defined Name Summary for '" & ActiveWorkbook.Name & "':" & vbNewLine & vbNewLine & _
          "Total Names: " & totalCount & vbNewLine & _
          "Visible Names: " & visibleCount & vbNewLine & _
          "Hidden Names: " & hiddenCount & vbNewLine & _
          "Names with #REF! errors: " & refErrorCount

    If refErrorCount > 0 Then
        msg = msg & vbNewLine & vbNewLine & "Names with #REF! errors:" & vbNewLine & refErrorNames
    End If

    MsgBox msg, vbInformation, "Name Analysis"
End Sub


Sub RemoveAllNames(control As IRibbonControl)
    Dim nmArray() As Name
    Dim i As Long
    Dim removedCount As Long
    Dim failedCount As Long
    Dim failedNames As String
    Dim nameCount As Long
    Dim msg As String

    ' Disable screen updates and alerts to speed up
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .StatusBar = False ' Reset status bar
    End With

    On Error GoTo CleanupWithError

    ' Check workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is currently active.", vbExclamation
        GoTo Cleanup
    End If

    ' Confirm deletion
    If MsgBox("Are you sure you want to remove ALL names (including hidden) from '" & ActiveWorkbook.Name & "'?", _
              vbYesNo + vbQuestion, "Confirm Name Removal") <> vbYes Then
        GoTo Cleanup
    End If

    ' Cache names to array
    nameCount = ActiveWorkbook.Names.count
    If nameCount = 0 Then
        MsgBox "No defined names found.", vbInformation
        GoTo Cleanup
    End If

    ReDim nmArray(1 To nameCount)
    For i = 1 To nameCount
        Set nmArray(i) = ActiveWorkbook.Names(i)
    Next i

    ' Delete names from cached array
    For i = nameCount To 1 Step -1
        ' Update progress in status bar every 10 items (adjust as needed)
        If i Mod 10 = 0 Or i = nameCount Then
            Application.StatusBar = "Removing names... " & (nameCount - i + 1) & " of " & nameCount
        End If

        On Error Resume Next
        nmArray(i).Visible = True ' Attempt to unhide
        nmArray(i).Delete
        If Err.Number <> 0 Then
            failedNames = failedNames & nmArray(i).Name & " - " & Err.Description & vbNewLine
            failedCount = failedCount + 1
            Err.Clear
        Else
            removedCount = removedCount + 1
        End If
        On Error GoTo 0
    Next i

    ' Final message
    msg = removedCount & " name(s) successfully removed."
    If failedCount > 0 Then
        msg = msg & vbNewLine & failedCount & " name(s) failed to delete:" & vbNewLine & failedNames
    End If
    MsgBox msg, vbInformation, "Name Removal Summary"

Cleanup:
    ' Restore Excel settings
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False ' Reset status bar
    End With
    Exit Sub

CleanupWithError:
    MsgBox "An unexpected error occurred: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub


Sub RemoveBrokenNames(control As IRibbonControl)
    Dim nmArray() As Name
    Dim i As Long
    Dim nameCount As Long
    Dim removedCount As Long
    Dim brokenNameList() As String
    Dim brokenIndex As Long
    Dim tempRef As String

    ' Speed up Excel
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    On Error GoTo CleanupWithError

    ' Check active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook is currently active.", vbExclamation
        GoTo Cleanup
    End If

    nameCount = ActiveWorkbook.Names.count
    If nameCount = 0 Then
        MsgBox "No defined names found.", vbInformation
        GoTo Cleanup
    End If

    ' Cache names
    ReDim nmArray(1 To nameCount)
    For i = 1 To nameCount
        Set nmArray(i) = ActiveWorkbook.Names(i)
    Next i

    ' Prepare broken names list
    ReDim brokenNameList(1 To nameCount)
    brokenIndex = 0

    ' Loop backwards and delete broken names
    For i = nameCount To 1 Step -1
        On Error Resume Next
        tempRef = nmArray(i).RefersTo
        If Err.Number <> 0 Or InStr(1, tempRef, "#REF!", vbTextCompare) > 0 Then
            Err.Clear
            nmArray(i).Delete
            removedCount = removedCount + 1
            brokenIndex = brokenIndex + 1
            brokenNameList(brokenIndex) = nmArray(i).Name
        End If
        On Error GoTo 0
    Next i

    ' Show result
    If removedCount > 0 Then
        MsgBox removedCount & " broken name(s) removed:" & vbNewLine & Join(SliceArray(brokenNameList, brokenIndex), vbNewLine), _
               vbInformation, "Broken Names Removed"
    Else
        MsgBox "No broken names found in '" & ActiveWorkbook.Name & "'.", vbInformation
    End If

Cleanup:
    ' Restore settings
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub

CleanupWithError:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Function SliceArray(arr() As String, count As Long) As Variant
    ' Returns first N elements of array
    If count = 0 Then
        SliceArray = Array()
    Else
        ReDim temp(1 To count) As String
        Dim i As Long
        For i = 1 To count
            temp(i) = arr(i)
        Next i
        SliceArray = temp
    End If
End Function


Sub DeleteAllNames_Fastest()
    Dim nmArray() As Name
    Dim i As Long
    Dim nameCount As Long

    ' Performance tuning
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Removing all names..."
    End With

    On Error GoTo CleanupWithError

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation
        GoTo Cleanup
    End If

    nameCount = ActiveWorkbook.Names.count
    If nameCount = 0 Then GoTo Cleanup

    ' Cache names
    ReDim nmArray(1 To nameCount)
    For i = 1 To nameCount
        Set nmArray(i) = ActiveWorkbook.Names(i)
    Next i

    ' Delete from cached array (no visibility check = faster)
    For i = nameCount To 1 Step -1
        On Error Resume Next
        nmArray(i).Delete
        Err.Clear
        On Error GoTo 0
    Next i

Cleanup:
    ' Restore settings
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub

CleanupWithError:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub


