Attribute VB_Name = "m_Modelling"
Option Explicit

Sub UnhideRowsColumns(control As IRibbonControl)
    Columns.EntireColumn.Hidden = False
    rows.EntireRow.Hidden = False
End Sub

Sub Tab_name(control As IRibbonControl)
'Adds tab name to selected cell

If ActiveWorkbook.path = "" Then
    MsgBox "Please save the workbook before inserting the tab name.", vbExclamation
    Exit Sub
End If

ActiveCell.value = "MID(CELL(" & Chr(34) & "filename" & Chr(34) & ",A1),FIND(" & Chr(34) & "]" & Chr(34) & ",CELL(" & Chr(34) & "filename" & Chr(34) & ",A1))+1,256)"

ActiveCell.Resize(2, 2).Select

ActiveCell.Replace What:="MID(CELL(", Replacement:="=MID(CELL(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

ActiveCell.Select

End Sub

Sub FillSequentialNumbers(control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection ' Or specify a range like Range("A1:A10")
    
    Dim cell As Range
    Dim counter As Long
    counter = 1
    
    ' Loop through each cell in the selection in order
    For Each cell In rng.Cells
        cell.value = counter
        counter = counter + 1
    Next cell
End Sub


Sub FillSequentialLetters(control As IRibbonControl)
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim count As Long
    
    Set rng = Selection
    count = 0
    
    Application.ScreenUpdating = False

    For Each cell In rng.Cells
        cell.value = GetLetterFromIndex(count)
        count = count + 1
    Next cell

    Application.ScreenUpdating = True
End Sub

' Converts 0-based index to Excel column-style letters
Function GetLetterFromIndex(index As Long) As String
    Dim result As String
    Dim n As Long
    n = index + 1 ' Convert to 1-based

    Do
        n = n - 1
        result = Chr(65 + (n Mod 26)) & result
        n = n \ 26
    Loop While n > 0

    GetLetterFromIndex = result
End Function


Sub FlagFormulasWithPlugs(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim cell As Range
    Dim regexCellRef As Object, regexNumber As Object
    Dim cellRefs As Object, numberMatches As Object
    Dim formula As String, match, val
    Dim hasPlug As Boolean

    ' Regex for cell refs like A1, $A$1
    Set regexCellRef = CreateObject("VBScript.RegExp")
    regexCellRef.pattern = "\$?[A-Z]{1,3}\$?[0-9]+"
    regexCellRef.IgnoreCase = True
    regexCellRef.Global = True

    ' Regex for standalone numbers
    Set regexNumber = CreateObject("VBScript.RegExp")
    regexNumber.pattern = "[-+]?\b\d+(\.\d+)?\b"
    regexNumber.IgnoreCase = True
    regexNumber.Global = True

    ' Functions where hardcoded numbers are acceptable (logic/flags/criteria)
    Dim ignoreFunctions As Variant
    ignoreFunctions = Array( _
    "XLOOKUP", "VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "IF", "IFS", "IFERROR", "LAMBDA", "SWITCH", _
    "SUMIFS", "SUBTOTAL", "COUNTIFS", "AVERAGEIFS", "INDEX", "CHOOSE", "LET", "RRI", "PV", "NPV", _
    "PMT", "FV", "IRR", "XIRR", "XNPV", "NPER", "RATE", "SLN", "SYD", "DB", _
    "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "CEILING", "FLOOR", "MOD", "ABS", _
    "RANK", "PERCENTILE", "QUARTILE", "STDEV", "VAR", "MEDIAN", _
    "LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH", "TEXT", "CONCAT", "CONCATENATE", "TEXTJOIN", _
    "REPT", "SUBSTITUTE", "REPLACE", "VALUE", "LOWER", "UPPER", "PROPER", "TRIM", _
    "OFFSET", "INDIRECT", "ADDRESS", "ROW", "ROWS", "COLUMN", "COLUMNS", _
    "SEQUENCE", "SORT", "FILTER", "UNIQUE", "TRANSPOSE", "XMATCH", _
    "DATE", "DATEDIF", "EDATE", "EOMONTH", "YEAR", "MONTH", "DAY", "WEEKDAY", _
    "HOUR", "MINUTE", "SECOND", "TODAY", "NOW", "TIME", "DATEVALUE", "TIMEVALUE")

    For Each cell In ws.UsedRange
        If cell.HasFormula Then
            formula = Mid(cell.formula, 2) ' Strip leading '='

            ' Skip very short formulas
            If Len(formula) < 3 Then GoTo NextCell

            ' Get all cell references
            Set cellRefs = CreateObject("Scripting.Dictionary")
            For Each match In regexCellRef.Execute(formula)
                cellRefs(match.value) = True
            Next match

            ' Get all numeric matches
            Set numberMatches = regexNumber.Execute(formula)
            hasPlug = False

            For Each match In numberMatches
                val = match.value

                ' Ignore if part of a cell reference
                Dim isInRef As Boolean: isInRef = False
                Dim r
                For Each r In cellRefs.keys
                    If InStr(r, val) > 0 Then
                        isInRef = True
                        Exit For
                    End If
                Next r
                If isInRef Then GoTo ContinueLoop

                ' Ignore if inside an allowed function
                Dim func
                For Each func In ignoreFunctions
                    If InStr(1, UCase(formula), func & "(", vbTextCompare) > 0 Then
                        Dim funcStart As Long
                        funcStart = InStr(1, UCase(formula), func & "(", vbTextCompare)
                        Dim afterFunc As String
                        afterFunc = Mid(formula, funcStart)

                        ' Check if the number appears inside the function's parentheses
                        Dim parenDepth As Long: parenDepth = 0
                        Dim i As Long
                        For i = 1 To Len(afterFunc)
                            Dim ch As String: ch = Mid(afterFunc, i, 1)
                            If ch = "(" Then
                                parenDepth = parenDepth + 1
                            ElseIf ch = ")" Then
                                parenDepth = parenDepth - 1
                                If parenDepth = 0 Then Exit For
                            ElseIf parenDepth > 0 And InStr(Mid(afterFunc, i), val) = 1 Then
                                GoTo ContinueLoop ' number is inside the function
                            End If
                        Next i
                    End If
                Next func

                ' If it wasn't excluded, it's a plug
                hasPlug = True
                Exit For

ContinueLoop:
            Next match

            If hasPlug Then
                With cell.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                    .color = vbRed
                End With
                With cell.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                    .color = vbRed
                End With
                With cell.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                    .color = vbRed
                End With
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                    .color = vbRed
                End With
            End If
        End If
NextCell:
    Next cell

    MsgBox "Finished checking for plug values. These will be highlighted with a thick red border.", vbInformation
End Sub

Sub FlagFormulasWithPlugsAndReport(control As IRibbonControl)
    Dim ws As Worksheet, summaryWS As Worksheet
    Dim cell As Range
    Dim regexCellRef As Object, regexNumber As Object
    Dim cellRefs As Object, numberMatches As Object
    Dim summaryRow As Long, f As String
    Dim Wb As Workbook, match, val, targetAddress As String
    Dim hasPlug As Boolean

    Set Wb = ActiveWorkbook

    ' Delete old summary if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Wb.Worksheets("Plug Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new summary sheet
    Set summaryWS = Wb.Worksheets.Add(before:=Worksheets(1))
    summaryWS.Name = "Plug Summary"

    ' Add headers
    summaryWS.Range("A1:D1").value = Array("Sheet Name", "Cell Address", "Formula", "Comment")
    summaryWS.Range("A1:D1").Font.Bold = True
    summaryWS.Range("A1:D1").Borders(xlEdgeBottom).LineStyle = xlContinuous
    summaryWS.Range("A1:D1").Borders(xlEdgeBottom).Weight = xlMedium
    summaryWS.Columns("A:D").ColumnWidth = 30
    summaryWS.Range(Cells(1, 5), Cells(1, Columns.count)).EntireColumn.Hidden = True
    
    summaryRow = 2

    ' Regex for cell refs like $A$1 or A1
    Set regexCellRef = CreateObject("VBScript.RegExp")
    regexCellRef.pattern = "\$?[A-Z]{1,3}\$?[0-9]+"
    regexCellRef.IgnoreCase = True
    regexCellRef.Global = True

    ' Regex for standalone numbers
    Set regexNumber = CreateObject("VBScript.RegExp")
    regexNumber.pattern = "[-+]?\b\d+(\.\d+)?\b"
    regexNumber.IgnoreCase = True
    regexNumber.Global = True

    ' Functions to ignore plugs in
    Dim ignoreFunctions As Variant
    ignoreFunctions = Array( _
    "XLOOKUP", "VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "IF", "IFS", "IFERROR", "LAMBDA", "SWITCH", _
    "SUMIFS", "SUBTOTAL", "COUNTIFS", "AVERAGEIFS", "INDEX", "CHOOSE", "LET", "RRI", "PV", "NPV", _
    "PMT", "FV", "IRR", "XIRR", "XNPV", "NPER", "RATE", "SLN", "SYD", "DB", _
    "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "CEILING", "FLOOR", "MOD", "ABS", _
    "RANK", "PERCENTILE", "QUARTILE", "STDEV", "VAR", "MEDIAN", _
    "LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH", "TEXT", "CONCAT", "CONCATENATE", "TEXTJOIN", _
    "REPT", "SUBSTITUTE", "REPLACE", "VALUE", "LOWER", "UPPER", "PROPER", "TRIM", _
    "OFFSET", "INDIRECT", "ADDRESS", "ROW", "ROWS", "COLUMN", "COLUMNS", _
    "SEQUENCE", "SORT", "FILTER", "UNIQUE", "TRANSPOSE", "XMATCH", _
    "DATE", "DATEDIF", "EDATE", "EOMONTH", "YEAR", "MONTH", "DAY", "WEEKDAY", _
    "HOUR", "MINUTE", "SECOND", "TODAY", "NOW", "TIME", "DATEVALUE", "TIMEVALUE")
    
    ' Loop through all worksheets
    For Each ws In Wb.Worksheets
        If ws.Name <> "Plug Summary" Then
            For Each cell In ws.UsedRange
                If cell.HasFormula Then
                    f = Mid(cell.formula, 2) ' Strip leading "="
                    Set cellRefs = CreateObject("Scripting.Dictionary")

                    ' Get cell refs
                    For Each match In regexCellRef.Execute(f)
                        cellRefs(match.value) = True
                    Next match

                    ' Get number matches
                    Set numberMatches = regexNumber.Execute(f)
                    hasPlug = False

                    For Each match In numberMatches
                        val = match.value

                        ' Skip if part of cell ref
                        Dim inRef As Boolean: inRef = False
                        Dim r
                        For Each r In cellRefs.keys
                            If InStr(r, val) > 0 Then
                                inRef = True
                                Exit For
                            End If
                        Next r
                        If inRef Then GoTo NextMatch

                        ' Skip if number is inside ignored function
                        Dim func
                        For Each func In ignoreFunctions
                            If InStr(1, UCase(f), func & "(", vbTextCompare) > 0 Then
                                Dim funcStart As Long
                                funcStart = InStr(1, UCase(f), func & "(", vbTextCompare)
                                Dim afterFunc As String
                                afterFunc = Mid(f, funcStart)

                                Dim parenDepth As Long: parenDepth = 0
                                Dim i As Long
                                For i = 1 To Len(afterFunc)
                                    Dim ch As String: ch = Mid(afterFunc, i, 1)
                                    If ch = "(" Then
                                        parenDepth = parenDepth + 1
                                    ElseIf ch = ")" Then
                                        parenDepth = parenDepth - 1
                                        If parenDepth = 0 Then Exit For
                                    ElseIf parenDepth > 0 And InStr(Mid(afterFunc, i), val) = 1 Then
                                        GoTo NextMatch
                                    End If
                                Next i
                            End If
                        Next func

                        ' Otherwise it's a plug
                        hasPlug = True
                        Exit For

NextMatch:
                    Next match

                    If hasPlug Then
                        With summaryWS
                            .Cells(summaryRow, 1).value = ws.Name
                            targetAddress = "'" & ws.Name & "'!" & cell.Address
                            .Hyperlinks.Add _
                                Anchor:=.Cells(summaryRow, 2), _
                                Address:="", _
                                SubAddress:=targetAddress, _
                                TextToDisplay:=cell.Address
                            .Cells(summaryRow, 3).value = "'" & cell.formula
                            .Cells(summaryRow, 4).value = GetPlugComment(cell.formula)
                        End With
                        summaryRow = summaryRow + 1
                    End If
                End If
            Next cell
        End If
    Next ws

    MsgBox "Plug formulas flagged and summary with hyperlinks created.", vbInformation
End Sub


Function GetPlugComment(formula As String) As String
    Dim regexNumber As Object, regexCellRef As Object
    Dim matches, match
    Dim cellRefs As Object, ignoreFunctions As Variant
    Dim f As String: f = Mid(formula, 2) ' Remove leading "="

    ' Setup regex for numbers
    Set regexNumber = CreateObject("VBScript.RegExp")
    regexNumber.pattern = "[-+]?\b\d+(\.\d+)?\b"
    regexNumber.IgnoreCase = True
    regexNumber.Global = True

    ' Setup regex for cell references like A1 or $B$2
    Set regexCellRef = CreateObject("VBScript.RegExp")
    regexCellRef.pattern = "\$?[A-Z]{1,3}\$?[0-9]+"
    regexCellRef.IgnoreCase = True
    regexCellRef.Global = True

    ' Store all found cell references
    Set cellRefs = CreateObject("Scripting.Dictionary")
    For Each match In regexCellRef.Execute(f)
        cellRefs(match.value) = True
    Next match

    ' Functions to ignore plugs inside
    ignoreFunctions = Array("XLOOKUP", "VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "IF", "IFS", "SWITCH", "SUMIFS", "COUNTIFS", "AVERAGEIFS", "INDEX", "CHOOSE")

    Set matches = regexNumber.Execute(f)
    If matches.count = 0 Then
        GetPlugComment = "No hardcoded value"
        Exit Function
    End If

    ' Check each match to see if it's a real plug
    Dim plugs As Collection: Set plugs = New Collection
    Dim val, inRef As Boolean, func, funcStart As Long, afterFunc As String
    Dim i As Long, ch As String, parenDepth As Long

    For Each match In matches
        val = match.value
        inRef = False

        ' Ignore if part of a cell reference
        For Each key In cellRefs.keys
            If InStr(key, val) > 0 Then
                inRef = True
                Exit For
            End If
        Next key
        If inRef Then GoTo SkipMatch

        ' Ignore if within one of the flagged functions
        For Each func In ignoreFunctions
            If InStr(1, UCase(f), func & "(", vbTextCompare) > 0 Then
                funcStart = InStr(1, UCase(f), func & "(", vbTextCompare)
                afterFunc = Mid(f, funcStart)

                parenDepth = 0
                For i = 1 To Len(afterFunc)
                    ch = Mid(afterFunc, i, 1)
                    If ch = "(" Then
                        parenDepth = parenDepth + 1
                    ElseIf ch = ")" Then
                        parenDepth = parenDepth - 1
                        If parenDepth = 0 Then Exit For
                    ElseIf parenDepth > 0 And InStr(Mid(afterFunc, i), val) = 1 Then
                        GoTo SkipMatch
                    End If
                Next i
            End If
        Next func

        ' Passed all checks, count it as a plug
        On Error Resume Next
        If val <> "" Then plugs.Add val
        On Error GoTo 0

SkipMatch:
    Next match

    If plugs.count = 0 Then
        GetPlugComment = "No plug values detected"
        Exit Function
    End If

    ' Create plug list string
    Dim plugList As String
    For i = 1 To plugs.count
        plugList = plugList & plugs(i)
        If i < plugs.count Then plugList = plugList & ", "
    Next i

    If plugs.count = 1 Then
        GetPlugComment = "Plug detected: " & plugList
    Else
        GetPlugComment = "Plugs detected: " & plugList
    End If
End Function

Sub ResetUsedRange(control As IRibbonControl)
  Dim LastUsedRow As Long, LastUsedCol As Long
  LastUsedRow = Cells.Find(What:="*", SearchOrder:=xlRows, _
                SearchDirection:=xlPrevious, LookIn:=xlFormulas).row
  LastUsedCol = Cells.Find(What:="*", SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, LookIn:=xlFormulas).Column
  rows(LastUsedRow + 1).Resize(rows.count - LastUsedRow).Delete
  Columns(LastUsedCol + 1).Resize(, Columns.count - LastUsedCol).Delete
End Sub

Sub QuickUnpivot(control As IRibbonControl)
    Dim rng As Range
    Dim idCols As Variant
    Dim r As Long, c As Long, outRow As Long
    Dim outputSheet As Worksheet
    Dim idColumnCount As Integer
    Dim i As Integer
    
    ' Ask user to select the range
    On Error Resume Next
    Set rng = Application.InputBox("Select your data table (including headers):", "Quick Unpivot", Type:=8)
    If rng Is Nothing Then Exit Sub
    On Error GoTo 0
    
    

    ' Ask how many ID columns to keep
    idCols = Application.InputBox("How many leftmost columns are IDs (not to unpivot)?", "ID Columns", 1, Type:=1)
    If idCols = False Or Not IsNumeric(idCols) Then Exit Sub
    idColumnCount = CInt(idCols)
    
    If idColumnCount >= rng.Columns.count Then
        MsgBox "Number of ID columns must be less than total number of columns.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False ' Speed up execution
    
    ' Create output sheet
    Set outputSheet = ActiveWorkbook.Sheets.Add
    outputSheet.Name = "Unpivoted_" & Format(Now, "hhmmss")
    outRow = 2

    ' Write headers
    For i = 1 To idColumnCount
        outputSheet.Cells(1, i).value = rng.Cells(1, i).value
    Next i
    outputSheet.Cells(1, idColumnCount + 1).value = "Attribute"
    outputSheet.Cells(1, idColumnCount + 2).value = "Value"

    ' Loop through data
    For r = 2 To rng.rows.count
        For c = idColumnCount + 1 To rng.Columns.count
            ' Copy ID columns
            For i = 1 To idColumnCount
                outputSheet.Cells(outRow, i).value = rng.Cells(r, i).value
            Next i
            ' Copy attribute name and value
            outputSheet.Cells(outRow, idColumnCount + 1).value = rng.Cells(1, c).value
            outputSheet.Cells(outRow, idColumnCount + 2).value = rng.Cells(r, c).value
            outRow = outRow + 1
        Next c
    Next r

    Application.ScreenUpdating = True ' Speed up execution
    
    outputSheet.Columns.AutoFit
    MsgBox "Unpivot complete. Output is in sheet: " & outputSheet.Name, vbInformation
End Sub



Sub List_Unique_Values()
'Create a list of unique values from the selected columns

Dim rSelection As Range
Dim ws As Worksheet
Dim vArray() As Long
Dim i As Long
Dim iColCount As Long

  'Check that a range is selected
  If TypeName(Selection) <> "Range" Then
    MsgBox "Please select a range first.", vbOKOnly, "List Unique Values Macro"
    Exit Sub
  End If
  
  'Store the selected range
  Set rSelection = Selection

  'Add a new worksheet
  Set ws = Worksheets.Add
  
  'Copy/paste selection to the new sheet
  rSelection.Copy
  
  With ws.Range("A1")
    .PasteSpecial xlPasteValues
    .PasteSpecial xlPasteFormats
    '.PasteSpecial xlPasteValuesAndNumberFormats
  End With
  
  'Load array with column count
  'For use when multiple columns are selected
  iColCount = rSelection.Columns.count
  ReDim vArray(1 To iColCount)
  For i = 1 To iColCount
    vArray(i) = i
  Next i
  
  'Remove duplicates
  ws.UsedRange.RemoveDuplicates Columns:=vArray(i - 1), Header:=xlGuess
  
  'Remove blank cells (optional)
  On Error Resume Next
    ws.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlShiftUp
  On Error GoTo 0
  
  'Autofit column
  ws.Columns("A").AutoFit
  
  'Exit CutCopyMode
  Application.CutCopyMode = False
    
End Sub
