Attribute VB_Name = "m_TraceLinks"
Option Explicit

Private Const INDENT_CHAR As String = "  "
Private Const MAX_TRACE_ARROWS As Long = 1000
Private Const MAX_TRACE_LINKS As Long = 1000


'==========================================================
' PUBLIC RIBBON CALLBACKS
'==========================================================

Public Sub ShowTracePrecedents(control As IRibbonControl)

    Dim selectedRange As Range
    Dim precedents As Collection
    Dim cell As Range
    Dim pr As Range
    Dim frm As frmPrecedents

    'Ensure user has selected a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range.", vbInformation
        Exit Sub
    End If

    Set selectedRange = Selection

    'Create form
    Set frm = New frmPrecedents

    'Clear and configure list
    frm.lstPrecedents.Clear
    frm.lstPrecedents.ColumnCount = 3
    frm.lstPrecedents.ColumnWidths = frm.GetColumnWidths()

    'Add headers
    frm.AddHeaders

    'Display formula of first selected cell
    frm.Controls("txtFormula").text = _
        GetFormulaAsString(selectedRange)

    '--------------------------------------------------
    ' Add selected cells
    '--------------------------------------------------
    For Each cell In selectedRange.Cells

        frm.lstPrecedents.AddItem

        With frm.lstPrecedents

            .List(.ListCount - 1, 0) = _
                displayAddress(cell, selectedRange.Cells(1, 1))

            .List(.ListCount - 1, 1) = _
                GetCellValueAsString(cell)

            .List(.ListCount - 1, 2) = _
                GetFormulaAsString(cell)

        End With

        'Keep target collection aligned with ListBox
        frm.AddTarget cell

    Next cell

    '--------------------------------------------------
    ' Find precedents
    '--------------------------------------------------
    Set precedents = FindRelatedCells( _
        inRange:=selectedRange, _
        towardPrecedent:=True)

    '--------------------------------------------------
    ' Add precedents
    '--------------------------------------------------
    For Each pr In precedents

        frm.lstPrecedents.AddItem

        With frm.lstPrecedents

            .List(.ListCount - 1, 0) = _
                INDENT_CHAR & _
                displayAddress(pr, selectedRange.Cells(1, 1))

            .List(.ListCount - 1, 1) = _
                GetCellValueAsString(pr)

            .List(.ListCount - 1, 2) = _
                GetFormulaAsString(pr)

        End With

        'Keep target collection aligned with ListBox
        frm.AddTarget pr

    Next pr

    'Select first row
    If frm.lstPrecedents.ListCount > 0 Then
        frm.lstPrecedents.ListIndex = 0
    End If

    'Position form
    frm.StartUpPosition = 0

    frm.Left = Application.Left + _
               (Application.width - frm.width) / 2

    frm.Top = Application.Top + _
              (Application.Height - frm.Height) / 2

    frm.Show vbModeless

End Sub


Public Sub ShowTraceDependents(control As IRibbonControl)

    Dim selectedRange As Range
    Dim dependents As Collection
    Dim cell As Range
    Dim dependentCell As Range
    Dim frm As frmDependents

    If Not TryGetSelectedRange(selectedRange) Then Exit Sub

    Set dependents = FindRelatedCells( _
        inRange:=selectedRange, _
        towardPrecedent:=False)

    Set frm = New frmDependents

    With frm

        .lstDependents.Clear
        .lstDependents.ColumnCount = 3
        .lstDependents.ColumnWidths = .GetColumnWidths()

        .AddHeaders

        .Controls("txtFormula").text = _
            GetFormulaAsString(selectedRange)

        'Add the selected cells first
        For Each cell In selectedRange.Cells

            AddListBoxRow _
                listBox:=.lstDependents, _
                addressText:=displayAddress( _
                    cell, selectedRange.Cells(1, 1)), _
                valueText:=GetCellValueAsString(cell), _
                formulaText:=GetFormulaAsString(cell)

            .AddTarget cell

        Next cell

        'Add the traced dependents
        For Each dependentCell In dependents

            AddListBoxRow _
                listBox:=.lstDependents, _
                addressText:=INDENT_CHAR & displayAddress( _
                    dependentCell, selectedRange.Cells(1, 1)), _
                valueText:=GetCellValueAsString(dependentCell), _
                formulaText:=GetFormulaAsString(dependentCell)

            .AddTarget dependentCell

        Next dependentCell

        'Optional status row where no dependents were found
        If dependents.count = 0 Then

            AddListBoxRow _
                listBox:=.lstDependents, _
                addressText:=INDENT_CHAR & _
                    "No traceable dependents found", _
                valueText:=vbNullString, _
                formulaText:=vbNullString

        End If

        If .lstDependents.ListCount > 0 Then
            .lstDependents.ListIndex = 0
        End If

        CentreForm frm
        .Show vbModeless

    End With

End Sub


'==========================================================
' SHARED TRACING ENGINE
'==========================================================

Private Function FindRelatedCells( _
    ByVal inRange As Range, _
    ByVal towardPrecedent As Boolean) As Collection

    Dim results As New Collection
    Dim sourceCell As Range

    Dim returnWorkbook As Workbook
    Dim returnWorksheet As Worksheet
    Dim returnSelection As Range

    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean

    On Error GoTo CleanFail

    Set returnWorkbook = ActiveWorkbook
    Set returnWorksheet = ActiveSheet
    Set returnSelection = Selection

    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Trace each selected cell separately.
    'ShowPrecedents and ShowDependents are more reliable when
    'called against a single cell rather than a multi-cell range.
    For Each sourceCell In inRange.Cells

        TraceOneCell _
            sourceCell:=sourceCell, _
            towardPrecedent:=towardPrecedent, _
            results:=results

    Next sourceCell

CleanExit:

    On Error Resume Next

    ClearAllTraceArrows inRange

    If Not returnWorkbook Is Nothing Then
        returnWorkbook.Activate
    End If

    If Not returnWorksheet Is Nothing Then
        returnWorksheet.Activate
    End If

    If Not returnSelection Is Nothing Then
        returnSelection.Select
    End If

    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating

    On Error GoTo 0

    Set FindRelatedCells = results
    Exit Function

CleanFail:

    'Return any relationships found before the unexpected error.
    Resume CleanExit

End Function


Private Sub TraceOneCell( _
    ByVal sourceCell As Range, _
    ByVal towardPrecedent As Boolean, _
    ByRef results As Collection)

    Dim arrowNumber As Long
    Dim linkNumber As Long

    Dim destination As Range
    Dim firstDestination As Range

    Dim showSucceeded As Boolean
    Dim navigateSucceeded As Boolean

    On Error GoTo SafeExit

    sourceCell.Parent.Parent.Activate
    sourceCell.Parent.Activate
    sourceCell.Select

    sourceCell.Parent.ClearArrows

    'ShowPrecedents or ShowDependents raises error 1004 when
    'Excel cannot find a traceable relationship.
    On Error Resume Next
    Err.Clear

    If towardPrecedent Then
        sourceCell.ShowPrecedents
    Else
        sourceCell.ShowDependents
    End If

    showSucceeded = (Err.Number = 0)

    Err.Clear
    On Error GoTo SafeExit

    If Not showSucceeded Then GoTo SafeExit

    For arrowNumber = 1 To MAX_TRACE_ARROWS

        Set destination = Nothing
        Set firstDestination = Nothing

        sourceCell.Parent.Parent.Activate
        sourceCell.Parent.Activate
        sourceCell.Select

        On Error Resume Next
        Err.Clear

        sourceCell.NavigateArrow _
            towardPrecedent:=towardPrecedent, _
            arrowNumber:=arrowNumber

        navigateSucceeded = (Err.Number = 0)

        If navigateSucceeded Then
            Set destination = ActiveCell
        End If

        Err.Clear
        On Error GoTo SafeExit

        'No arrow at this number means all arrows have been read.
        If Not navigateSucceeded Then Exit For

        Set firstDestination = destination

        If Not destination Is Nothing Then

            If Not SameCell(destination, sourceCell) Then
                AddUniqueRange results, destination
            End If

        End If

        'If NavigateArrow moved to another worksheet, the arrow
        'may represent more than one external link destination.
        If Not destination Is Nothing Then

            If Not destination.Parent Is sourceCell.Parent Then

                For linkNumber = 1 To MAX_TRACE_LINKS

                    Set destination = Nothing

                    sourceCell.Parent.Parent.Activate
                    sourceCell.Parent.Activate
                    sourceCell.Select

                    On Error Resume Next
                    Err.Clear

                    sourceCell.NavigateArrow _
                        towardPrecedent:=towardPrecedent, _
                        arrowNumber:=arrowNumber, _
                        linkNumber:=linkNumber

                    navigateSucceeded = (Err.Number = 0)

                    If navigateSucceeded Then
                        Set destination = ActiveCell
                    End If

                    Err.Clear
                    On Error GoTo SafeExit

                    If Not navigateSucceeded Then Exit For

                    If Not destination Is Nothing Then

                        If Not SameCell(destination, sourceCell) Then
                            AddUniqueRange results, destination
                        End If

                    End If

                Next linkNumber

            End If

        End If

    Next arrowNumber

SafeExit:

    On Error Resume Next

    sourceCell.Parent.ClearArrows
    sourceCell.Parent.Parent.Activate
    sourceCell.Parent.Activate
    sourceCell.Select

    On Error GoTo 0

End Sub


'==========================================================
' COLLECTION AND COMPARISON HELPERS
'==========================================================

Private Sub AddUniqueRange( _
    ByRef results As Collection, _
    ByVal target As Range)

    Dim itemKey As String

    If target Is Nothing Then Exit Sub

    itemKey = RangeKey(target)

    On Error Resume Next

    'Using the address as a Collection key prevents duplicates.
    results.Add item:=target, key:=itemKey

    Err.Clear
    On Error GoTo 0

End Sub


Private Function RangeKey(ByVal target As Range) As String

    RangeKey = target.Parent.Parent.Name & "|" & _
               target.Parent.CodeName & "|" & _
               target.address( _
                    RowAbsolute:=True, _
                    ColumnAbsolute:=True, _
                    ReferenceStyle:=xlA1, _
                    External:=False)

End Function


Private Function SameCell( _
    ByVal firstCell As Range, _
    ByVal secondCell As Range) As Boolean

    If firstCell Is Nothing Then Exit Function
    If secondCell Is Nothing Then Exit Function

    If Not firstCell.Parent.Parent Is secondCell.Parent.Parent Then
        Exit Function
    End If

    If Not firstCell.Parent Is secondCell.Parent Then
        Exit Function
    End If

    SameCell = _
        (firstCell.address = secondCell.address)

End Function


Private Sub ClearAllTraceArrows(ByVal inRange As Range)

    Dim worksheetItem As Worksheet
    Dim workbookItem As Workbook

    On Error Resume Next

    Set workbookItem = inRange.Parent.Parent

    For Each worksheetItem In workbookItem.Worksheets
        worksheetItem.ClearArrows
    Next worksheetItem

    On Error GoTo 0

End Sub


'==========================================================
' ADDRESS HELPERS
'==========================================================

Private Function fullAddress(ByVal inCell As Range) As String

    If inCell Is Nothing Then Exit Function

    fullAddress = inCell.address( _
        RowAbsolute:=True, _
        ColumnAbsolute:=True, _
        ReferenceStyle:=xlA1, _
        External:=True)

End Function


Private Function displayAddress( _
    ByVal inCell As Range, _
    ByVal baseCell As Range) As String

    Dim addressText As String
    Dim workbookName As String

    If inCell Is Nothing Then Exit Function
    If baseCell Is Nothing Then Exit Function

    addressText = inCell.address( _
        RowAbsolute:=True, _
        ColumnAbsolute:=True, _
        ReferenceStyle:=xlA1, _
        External:=True)

    workbookName = inCell.Parent.Parent.Name

    'Remove the workbook portion when both cells are
    'in the same workbook.
    If inCell.Parent.Parent Is baseCell.Parent.Parent Then

        addressText = Replace( _
            Expression:=addressText, _
            Find:="[" & workbookName & "]", _
            Replace:=vbNullString, _
            Compare:=vbTextCompare)

    End If

    displayAddress = addressText

End Function


'==========================================================
' VALUE AND FORMULA HELPERS
'==========================================================

Private Function GetCellValueAsString( _
    ByVal cell As Range) As String

    Dim cellValue As Variant

    On Error GoTo ValueError

    If cell Is Nothing Then Exit Function

    If cell.CountLarge > 1 Then
        GetCellValueAsString = "{range}"
        Exit Function
    End If

    cellValue = cell.value

    If IsError(cellValue) Then

        GetCellValueAsString = CellErrorAsString(cellValue)

    ElseIf IsEmpty(cellValue) Then

        GetCellValueAsString = vbNullString

    Else

        GetCellValueAsString = CStr(cellValue)

    End If

    Exit Function

ValueError:
    GetCellValueAsString = "#ERROR"

End Function


Private Function GetFormulaAsString( _
    ByVal cell As Range) As String

    Dim formulaValue As Variant

    On Error GoTo FormulaError

    If cell Is Nothing Then Exit Function

    If cell.CountLarge > 1 Then
        formulaValue = cell.Cells(1, 1).formula
    Else
        formulaValue = cell.formula
    End If

    If IsError(formulaValue) Then
        GetFormulaAsString = CellErrorAsString(formulaValue)
    Else
        GetFormulaAsString = CStr(formulaValue)
    End If

    Exit Function

FormulaError:
    GetFormulaAsString = "#ERROR"

End Function


Private Function CellErrorAsString( _
    ByVal errorValue As Variant) As String

    Select Case CLng(errorValue)

        Case xlErrDiv0
            CellErrorAsString = "#DIV/0!"

        Case xlErrNA
            CellErrorAsString = "#N/A"

        Case xlErrName
            CellErrorAsString = "#NAME?"

        Case xlErrNull
            CellErrorAsString = "#NULL!"

        Case xlErrNum
            CellErrorAsString = "#NUM!"

        Case xlErrRef
            CellErrorAsString = "#REF!"

        Case xlErrValue
            CellErrorAsString = "#VALUE!"

        Case Else
            CellErrorAsString = "#ERROR"

    End Select

End Function


'==========================================================
' FORM AND SELECTION HELPERS
'==========================================================

Private Function TryGetSelectedRange( _
    ByRef selectedRange As Range) As Boolean

    If Not TypeOf Selection Is Excel.Range Then

        MsgBox _
            Prompt:="Please select a cell or range.", _
            Buttons:=vbInformation, _
            Title:="Trace formulas"

        Exit Function

    End If

    Set selectedRange = Selection

    If selectedRange Is Nothing Then Exit Function

    TryGetSelectedRange = True

End Function


Private Sub AddListBoxRow( _
    ByVal listBox As Object, _
    ByVal addressText As String, _
    ByVal valueText As String, _
    ByVal formulaText As String)

    Dim rowNumber As Long

    listBox.AddItem
    rowNumber = listBox.ListCount - 1

    listBox.List(rowNumber, 0) = addressText
    listBox.List(rowNumber, 1) = valueText
    listBox.List(rowNumber, 2) = formulaText

End Sub


Private Sub CentreForm(ByVal formObject As Object)

    With formObject

        .StartUpPosition = 0

        .Left = Application.Left + _
                (Application.width - .width) / 2

        .Top = Application.Top + _
               (Application.Height - .Height) / 2

    End With

End Sub

