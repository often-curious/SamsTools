Attribute VB_Name = "m_SwapPositions"
Sub Swap_Two_Selected_Items(control As IRibbonControl)
    Dim selType As String
    Dim swapped As Boolean
    
    On Error GoTo FailSafe
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    selType = TypeName(Selection)
    
    '======== CASE 1: RANGES / CELLS ========
    If selType = "Range" Then
        swapped = SwapTwoRanges_NoSheetDelete(Selection)
        GoTo Done
    End If
    
    '======== CASE 2: SHAPES (incl. Charts as shapes) ========
    If selType = "ShapeRange" Or selType = "DrawingObjects" Or selType = "GroupObject" Then
        swapped = SwapTwoShapes()
        GoTo Done
    End If
    
    '======== CASE 3: Chart part selected (e.g., axis, plot area) ========
    If Not ActiveChart Is Nothing Then
        MsgBox "Please select the chart OBJECTS (Shift+Click their borders) so two are highlighted, then run again.", vbInformation, "Select Two Charts"
        GoTo Done
    End If
    
    MsgBox "Select exactly two cells/ranges OR exactly two shapes/charts, then run again.", vbExclamation, "Need Exactly Two Items"
    GoTo Done

Done:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    'If swapped Then MsgBox "Swapped the positions of the two selected items.", vbInformation
    Exit Sub

FailSafe:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Couldn't complete the swap. Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

'===============================================================
' Swap two ranges while preserving exact formulas and formatting.
' Formulas are transferred literally (no reference shifts).
'===============================================================
Private Function SwapTwoRanges_NoSheetDelete(ByVal sel As Range) As Boolean
    Dim r1 As Range, r2 As Range
    Dim v1 As Variant, v2 As Variant
    Dim f1 As Variant, f2 As Variant
    
    ' Identify the two targets
    If sel.Areas.count = 2 Then
        Set r1 = sel.Areas(1)
        Set r2 = sel.Areas(2)
    ElseIf sel.Areas.count = 1 And sel.CountLarge = 2 Then
        Set r1 = sel.Cells(1, 1)
        Set r2 = sel.Cells(2, 1)
    Else
        MsgBox "For ranges: select exactly two non-contiguous areas, or exactly two cells.", vbExclamation, "Need Two Ranges/Cells"
        Exit Function
    End If
    
    ' Require same size
    If r1.rows.count <> r2.rows.count Or r1.Columns.count <> r2.Columns.count Then
        MsgBox "The two ranges must be the same size to swap. Adjust selection and try again.", vbExclamation, "Mismatched Sizes"
        Exit Function
    End If
    
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' --- Capture both formulas and values ---
    f1 = r1.formula
    f2 = r2.formula
    v1 = r1.value
    v2 = r2.value
    
    ' --- Swap content (formulas copied literally) ---
    ' Use Formula first so Excel recognizes formula syntax
    r1.formula = f2
    r2.formula = f1
    
    ' --- Swap formats separately ---
    r1.Copy
    r2.PasteSpecial xlPasteFormats
    r2.Copy
    r1.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    SwapTwoRanges_NoSheetDelete = True
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Function

CleanFail:
    MsgBox "Swap failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Function


'===============================================================
' Swap two shapes' positions (Left/Top). Works for charts as shapes.
'===============================================================
Private Function SwapTwoShapes() As Boolean
    Dim sr As ShapeRange
    Dim sh1 As Shape, sh2 As Shape
    Dim t As Double, l As Double
    
    On Error Resume Next
    Set sr = Selection.ShapeRange
    On Error GoTo 0
    
    If sr Is Nothing Then
        MsgBox "For shapes/charts: multi-select exactly two shapes or chart objects, then run again.", vbExclamation, "Need Two Shapes"
        Exit Function
    End If
    
    If sr.count <> 2 Then
        MsgBox "Select exactly two shapes/charts (no more, no less).", vbExclamation, "Need Exactly Two"
        Exit Function
    End If
    
    Set sh1 = sr(1)
    Set sh2 = sr(2)
    
    t = sh1.Top: l = sh1.Left
    sh1.Top = sh2.Top: sh1.Left = sh2.Left
    sh2.Top = t:       sh2.Left = l
    
    SwapTwoShapes = True
End Function

