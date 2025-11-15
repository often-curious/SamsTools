Attribute VB_Name = "m_UndoManager"
Public UndoValues As Object ' Dictionary

Sub SaveUndoState()
    Dim cell As Range
    Set UndoValues = CreateObject("Scripting.Dictionary")

    For Each cell In Selection
        If Not cell.HasFormula Then
            UndoValues(cell.Address) = cell.value
        End If
    Next cell
End Sub

Sub UndoLastChange()
    Dim addr As Variant
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ' Check if UndoValues is initialized and has items
    If UndoValues Is Nothing Then
        MsgBox "Nothing to undo.", vbExclamation
        GoTo CleanExit
    ElseIf UndoValues.count = 0 Then
        MsgBox "Nothing to undo.", vbExclamation
        GoTo CleanExit
    End If

    ' Undo changes
    For Each addr In UndoValues.keys
        Range(addr).value = UndoValues(addr)
    Next addr

    MsgBox "Changes undone.", vbInformation

CleanExit:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    Set UndoValues = Nothing
End Sub

