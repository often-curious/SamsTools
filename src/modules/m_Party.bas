Attribute VB_Name = "m_Party"
Option Explicit

' Flashes random visible cells for a few seconds, then restores their fills.

Private Const FLASH_CELLS As Long = 30
Private Const FLASH_CYCLES As Long = 7
Private Const FLASH_SECONDS As Double = 0.5

Private mFrm As frmParty

Public Sub TriggerEasterEgg()
    Dim colors(1 To 3) As Long
    Dim flashCells(1 To FLASH_CELLS) As Range
    Dim origColor(1 To FLASH_CELLS) As Long
    Dim origNoFill(1 To FLASH_CELLS) As Boolean
    Dim lastIdx(1 To FLASH_CELLS) As Long
    Dim seen As Collection
    Dim vis As Range, r As Range
    Dim wb As Workbook
    Dim n As Long, i As Long, cycle As Long, idx As Long, attempt As Long
    Dim wasSaved As Boolean, prevScreenUpdating As Boolean
    Dim errNum As Long, errDesc As String

    ' Guards: nothing to flash on, or nothing we are allowed to write to
    If ActiveWindow Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.ProtectContents Then Exit Sub

    Set wb = ActiveWorkbook
    wasSaved = wb.Saved
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True

    colors(1) = RGB(250, 250, 0)   ' Yellow
    colors(2) = RGB(41, 209, 213)  ' Teal
    colors(3) = RGB(240, 14, 132)  ' Pink

    On Error GoTo CleanUp

    Set mFrm = New frmParty
    mFrm.btnEndParty.Enabled = False
    mFrm.Show vbModeless

    ' Pick up to FLASH_CELLS distinct, visible cells and record their fills
    Set seen = New Collection
    Set vis = ActiveWindow.visibleRange
    Randomize
    For attempt = 1 To FLASH_CELLS * 5
        If n >= FLASH_CELLS Then Exit For
        Set r = vis.Cells(Int(Rnd() * vis.rows.count) + 1, _
                          Int(Rnd() * vis.Columns.count) + 1)
        If Not r.EntireRow.Hidden And Not r.EntireColumn.Hidden Then
            If TryAddKey(seen, r.address) Then
                n = n + 1
                Set flashCells(n) = r
                With r.Interior
                    origNoFill(n) = (.pattern = xlNone)
                    origColor(n) = .color
                End With
            End If
        End If
    Next attempt

    ' Flash; each cell gets a different colour from its previous one
    For cycle = 1 To FLASH_CYCLES
        If Not IsFormLoaded() Then Exit For
        For i = 1 To n
            Do
                idx = Int(Rnd() * 3) + 1
            Loop While idx = lastIdx(i)
            lastIdx(i) = idx
            flashCells(i).Interior.color = colors(idx)
        Next i
        PauseFor FLASH_SECONDS
    Next cycle

CleanUp:
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    For i = 1 To n
        RestoreFill flashCells(i), origColor(i), origNoFill(i)
    Next i

    If IsFormLoaded() Then mFrm.btnEndParty.Enabled = True

    wb.Saved = wasSaved
    Application.ScreenUpdating = prevScreenUpdating

    If errNum <> 0 Then Debug.Print "TriggerEasterEgg error " & errNum & ": " & errDesc
End Sub

' Adds a key to the collection. Returns False if the key already exists.
Private Function TryAddKey(ByVal keys As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    keys.Add True, key
    TryAddKey = (Err.Number = 0)
    Err.Clear
End Function

' Restores one cell's fill. Errors are skipped so one bad cell cannot
' block the rest of the restore.
Private Sub RestoreFill(ByVal r As Range, ByVal clr As Long, ByVal noFill As Boolean)
    On Error Resume Next
    If noFill Then
        r.Interior.pattern = xlNone
    Else
        r.Interior.color = clr
    End If
End Sub

' True while the party form instance is still loaded.
Private Function IsFormLoaded() As Boolean
    Dim f As Object
    If mFrm Is Nothing Then Exit Function
    For Each f In VBA.UserForms
        If f Is mFrm Then
            IsFormLoaded = True
            Exit Function
        End If
    Next f
End Function

' Sub-second pause that keeps Excel responsive. Handles Timer reset at midnight.
Private Sub PauseFor(ByVal seconds As Double)
    Dim t0 As Double
    t0 = Timer
    Do While Timer - t0 < seconds
        If Timer < t0 Then Exit Do
        DoEvents
    Loop
End Sub


