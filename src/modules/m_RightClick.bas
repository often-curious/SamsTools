Attribute VB_Name = "m_RightClick"
Option Explicit

Private Const POPUP_CAPTION As String = "Apply Number Formatting"
Private Const POPUP_TAG As String = "NumFmtPopup_v1"
' Self-contained constants (no Office reference needed)
Private Const msoControlButton As Long = 1
Private Const msoControlPopup As Long = 10

' Right-click CommandBars to target
Private Function TargetBars() As Variant
    TargetBars = Array( _
        "Cell", _
        "List Range Popup", _
        "PivotTable Context Menu", _
        "Chart Area", "Plot Area", _
        "Series", "Axis", _
        "Data Labels", "Legend" _
    )
End Function

' Call on Workbook_Open
Public Sub EnsureNumberFormatMenus()
    Dim oldUpd As Boolean, oldEvt As Boolean, oldDisp As Boolean
    oldUpd = Application.ScreenUpdating
    oldEvt = Application.EnableEvents
    oldDisp = Application.DisplayStatusBar
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

    Dim bars As Variant, i As Long
    bars = TargetBars()
    For i = LBound(bars) To UBound(bars)
        EnsureMenuOnBar CStr(bars(i))
    Next i

    Application.ScreenUpdating = oldUpd
    Application.EnableEvents = oldEvt
    Application.DisplayStatusBar = oldDisp
End Sub

' Call on Workbook_BeforeClose only
Public Sub RemoveAllNumberFormatMenus()
    Dim bars As Variant, i As Long
    bars = TargetBars()
    For i = LBound(bars) To UBound(bars)
        DeleteMenuOnBar CStr(bars(i))
    Next i
End Sub

Private Sub EnsureMenuOnBar(barName As String)
    Dim cb As CommandBar, popupCtrl As CommandBarControl

    On Error Resume Next
    Set cb = Application.CommandBars(barName)
    On Error GoTo 0
    If cb Is Nothing Then Exit Sub

    ' If already present (by Tag), skip
    Set popupCtrl = cb.FindControl(Type:=msoControlPopup, ID:=0, tag:=POPUP_TAG, Visible:=True)
    If Not popupCtrl Is Nothing Then Exit Sub

    ' Create the popup and add buttons; use BeginGroup for visual separators
    Set popupCtrl = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With popupCtrl
        .caption = POPUP_CAPTION
        .tag = POPUP_TAG
    End With

    ' Group 1
    AddFmtButton popupCtrl, "1,000", "#,##0;[Red](#,##0);-", True
    AddFmtButton popupCtrl, "$1,000", "$#,##0;[Red]($#,##0);-"
    AddFmtButton popupCtrl, "$1000.00", "$#,##0.00;[Red]($#,##0.00);-"

    ' Group 2
    AddFmtButton popupCtrl, "$100k", "$#,##0,""k"";[Red]-$#,##0,""k"";-", True
    AddFmtButton popupCtrl, "$10.0m", "$#,##0.0,,""m"";[Red]-$#,##0.0,,""m"";-"

    ' Group 3
    AddFmtButton popupCtrl, "m²", "#,##0""m²"";[Red](#,##0""m²"");-", True
    AddFmtButton popupCtrl, "m³", "#,##0""m³"";[Red](#,##0""m³"");-"
    
    ' Group 4
    AddFmtButton popupCtrl, "0%", "#,##0%;[Red](#,##0%);-", True
    AddFmtButton popupCtrl, "+0% ; -0%", "+#,##0%;[Red]-#,##0%;-"
    
    ' Group 5
    AddFmtButton popupCtrl, "x times", "#,##0x;[Red]-#,##0x;-", True
    AddFmtButton popupCtrl, "Jan-25", "mmm-yy;;-"
    
    ' Group 6
    AddFmtButton popupCtrl, "Clear number format", "General", True
End Sub

Private Sub DeleteMenuOnBar(barName As String)
    Dim cb As CommandBar, ctrl As CommandBarControl
    On Error Resume Next
    Set cb = Application.CommandBars(barName)
    If cb Is Nothing Then Exit Sub
    Set ctrl = cb.FindControl(Type:=msoControlPopup, ID:=0, tag:=POPUP_TAG, Visible:=True)
    If Not ctrl Is Nothing Then ctrl.Delete
    On Error GoTo 0
End Sub

Public Sub ApplyNumberFormat_FromMenu()
    On Error GoTo Done
    Dim ac As CommandBarControl, fmt As String
    Set ac = Application.CommandBars.ActionControl
    If ac Is Nothing Then GoTo Done
    fmt = ac.tag: If Len(fmt) = 0 Then GoTo Done

    Application.ScreenUpdating = False

    If TypeName(Selection) = "Range" Then
        Dim rng As Range: Set rng = Selection
        Dim pc As PivotCell
        On Error Resume Next
        Set pc = rng.Cells(1, 1).PivotCell
        On Error GoTo 0
        If Not pc Is Nothing Then
            If pc.PivotCellType = xlPivotCellValue Or pc.PivotCellType = xlPivotCellSubtotal Then
                If Not pc.PivotField Is Nothing Then
                    pc.PivotField.numberFormat = fmt
                    GoTo Done
                End If
            End If
        End If
        rng.numberFormat = fmt
        GoTo Done
    End If

    If Not ActiveChart Is Nothing Then
        Select Case TypeName(Selection)
            Case "Axis": On Error Resume Next: Selection.TickLabels.numberFormat = fmt: On Error GoTo 0
            Case "DataLabels", "DataLabel": On Error Resume Next: Selection.numberFormat = fmt: On Error GoTo 0
            Case "Series"
                On Error Resume Next
                If Not Selection.HasDataLabels Then Selection.ApplyDataLabels
                Selection.DataLabels.numberFormat = fmt
                On Error GoTo 0
            Case "ChartArea", "PlotArea"
                On Error Resume Next
                ActiveChart.Axes(xlCategory).TickLabels.numberFormat = fmt
                ActiveChart.Axes(xlValue).TickLabels.numberFormat = fmt
                On Error GoTo 0
            Case Else
                On Error Resume Next: Selection.numberFormat = fmt: On Error GoTo 0
        End Select
    End If
Done:
    Application.ScreenUpdating = True
End Sub


' beginGroup := True puts a divider line before the button
Private Sub AddFmtButton(parentPopup As CommandBarControl, caption As String, numberFormat As String, Optional beginGroup As Boolean = False)
    Dim btn As CommandBarButton
    Set btn = parentPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With btn
        .caption = caption
        .Style = 1 ' msoButtonCaption
        .tag = numberFormat
        .OnAction = "ApplyNumberFormat_FromMenu"
        If beginGroup Then .beginGroup = True
    End With
End Sub

