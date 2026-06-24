Attribute VB_Name = "m_RightClick"
Option Explicit

Private Const POPUP_CAPTION As String = "Apply Number Formatting"
Private Const POPUP_TAG As String = "NumFmtPopup_v1"

Private Const msoControlButton As Long = 1
Private Const msoControlPopup As Long = 10
Private Const msoButtonCaption As Long = 2

Private gMenusLoaded As Boolean

Private Function TargetBars() As Variant
    TargetBars = Array( _
        "Cell", _
        "List Range Popup", _
        "PivotTable Context Menu" _
    )
End Function

Public Sub EnsureNumberFormatMenus()
    If gMenusLoaded Then Exit Sub
    gMenusLoaded = True

    Dim oldUpd As Boolean
    Dim oldEvt As Boolean
    Dim oldDisp As Boolean

    oldUpd = Application.ScreenUpdating
    oldEvt = Application.EnableEvents
    oldDisp = Application.DisplayStatusBar

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

    Dim bars As Variant
    Dim i As Long

    bars = TargetBars()

    For i = LBound(bars) To UBound(bars)
        EnsureMenuOnBar CStr(bars(i))
    Next i

CleanUp:
    Application.ScreenUpdating = oldUpd
    Application.EnableEvents = oldEvt
    Application.DisplayStatusBar = oldDisp
End Sub

Public Sub RemoveAllNumberFormatMenus()
    Dim bars As Variant
    Dim i As Long

    bars = TargetBars()

    For i = LBound(bars) To UBound(bars)
        DeleteMenuOnBar CStr(bars(i))
    Next i

    gMenusLoaded = False
End Sub

Private Sub EnsureMenuOnBar(ByVal barName As String)
    Dim cb As CommandBar
    Dim popupCtrl As CommandBarControl

    On Error Resume Next
    Set cb = Application.CommandBars(barName)
    On Error GoTo 0

    If cb Is Nothing Then Exit Sub

    Set popupCtrl = cb.FindControl( _
        Type:=msoControlPopup, _
        ID:=0, _
        tag:=POPUP_TAG, _
        Visible:=True _
    )

    If Not popupCtrl Is Nothing Then Exit Sub

    Set popupCtrl = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)

    With popupCtrl
        .caption = POPUP_CAPTION
        .tag = POPUP_TAG
    End With

    AddFmtButton popupCtrl, "1,000", "#,##0;[Red](#,##0);-", True
    AddFmtButton popupCtrl, "$1,000", "$#,##0;[Red]($#,##0);-"
    AddFmtButton popupCtrl, "$1000.00", "$#,##0.00;[Red]($#,##0.00);-"

    AddFmtButton popupCtrl, "$100k", "$#,##0,""k"";[Red]-$#,##0,""k"";-", True
    AddFmtButton popupCtrl, "$10.0m", "$#,##0.0,,""m"";[Red]-$#,##0.0,,""m"";-"

    AddFmtButton popupCtrl, "m²", "#,##0""m²"";[Red](#,##0""m²"");-", True
    AddFmtButton popupCtrl, "m³", "#,##0""m³"";[Red](#,##0""m³"");-"

    AddFmtButton popupCtrl, "0%", "#,##0%;[Red](#,##0%);-", True
    AddFmtButton popupCtrl, "+0% ; -0%", "+#,##0%;[Red]-#,##0%;-"

    AddFmtButton popupCtrl, "x times", "#,##0x;[Red]-#,##0x;-", True
    AddFmtButton popupCtrl, "Jan-25", "mmm-yy;;-"

    AddFmtButton popupCtrl, "Clear number format", "General", True
End Sub

Private Sub DeleteMenuOnBar(ByVal barName As String)
    Dim cb As CommandBar
    Dim ctrl As CommandBarControl

    On Error Resume Next
    Set cb = Application.CommandBars(barName)

    If cb Is Nothing Then Exit Sub

    Set ctrl = cb.FindControl( _
        Type:=msoControlPopup, _
        ID:=0, _
        tag:=POPUP_TAG, _
        Visible:=True _
    )

    If Not ctrl Is Nothing Then ctrl.Delete

    On Error GoTo 0
End Sub

Public Sub ApplyNumberFormat_FromMenu()
    Dim oldUpd As Boolean
    oldUpd = Application.ScreenUpdating

    On Error GoTo CleanUp

    Dim ac As CommandBarControl
    Dim fmt As String

    Set ac = Application.CommandBars.ActionControl
    If ac Is Nothing Then GoTo CleanUp

    fmt = ac.tag
    If Len(fmt) = 0 Then GoTo CleanUp

    Application.ScreenUpdating = False

    If TypeName(Selection) = "Range" Then
        ApplyFormatToRange Selection, fmt
        GoTo CleanUp
    End If

    If Not ActiveChart Is Nothing Then
        ApplyFormatToChartSelection fmt
    End If

CleanUp:
    Application.ScreenUpdating = oldUpd
End Sub

Private Sub ApplyFormatToRange(ByVal rng As Range, ByVal fmt As String)
    Dim pc As PivotCell

    On Error Resume Next
    Set pc = rng.Cells(1, 1).PivotCell
    On Error GoTo 0

    If Not pc Is Nothing Then
        If pc.PivotCellType = xlPivotCellValue Or pc.PivotCellType = xlPivotCellSubtotal Then
            If Not pc.PivotField Is Nothing Then
                pc.PivotField.numberFormat = fmt
                Exit Sub
            End If
        End If
    End If

    rng.numberFormat = fmt
End Sub

Private Sub ApplyFormatToChartSelection(ByVal fmt As String)
    On Error Resume Next

    Select Case TypeName(Selection)

        Case "Axis"
            Selection.TickLabels.numberFormat = fmt

        Case "DataLabels", "DataLabel"
            Selection.numberFormat = fmt

        Case "Series"
            If Not Selection.HasDataLabels Then Selection.ApplyDataLabels
            Selection.DataLabels.numberFormat = fmt

        Case "ChartArea", "PlotArea"
            ActiveChart.Axes(xlCategory).TickLabels.numberFormat = fmt
            ActiveChart.Axes(xlValue).TickLabels.numberFormat = fmt

        Case Else
            Selection.numberFormat = fmt

    End Select

    On Error GoTo 0
End Sub

Private Sub AddFmtButton( _
    ByVal parentPopup As CommandBarControl, _
    ByVal caption As String, _
    ByVal numberFormat As String, _
    Optional ByVal beginGroup As Boolean = False _
)
    Dim btn As CommandBarButton

    Set btn = parentPopup.Controls.Add(Type:=msoControlButton, Temporary:=True)

    With btn
        .caption = caption
        .Style = msoButtonCaption
        .tag = numberFormat
        .OnAction = "ApplyNumberFormat_FromMenu"
        If beginGroup Then .beginGroup = True
    End With
End Sub



