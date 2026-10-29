VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDependents 
   Caption         =   "Trace Dependents"
   ClientHeight    =   3495
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8970.001
   OleObjectBlob   =   "frmDependents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDependents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Navigation targets, filled in lock-step with the list rows
Private mTargets As Collection

'Form-level dimension constants
Private Const FORM_PADDING As Long = 5

'Column width constants
Private Const COLUMN_WIDTH_ADDRESS As Long = 250
Private Const COLUMN_WIDTH_VALUE As Long = 75
Private Const COLUMN_WIDTH_FORMULA As Long = 200

'Control spacing constants
Private Const CONTROL_SPACING As Long = 15
Private Const HEADER_OFFSET As Long = 12
Private Const SCROLLBAR_WIDTH As Long = 20

Private Const INNER_WIDTH As Long = _
    COLUMN_WIDTH_ADDRESS + _
    COLUMN_WIDTH_VALUE + _
    COLUMN_WIDTH_FORMULA + _
    SCROLLBAR_WIDTH

Private Const FORM_WIDTH As Long = _
    INNER_WIDTH + (4 * FORM_PADDING)

Private Const FORM_HEIGHT As Long = 250

'Formula box constants
Private Const FORMULA_HEIGHT As Long = 20
Private Const FORMULA_TOP As Long = FORM_PADDING
Private Const FORMULA_WIDTH As Long = INNER_WIDTH

'ListBox constants
Private Const LISTBOX_TOP As Long = _
    FORMULA_TOP + FORMULA_HEIGHT + CONTROL_SPACING

Private Const LISTBOX_WIDTH As Long = INNER_WIDTH

Private Const LISTBOX_HEIGHT As Long = _
    FORM_HEIGHT - LISTBOX_TOP - FORM_PADDING


'==========================================================
' COLUMN WIDTHS
'==========================================================

Public Function GetColumnWidths() As String

    GetColumnWidths = _
        COLUMN_WIDTH_ADDRESS & ";" & _
        COLUMN_WIDTH_VALUE & ";" & _
        COLUMN_WIDTH_FORMULA

End Function


'==========================================================
' NAVIGATION TARGETS
'==========================================================

Public Sub AddTarget(ByVal target As Range)

    If target Is Nothing Then Exit Sub

    If mTargets Is Nothing Then
        Set mTargets = New Collection
    End If

    mTargets.Add target

End Sub


'==========================================================
' HEADERS
'==========================================================

Public Sub AddHeaders()

    Dim headerLabel1 As MSForms.Label
    Dim headerLabel2 As MSForms.Label
    Dim headerLabel3 As MSForms.Label

    'Prevent duplicate headers if method is called again
    On Error Resume Next

    Me.Controls.Remove "lblHeader1"
    Me.Controls.Remove "lblHeader2"
    Me.Controls.Remove "lblHeader3"

    On Error GoTo 0

    Set headerLabel1 = _
        Me.Controls.Add("Forms.Label.1", "lblHeader1")

    With headerLabel1

        .Top = lstDependents.Top - HEADER_OFFSET
        .Left = lstDependents.Left + FORM_PADDING
        .caption = "Address"
        .width = COLUMN_WIDTH_ADDRESS

    End With


    Set headerLabel2 = _
        Me.Controls.Add("Forms.Label.1", "lblHeader2")

    With headerLabel2

        .Top = lstDependents.Top - HEADER_OFFSET

        .Left = _
            lstDependents.Left + _
            COLUMN_WIDTH_ADDRESS + _
            FORM_PADDING

        .caption = "Value"
        .width = COLUMN_WIDTH_VALUE

    End With


    Set headerLabel3 = _
        Me.Controls.Add("Forms.Label.1", "lblHeader3")

    With headerLabel3

        .Top = lstDependents.Top - HEADER_OFFSET

        .Left = _
            lstDependents.Left + _
            COLUMN_WIDTH_ADDRESS + _
            COLUMN_WIDTH_VALUE + _
            FORM_PADDING

        .caption = "Formula"
        .width = COLUMN_WIDTH_FORMULA

    End With

End Sub


'==========================================================
' INITIALISE FORM
'==========================================================

Private Sub UserForm_Initialize()

    Set mTargets = New Collection

    'Set form caption and size
    Me.caption = "Trace Dependents"
    Me.width = FORM_WIDTH
    Me.Height = FORM_HEIGHT + (4 * FORM_PADDING) - 2


    'Create formula text box
    With Me.Controls.Add("Forms.TextBox.1", "txtFormula")

        .Top = FORMULA_TOP
        .Left = FORM_PADDING
        .width = FORMULA_WIDTH
        .Height = FORMULA_HEIGHT

        .BackColor = RGB(240, 240, 240)

        .Locked = True
        .MultiLine = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsHorizontal

        .Font.Size = 10

    End With


    'Configure list box
    With lstDependents

        .Top = LISTBOX_TOP
        .Left = FORM_PADDING
        .width = LISTBOX_WIDTH
        .Height = LISTBOX_HEIGHT

        .ColumnCount = 3
        .ColumnWidths = GetColumnWidths()

        .Font.Size = 10

    End With

End Sub


'==========================================================
' LISTBOX NAVIGATION
'==========================================================

Private Sub lstDependents_Click()

    Dim i As Long
    Dim target As Range

    i = lstDependents.ListIndex

    If i < 0 Then Exit Sub
    If mTargets Is Nothing Then Exit Sub

    'ListBox is zero-based.
    'Collection is one-based.
    If i + 1 > mTargets.count Then Exit Sub

    Set target = mTargets(i + 1)

    If target Is Nothing Then Exit Sub

    'Navigate directly to the stored Range.
    'No address parsing is required.
    On Error GoTo NavigationError

    Application.GoTo _
        Reference:=target, _
        Scroll:=True

    'Update formula display
    If target.CountLarge > 1 Then

        Me.Controls("txtFormula").text = _
            CStr(target.Cells(1, 1).formula)

    Else

        Me.Controls("txtFormula").text = _
            CStr(target.formula)

    End If

    Exit Sub


NavigationError:

    MsgBox _
        "Excel could not navigate to this dependent.", _
        vbExclamation, _
        "Trace Dependents"

End Sub


'==========================================================
' KEYBOARD
'==========================================================

Private Sub lstDependents_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

