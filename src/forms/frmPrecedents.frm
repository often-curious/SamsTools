VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrecedents 
   Caption         =   "Trace Precedents"
   ClientHeight    =   4425
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10260
   OleObjectBlob   =   "frmPrecedents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrecedents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Navigation targets, filled in lock-step with the list rows
Private mTargets As Collection

' Form-level dimension constants
Private Const FORM_PADDING As Long = 5        ' Consistent padding around all edges

' Column width constants
Private Const COLUMN_WIDTH_ADDRESS As Long = 250
Private Const COLUMN_WIDTH_VALUE As Long = 75
Private Const COLUMN_WIDTH_FORMULA As Long = 200

' Control spacing constants
Private Const CONTROL_SPACING As Long = 15     ' Vertical space between controls
Private Const HEADER_OFFSET As Long = 12       ' Space between headers and listbox
Private Const SCROLLBAR_WIDTH As Long = 20     ' Width of vertical scrollbar

Private Const INNER_WIDTH As Long = COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + COLUMN_WIDTH_FORMULA + SCROLLBAR_WIDTH  ' Width of content
Private Const FORM_WIDTH As Long = INNER_WIDTH + (4 * FORM_PADDING)  ' Total form width including padding
Private Const FORM_HEIGHT As Long = 250       ' Total form height

' Formula box constants
Private Const FORMULA_HEIGHT As Long = 20
Private Const FORMULA_TOP As Long = FORM_PADDING
Private Const FORMULA_WIDTH As Long = INNER_WIDTH  ' Width matches the listbox content

' ListBox constants
Private Const LISTBOX_TOP As Long = FORMULA_TOP + FORMULA_HEIGHT + CONTROL_SPACING
Private Const LISTBOX_WIDTH As Long = INNER_WIDTH
Private Const LISTBOX_HEIGHT As Long = FORM_HEIGHT - LISTBOX_TOP - FORM_PADDING  ' Account for bottom padding


' Function to get column widths string
Public Function GetColumnWidths() As String
    GetColumnWidths = COLUMN_WIDTH_ADDRESS & ";" & _
                     COLUMN_WIDTH_VALUE & ";" & _
                     COLUMN_WIDTH_FORMULA
End Function

' Store a navigation target for the row just added to the list
Public Sub AddTarget(ByVal target As Range)
    If mTargets Is Nothing Then Set mTargets = New Collection
    mTargets.Add target
End Sub

Public Sub AddHeaders()
    With lstPrecedents
        ' Add headers using a Label control for each column
        Dim headerLabel1 As MSForms.Label
        Set headerLabel1 = Me.Controls.Add("Forms.Label.1", "lblHeader1")
        With headerLabel1
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + FORM_PADDING
            .caption = "Address"
            .width = COLUMN_WIDTH_ADDRESS
        End With

        Dim headerLabel2 As MSForms.Label
        Set headerLabel2 = Me.Controls.Add("Forms.Label.1", "lblHeader2")
        With headerLabel2
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + COLUMN_WIDTH_ADDRESS + FORM_PADDING
            .caption = "Value"
            .width = COLUMN_WIDTH_VALUE
        End With

        Dim headerLabel3 As MSForms.Label
        Set headerLabel3 = Me.Controls.Add("Forms.Label.1", "lblHeader3")
        With headerLabel3
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + FORM_PADDING
            .caption = "Formula"
            .width = COLUMN_WIDTH_FORMULA
        End With
    End With
End Sub

Private Sub UserForm_Initialize()
    Set mTargets = New Collection

    ' Set form caption and size
    Me.caption = "Trace Precedents"
    Me.width = FORM_WIDTH
    Me.Height = FORM_HEIGHT + (4 * FORM_PADDING) - 2

    ' Initialize the formula text box
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

    ' Initialize the list box with adjusted positioning
    With lstPrecedents
        .Top = LISTBOX_TOP
        .Left = FORM_PADDING
        .width = LISTBOX_WIDTH
        .Height = LISTBOX_HEIGHT
        .ColumnCount = 3
        .ColumnWidths = GetColumnWidths()
        .Font.Size = 10
    End With
End Sub

Private Sub lstPrecedents_Click()
    Dim i As Long
    i = lstPrecedents.ListIndex
    If i < 0 Then Exit Sub
    If mTargets Is Nothing Then Exit Sub
    If i + 1 > mTargets.count Then Exit Sub

    Dim target As Range
    Set target = mTargets(i + 1)
    If target Is Nothing Then Exit Sub

    Application.GoTo Reference:=target, Scroll:=True

    If target.count > 1 Then
        Me.Controls("txtFormula").text = CStr(target.Cells(1, 1).formula)
    Else
        Me.Controls("txtFormula").text = CStr(target.formula)
    End If
End Sub

Private Sub lstPrecedents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

