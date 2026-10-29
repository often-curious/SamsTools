VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShortcuts 
   Caption         =   "Sam's Tools - Keyboard Shortcuts"
   ClientHeight    =   3876
   ClientLeft      =   15
   ClientTop       =   75
   ClientWidth     =   2775
   OleObjectBlob   =   "frmShortcuts.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstShortcuts_Click()

End Sub

Private Sub UserForm_Initialize()

    Me.StartUpPosition = 2
    Me.Height = 410
    Me.width = 355

    Dim item
    
    lstShortcuts.ColumnCount = 2
    lstShortcuts.ColumnWidths = "100 pt;200 pt"

    For Each item In ShortcutList

        lstShortcuts.AddItem
        lstShortcuts.List(lstShortcuts.ListCount - 1, 0) = item(0)
        lstShortcuts.List(lstShortcuts.ListCount - 1, 1) = item(1)

    Next item

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub
