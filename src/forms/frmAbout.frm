VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "Sam's Tools"
   ClientHeight    =   4008
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5010
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const SamsToolsVersion As String = "2.0"

Private Sub btnOK_Click()
    Unload Me
End Sub


Private Sub btnParty_Click()
    Unload Me
    Call TriggerEasterEgg
End Sub

Private Sub UserForm_Initialize()
    lblVersionNumber.caption = SamsToolsVersion
End Sub
