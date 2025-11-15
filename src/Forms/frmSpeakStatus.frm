VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpeakStatus 
   Caption         =   "Speak Selection"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3930
   OleObjectBlob   =   "frmSpeakStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpeakStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isCancelled As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = isCancelled
End Property

Private Sub btnCancel_Click()
    ' Send ESC key to attempt stopping speech
    Application.SendKeys "{ESC}"
    Me.Hide
End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub UserForm_Initialize()
    isCancelled = False
End Sub

