VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmParty 
   Caption         =   "Surprise!"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6210
   OleObjectBlob   =   "frmParty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Dim userName As String
    userName = Application.userName 'Environ("USERNAME") ' Retrieves the Windows username

    lblUsername.caption = userName
    
    Application.Speech.Speak "Congratulations" & userName & ", you found the Easter Egg - enjoy the party!", SpeakAsync:=True
    
End Sub

Private Sub btnEndParty_Click()
    Unload Me
End Sub
