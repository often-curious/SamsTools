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
Option Explicit

Private Sub UserForm_Initialize()
    Dim who As String

    who = Trim$(Application.userName)
    If Len(who) = 0 Then who = Environ$("USERNAME")
    lblUsername.caption = who

    SpeakSafe "Congratulations" & IIf(Len(who) > 0, " " & who, "") & _
              ", you found the Easter Egg - enjoy the party!"
End Sub

Private Sub btnEndParty_Click()
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    StopSpeech
End Sub

' Speech is unavailable on Mac and some locked-down PCs, so never let it
' block the form from opening.
Private Sub SpeakSafe(ByVal msg As String)
    On Error Resume Next
    Application.Speech.Speak msg, SpeakAsync:=True
End Sub

Private Sub StopSpeech()
    On Error Resume Next
    Application.Speech.Speak "", SpeakAsync:=True, Purge:=True
End Sub

