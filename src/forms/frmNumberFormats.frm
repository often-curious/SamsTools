VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNumberFormats 
   Caption         =   "Set Custom Value Formats"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12345
   OleObjectBlob   =   "frmNumberFormats.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNumberFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    txtFormat1 = GetSetting("SamsTools", "FormatToggle", "Format1", "")
    txtFormat2 = GetSetting("SamsTools", "FormatToggle", "Format2", "")
    txtFormat3 = GetSetting("SamsTools", "FormatToggle", "Format3", "")
    txtFormat4 = GetSetting("SamsTools", "FormatToggle", "Format4", "")
    txtFormat5 = GetSetting("SamsTools", "FormatToggle", "Format5", "")
    txtFormat6 = GetSetting("SamsTools", "FormatToggle", "Format6", "")

    UpdatePreview txtFormat1, lblPreview1
    UpdatePreview txtFormat2, lblPreview2
    UpdatePreview txtFormat3, lblPreview3
    UpdatePreview txtFormat4, lblPreview4
    UpdatePreview txtFormat5, lblPreview5
    UpdatePreview txtFormat6, lblPreview6

End Sub

Private Sub cmdSave_Click()

    SaveSetting "SamsTools", "FormatToggle", "Format1", txtFormat1.value
    SaveSetting "SamsTools", "FormatToggle", "Format2", txtFormat2.value
    SaveSetting "SamsTools", "FormatToggle", "Format3", txtFormat3.value
    SaveSetting "SamsTools", "FormatToggle", "Format4", txtFormat4.value
    SaveSetting "SamsTools", "FormatToggle", "Format5", txtFormat5.value
    SaveSetting "SamsTools", "FormatToggle", "Format6", txtFormat6.value

    MsgBox "Formats saved successfully.", vbInformation

    Unload Me

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub UpdatePreview(txt As MSForms.TextBox, lblPreview As MSForms.Label)

    Dim c As Range
    Dim PositiveText As String
    Dim ZeroText As String
    Dim NegativeText As String
    Dim TextText As String

    Set c = ThisWorkbook.Worksheets("Hidden").Range("A1")

    On Error GoTo InvalidFormat

    c.numberFormat = txt.text

    c.value = 12345.678
    PositiveText = c.text

    c.value = 0
    ZeroText = c.text

    c.value = -12345.678
    NegativeText = c.text
    
    c.value = "Text"
    TextText = c.text

    lblPreview.caption = _
        PositiveText & " ; " & _
        NegativeText & " ; " & _
        ZeroText & " ; " & _
        TextText

    Exit Sub

InvalidFormat:
    lblPreview.caption = "Invalid format"

End Sub

Private Sub txtFormat1_Change()
    UpdatePreview txtFormat1, lblPreview1
End Sub

Private Sub txtFormat2_Change()
    UpdatePreview txtFormat2, lblPreview2
End Sub

Private Sub txtFormat3_Change()
    UpdatePreview txtFormat3, lblPreview3
End Sub

Private Sub txtFormat4_Change()
    UpdatePreview txtFormat4, lblPreview4
End Sub

Private Sub txtFormat5_Change()
    UpdatePreview txtFormat5, lblPreview5
End Sub

Private Sub txtFormat6_Change()
    UpdatePreview txtFormat6, lblPreview6
End Sub
