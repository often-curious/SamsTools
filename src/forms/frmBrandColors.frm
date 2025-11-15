VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBrandColors 
   Caption         =   "Brand Colours"
   ClientHeight    =   4020
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "frmBrandColors.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBrandColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FillColors(1 To 3) As Long
Dim TextColors(1 To 3) As Long

Private Sub UserForm_Initialize()
    Dim i As Integer
    On Error Resume Next
    For i = 1 To 3
        FillColors(i) = Evaluate(ThisWorkbook.Names("BrandFillColor" & i).RefersTo)
        TextColors(i) = Evaluate(ThisWorkbook.Names("BrandTextColor" & i).RefersTo)
        Me.Controls("lblPreviewFill" & i).BackColor = FillColors(i)
        Me.Controls("lblPreviewText" & i).BackColor = TextColors(i)
    Next i
    On Error GoTo 0
End Sub


Private Sub cmdPickFill1_Click()
    Dim newColor As Long
    newColor = GetColor()
    If newColor <> -1 Then
        FillColors(1) = newColor
        lblPreviewFill1.BackColor = newColor
    End If
End Sub

Private Sub cmdPickText1_Click()
    Dim newColor As Long
    newColor = GetColor()
    If newColor <> -1 Then
        TextColors(1) = newColor
        lblPreviewText1.BackColor = newColor
    End If
End Sub

Private Sub cmdPickFill2_Click()
    Dim newColor As Long
    newColor = GetColor()
    If newColor <> -1 Then
        FillColors(2) = newColor
        lblPreviewFill2.BackColor = newColor
    End If
End Sub

Private Sub cmdPickText2_Click()
    Dim newColor As Long
    newColor = GetColor()
    If newColor <> -1 Then
        TextColors(2) = newColor
        lblPreviewText2.BackColor = newColor
    End If
End Sub

Private Sub cmdPickFill3_Click()
    Dim newColor As Long
    newColor = GetColor()
    If newColor <> -1 Then
        FillColors(3) = newColor
        lblPreviewFill3.BackColor = newColor
    End If
End Sub

Private Sub cmdPickText3_Click()
    Dim newColor As Long
    newColor = GetColor()
    If newColor <> -1 Then
        TextColors(3) = newColor
        lblPreviewText3.BackColor = newColor
    End If
End Sub


Private Sub cmdOK_Click()
    Dim i As Integer
    For i = 1 To 3
        ThisWorkbook.Names.Add Name:="BrandFillColor" & i, RefersTo:=FillColors(i)
        ThisWorkbook.Names.Add Name:="BrandTextColor" & i, RefersTo:=TextColors(i)
    Next i
    ThisWorkbook.Save
    Unload Me
    'MsgBox "Brand colours saved!", vbInformation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

