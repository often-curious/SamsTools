VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChartPicker 
   Caption         =   "Copy Chart Style"
   ClientHeight    =   3036
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "frmChartPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChartPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim co As ChartObject

    With Me.lstCharts
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "200 pt;0 pt;0 pt"
    End With

    For Each ws In ActiveWorkbook.Worksheets
        For Each co In ws.ChartObjects

            If Not co.Chart Is gSourceChart Then
                Dim chartTitle As String

                If co.Chart.HasTitle Then
                    chartTitle = co.Chart.chartTitle.text
                Else
                    chartTitle = "(No Title)"
                End If
                
                Me.lstCharts.AddItem ws.Name & " | " & co.Name & " | " & chartTitle
                
                Me.lstCharts.List(Me.lstCharts.ListCount - 1, 1) = ws.Name
                Me.lstCharts.List(Me.lstCharts.ListCount - 1, 2) = co.Name
            End If

        Next co
    Next ws

End Sub

Private Sub cmdApply_Click()

    Dim wsName As String
    Dim chartName As String

    If Me.lstCharts.ListIndex = -1 Then
        MsgBox "Please select a target chart.", vbExclamation
        Exit Sub
    End If

    wsName = Me.lstCharts.List(Me.lstCharts.ListIndex, 1)
    chartName = Me.lstCharts.List(Me.lstCharts.ListIndex, 2)

    ApplyFormattingToTargetChart wsName, chartName

    Unload Me

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
