Attribute VB_Name = "m_Copy"
Public Sub doMultiCopy(control As IRibbonControl)
  If TypeName(Selection) <> "Range" Then Exit Sub
  Dim rngDest As Excel.Range
  Dim i As Long
  Dim vRegions As Variant
  Dim rngRegions() As Excel.Range
 
  vRegions = Split(Selection.Address, ",")
 
  ReDim rngRegions(LBound(vRegions) To UBound(vRegions))
 
  Set rngDest = Application.InputBox("Select a destination cell", "Where to paste selections?", , , , , , 8)
 
  For i = LBound(vRegions) To UBound(vRegions)
    Set rngRegions(i) = Range(vRegions(i))
    rngRegions(i).Copy _
        Destination:=rngDest.Offset(rngRegions(i).row - rngRegions(LBound(rngRegions)).row, _
              rngRegions(i).Column - rngRegions(LBound(rngRegions)).Column)
  Next i
End Sub

Sub CopySelections(control As IRibbonControl)
Set cellranges = Application.Selection
Set ThisRng = Application.InputBox("Select a destination cell", "Where to paste selections?", Type:=8)
For Each cellRange In cellranges.Areas
    cellRange.Copy ThisRng.Offset(i)
    i = i + cellRange.rows.CountLarge
Next cellRange
End Sub


Sub CopyPicture(control As IRibbonControl)
    Dim ws As Worksheet
    Dim chtObj As ChartObject

    Set ws = ActiveSheet

    On Error Resume Next

    ' Case 1: Range is selected
    If TypeOf Selection Is Range Then
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        ws.PasteSpecial Format:="Picture (Enhanced Metafile)"
        Exit Sub
    End If

    ' Case 2: ChartObject is selected directly (e.g. clicked on border)
    If TypeOf Selection Is ChartObject Then
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        ws.PasteSpecial Format:="Picture (Enhanced Metafile)"
        Exit Sub
    End If

    ' Case 3: A part inside the chart is selected (ChartArea, Series, etc.)
    If TypeOf Selection Is Chart Then
        ' In this case, Selection is a Chart, but we need to get the ChartObject
        Set chtObj = Selection.Parent
        If Not chtObj Is Nothing Then
            chtObj.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            ws.PasteSpecial Format:="Picture (Enhanced Metafile)"
            Exit Sub
        End If
    End If

    ' Case 4: A chart part like the plot area or data label is selected
    If Not Selection Is Nothing Then
        If Not Selection.Parent Is Nothing Then
            If TypeOf Selection.Parent Is Chart Then
                Set chtObj = Selection.Parent.Parent
                If TypeOf chtObj Is ChartObject Then
                    chtObj.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                    ws.PasteSpecial Format:="Picture (Enhanced Metafile)"
                    Exit Sub
                End If
            End If
        End If
    End If

    ' If none matched
    MsgBox "Please select a range or a chart (not a chart sheet).", vbExclamation
End Sub


Sub LinkedPicture(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range before running this.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    Selection.Copy
    ActiveSheet.Pictures.Paste Link:=True
End Sub
