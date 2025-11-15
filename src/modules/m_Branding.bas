Attribute VB_Name = "m_Branding"
#If VBA7 Then
    Private Declare PtrSafe Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
    Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, pColorRef As Long) As Long
#Else
    Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
    Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, pColorRef As Long) As Long
#End If

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Private Type hsl
    h As Double
    s As Double
    l As Double
End Type

Public TemporaryFillColor As Long


Public Function GetColor(Optional DefaultColor As Long = 0) As Long
    Dim cc As CHOOSECOLOR
    Dim CustColors(15) As Long
    
    cc.lStructSize = LenB(cc)
    cc.hwndOwner = Application.Hwnd
    cc.lpCustColors = VarPtr(CustColors(0))
    cc.rgbResult = DefaultColor
    cc.flags = 0

    If CHOOSECOLOR(cc) <> 0 Then
        GetColor = cc.rgbResult
    Else
        GetColor = -1 ' Cancelled
    End If
End Function

Sub ShowBrandColorPicker(control As IRibbonControl)
    With frmBrandColors
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    
End Sub

' OnKey wrapper with no parameters
Public Sub Key_ToggleBrandColour1()
    ToggleBrandStyle 1
End Sub

Public Sub Key_ToggleBrandColour2()
    ToggleBrandStyle 2
End Sub

Public Sub Key_ToggleBrandColour3()
    ToggleBrandStyle 3
End Sub

Sub ToggleBrandColour1(control As IRibbonControl)
    ToggleBrandStyle 1
End Sub

Sub ToggleBrandColour2(control As IRibbonControl)
    ToggleBrandStyle 2
End Sub

Sub ToggleBrandColour3(control As IRibbonControl)
    ToggleBrandStyle 3
End Sub

Private Sub ToggleBrandStyle(index As Integer)
    Dim fillColor As Long
    Dim textColor As Long
    Dim rng As Range
    Dim selChart As Chart
    Dim selSeries As Series
    Dim chartStateName As String
    Dim chartState As Long
    Dim table As ListObject

    On Error Resume Next
    fillColor = Evaluate(ThisWorkbook.Names("BrandFillColor" & index).RefersTo)
    textColor = Evaluate(ThisWorkbook.Names("BrandTextColor" & index).RefersTo)
    On Error GoTo 0

    If fillColor = 0 And textColor = 0 Then
        MsgBox "Brand colours not set. Run the brand setup first under Formatting Extras.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' --- Handle Chart Selection ---
    If TypeName(Selection) = "ChartArea" Then
        Set selChart = ActiveChart
        If Not selChart Is Nothing Then
            Dim i As Integer
            For i = 1 To selChart.SeriesCollection.count
                Dim seriesColor As Long
                seriesColor = AdjustBrightness(fillColor, (i - 1) * 0.2) ' Lighter with each series
                selChart.SeriesCollection(i).Format.Fill.ForeColor.RGB = seriesColor
                selChart.SeriesCollection(i).Format.Fill.Visible = msoTrue
                selChart.SeriesCollection(i).Format.line.Visible = msoFalse
            Next i
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If
    
    If TypeName(Selection) = "Series" Or TypeName(Selection) = "Point" Then
        On Error Resume Next
        Set selSeries = Selection
        On Error GoTo 0
    
        If Not selSeries Is Nothing Then
            With selSeries.Format
                Dim hasFill As Boolean
                Dim hasOutline As Boolean
                Dim fillRGB As Long
                Dim outlineRGB As Long
    
                hasFill = .Fill.Visible And .Fill.ForeColor.RGB = fillColor
                hasOutline = .line.Visible
    
                If hasFill And Not hasOutline Then
                    ' Change to: no fill, thick outline in brand colour
                    .Fill.Visible = msoFalse
                    .line.Visible = msoTrue
                    .line.ForeColor.RGB = fillColor
                    .line.Weight = 2.25
                    .line.DashStyle = msoLineSolid
    
                ElseIf Not hasFill And hasOutline And .line.ForeColor.RGB = fillColor Then
                    ' Change to: no fill, black outline
                    .line.ForeColor.RGB = RGB(0, 0, 0)
                    .line.Weight = 1
    
                Else
                    ' Default back to: solid fill, no outline
                    .Fill.Visible = msoTrue
                    .Fill.Solid
                    .Fill.ForeColor.RGB = fillColor
                    .Fill.Transparency = 0
                    .line.Visible = msoFalse
                End If
            End With
    
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If




    ' --- Handle Table Selection ---
    If TypeName(Selection) = "Range" Then
        Set rng = Selection

        On Error Resume Next
        Set table = rng.Cells(1).ListObject
        On Error GoTo 0

        If Not table Is Nothing Then
            Dim tableStateName As String
            Dim tableState As Long
            tableStateName = "Brand" & index & "_TableState_" & table.Name
            tableState = GetStoredState(tableStateName)
            
            With table
                Select Case tableState
                    Case 0
                    ' Clear default style to remove built-in borders
                    .TableStyle = ""
                    ' Disable default table styling
                    .ShowTableStyleFirstColumn = False
                    .ShowTableStyleLastColumn = False
                    .ShowTableStyleRowStripes = False
                    .ShowTableStyleColumnStripes = False
                    
                
                    ' Apply header style
                    With .HeaderRowRange
                        .Interior.color = fillColor
                        .Font.color = textColor
                        With .Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .color = fillColor
                        End With
                    End With
                
                    ' Clear row fill & borders
                    If Not .DataBodyRange Is Nothing Then
                        With .DataBodyRange
                            .Interior.ColorIndex = xlNone
                            Dim b As Variant
                            For Each b In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal)
                                .Borders(b).LineStyle = xlNone
                            Next b
                        End With
                    End If
                
                    ' Apply outer border
                    ApplyTableBorder .Range, fillColor, xlThick
                
                    SaveStoredState tableStateName, 1

        
                    Case 1
                        ' Add horizontal row borders
                        If Not .DataBodyRange Is Nothing Then
                            With .DataBodyRange.Borders(xlInsideHorizontal)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .color = fillColor
                            End With
                        End If
        
                        SaveStoredState tableStateName, 2
        
                    Case Else
                        ' Reset header
                        .HeaderRowRange.Interior.ColorIndex = xlNone
                        .HeaderRowRange.Font.color = vbBlack
                        
                        .TableStyle = "TableStyleLight1" ' or "" if you want no built-in style
        
                        ' Clear borders
                        ClearTableBorders .Range
                        
        
                        If Not .DataBodyRange Is Nothing Then
                            .DataBodyRange.Interior.ColorIndex = xlNone
                            For Each b In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal)
                                .DataBodyRange.Borders(b).LineStyle = xlNone
                            Next b
                        End If
        
                        SaveStoredState tableStateName, 0
                End Select
            End With
        
            Application.ScreenUpdating = True
            Exit Sub
        End If


        
        ' --- Regular Cell Toggle Logic ---
        Dim c As Range
        Dim allHaveBrandFillAndText As Boolean
        allHaveBrandFillAndText = True
        
        ' Step 1: Do all selected cells have brand fill + text?
        For Each c In rng.Cells
            If Not (c.Interior.color = fillColor And c.Font.color = textColor) Then
                allHaveBrandFillAndText = False
                Exit For
            End If
        Next c
        
        If Not allHaveBrandFillAndText Then
            ' Apply brand fill + brand text to all selected cells
            For Each c In rng.Cells
                c.Interior.color = fillColor
                c.Font.color = textColor
            Next c
        
        Else
            ' All cells already have brand fill + text, so toggle to next states
            For Each c In rng.Cells
                Dim hasNoFill As Boolean
                hasNoFill = (c.Interior.ColorIndex = xlColorIndexNone)
        
                If c.Interior.color = fillColor And c.Font.color = textColor Then
                    ' ? no fill + brand text
                    c.Interior.ColorIndex = xlColorIndexNone
                    c.Font.color = fillColor
        
                ElseIf hasNoFill And c.Font.color = fillColor Then
                    ' ? no fill + black text
                    c.Font.color = vbBlack
        
                Else
                    ' ? fallback: brand fill + brand text
                    c.Interior.color = fillColor
                    c.Font.color = textColor
                End If
            Next c
        End If
    End If

    Application.ScreenUpdating = True
End Sub

Private Function GetStoredState(key As String) As Long
    On Error Resume Next
    GetStoredState = CLng(Replace(ThisWorkbook.Names(key).RefersTo, "=", ""))
    On Error GoTo 0
End Function

Private Sub SaveStoredState(key As String, value As Long)
    On Error Resume Next
    If Not NameExists(key) Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:="=" & value
    Else
        ThisWorkbook.Names(key).RefersTo = "=" & value
    End If
    On Error GoTo 0
End Sub

Private Function NameExists(nameStr As String) As Boolean
    Dim n As Name
    On Error Resume Next
    NameExists = Not ThisWorkbook.Names(nameStr) Is Nothing
    On Error GoTo 0
End Function


Private Sub ApplyTableBorder(tblRange As Range, borderColor As Long, borderWeight As XlBorderWeight)
    Dim b As Variant
    For Each b In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        With tblRange.Borders(b)
            .LineStyle = xlContinuous
            .Weight = borderWeight
            .color = borderColor
        End With
    Next b
End Sub

Private Sub ClearTableBorders(tblRange As Range)
    Dim b As Variant
    For Each b In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal)
        With tblRange.Borders(b)
            .LineStyle = xlNone
        End With
    Next b
End Sub

Sub SetTemporaryFillFromSelection(control As IRibbonControl)
    Dim rng As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell with a fill color.", vbExclamation
        Exit Sub
    End If

    Set rng = Selection.Cells(1)
    TemporaryFillColor = rng.Interior.color

    'MsgBox "Temporary fill color set.", vbInformation
End Sub


Sub ApplyTemporaryFill(control As IRibbonControl)
    If TemporaryFillColor = 0 Then
        MsgBox "No temporary fill color has been set.", vbExclamation
        Exit Sub
    End If
    Selection.Interior.color = TemporaryFillColor
End Sub

Sub WhiteDividers(control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection

    If rng Is Nothing Then Exit Sub

    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .color = RGB(255, 255, 255)
        .Weight = xlThin
    End With

    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .color = RGB(255, 255, 255)
        .Weight = xlThin
    End With
End Sub

Sub LightenColor(control As IRibbonControl)
    Dim hsl As hsl
    Dim r As Long, g As Long, b As Long
    Dim colorVal As Long

    On Error Resume Next

    ' --- Handle Range (Cells) ---
    If TypeName(Selection) = "Range" Then
        Dim cell As Range
        For Each cell In Selection
            If cell.Interior.ColorIndex = xlNone Then
                colorVal = cell.Font.color
                RGBComponents colorVal, r, g, b
                hsl = RGBToHSL(r, g, b)
                hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1) ' lighten
                cell.Font.color = HSLToRGB(hsl)
            Else
                colorVal = cell.Interior.color
                RGBComponents colorVal, r, g, b
                hsl = RGBToHSL(r, g, b)
                hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1) ' lighten
                cell.Interior.color = HSLToRGB(hsl)
            End If
        Next cell

    ' --- Handle Chart Series or Full Chart ---
    ElseIf Not ActiveChart Is Nothing Then
        Dim cht As Chart: Set cht = ActiveChart
        Dim srs As Series
        Dim i As Long
        
        Dim selectedSeries As Series
        Set selectedSeries = Nothing
        On Error Resume Next
        Set selectedSeries = Selection
        On Error GoTo 0
        
        If Not selectedSeries Is Nothing Then
            ' Single series selected
            With selectedSeries.Format
                If .Fill.Visible And .Fill.ForeColor.RGB <> 0 Then
                    RGBComponents .Fill.ForeColor.RGB, r, g, b
                    hsl = RGBToHSL(r, g, b)
                    hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
                    .Fill.ForeColor.RGB = HSLToRGB(hsl)
                ElseIf .line.Visible And .line.ForeColor.RGB <> 0 Then
                    RGBComponents .line.ForeColor.RGB, r, g, b
                    hsl = RGBToHSL(r, g, b)
                    hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
                    .line.ForeColor.RGB = HSLToRGB(hsl)
                End If
            End With
        Else
            ' Whole chart selected? loop through all series
            For i = 1 To cht.SeriesCollection.count
                Set srs = cht.SeriesCollection(i)
                With srs.Format
                    If .Fill.Visible And .Fill.ForeColor.RGB <> 0 Then
                        RGBComponents .Fill.ForeColor.RGB, r, g, b
                        hsl = RGBToHSL(r, g, b)
                        hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
                        .Fill.ForeColor.RGB = HSLToRGB(hsl)
                    ElseIf .line.Visible And .line.ForeColor.RGB <> 0 Then
                        RGBComponents .line.ForeColor.RGB, r, g, b
                        hsl = RGBToHSL(r, g, b)
                        hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
                        .line.ForeColor.RGB = HSLToRGB(hsl)
                    End If
                End With
            Next i
        End If

    ' --- Handle Shapes or Objects with Fill or Line ---
    ElseIf TypeName(Selection) = "Shape" Or TypeName(Selection) = "ShapeRange" Then
        Dim shp As Shape
        Dim shpRange As ShapeRange
        Dim idx As Long

        If TypeName(Selection) = "Shape" Then
            Set shp = Selection
            Call ProcessShapeLighten(shp)
        ElseIf TypeName(Selection) = "ShapeRange" Then
            Set shpRange = Selection
            For idx = 1 To shpRange.count
                Call ProcessShapeLighten(shpRange.Item(idx))
            Next idx
        End If

    Else
        MsgBox "Selected object is not supported.", vbExclamation
    End If

    On Error GoTo 0
End Sub

' Helper sub for shapes to lighten fill, line, or text fill
Sub ProcessShapeLighten(shp As Shape)
    Dim hsl As hsl
    Dim r As Long, g As Long, b As Long
    Dim colorVal As Long

    On Error Resume Next
    If shp.Fill.Visible Then
        colorVal = shp.Fill.ForeColor.RGB
        If colorVal = 0 Then ' No fill color, try line
            If shp.line.Visible Then
                colorVal = shp.line.ForeColor.RGB
                RGBComponents colorVal, r, g, b
                hsl = RGBToHSL(r, g, b)
                hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
                shp.line.ForeColor.RGB = HSLToRGB(hsl)
            End If
        Else
            RGBComponents colorVal, r, g, b
            hsl = RGBToHSL(r, g, b)
            hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
            shp.Fill.ForeColor.RGB = HSLToRGB(hsl)
        End If
    ElseIf shp.line.Visible Then
        colorVal = shp.line.ForeColor.RGB
        RGBComponents colorVal, r, g, b
        hsl = RGBToHSL(r, g, b)
        hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
        shp.line.ForeColor.RGB = HSLToRGB(hsl)
    ElseIf shp.TextFrame2.HasText Then
        With shp.TextFrame2.TextRange.Font.Fill.ForeColor
            colorVal = .RGB
            RGBComponents colorVal, r, g, b
            hsl = RGBToHSL(r, g, b)
            hsl.l = WorksheetFunction.Min(1, hsl.l + 0.1)
            .RGB = HSLToRGB(hsl)
        End With
    Else
        MsgBox "Shape has no fill, line, or text to lighten.", vbInformation
    End If
    On Error GoTo 0
End Sub


Sub DarkenColor(control As IRibbonControl)
    Dim hsl As hsl
    Dim r As Long, g As Long, b As Long
    Dim colorVal As Long

    On Error Resume Next

    ' --- Handle Range (Cells) ---
    If TypeName(Selection) = "Range" Then
        Dim cell As Range
        For Each cell In Selection
            If cell.Interior.ColorIndex = xlNone Then
                colorVal = cell.Font.color
                RGBComponents colorVal, r, g, b
                hsl = RGBToHSL(r, g, b)
                hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                cell.Font.color = HSLToRGB(hsl)
            Else
                colorVal = cell.Interior.color
                RGBComponents colorVal, r, g, b
                hsl = RGBToHSL(r, g, b)
                hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                cell.Interior.color = HSLToRGB(hsl)
            End If
        Next cell

    ' --- Handle Chart Series or Full Chart ---
        ElseIf Not ActiveChart Is Nothing Then
            Dim cht As Chart: Set cht = ActiveChart
            Dim srs As Series
            Dim i As Long
        
            Dim selectedSeries As Series
            Set selectedSeries = Nothing
            On Error Resume Next
            Set selectedSeries = Selection
            On Error GoTo 0
        
            If Not selectedSeries Is Nothing Then
                ' Single series selected
                With selectedSeries.Format
                    If .Fill.Visible And .Fill.ForeColor.RGB <> 0 Then
                        RGBComponents .Fill.ForeColor.RGB, r, g, b
                        hsl = RGBToHSL(r, g, b)
                        hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                        .Fill.ForeColor.RGB = HSLToRGB(hsl)
                    ElseIf .line.Visible And .line.ForeColor.RGB <> 0 Then
                        RGBComponents .line.ForeColor.RGB, r, g, b
                        hsl = RGBToHSL(r, g, b)
                        hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                        .line.ForeColor.RGB = HSLToRGB(hsl)
                    End If
                End With
            Else
                ' Whole chart selected ? loop through all series
                For i = 1 To cht.SeriesCollection.count
                    Set srs = cht.SeriesCollection(i)
                    With srs.Format
                        If .Fill.Visible And .Fill.ForeColor.RGB <> 0 Then
                            RGBComponents .Fill.ForeColor.RGB, r, g, b
                            hsl = RGBToHSL(r, g, b)
                            hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                            .Fill.ForeColor.RGB = HSLToRGB(hsl)
                        ElseIf .line.Visible And .line.ForeColor.RGB <> 0 Then
                            RGBComponents .line.ForeColor.RGB, r, g, b
                            hsl = RGBToHSL(r, g, b)
                            hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                            .line.ForeColor.RGB = HSLToRGB(hsl)
                        End If
                    End With
                Next i
            End If


    ' --- Handle Shapes or Objects with Fill ---
    ElseIf TypeName(Selection) = "Picture" Or TypeName(Selection) = "Shape" Or TypeName(Selection) = "TextBox" Then
        Dim shp As Shape
        Set shp = Selection

        If shp.Fill.Visible Then
            colorVal = shp.Fill.ForeColor.RGB
            RGBComponents colorVal, r, g, b
            hsl = RGBToHSL(r, g, b)
            hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
            shp.Fill.ForeColor.RGB = HSLToRGB(hsl)
        ElseIf shp.TextFrame2.HasText Then
            With shp.TextFrame2.TextRange.Font.Fill.ForeColor
                colorVal = .RGB
                RGBComponents colorVal, r, g, b
                hsl = RGBToHSL(r, g, b)
                hsl.l = WorksheetFunction.Max(0, hsl.l - 0.1)
                .RGB = HSLToRGB(hsl)
            End With
        End If

    Else
        MsgBox "Selected object is not supported.", vbExclamation
    End If

    On Error GoTo 0
End Sub




Private Sub RGBComponents(colorVal As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    r = colorVal And 255
    g = (colorVal \ 256) And 255
    b = (colorVal \ 65536) And 255
End Sub

' Convert RGB to HSL
Private Function RGBToHSL(r As Long, g As Long, b As Long) As hsl
    Dim rf As Double, gf As Double, bf As Double
    Dim maxc As Double, minc As Double, delta As Double
    Dim hslColor As hsl
    
    rf = r / 255#
    gf = g / 255#
    bf = b / 255#
    
    maxc = WorksheetFunction.Max(rf, gf, bf)
    minc = WorksheetFunction.Min(rf, gf, bf)
    delta = maxc - minc
    
    ' Lightness
    hslColor.l = (maxc + minc) / 2#
    
    ' Saturation
    If delta = 0 Then
        hslColor.s = 0
        hslColor.h = 0 ' Undefined hue
    Else
        If hslColor.l < 0.5 Then
            hslColor.s = delta / (maxc + minc)
        Else
            hslColor.s = delta / (2# - maxc - minc)
        End If
        
        ' Hue
        If maxc = rf Then
            hslColor.h = (gf - bf) / delta
        ElseIf maxc = gf Then
            hslColor.h = 2# + (bf - rf) / delta
        Else
            hslColor.h = 4# + (rf - gf) / delta
        End If
        hslColor.h = hslColor.h * 60
        If hslColor.h < 0 Then hslColor.h = hslColor.h + 360
    End If
    
    RGBToHSL = hslColor
End Function

' Convert HSL to RGB
Private Function HSLToRGB(hsl As hsl) As Long
    Dim r As Double, g As Double, b As Double
    Dim q As Double, p As Double
    Dim hk As Double, t(0 To 2) As Double
    Dim i As Integer

    If hsl.s = 0 Then
        r = hsl.l: g = hsl.l: b = hsl.l
    Else
        If hsl.l < 0.5 Then
            q = hsl.l * (1 + hsl.s)
        Else
            q = hsl.l + hsl.s - hsl.l * hsl.s
        End If
        p = 2 * hsl.l - q
        hk = hsl.h / 360
        t(0) = hk + 1 / 3
        t(1) = hk
        t(2) = hk - 1 / 3
        For i = 0 To 2
            If t(i) < 0 Then t(i) = t(i) + 1
            If t(i) > 1 Then t(i) = t(i) - 1
            If t(i) < 1 / 6 Then
                t(i) = p + (q - p) * 6 * t(i)
            ElseIf t(i) < 1 / 2 Then
                t(i) = q
            ElseIf t(i) < 2 / 3 Then
                t(i) = p + (q - p) * (2 / 3 - t(i)) * 6
            Else
                t(i) = p
            End If
        Next i
        r = t(0): g = t(1): b = t(2)
    End If
    HSLToRGB = RGB(Int(r * 255), Int(g * 255), Int(b * 255))
End Function

Function AdjustBrightness(color As Long, factor As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = (color Mod 256)
    g = (color \ 256) Mod 256
    b = (color \ 65536) Mod 256

    r = Application.Min(255, Application.Max(0, r + (255 - r) * factor))
    g = Application.Min(255, Application.Max(0, g + (255 - g) * factor))
    b = Application.Min(255, Application.Max(0, b + (255 - b) * factor))

    AdjustBrightness = RGB(r, g, b)
End Function


