Attribute VB_Name = "m_ExamplePL"
Option Explicit

Public Sub Create_Example_PL(control As IRibbonControl)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tmpName As String
    Dim oldAlerts As Boolean, oldScreen As Boolean, oldEvents As Boolean
    Dim oldCalc As XlCalculation

    Set wb = ActiveWorkbook
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation

    ShowLoading "Creating Example P&L..."
    
    On Error GoTo ErrHandler
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    tmpName = NextTempSheetName(wb, "__Example_PL_tmp")
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = tmpName

    If SheetExists(wb, "Example P&L") Then ws.Name = "Example P&L 2"

    ConfigureExamplePLSheet ws
    WriteExamplePLCells ws
    ApplyExamplePLStyles ws
    MergeExamplePLHeaders ws
    ApplyExamplePLRichText ws
    ApplyExamplePLConditionalFormatting ws

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    ws.Range("B2").Select

    ws.Calculate

CleanExit:
    HideLoading
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Exit Sub

ErrHandler:
    MsgBox "Unable to create Example P&L: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

Private Function NextTempSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim i As Long
    Dim candidate As String
    i = 1
    Do
        candidate = Left$(baseName & "_" & CStr(i), 31)
        If Not SheetExists(wb, candidate) Then
            NextTempSheetName = candidate
            Exit Function
        End If
        i = i + 1
    Loop
End Function

Private Sub ConfigureExamplePLSheet(ByVal ws As Worksheet)
    With ws
        .Cells.Clear
        .Cells.Font.Name = "Aptos Narrow"
        .Cells.Font.SIZE = 11
        .Cells.ColumnWidth = 8.43
        .rows.RowHeight = 14.4
        .Columns(1).ColumnWidth = 2.8867188
        .Columns(2).ColumnWidth = 2.4414062
        .Columns(3).ColumnWidth = 24.332031
        .Columns(4).ColumnWidth = 1.4414062
        .Columns(6).ColumnWidth = 1.109375
        .Columns(7).ColumnWidth = 8
        .Columns(8).ColumnWidth = 1.109375
        .Columns(16).ColumnWidth = 1.4414062
        .Range(.Columns(17), .Columns(22)).ColumnWidth = 7.21875
        .Columns(23).ColumnWidth = 1.4414062
        .Columns(25).ColumnWidth = 1
        .Range(.Columns(26), .Columns(28)).ColumnWidth = 7.21875
        .Columns(30).ColumnWidth = 3.3320312
        .rows(2).RowHeight = 21
        .rows(6).RowHeight = 24
        .rows(7).RowHeight = 6
        .rows(6).VerticalAlignment = xlCenter
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
        End With
    End With
End Sub

Private Sub WriteExamplePLCells(ByVal ws As Worksheet)
    PutText ws, "B2", "[Example Business Name]"
    PutText ws, "B3", "Income Statement ($000)"
    PutText ws, "B4", "2026"
    PutText ws, "I5", "AC"
    PutText ws, "X5", "FC"
    PutText ws, "E6", "PY"
    PutText ws, "G6", "PL"
    PutText ws, "I6", "Q1"
    PutText ws, "J6", "Q2"
    PutText ws, "K6", "Q3"
    PutText ws, "L6", "Q4"
    PutText ws, "M6", "of which Dec"
    PutFormula ws, "N6", "=B4"
    PutText ws, "O6", "in % of sales"
    PutText ws, "Q6", ChrW$(&H2206) & " PY"
    PutText ws, "T6", ChrW$(&H2206) & " PL"
    PutText ws, "U6", ChrW$(&H2206) & " PL"
    PutFormula ws, "X6", "=VALUE(N6)+1"
    PutText ws, "Z6", ChrW$(&H2206) & " FC"
    PutText ws, "AD6", "Key"
    PutText ws, "B8", "+"
    PutText ws, "C8", "Software"
    PutFormula ws, "E8", "=RANDBETWEEN(200,300)"
    PutFormula ws, "G8", "=RANDBETWEEN(200,300)"
    PutFormula ws, "I8", "=RANDBETWEEN(200,300)/4"
    PutFormula ws, "J8", "=RANDBETWEEN(200,300)/4"
    PutFormula ws, "K8", "=RANDBETWEEN(200,300)/4"
    PutFormula ws, "L8", "=RANDBETWEEN(200,300)/4"
    PutFormula ws, "M8", "=L8*40%"
    PutFormula ws, "N8", "=SUM(I8,J8,K8,L8)"
    PutFormula ws, "O8", "=N8/$N$12"
    PutFormula ws, "Q8", "=IF(R8>0,0,R8)"
    PutFormula ws, "R8", "=IFERROR(N8/E8-1,0)"
    PutFormula ws, "S8", "=IF(R8<0,0,R8)"
    PutFormula ws, "T8", "=IF(U8>0,0,U8)"
    PutFormula ws, "U8", "=IFERROR(N8/G8-1,0)"
    PutFormula ws, "V8", "=IF(U8<0,0,U8)"
    PutFormula ws, "X8", "=RANDBETWEEN(200,300)"
    PutFormula ws, "Z8", "=IF(AA8>0,0,AA8)"
    PutFormula ws, "AA8", "=IFERROR(X8/N8-1,0)"
    PutFormula ws, "AB8", "=IF(AA8<0,0,AA8)"
    PutText ws, "AE8", "Prior Year"
    PutText ws, "B9", "+"
    PutText ws, "C9", "Hardware"
    PutFormula ws, "E9", "=RANDBETWEEN(100,200)"
    PutFormula ws, "G9", "=RANDBETWEEN(100,200)"
    PutFormula ws, "I9", "=RANDBETWEEN(100,200)/4"
    PutFormula ws, "J9", "=RANDBETWEEN(100,200)/4"
    PutFormula ws, "K9", "=RANDBETWEEN(100,200)/4"
    PutFormula ws, "L9", "=RANDBETWEEN(100,200)/4"
    PutFormula ws, "M9", "=L9*40%"
    PutFormula ws, "N9", "=SUM(I9,J9,K9,L9)"
    PutFormula ws, "O9", "=N9/$N$12"
    PutFormula ws, "Q9", "=IF(R9>0,0,R9)"
    PutFormula ws, "R9", "=IFERROR(N9/E9-1,0)"
    PutFormula ws, "S9", "=IF(R9<0,0,R9)"
    PutFormula ws, "T9", "=IF(U9>0,0,U9)"
    PutFormula ws, "U9", "=IFERROR(N9/G9-1,0)"
    PutFormula ws, "V9", "=IF(U9<0,0,U9)"
    PutFormula ws, "X9", "=RANDBETWEEN(100,200)"
    PutFormula ws, "Z9", "=IF(AA9>0,0,AA9)"
    PutFormula ws, "AA9", "=IFERROR(X9/N9-1,0)"
    PutFormula ws, "AB9", "=IF(AA9<0,0,AA9)"
    PutText ws, "AE9", "Actual"
    PutText ws, "B10", "+"
    PutText ws, "C10", "Consulting"
    PutFormula ws, "E10", "=RANDBETWEEN(150,500)"
    PutFormula ws, "G10", "=RANDBETWEEN(150,500)"
    PutFormula ws, "I10", "=RANDBETWEEN(150,500)/4"
    PutFormula ws, "J10", "=RANDBETWEEN(150,500)/4"
    PutFormula ws, "K10", "=RANDBETWEEN(150,500)/4"
    PutFormula ws, "L10", "=RANDBETWEEN(150,500)/4"
    PutFormula ws, "M10", "=L10*40%"
    PutFormula ws, "N10", "=SUM(I10,J10,K10,L10)"
    PutFormula ws, "O10", "=N10/$N$12"
    PutFormula ws, "Q10", "=IF(R10>0,0,R10)"
    PutFormula ws, "R10", "=IFERROR(N10/E10-1,0)"
    PutFormula ws, "S10", "=IF(R10<0,0,R10)"
    PutFormula ws, "T10", "=IF(U10>0,0,U10)"
    PutFormula ws, "U10", "=IFERROR(N10/G10-1,0)"
    PutFormula ws, "V10", "=IF(U10<0,0,U10)"
    PutFormula ws, "X10", "=RANDBETWEEN(150,500)"
    PutFormula ws, "Z10", "=IF(AA10>0,0,AA10)"
    PutFormula ws, "AA10", "=IFERROR(X10/N10-1,0)"
    PutFormula ws, "AB10", "=IF(AA10<0,0,AA10)"
    PutText ws, "AE10", "Forecast"
    PutText ws, "C11", "of which Mgmt-Cons."
    PutFormula ws, "E11", "=E10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "G11", "=G10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "I11", "=I10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "J11", "=J10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "K11", "=K10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "L11", "=L10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "M11", "=M10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "N11", "=N10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "O11", "=N11/$N$12"
    PutFormula ws, "Q11", "=IF(R11>0,0,R11)"
    PutFormula ws, "R11", "=IFERROR(N11/E11-1,0)"
    PutFormula ws, "S11", "=IF(R11<0,0,R11)"
    PutFormula ws, "T11", "=IF(U11>0,0,U11)"
    PutFormula ws, "U11", "=IFERROR(N11/G11-1,0)"
    PutFormula ws, "V11", "=IF(U11<0,0,U11)"
    PutFormula ws, "X11", "=X10*RANDBETWEEN(65,85)/100"
    PutFormula ws, "Z11", "=IF(AA11>0,0,AA11)"
    PutFormula ws, "AA11", "=IFERROR(X11/N11-1,0)"
    PutFormula ws, "AB11", "=IF(AA11<0,0,AA11)"
    PutText ws, "AE11", "Plan"
    PutText ws, "B12", "="
    PutText ws, "C12", "Sales"
    PutFormula ws, "E12", "=SUM(E8:E10)"
    PutFormula ws, "G12", "=SUM(G8:G10)"
    PutFormula ws, "I12", "=SUM(I8:I10)"
    PutFormula ws, "J12", "=SUM(J8:J10)"
    PutFormula ws, "K12", "=SUM(K8:K10)"
    PutFormula ws, "L12", "=SUM(L8:L10)"
    PutFormula ws, "M12", "=SUM(M8:M10)"
    PutFormula ws, "N12", "=SUM(N8:N10)"
    PutFormula ws, "O12", "=N12/$N$12"
    PutFormula ws, "Q12", "=IF(R12>0,0,R12)"
    PutFormula ws, "R12", "=IFERROR(N12/E12-1,0)"
    PutFormula ws, "S12", "=IF(R12<0,0,R12)"
    PutFormula ws, "T12", "=IF(U12>0,0,U12)"
    PutFormula ws, "U12", "=IFERROR(N12/G12-1,0)"
    PutFormula ws, "V12", "=IF(U12<0,0,U12)"
    PutFormula ws, "X12", "=SUM(X8:X10)"
    PutFormula ws, "Z12", "=IF(AA12>0,0,AA12)"
    PutFormula ws, "AA12", "=IFERROR(X12/N12-1,0)"
    PutFormula ws, "AB12", "=IF(AA12<0,0,AA12)"
    PutText ws, "B14", "+"
    PutText ws, "C14", "Other income"
    PutFormula ws, "E14", "=RANDBETWEEN(10,50)"
    PutFormula ws, "G14", "=RANDBETWEEN(10,50)"
    PutFormula ws, "I14", "=RANDBETWEEN(10,50)/4"
    PutFormula ws, "J14", "=RANDBETWEEN(10,50)/4"
    PutFormula ws, "K14", "=RANDBETWEEN(10,50)/4"
    PutFormula ws, "L14", "=RANDBETWEEN(10,50)/4"
    PutFormula ws, "M14", "=L14*40%"
    PutFormula ws, "N14", "=SUM(I14,J14,K14,L14)"
    PutFormula ws, "O14", "=N14/$N$12"
    PutFormula ws, "Q14", "=IF(R14>0,0,R14)"
    PutFormula ws, "R14", "=IFERROR(N14/E14-1,0)"
    PutFormula ws, "S14", "=IF(R14<0,0,R14)"
    PutFormula ws, "T14", "=IF(U14>0,0,U14)"
    PutFormula ws, "U14", "=IFERROR(N14/G14-1,0)"
    PutFormula ws, "V14", "=IF(U14<0,0,U14)"
    PutFormula ws, "X14", "=RANDBETWEEN(10,50)"
    PutFormula ws, "Z14", "=IF(AA14>0,0,AA14)"
    PutFormula ws, "AA14", "=IFERROR(X14/N14-1,0)"
    PutFormula ws, "AB14", "=IF(AA14<0,0,AA14)"
    PutText ws, "B15", "-"
    PutText ws, "C15", "Direct costs"
    PutFormula ws, "E15", "=RANDBETWEEN(300,400)"
    PutFormula ws, "G15", "=RANDBETWEEN(300,400)"
    PutFormula ws, "I15", "=RANDBETWEEN(300,400)/4"
    PutFormula ws, "J15", "=RANDBETWEEN(300,400)/4"
    PutFormula ws, "K15", "=RANDBETWEEN(300,400)/4"
    PutFormula ws, "L15", "=RANDBETWEEN(300,400)/4"
    PutFormula ws, "M15", "=L15*40%"
    PutFormula ws, "N15", "=SUM(I15,J15,K15,L15)"
    PutFormula ws, "O15", "=N15/$N$12"
    PutFormula ws, "Q15", "=IF(R15>0,0,R15)"
    PutFormula ws, "R15", "=IFERROR(N15/E15-1,0)"
    PutFormula ws, "S15", "=IF(R15<0,0,R15)"
    PutFormula ws, "T15", "=IF(U15>0,0,U15)"
    PutFormula ws, "U15", "=IFERROR(N15/G15-1,0)"
    PutFormula ws, "V15", "=IF(U15<0,0,U15)"
    PutFormula ws, "X15", "=RANDBETWEEN(300,400)"
    PutFormula ws, "Z15", "=IF(AA15>0,0,AA15)"
    PutFormula ws, "AA15", "=IFERROR(X15/N15-1,0)"
    PutFormula ws, "AB15", "=IF(AA15<0,0,AA15)"
    PutText ws, "B16", "="
    PutText ws, "C16", "Gross profit"
    PutFormula ws, "E16", "=E12+E14-E15"
    PutFormula ws, "G16", "=G12+G14-G15"
    PutFormula ws, "I16", "=I12+I14-I15"
    PutFormula ws, "J16", "=J12+J14-J15"
    PutFormula ws, "K16", "=K12+K14-K15"
    PutFormula ws, "L16", "=L12+L14-L15"
    PutFormula ws, "M16", "=M12+M14-M15"
    PutFormula ws, "N16", "=N12+N14-N15"
    PutFormula ws, "O16", "=N16/$N$12"
    PutFormula ws, "Q16", "=IF(R16>0,0,R16)"
    PutFormula ws, "R16", "=IFERROR(N16/E16-1,0)"
    PutFormula ws, "S16", "=IF(R16<0,0,R16)"
    PutFormula ws, "T16", "=IF(U16>0,0,U16)"
    PutFormula ws, "U16", "=IFERROR(N16/G16-1,0)"
    PutFormula ws, "V16", "=IF(U16<0,0,U16)"
    PutFormula ws, "X16", "=X12+X14-X15"
    PutFormula ws, "Z16", "=IF(AA16>0,0,AA16)"
    PutFormula ws, "AA16", "=IFERROR(X16/N16-1,0)"
    PutFormula ws, "AB16", "=IF(AA16<0,0,AA16)"
    PutText ws, "C17", "in % of sales"
    PutFormula ws, "E17", "=E16/E12"
    PutFormula ws, "G17", "=G16/G12"
    PutFormula ws, "I17", "=I16/I12"
    PutFormula ws, "J17", "=J16/J12"
    PutFormula ws, "K17", "=K16/K12"
    PutFormula ws, "L17", "=L16/L12"
    PutFormula ws, "M17", "=M16/M12"
    PutFormula ws, "N17", "=N16/N12"
    PutFormula ws, "X17", "=X16/X12"
    PutText ws, "B19", "-"
    PutText ws, "C19", "Personal costs"
    PutFormula ws, "E19", "=RANDBETWEEN(50,100)"
    PutFormula ws, "G19", "=RANDBETWEEN(50,100)"
    PutFormula ws, "I19", "=RANDBETWEEN(50,100)/4"
    PutFormula ws, "J19", "=RANDBETWEEN(50,100)/4"
    PutFormula ws, "K19", "=RANDBETWEEN(50,100)/4"
    PutFormula ws, "L19", "=RANDBETWEEN(50,100)/4"
    PutFormula ws, "M19", "=L19*40%"
    PutFormula ws, "N19", "=SUM(I19,J19,K19,L19)"
    PutFormula ws, "O19", "=N19/$N$12"
    PutFormula ws, "Q19", "=IF(R19>0,0,R19)"
    PutFormula ws, "R19", "=IFERROR(N19/E19-1,0)"
    PutFormula ws, "S19", "=IF(R19<0,0,R19)"
    PutFormula ws, "T19", "=IF(U19>0,0,U19)"
    PutFormula ws, "U19", "=IFERROR(N19/G19-1,0)"
    PutFormula ws, "V19", "=IF(U19<0,0,U19)"
    PutFormula ws, "X19", "=RANDBETWEEN(50,100)"
    PutFormula ws, "Z19", "=IF(AA19>0,0,AA19)"
    PutFormula ws, "AA19", "=IFERROR(X19/N19-1,0)"
    PutFormula ws, "AB19", "=IF(AA19<0,0,AA19)"
    PutText ws, "B20", "-"
    PutText ws, "C20", "Operating costs"
    PutFormula ws, "E20", "=RANDBETWEEN(55,75)"
    PutFormula ws, "G20", "=RANDBETWEEN(55,75)"
    PutFormula ws, "I20", "=RANDBETWEEN(55,75)/4"
    PutFormula ws, "J20", "=RANDBETWEEN(55,75)/4"
    PutFormula ws, "K20", "=RANDBETWEEN(55,75)/4"
    PutFormula ws, "L20", "=RANDBETWEEN(55,75)/4"
    PutFormula ws, "M20", "=L20*40%"
    PutFormula ws, "N20", "=SUM(I20,J20,K20,L20)"
    PutFormula ws, "O20", "=N20/$N$12"
    PutFormula ws, "Q20", "=IF(R20>0,0,R20)"
    PutFormula ws, "R20", "=IFERROR(N20/E20-1,0)"
    PutFormula ws, "S20", "=IF(R20<0,0,R20)"
    PutFormula ws, "T20", "=IF(U20>0,0,U20)"
    PutFormula ws, "U20", "=IFERROR(N20/G20-1,0)"
    PutFormula ws, "V20", "=IF(U20<0,0,U20)"
    PutFormula ws, "X20", "=RANDBETWEEN(55,75)"
    PutFormula ws, "Z20", "=IF(AA20>0,0,AA20)"
    PutFormula ws, "AA20", "=IFERROR(X20/N20-1,0)"
    PutFormula ws, "AB20", "=IF(AA20<0,0,AA20)"
    PutText ws, "B21", "-"
    PutText ws, "C21", "Other costs"
    PutFormula ws, "E21", "=RANDBETWEEN(5,25)"
    PutFormula ws, "G21", "=RANDBETWEEN(5,25)"
    PutFormula ws, "I21", "=RANDBETWEEN(5,25)/4"
    PutFormula ws, "J21", "=RANDBETWEEN(5,25)/4"
    PutFormula ws, "K21", "=RANDBETWEEN(5,25)/4"
    PutFormula ws, "L21", "=RANDBETWEEN(5,25)/4"
    PutFormula ws, "M21", "=L21*40%"
    PutFormula ws, "N21", "=SUM(I21,J21,K21,L21)"
    PutFormula ws, "O21", "=N21/$N$12"
    PutFormula ws, "Q21", "=IF(R21>0,0,R21)"
    PutFormula ws, "R21", "=IFERROR(N21/E21-1,0)"
    PutFormula ws, "S21", "=IF(R21<0,0,R21)"
    PutFormula ws, "T21", "=IF(U21>0,0,U21)"
    PutFormula ws, "U21", "=IFERROR(N21/G21-1,0)"
    PutFormula ws, "V21", "=IF(U21<0,0,U21)"
    PutFormula ws, "X21", "=RANDBETWEEN(5,25)"
    PutFormula ws, "Z21", "=IF(AA21>0,0,AA21)"
    PutFormula ws, "AA21", "=IFERROR(X21/N21-1,0)"
    PutFormula ws, "AB21", "=IF(AA21<0,0,AA21)"
    PutText ws, "B22", "="
    PutText ws, "C22", "Operating profit"
    PutFormula ws, "E22", "=E16-SUM(E19:E21)"
    PutFormula ws, "G22", "=G16-SUM(G19:G21)"
    PutFormula ws, "I22", "=I16-SUM(I19:I21)"
    PutFormula ws, "J22", "=J16-SUM(J19:J21)"
    PutFormula ws, "K22", "=K16-SUM(K19:K21)"
    PutFormula ws, "L22", "=L16-SUM(L19:L21)"
    PutFormula ws, "M22", "=M16-SUM(M19:M21)"
    PutFormula ws, "N22", "=N16-SUM(N19:N21)"
    PutFormula ws, "O22", "=N22/$N$12"
    PutFormula ws, "Q22", "=IF(R22>0,0,R22)"
    PutFormula ws, "R22", "=IFERROR(N22/E22-1,0)"
    PutFormula ws, "S22", "=IF(R22<0,0,R22)"
    PutFormula ws, "T22", "=IF(U22>0,0,U22)"
    PutFormula ws, "U22", "=IFERROR(N22/G22-1,0)"
    PutFormula ws, "V22", "=IF(U22<0,0,U22)"
    PutFormula ws, "X22", "=X16-SUM(X19:X21)"
    PutFormula ws, "Z22", "=IF(AA22>0,0,AA22)"
    PutFormula ws, "AA22", "=IFERROR(X22/N22-1,0)"
    PutFormula ws, "AB22", "=IF(AA22<0,0,AA22)"
    PutText ws, "C23", "in % of sales"
    PutFormula ws, "E23", "=E22/E12"
    PutFormula ws, "G23", "=G22/G12"
    PutFormula ws, "I23", "=I22/I12"
    PutFormula ws, "J23", "=J22/J12"
    PutFormula ws, "K23", "=K22/K12"
    PutFormula ws, "L23", "=L22/L12"
    PutFormula ws, "M23", "=M22/M12"
    PutFormula ws, "N23", "=N22/N12"
    PutFormula ws, "X23", "=X22/X12"
End Sub

Private Sub PutText(ByVal ws As Worksheet, ByVal addressText As String, ByVal textValue As String)
    With ws.Range(addressText)
        .numberFormat = "@"
        .Value2 = textValue
    End With
End Sub

Private Sub PutFormula(ByVal ws As Worksheet, ByVal addressText As String, ByVal formulaText As String)
    ws.Range(addressText).formula = formulaText
End Sub

Private Sub ApplyExamplePLStyles(ByVal ws As Worksheet)
    ApplyStyleList ws, "B3,AE8,AE9,AE10,AE11", 0
    ApplyStyleList ws, "B2", 1
    ApplyStyleList ws, "B4", 2
    ApplyStyleList ws, "I5,J5,K5,L5,M5,N5,O5,X5", 3
    ApplyStyleList ws, "B6,C6", 4
    ApplyStyleList ws, "E6,F6,G6,H6,I6,J6,K6,L6", 6
    ApplyStyleList ws, "M6", 7
    ApplyStyleList ws, "N6,X6", 8
    ApplyStyleList ws, "O6", 9
    ApplyStyleList ws, "Q6,R6,T6,U6,Z6,AA6,AB6", 10
    ApplyStyleList ws, "S6,V6", 11
    ApplyStyleList ws, "AD6", 12
    ApplyStyleList ws, "E7", 13
    ApplyStyleList ws, "G7", 14
    ApplyStyleList ws, "I7,J7,K7", 15
    ApplyStyleList ws, "L7,N7", 16
    ApplyStyleList ws, "M7,O7", 17
    ApplyStyleList ws, "Q7,R7,U7,V7", 18
    ApplyStyleList ws, "S7", 19
    ApplyStyleList ws, "T7", 20
    ApplyStyleList ws, "X7,AD10", 21
    ApplyStyleList ws, "B8,B14,B19,B20", 22
    ApplyStyleList ws, "C8,C14,C19,C20", 23
    ApplyStyleList ws, "E8,G8,E14,G14,E19,G19,E20,G20", 24
    ApplyStyleList ws, "F8,H8,F9,H9,F10,H10,E13,F13,G13,H13,I13,J13,K13,L13,M13,N13,X13,F14,H14,F15,H15,F19,H19,F20,H20,F21,H21", 25
    ApplyStyleList ws, "I8,J8,K8,I14,J14,K14,I19,J19,K19,I20,J20,K20", 26
    ApplyStyleList ws, "L8,N8,X8,L14,N14,X14,L19,N19,X19,L20,N20,X20", 27
    ApplyStyleList ws, "M8,M14,M19,M20", 28
    ApplyStyleList ws, "O8", 29
    ApplyStyleList ws, "Q8,R8,U8,AA8", 30
    ApplyStyleList ws, "S8,V8,AB8", 31
    ApplyStyleList ws, "T8,Z8", 32
    ApplyStyleList ws, "AD8", 33
    ApplyStyleList ws, "B9,B15,B21", 34
    ApplyStyleList ws, "C9,C15,C21", 35
    ApplyStyleList ws, "E9,G9,E15,G15,E21,G21", 36
    ApplyStyleList ws, "I9,J9,K9,I15,J15,K15,I21,J21,K21", 37
    ApplyStyleList ws, "L9,N9,X9,L15,N15,X15,L21,N21,X21", 38
    ApplyStyleList ws, "M9,M15,M21", 39
    ApplyStyleList ws, "O9", 40
    ApplyStyleList ws, "Q9,R9,U9,AA9", 41
    ApplyStyleList ws, "S9,V9,AB9", 42
    ApplyStyleList ws, "T9,Z9", 43
    ApplyStyleList ws, "AD9", 44
    ApplyStyleList ws, "B10", 45
    ApplyStyleList ws, "C10", 46
    ApplyStyleList ws, "E10,G10", 47
    ApplyStyleList ws, "I10,J10,K10", 48
    ApplyStyleList ws, "L10,N10,X10", 49
    ApplyStyleList ws, "M10", 50
    ApplyStyleList ws, "O10", 51
    ApplyStyleList ws, "Q10,R10,U10,AA10", 52
    ApplyStyleList ws, "S10,V10,AB10", 53
    ApplyStyleList ws, "T10,Z10", 54
    ApplyStyleList ws, "C11,P11,W11,Y11,D17,F17,H17,P17,W17,Y17", 55
    ApplyStyleList ws, "E11,G11,I11,J11,K11", 56
    ApplyStyleList ws, "F11,H11", 57
    ApplyStyleList ws, "L11,N11,X11", 58
    ApplyStyleList ws, "M11", 59
    ApplyStyleList ws, "O11", 60
    ApplyStyleList ws, "Q11,T11,Z11", 61
    ApplyStyleList ws, "R11,U11,AA11", 62
    ApplyStyleList ws, "S11,V11,AB11", 63
    ApplyStyleList ws, "AD11", 64
    ApplyStyleList ws, "B12,C12,B16,C16,B22,C22", 65
    ApplyStyleList ws, "E12,G12,I12,J12,K12,E16,G16,I16,J16,K16,X16,E22,G22,I22,J22,K22,X22", 66
    ApplyStyleList ws, "F12,H12,F16,H16,F22,H22", 67
    ApplyStyleList ws, "L12,N12,X12,L16,N16,L22,N22", 68
    ApplyStyleList ws, "M12,M16,M22", 69
    ApplyStyleList ws, "Q12,T12,Z12", 71
    ApplyStyleList ws, "R12,U12,AA12", 72
    ApplyStyleList ws, "S12,V12,AB12", 73
    ApplyStyleList ws, "S13,V13,AB13", 74
    ApplyStyleList ws, "T13,Z13", 75
    ApplyStyleList ws, "O14,O15,O19,O20", 76
    ApplyStyleList ws, "Q14,R14,U14,AA14,Q19,R19,U19,AA19,Q20,R20,U20,AA20", 77
    ApplyStyleList ws, "S14,V14,AB14,S19,V19,AB19,S20,V20,AB20", 78
    ApplyStyleList ws, "T14,Z14,T19,Z19,T20,Z20", 79
    ApplyStyleList ws, "Q15,R15,U15,AA15,Q21,R21,U21,AA21", 80
    ApplyStyleList ws, "S15,V15,AB15,S21,V21,AB21", 81
    ApplyStyleList ws, "T15,Z15,T21,Z21", 82
    ApplyStyleList ws, "P16,W16,Y16,P22,W22,Y22", 83
    ApplyStyleList ws, "Q16,T16,Z16,Q22,T22,Z22", 84
    ApplyStyleList ws, "R16,U16,AA16,R22,U22,AA22", 85
    ApplyStyleList ws, "S16,V16,AB16,S22,V22,AB22", 86
    ApplyStyleList ws, "C17,C23", 87
    ApplyStyleList ws, "E17,G17,I17,J17,K17,L17,M17,N17,X17,E23,G23,I23,J23,K23,L23,M23,N23,X23", 88
    ApplyStyleList ws, "O17", 89
    ApplyStyleList ws, "Q17,R17,U17,AA17", 90
    ApplyStyleList ws, "S17,V17,AB17", 91
    ApplyStyleList ws, "T17,Z17", 92
    ApplyStyleList ws, "O18", 93
    ApplyStyleList ws, "Q18,R18,U18,AA18", 94
    ApplyStyleList ws, "S18,V18,AB18", 95
    ApplyStyleList ws, "T18,Z18", 96
    ApplyStyleList ws, "O21", 97
    ApplyStyleList ws, "O12,O16,O22", 70
End Sub

Private Sub ApplyStyleList(ByVal ws As Worksheet, ByVal addressList As String, ByVal styleId As Long)
    Dim item As Variant
    For Each item In Split(addressList, ",")
        ApplyStyleId ws.Range(CStr(item)), styleId
    Next item
End Sub

Private Sub ApplyStyleId(ByVal r As Range, ByVal styleId As Long)
    Select Case styleId
        Case 0
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 1
            SetBaseFormat r, "General", 16, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 2
            SetBaseFormat r, "@", 11, False, False, xlUnderlineStyleNone, xlLeft, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 3
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlCenterAcrossSelection, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 4
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlCenter, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 6
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlRight, xlCenter, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 7
            SetBaseFormat r, "General", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlCenter, True, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 8
            SetBaseFormat r, "@", 11, False, False, xlUnderlineStyleNone, xlRight, xlCenter, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 9
            SetBaseFormat r, "General", 9, False, True, xlUnderlineStyleNone, xlGeneral, xlCenter, True, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 10
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlCenter, xlCenter, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 11
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlCenter, xlCenter, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 12
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleSingle, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 13
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlSolid, True, 191, 191, 191, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThin, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThin, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 255, 255, 255
        Case 14
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeRight, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 15
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlSolid, True, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
        Case 16
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlSolid, True, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
        Case 17
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlSolid, True, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
        Case 18
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 19
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 20
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 21
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlPatternLightUp, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeRight, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 22
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 23
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 24
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 25
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 26
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 27
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 28
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 29
            SetBaseFormat r, "0%", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 30
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 31
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 32
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 33
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlSolid, True, 191, 191, 191, 0, 0, 0, True
        Case 34
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 35
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 36
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 37
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 38
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 39
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 40
            SetBaseFormat r, "0%", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 41
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 42
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 43
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 44
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlSolid, True, 0, 0, 0, 0, 0, 0, True
        Case 45
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 46
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 47
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 48
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 49
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 50
            SetBaseFormat r, "0", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 51
            SetBaseFormat r, "0%", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 52
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 53
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 54
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
        Case 55
            SetBaseFormat r, "General", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 56
            SetBaseFormat r, "0", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 57
            SetBaseFormat r, "0", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 58
            SetBaseFormat r, "0", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 59
            SetBaseFormat r, "0", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 60
            SetBaseFormat r, "0%", 8, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 61
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 62
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 63
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 64
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeRight, xlThin, False, 0, 0, 0
            SetEdgeBorder r, xlEdgeBottom, xlThin, False, 0, 0, 0
        Case 65
            SetBaseFormat r, "General", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 66
            SetBaseFormat r, "0", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 67
            SetBaseFormat r, "0", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 68
            SetBaseFormat r, "0", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 69
            SetBaseFormat r, "0", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 70
            SetBaseFormat r, "0%", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 71
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 72
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 73
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 74
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
        Case 75
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
        Case 76
            SetBaseFormat r, "0%", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 77
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 78
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 79
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 80
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 81
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 82
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case 83
            SetBaseFormat r, "General", 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 84
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 85
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 86
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, True, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
            SetEdgeBorder r, xlEdgeTop, xlThin, False, 0, 0, 0
        Case 87
            SetBaseFormat r, "General", 9, False, True, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 88
            SetBaseFormat r, "0%", 9, False, True, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 89
            SetBaseFormat r, "0%", 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 90
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 91
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
        Case 92
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 9, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
        Case 93
            SetBaseFormat r, "0%", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 94
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
        Case 95
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeRight, xlThick, True, 255, 255, 255
        Case 96
            SetBaseFormat r, "\+#,##0%;[Red]\-#,##0%;"" - """, 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeLeft, xlThick, True, 255, 255, 255
        Case 97
            SetBaseFormat r, "0%", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, True
            SetEdgeBorder r, xlEdgeTop, xlThin, True, 191, 191, 191
            SetEdgeBorder r, xlEdgeBottom, xlThin, True, 191, 191, 191
        Case Else
            SetBaseFormat r, "General", 11, False, False, xlUnderlineStyleNone, xlGeneral, xlBottom, False, False, xlNone, False, 0, 0, 0, 0, 0, 0, False
    End Select
End Sub

Private Sub SetBaseFormat(ByVal r As Range, ByVal numFmt As String, ByVal fontSize As Double, _
                          ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal underlineStyle As Long, _
                          ByVal hAlign As Long, ByVal vAlign As Long, ByVal wrapIt As Boolean, ByVal shrinkIt As Boolean, _
                          ByVal fillPattern As Long, ByVal hasFillColor As Boolean, ByVal fillR As Long, ByVal fillG As Long, ByVal fillB As Long, _
                          ByVal fontR As Long, ByVal fontG As Long, ByVal fontB As Long, ByVal hasFontColor As Boolean)
    With r
        .numberFormat = numFmt
        .HorizontalAlignment = hAlign
        .VerticalAlignment = vAlign
        .WrapText = wrapIt
        .ShrinkToFit = shrinkIt
        .Font.Name = "Aptos Narrow"
        .Font.SIZE = fontSize
        .Font.Bold = isBold
        .Font.Italic = isItalic
        .Font.Underline = underlineStyle
        If hasFontColor Then
            .Font.color = RGB(fontR, fontG, fontB)
        Else
            .Font.colorIndex = xlAutomatic
        End If
        .Interior.pattern = fillPattern
        If fillPattern = xlNone Then
            .Interior.pattern = xlNone
        ElseIf hasFillColor Then
            .Interior.color = RGB(fillR, fillG, fillB)
        End If
        .Borders.LineStyle = xlNone
    End With
End Sub

Private Sub SetEdgeBorder(ByVal r As Range, ByVal edgeIndex As Long, ByVal weightValue As Long, _
                          ByVal hasRgbColor As Boolean, ByVal redValue As Long, ByVal greenValue As Long, ByVal blueValue As Long)
    With r.Borders(edgeIndex)
        .LineStyle = xlContinuous
        .Weight = weightValue
        If hasRgbColor Then
            .color = RGB(redValue, greenValue, blueValue)
        Else
            .colorIndex = xlAutomatic
        End If
    End With
End Sub

Private Sub MergeExamplePLHeaders(ByVal ws As Worksheet)
    ws.Range("Q6:S6").Merge
    ws.Range("T6:V6").Merge
    ws.Range("Z6:AB6").Merge
    ws.Range("Q6:S6,T6:V6,Z6:AB6").HorizontalAlignment = xlCenter
    ws.Range("Q6:S6,T6:V6,Z6:AB6").VerticalAlignment = xlCenter
End Sub

Private Sub ApplyExamplePLRichText(ByVal ws As Worksheet)
    ws.Range("Q6").Value2 = ChrW$(&H2206) & " PY"
    With ws.Range("Q6")
        .Characters(1, 1).Font.SIZE = 11
        .Characters(2, Len(.Value2) - 1).Font.SIZE = 9.9
    End With
    ws.Range("Z6").Value2 = ChrW$(&H2206) & " FC"
    With ws.Range("Z6")
        .Characters(1, 1).Font.SIZE = 11
        .Characters(2, Len(.Value2) - 1).Font.SIZE = 9.9
    End With
End Sub

Private Sub ApplyExamplePLConditionalFormatting(ByVal ws As Worksheet)
    ws.Range("R8:R22,U8:U22,AA8:AA22").FormatConditions.Delete
    AddDataBar ws.Range("U8:U22"), RGB(99, 195, 132), False
    AddDataBar ws.Range("R8:R22"), RGB(99, 195, 132), False
    AddDataBar ws.Range("AA8:AA22"), RGB(99, 195, 132), False
End Sub

Private Sub AddDataBar(ByVal rng As Range, ByVal barRgb As Long, ByVal showDataValue As Boolean)
    Dim db As Object
    Set db = rng.FormatConditions.AddDataBar
    With db
        .ShowValue = showDataValue
        .BarColor.color = barRgb
        .BarFillType = xlDataBarFillSolid
        .BarBorder.Type = xlDataBarBorderNone
        .AxisPosition = xlDataBarAxisAutomatic
        .AxisColor.color = RGB(0, 0, 0)
        .MinPoint.Modify newtype:=xlConditionValueLowestValue
        .MaxPoint.Modify newtype:=xlConditionValueHighestValue
         With .NegativeBarFormat
            .ColorType = xlDataBarColor
            .BorderColorType = xlDataBarColor
        End With
        
        .NegativeBarFormat.color.color = RGB(255, 0, 0)
        .Direction = xlContext
        
    End With
End Sub

