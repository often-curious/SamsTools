Attribute VB_Name = "m_ListDefinedNames"
Option Explicit

Public Sub AuditDefinedNames(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name
    Dim hasNames As Boolean
    Dim outputRow As Long
    
    Set wb = ActiveWorkbook
    
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If
    
    hasNames = False

    For Each nm In wb.Names
        If Not IsSystemName(nm) Then
            hasNames = True
            Exit For
        End If
    Next nm
    
    If Not hasNames Then
        MsgBox "No user-defined names were found in '" & wb.Name & "'.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    
    'Remove existing audit sheet
    On Error Resume Next
    wb.Worksheets("_Names").Delete
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    
    
    'Create audit sheet in active workbook
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = "_Names"
    
    
    ws.Range("A1:K1").value = Array( _
        "Name", _
        "Value", _
        "Refers To", _
        "Scope", _
        "Comment", _
        "Visible?", _
        "Type", _
        "Used In Formula Count", _
        "Referenced By Names", _
        "Status", _
        "Error")
        
    ws.rows(1).Font.Bold = True
    
    outputRow = 2
    
    
    'Only names from active workbook
    For Each nm In wb.Names
        
        If Not IsSystemName(nm) Then
            
            WriteNameAudit ws, nm, outputRow, wb
            
            outputRow = outputRow + 1
            
        End If
        
    Next nm
    
    
    FormatNameAudit ws, outputRow - 1
    
    Application.ScreenUpdating = True
    
    MsgBox "Defined Name Audit complete for: " & wb.Name, vbInformation

End Sub

Private Sub WriteNameAudit(ws As Worksheet, nm As Name, row As Long, wb As Workbook)

    Dim rng As Range
    Dim status As String
    Dim issue As String
    
    Dim nameShort As String
    
    nameShort = GetShortName(nm.Name)
    
    
    ws.Cells(row, 1).value = nm.Name
    ws.Cells(row, 3).value = nm.RefersTo
    
    
    'Value
    On Error Resume Next
    Set rng = nm.RefersToRange
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        ws.Cells(row, 2).value = RangeValuesToString(rng)
    Else
        ws.Cells(row, 2).value = nm.RefersTo
    End If
    
    
    'Scope
    ws.Cells(row, 4).value = GetScope(nm)
    
    
    'Comment
    On Error Resume Next
    ws.Cells(row, 5).value = nm.Comment
    On Error GoTo 0
    
    
    'Visibility
    If nm.Visible Then
        ws.Cells(row, 6).value = "Yes"
    Else
        ws.Cells(row, 6).value = "No"
    End If
    
    
    'Type
    ws.Cells(row, 7).value = GetDefinedNameType(nm)
    
    
    'Formula usage
    ws.Cells(row, 8).value = GetNameUsageCount(nameShort, wb)
    
    
    'Dependencies
    ws.Cells(row, 9).value = GetReferencingNames(nameShort)
    
    
    'Status
    CheckNameStatus nm, status, issue, wb
    
    ws.Cells(row, 10).value = status
    ws.Cells(row, 11).value = issue
    
    
    'Refers To with hyperlink formula
    If Not rng Is Nothing Then
        
        ws.Cells(row, 3).Formula2 = _
            "=HYPERLINK(""#'" & rng.Parent.Name & "'!" & rng.address & """,""" & Replace(nm.RefersTo, """", """""") & """)"
    
    Else
        
        ws.Cells(row, 3).value = nm.RefersTo
        
    End If

End Sub

Private Function RangeValuesToString(rng As Range) As String

    Dim cell As Range
    Dim result As String
    Dim valueText As String
    
    For Each cell In rng.Cells
        
        If Len(cell.value) > 0 Then
            
            valueText = cell.value
            
            If result = "" Then
                result = valueText
            Else
                result = result & ", " & valueText
            End If
            
        End If
        
    Next cell
    
    RangeValuesToString = result

End Function

Private Function GetDefinedNameType(nm As Name) As String

    Dim rng As Range
    
    If InStr(nm.RefersTo, "[") > 0 Then
        GetDefinedNameType = "External Link"
        Exit Function
    End If
    
    
    On Error Resume Next
    Set rng = nm.RefersToRange
    On Error GoTo 0
    
    
    If Not rng Is Nothing Then
        GetDefinedNameType = "Range"
        
    ElseIf IsNumeric(Replace(nm.RefersTo, "=", "")) Then
        GetDefinedNameType = "Constant"
        
    Else
        GetDefinedNameType = "Formula"
        
    End If

End Function

Private Sub CheckNameStatus(nm As Name, _
                            ByRef status As String, _
                            ByRef issue As String, _
                            wb As Workbook)

    If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
        
        status = "Broken"
        issue = "Contains #REF!"
        
    ElseIf GetNameUsageCount(GetShortName(nm.Name), wb) = 0 _
        And Left(GetShortName(nm.Name), 1) <> "_" Then
        
        status = "Unused"
        issue = "No formulas reference this name"
        
    Else
        
        status = "OK"
        issue = ""
        
    End If

End Sub

Private Function GetNameUsageCount(nameText As String, wb As Workbook) As Long

    Dim ws As Worksheet
    Dim rng As Range
    Dim c As Range
    Dim count As Long
    
    For Each ws In wb.Worksheets
        
        If ws.Name <> "_Names" Then
            
            On Error Resume Next
            Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            
            If Not rng Is Nothing Then
                
                For Each c In rng
                    
                    If InStr(1, c.formula, nameText, vbTextCompare) > 0 Then
                        count = count + 1
                    End If
                    
                Next
                
            End If
            
        End If
        
        Set rng = Nothing
        
    Next
    
    GetNameUsageCount = count

End Function

Private Function GetReferencingNames(targetName As String) As String

    Dim nm As Name
    Dim result As String
    
    For Each nm In ThisWorkbook.Names
        
        If nm.Name <> targetName Then
            
            If InStr(1, nm.RefersTo, targetName, vbTextCompare) > 0 Then
                
                result = result & GetShortName(nm.Name) & ", "
                
            End If
            
        End If
        
    Next
    
    
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 2)
    End If
    
    GetReferencingNames = result

End Function

Private Function GetShortName(fullName As String) As String

    If InStr(fullName, "!") > 0 Then
        GetShortName = Mid(fullName, InStrRev(fullName, "!") + 1)
    Else
        GetShortName = fullName
    End If

End Function


Private Function GetScope(nm As Name) As String

    If InStr(nm.Name, "!") > 0 Then
        GetScope = Replace(Split(nm.Name, "!")(0), "'", "")
    Else
        GetScope = "Workbook"
    End If

End Function

Private Sub FormatNameAudit(ws As Worksheet, lastRow As Long)

    With ws
        
        .ListObjects.Add _
            xlSrcRange, _
            .Range("A1:K" & lastRow), _
            , xlYes
            
        .ListObjects(1).Name = "tblNameAudit"
        
        .Columns.AutoFit
        
        'Freeze header row
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        
    End With
    
    
    'Status colours
    With ws.Range("J2:J" & lastRow)
        
        .FormatConditions.Add _
            Type:=xlTextString, _
            String:="Broken", _
            TextOperator:=xlContains
            
        .FormatConditions.Add _
            Type:=xlTextString, _
            String:="Unused", _
            TextOperator:=xlContains
            
    End With

End Sub

Private Function IsSystemName(nm As Name) As Boolean

    Dim nmText As String
    
    nmText = nm.Name
    
    
    'Excel internal names
    If Left(nmText, 6) = "_xlfn." Then
        IsSystemName = True
        Exit Function
    End If
    
    
    'Excel hidden/system names
    If Left(nmText, 1) = "_" Then
        
        Select Case True
            
            Case nmText Like "_FilterDatabase*"
                IsSystemName = True
                
            Case nmText Like "_xlnm.Print_Area*"
                IsSystemName = True
                
            Case nmText Like "_xlnm.Print_Titles*"
                IsSystemName = True
                
        End Select
        
    End If

End Function

