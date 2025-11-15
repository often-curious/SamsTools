Attribute VB_Name = "m_ModelExport"
' ====================================================================================
' EXPORT FORMULA AND NAME MANAGER MAP (Version 4)
'
' Purpose: Exports all formulas, defined names, and tables from the active workbook
'          to a text file. It specifically identifies dependencies on tables.
' Author:  MBH
' Date:    Jul 2025
'
' v4 Changes:
' - Implemented robust error handling inside the .Precedents loop to prevent
'   crashes when a precedent is an invalid range (e.g., from a closed workbook).
' ====================================================================================

Option Explicit

Public Sub ExportFormulaMap(control As IRibbonControl)

    ' --- Variable Declaration ---
    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim formulaCells As Range
    Dim precedentCell As Range
    Dim nm As Name
    Dim tbl As ListObject
    Dim nameCounter As Long
    Dim tableCounter As Long
    
    Dim fso As Object ' FileSystemObject for file operations
    Dim fileStream As Object ' TextStream for writing to the file
    Dim filePath As String
    Dim fileName As String
    Dim downloadsFolder As String
    
    ' --- Setup & Error Handling ---
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False ' Speed up execution
    
    Set Wb = ActiveWorkbook
    
    ' --- 1. Select Folder and Create File Path ---
    fileName = "Formula_Log_" & Replace(Wb.Name, ".xlsm", "") & "_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    filePath = Application.GetSaveAsFilename(InitialFileName:=fileName, _
        FileFilter:="Text Files (*.txt), *.txt", Title:="Save Formula Log As")
    
    If filePath = "False" Then
        MsgBox "Export cancelled by the user.", vbExclamation, "Export Cancelled"
        Exit Sub
    End If
    
    ' --- 2. Create and Open the Text File for Writing ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.CreateTextFile(filePath, True, True)
    
    ' --- 3. Write Header Information ---
    fileStream.WriteLine "====================================================================="
    fileStream.WriteLine "             Excel Formula & Name Manager Dependency Log"
    fileStream.WriteLine "====================================================================="
    fileStream.WriteLine "Workbook: " & Wb.FullName
    fileStream.WriteLine "Exported On: " & Now()
    fileStream.WriteLine ""
    
    ' --- 4a. Write Name Manager Details ---
    fileStream.WriteLine "---------------------------------------------------------------------"
    fileStream.WriteLine "                   USER-DEFINED NAMES (NAME MANAGER)"
    fileStream.WriteLine "---------------------------------------------------------------------"
    
    nameCounter = 0
    If Wb.Names.count > 0 Then
        For Each nm In Wb.Names
            If Not (nm.Name Like "_xlfn.*" Or nm.Name Like "_xlpm.*") Then
                nameCounter = nameCounter + 1
                fileStream.WriteLine "Name:     " & nm.Name
                If TypeName(nm.Parent) = "Workbook" Then
                    fileStream.WriteLine "Scope:    Workbook"
                Else
                    fileStream.WriteLine "Scope:    Sheet: '" & nm.Parent.Name & "'"
                End If
                fileStream.WriteLine "Refers To: " & nm.RefersToLocal
                fileStream.WriteLine "---------------------------------"
            End If
        Next nm
    End If
    
    If nameCounter = 0 Then
        fileStream.WriteLine "No user-defined names found in this workbook."
    End If
    fileStream.WriteLine ""
    
    ' --- 4b. Write Table Details ---
    fileStream.WriteLine "---------------------------------------------------------------------"
    fileStream.WriteLine "                       TABLE DEFINITIONS"
    fileStream.WriteLine "---------------------------------------------------------------------"
    
    tableCounter = 0
    For Each ws In Wb.Worksheets
        For Each tbl In ws.ListObjects
            tableCounter = tableCounter + 1
            fileStream.WriteLine "Table Name: " & tbl.Name
            fileStream.WriteLine "  -> On Sheet: '" & ws.Name & "'"
            fileStream.WriteLine "  -> Covers Range: " & tbl.Range.Address(False, False)
            fileStream.WriteLine "---------------------------------"
        Next tbl
    Next ws
    
    If tableCounter = 0 Then
        fileStream.WriteLine "No tables found in this workbook."
    End If
    fileStream.WriteLine ""
    
    ' --- 5. Loop Through Each Worksheet and Export Formulas & Dependencies ---
    For Each ws In Wb.Worksheets
        
        fileStream.WriteLine "====================================================================="
        fileStream.WriteLine "          FORMULAS & DEPENDENCIES IN SHEET: '" & ws.Name & "'"
        fileStream.WriteLine "====================================================================="
        fileStream.WriteLine ""
        
        On Error Resume Next
        Set formulaCells = ws.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrorHandler
        
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells
                fileStream.WriteLine "TARGET CELL: " & "'" & ws.Name & "'!" & cell.Address(False, False)
                fileStream.WriteLine "  -> Formula: " & cell.FormulaLocal
                
                ' --- <<<< REVISED AND ROBUST ERROR HANDLING >>>> ---
                ' Turn on error skipping right before the loop. This handles cases where
                ' a precedent is not a valid range (e.g., external link, error value).
                On Error Resume Next
                
                For Each precedentCell In cell.Precedents
                    Set tbl = Nothing ' Reset table object for each loop
                    
                    ' The .ListObject property itself can error, but we are protected
                    Set tbl = precedentCell.ListObject
                    
                    If Not tbl Is Nothing Then
                        ' The precedent is part of a table, log the table's name.
                        fileStream.WriteLine "  <- DEPENDS ON (TABLE): " & tbl.Name & " (on sheet '" & tbl.Parent.Name & "')"
                    Else
                        ' It's a standard range. The .Address property might error if
                        ' precedentCell is invalid, but On Error Resume Next will skip it.
                        fileStream.WriteLine "  <- DEPENDS ON: " & precedentCell.Address(External:=True)
                    End If
                Next precedentCell
                
                ' IMPORTANT: Reset the error handler back to the main one after the loop.
                On Error GoTo ErrorHandler
                ' --- End of revised logic ---
                
                fileStream.WriteLine ""
            Next cell
        Else
            fileStream.WriteLine "No formulas found on this sheet."
            fileStream.WriteLine ""
        End If
        
        Set formulaCells = Nothing
    Next ws

    ' --- Cleanup and Final Message ---
    fileStream.Close
    Set fso = Nothing
    Set fileStream = Nothing
    Application.ScreenUpdating = True
    
    MsgBox "Formula map successfully exported to:" & vbCrLf & vbCrLf & filePath, vbInformation, "Export Complete"
    Exit Sub

' --- Error Handling Section ---
ErrorHandler:
    Application.ScreenUpdating = True
    If Not fileStream Is Nothing Then fileStream.Close
    Set fso = Nothing
    Set fileStream = Nothing
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical, "Script Error"

End Sub


Private Function GetDownloadsFolderPath() As String
    ' Note: This helper function is robust and does not need changes.
    Dim fso As Object, wsh As Object
    Dim folderPath As String
    
    On Error Resume Next
    folderPath = Environ("USERPROFILE") & Application.PathSeparator & "Downloads"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        GetDownloadsFolderPath = folderPath
    Else
        Set wsh = CreateObject("WScript.Shell")
        If Err.Number = 0 Then GetDownloadsFolderPath = wsh.SpecialFolders("Downloads")
    End If
    On Error GoTo 0
End Function
