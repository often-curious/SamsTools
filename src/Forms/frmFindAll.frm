VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFindAll 
   Caption         =   "Find All On All Sheets"
   ClientHeight    =   3396
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   OleObjectBlob   =   "frmFindAll.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFindAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------     ExcelCampus.com     ------------------------
'Find All User Form
'
'by Jon Acampora, jon@excelcampus.com
'
'Description: The form uses the the FindAll function by Chip Pearson
'             to find and return results to a listbox as the user types.
'             The user can click on a result to go to the cell listed in
'             the results.
'
'Date: 03/19/2013
'
'-------------------------------------------------------------------------

Dim strSearchAddress As String

Private Sub UserForm_Initialize()
'Define Search Address

Dim ws As Worksheet
Dim lRow As Long
Dim lCol As Long
Dim lMaxRow As Long
Dim lMaxCol As Long

    lMaxRow = 0
    lMaxCol = 0
    
    'Set range to search
    For Each ws In ActiveWorkbook.Worksheets
        lRow = ws.UsedRange.Cells.rows.count
        lCol = ws.UsedRange.Cells.Columns.count

        If lRow > lMaxRow Then lMaxRow = lRow
        If lCol > lMaxCol Then lMaxCol = lCol
    Next ws
    
    strSearchAddress = Range(Cells(1, 1), Cells(lMaxRow, lMaxCol)).Address

End Sub

Private Sub TextBox_Find_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Calls the FindAllMatches routine as user types text in the textbox

    Call FindAllMatches
    
End Sub

Private Sub Label_ClearFind_Click()
'Clears the find text box and sets focus

    Me.TextBox_Find.text = ""
    Me.TextBox_Find.SetFocus
    
End Sub

Sub FindAllMatches()
'Find all matches on activesheet
'Called by: TextBox_Find_KeyUp event

Dim FindWhat As Variant
Dim FoundCells As Variant
Dim FoundCell As Range
Dim arrResults() As Variant
Dim lFound As Long
Dim lSearchCol As Long
Dim lLastRow As Long
Dim lWS As Long
Dim lCount As Long
Dim ws As Worksheet
Dim lRow As Long
Dim lCol As Long
Dim lMaxRow As Long
Dim lMaxCol As Long
   
    If Len(frmFindAll.TextBox_Find.value) > 1 Then 'Do search if text in find box is longer than 1 character.
        
        FindWhat = frmFindAll.TextBox_Find.value
        'Calls the FindAll function
        FoundCells = FindAllOnWorksheets(Nothing, Empty, SearchAddress:=strSearchAddress, _
                                FindWhat:=FindWhat, _
                                LookIn:=xlValues, _
                                LookAt:=xlPart, _
                                SearchOrder:=xlByColumns, _
                                MatchCase:=False, _
                                BeginsWith:=vbNullString, _
                                EndsWith:=vbNullString, _
                                BeginEndCompare:=vbTextCompare)

        'Add results of FindAll to an array
        lCount = 0
        For lWS = LBound(FoundCells) To UBound(FoundCells)
            If Not FoundCells(lWS) Is Nothing Then
                lCount = lCount + FoundCells(lWS).count
            End If
        Next lWS
        
        If lCount = 0 Then
            ReDim arrResults(1 To 1, 1 To 2)
            arrResults(1, 1) = "No Results"
        
        Else
        
            ReDim arrResults(1 To lCount, 1 To 2)
            
            lFound = 1
            For lWS = LBound(FoundCells) To UBound(FoundCells)
                If Not FoundCells(lWS) Is Nothing Then
                    For Each FoundCell In FoundCells(lWS)
                        arrResults(lFound, 1) = FoundCell.value
                        arrResults(lFound, 2) = "'" & FoundCell.Parent.Name & "'!" & FoundCell.Address(External:=False)
                        lFound = lFound + 1
                    Next FoundCell
                End If
            Next lWS
        End If
        
        'Populate the listbox with the array
        Me.ListBox_Results.List = arrResults
        
    Else
        Me.ListBox_Results.Clear
    End If
        
End Sub

Private Sub ListBox_Results_Click()
'Go to selection on sheet when result is clicked

Dim strAddress As String
Dim strSheet As String
Dim strCell As String
Dim l As Long

    For l = 0 To ListBox_Results.ListCount
        If ListBox_Results.Selected(l) = True Then
            strAddress = ListBox_Results.List(l, 1)
            strSheet = Replace(Mid(strAddress, 1, InStr(1, strAddress, "!") - 1), "'", "")
            Worksheets(strSheet).Select
            Worksheets(strSheet).Range(strAddress).Select
            GoTo EndLoop
        End If
    Next l

EndLoop:
    
End Sub

Private Sub CommandButton_Close_Click()
'Close the userform

    Unload Me
    
End Sub

