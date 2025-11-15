VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTextTools 
   Caption         =   "Text Editing Tools"
   ClientHeight    =   8670.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "frmTextTools.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTextTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    UpdateDeleteTextboxStates
    UpdateInsertTextboxStates
End Sub


Private Sub cmdOK_Click()
    Select Case mpgTabs.value
        ' ------------------------------
        ' ?? Change Case Tab
        ' ------------------------------
        Case 0
            If optLower.value Or optUpper.value Or optSentence.value Or optProper.value Then
                Call SaveUndoState
            End If

            If optLower.value Then
                Call ConvertTextToLower
            ElseIf optUpper.value Then
                Call ConvertTextToUpper
            ElseIf optSentence.value Then
                Call ConvertTextToSentenceCase
            ElseIf optProper.value Then
                Call ConvertTextToProper
            End If

        ' ------------------------------
        ' ?? Delete Tab
        ' ------------------------------
        Case 1
            If optDelNone.value Then
                ' Do nothing
            Else
                Call SaveUndoState

                If optDelFirstChars.value Then
                    If Not IsNumeric(txtDelFirstChars.value) Then
                        MsgBox "Please enter a valid number of characters to delete from the start.", vbExclamation, "Invalid Input"
                        Exit Sub
                    End If
                    Call DeleteFirstChars(CInt(txtDelFirstChars.value))

                ElseIf optDelLastChars.value Then
                    If Not IsNumeric(txtDelLastChars.value) Then
                        MsgBox "Please enter a valid number of characters to delete from the end.", vbExclamation, "Invalid Input"
                        Exit Sub
                    End If
                    Call DeleteLastChars(CInt(txtDelLastChars.value))

                ElseIf optDelAtPosition.value Then
                    If Not IsNumeric(txtDelStart.value) Or Not IsNumeric(txtDelCount.value) Then
                        MsgBox "Please enter valid numbers for start position and count.", vbExclamation, "Invalid Input"
                        Exit Sub
                    End If
                    Call DeleteAtPosition(CInt(txtDelStart.value), CInt(txtDelCount.value))

                ElseIf optDelSpaces.value Then
                    Call DeleteExtraSpaces

                ElseIf optDelNonPrintable.value Then
                    Call DeleteNonPrintable

                ElseIf optDelApostrophes.value Then
                    Call DeleteInitialApostrophes

                ElseIf optDelNumbersOnly.value Then
                    Call DeleteAllExceptNumbers

                ElseIf optDelLettersOnly.value Then
                    Call DeleteAllExceptLettersAndSpaces

                ElseIf optDelBeforeText.value Then
                    If Trim(txtDelBeforeText.value) = "" Then
                        MsgBox "Please enter the reference text to delete before.", vbExclamation, "Missing Input"
                        Exit Sub
                    End If
                    Call DeleteBeforeText(txtDelBeforeText.value)

                ElseIf optDelAfterText.value Then
                    If Trim(txtDelAfterText.value) = "" Then
                        MsgBox "Please enter the reference text to delete after.", vbExclamation, "Missing Input"
                        Exit Sub
                    End If
                    Call DeleteAfterText(txtDelAfterText.value)

                ElseIf optDelLineBreaks.value Then
                    Call DeleteLineBreaks
                End If
            End If

        ' ------------------------------
        ' ?? Insert Tab (future)
        ' ------------------------------
        Case 2 ' Insert Tab
            Dim insertText As String
            insertText = txtInsertText.value
        
            If optInsertNone.value Then
                ' Do nothing
            ElseIf insertText = "" Then
                MsgBox "Please enter the text to insert.", vbExclamation, "Missing Input"
                Exit Sub
            Else
                Call SaveUndoState
        
                If optInsertBefore.value Then
                    Call InsertTextBefore(insertText)
                ElseIf optInsertAfter.value Then
                    Call InsertTextAfter(insertText)
                ElseIf optInsertAtPosition.value Then
                    If Not IsNumeric(txtInsertPosition.value) Then
                        MsgBox "Please enter a valid position number.", vbExclamation, "Invalid Input"
                        Exit Sub
                    End If
                    Call InsertTextAtPosition(insertText, CLng(txtInsertPosition.value))
                End If
            End If

    End Select

    Unload Me
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUndo_Click()
    Call UndoLastChange
End Sub


Private Sub UpdateDeleteTextboxStates()
    ' Default: disable all input boxes
    txtDelFirstChars.Enabled = False
    txtDelLastChars.Enabled = False
    txtDelCount.Enabled = False
    txtDelStart.Enabled = False
    txtDelBeforeText.Enabled = False
    txtDelAfterText.Enabled = False

    ' Enable boxes based on selected option
    If optDelFirstChars.value Then
        txtDelFirstChars.Enabled = True
    ElseIf optDelLastChars.value Then
        txtDelLastChars.Enabled = True
    ElseIf optDelAtPosition.value Then
        txtDelStart.Enabled = True
        txtDelCount.Enabled = True
    ElseIf optDelBeforeText.value Then
        txtDelBeforeText.Enabled = True
    ElseIf optDelAfterText.value Then
        txtDelAfterText.Enabled = True
    End If
End Sub

Private Sub optDelFirstChars_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelLastChars_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelAtPosition_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelBeforeText_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelAfterText_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelNone_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelSpaces_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelNonPrintable_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelApostrophes_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelNumbersOnly_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelLettersOnly_Click(): UpdateDeleteTextboxStates: End Sub
Private Sub optDelLineBreaks_Click(): UpdateDeleteTextboxStates: End Sub


Private Sub UpdateInsertTextboxStates()

    ' Default: disable all input boxes
    txtInsertPosition.Enabled = False
    
    If optInsertAtPosition.value Then
        txtInsertPosition.Enabled = True
    End If
    
End Sub

Private Sub optInsertNone_Click(): UpdateInsertTextboxStates: End Sub
Private Sub optInsertAfter_Click(): UpdateInsertTextboxStates: End Sub
Private Sub optInsertBefore_Click(): UpdateInsertTextboxStates: End Sub
Private Sub optInsertAtPosition_Click(): UpdateInsertTextboxStates: End Sub


