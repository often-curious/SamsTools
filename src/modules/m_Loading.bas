Attribute VB_Name = "m_Loading"
Public Sub ShowLoading(Optional ByVal message As String = "Processing, please wait...")
    With frmLoading
        .lblMessage.caption = message
        .Show vbModeless
        .Repaint
    End With
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    DoEvents
End Sub

Public Sub HideLoading()
    On Error Resume Next
    Unload frmLoading
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

'Make sure to add the below at the start and end of any macros
'ShowLoading "Creating Example P&L..."

'HideLoading
