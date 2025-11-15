Attribute VB_Name = "m_View"
Public FullScreen As Boolean
Public hasShownFullScreenMsg As Boolean

Sub Present_mode(control As IRibbonControl)
'Changes screen zoom to 85% for all worksheets and selects cell A1 and hides the ribbon

 Static hasShownFullScreenMsg As Boolean

FullScreen = Abs(FullScreen) - 1
Application.DisplayFullScreen = FullScreen

If Not hasShownFullScreenMsg Then
    MsgBox "To exit presentation mode press Esc key", vbExclamation
    hasShownFullScreenMsg = True
End If
        
End Sub

Sub FitSelectionToScreen(control As IRibbonControl)

'Zoom to selection
ActiveWindow.Zoom = True

'Select first cell on worksheet
'Range("A1").Select

End Sub
