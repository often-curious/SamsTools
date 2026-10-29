Attribute VB_Name = "m_Shortcuts"
Option Explicit

Public ShortcutList As Collection

Sub LoadShortcuts()

    Set ShortcutList = New Collection


    ShortcutList.Add Array("", "=== BRAND FORMATTING ===")
    ShortcutList.Add Array("Alt+Shift+1", "Quick colour #1")
    ShortcutList.Add Array("Alt+Shift+2", "Quick colour #2")
    ShortcutList.Add Array("Alt+Shift+3", "Quick colour #3")
    ShortcutList.Add Array("", "")
    
    ShortcutList.Add Array("", "=== MODEL & NUMBER FORMATTING ===")
    ShortcutList.Add Array("Ctrl+Shift+M", "Toggle cell model format")
    ShortcutList.Add Array("Ctrl+Shift+W", "Centre across selection")
    ShortcutList.Add Array("Ctrl+Shift+T", "Toggle total format")
    ShortcutList.Add Array("Ctrl+Shift+N", "Format number")
    ShortcutList.Add Array("Ctrl+Shift+D", "Format dollar")
    ShortcutList.Add Array("", "")
    
    ShortcutList.Add Array("", "=== CHARTS ===")
    ShortcutList.Add Array("Shift+Up Arrow", "Nudge chart shape or data label up")
    ShortcutList.Add Array("Shift+Right Arrow", "Nudge chart shape or data label right")
    ShortcutList.Add Array("Shift+Down Arrow", "Nudge chart shape or data label down")
    ShortcutList.Add Array("Shift+Left Arrow", "Nudge chart shape or data label left")
    ShortcutList.Add Array("", "")
    
    ShortcutList.Add Array("", "=== HELP ===")
    ShortcutList.Add Array("Ctrl+Shift+H", "Show Shortcut Help")
    ShortcutList.Add Array("F1", "Key has been disabled")
    

End Sub


Sub Show_Shortcut_Help(control As IRibbonControl)

    ShowShortcutHelp

End Sub

Sub ShowShortcutHelp()

    If ShortcutList Is Nothing Then
        LoadShortcuts
    End If

    frmShortcuts.Show vbModeless

End Sub
