Attribute VB_Name = "modRibbonTabConsole"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabConsole
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Callback bridge for the "Console" Ribbon Tab, providing controls for
'   message routing, console visibility, and console-driven utilities.
'
' RESPONSIBILITIES:
'   - Dispatch IRibbonControl callbacks for Console tab controls.
'   - Route user actions to console utilities (clear, save, copy).
'   - Manage append-mode and error-routing toggles.
'   - Provide macOS-specific visibility logic for clipboard controls.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Named Ranges: SETTINGS_APPEND_CONSOLE, SETTINGS_ERROR_TO_*.
'   - Worksheets: ConsoleSheet, SettingsSheet.
'
' CROSS-PLATFORM NOTES:
'   - Clipboard operations hidden on macOS.
'
' ERROR HANDLING:
'   - Callback signatures follow IRibbonControl requirements.
'
' RELATED WIKI PAGES:
'   - Console Worksheet
'   - Message Routing & Silent Mode
' =============================================================================

Option Explicit

Private Sub consoleClear_onAction(ByVal control As IRibbonControl)
    ClearConsoleWorksheet
End Sub

Private Sub consoleSave_onAction(ByVal control As IRibbonControl)
    SaveConsoleToFile
End Sub

Private Sub consoleClipboard_onAction(ByVal control As IRibbonControl)
    CopyConsoleToClipboard
End Sub

Private Sub consoleClipboard_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

' ===========================================================================
' Callbacks for toggleAppendMode

Private Sub toggleAppendMode_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_APPEND_CONSOLE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

Private Sub toggleAppendMode_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_APPEND_CONSOLE).value = TOGGLE_YES
End Sub

' ===========================================================================
' Callbacks for Help

Private Sub consoleHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLConsoleTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for toggleErrorToConsole

Private Sub toggleErrorToConsole_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_ERROR_TO_CONSOLE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

Private Sub toggleErrorToConsole_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_ERROR_TO_CONSOLE).value = TOGGLE_YES
End Sub

' ===========================================================================
' Callbacks for toggleErrorToMessageBox

Private Sub toggleErrorToMessageBox_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_ERROR_TO_MESSAGE_BOX).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

Private Sub toggleErrorToMessageBox_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_ERROR_TO_MESSAGE_BOX).value = TOGGLE_YES
End Sub

' ===========================================================================
' Callbacks for toggleErrorToStatusBar

Private Sub toggleErrorToStatusBar_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_ERROR_TO_STATUS_BAR).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

Private Sub toggleErrorToStatusBar_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_ERROR_TO_STATUS_BAR).value = TOGGLE_YES
End Sub

