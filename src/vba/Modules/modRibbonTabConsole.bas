Attribute VB_Name = "modRibbonTabConsole"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ParameterNotUsed

Option Explicit


'@Ignore ParameterNotUsed
Public Sub consoleClear_onAction(ByVal control As IRibbonControl)
    ClearConsoleWorksheet
End Sub

'@Ignore ParameterNotUsed
Public Sub consoleSave_onAction(ByVal control As IRibbonControl)
    SaveConsoleToFile
End Sub

'@Ignore ParameterNotUsed
Public Sub consoleClipboard_onAction(ByVal control As IRibbonControl)
    CopyConsoleToClipboard
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub consoleClipboard_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

' ===========================================================================
' Callbacks for toggleAppendMode

Public Sub toggleAppendMode_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_APPEND_CONSOLE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

Public Sub toggleAppendMode_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_APPEND_CONSOLE).value = TOGGLE_YES
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub consoleHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLConsoleTab").value, NewWindow:=True
End Sub

