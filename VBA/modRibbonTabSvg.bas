Attribute VB_Name = "modRibbonTabSvg"
'@IgnoreModule ProcedureNotUsed
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")

Option Explicit

' ===========================================================================
' Callbacks for svgPostprocess

'@Ignore ParameterNotUsed
Public Sub svgPostprocess_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_POST_PROCESS_SVG).Value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub svgPostprocess_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_POST_PROCESS_SVG)
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub svgHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSvgTab").Value, NewWindow:=True
End Sub

