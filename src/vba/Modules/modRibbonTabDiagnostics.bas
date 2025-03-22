Attribute VB_Name = "modRibbonTabDiagnostics"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")

Option Explicit

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub diagnosticsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLDiagnosticsTab").value, NewWindow:=True
End Sub


' ===========================================================================
' Callbacks for diagnosticsRefresh

'@Ignore ParameterNotUsed
Public Sub diagnosticsRefresh_onAction(ByVal control As IRibbonControl)
    ReportDiagnostics
End Sub

' ===========================================================================
' Callbacks for diagnosticsClearColors

'@Ignore ParameterNotUsed
Public Sub diagnosticsClearColors_onAction(ByVal control As IRibbonControl)
    ClearColorsImageFolder
End Sub

' ===========================================================================
' Callbacks for diagnosticsClearFonts

'@Ignore ParameterNotUsed
Public Sub diagnosticsClearFonts_onAction(ByVal control As IRibbonControl)
    ClearFontImageFolder
End Sub


