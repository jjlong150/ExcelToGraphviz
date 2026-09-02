Attribute VB_Name = "modRibbonTabDiagnostics"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabDiagnostics
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Callback bridge for the "Diagnostics" Ribbon Tab, exposing environment
'   reporting, cache clearing, and diagnostic utilities.
'
' RESPONSIBILITIES:
'   - Dispatch IRibbonControl callbacks for Diagnostics tab controls.
'   - Trigger diagnostic refresh, color/font cache clearing.
'   - Provide help-panel navigation.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Worksheets: DiagnosticsSheet, SettingsSheet.
'
' CROSS-PLATFORM NOTES:
'   - Fully supported on Windows and macOS.
'
' ERROR HANDLING:
'   - Minimal; operations are worksheet-level and safe.
'
' RELATED WIKI PAGES:
'   - Diagnostics Worksheet
'   - Troubleshooting & Environment Documentation
' =============================================================================

Option Explicit

' ===========================================================================
' Callbacks for Help

Private Sub diagnosticsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLDiagnosticsTab").value, NewWindow:=True
End Sub


' ===========================================================================
' Callbacks for diagnosticsRefresh

Private Sub diagnosticsRefresh_onAction(ByVal control As IRibbonControl)
    ReportDiagnostics
End Sub

' ===========================================================================
' Callbacks for diagnosticsClearColors

Private Sub diagnosticsClearColors_onAction(ByVal control As IRibbonControl)
    ClearColorsImageFolder
End Sub

' ===========================================================================
' Callbacks for diagnosticsClearFonts

Private Sub diagnosticsClearFonts_onAction(ByVal control As IRibbonControl)
    ClearFontImageFolder
End Sub


