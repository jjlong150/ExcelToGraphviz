Attribute VB_Name = "modRibbonTabSvg"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabSvg
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Callback bridge for the "SVG" Ribbon Tab, providing controls for editing,
'   post-processing, clipboard operations, and worksheet-level SVG utilities.
'   Acts as the UI surface for managing SVG replacement strings, animation
'   options, and post-processing directives.
'
' RESPONSIBILITIES:
'   - Dispatch IRibbonControl callbacks for all SVG tab controls.
'   - Persist post-processing settings (SETTINGS_POST_PROCESS_SVG).
'   - Launch the SVG edit form for large replacement strings.
'   - Support clipboard operations (Windows-only).
'   - Provide help-panel navigation for SVG documentation.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Named Ranges: SETTINGS_POST_PROCESS_SVG, HelpURLSvgTab.
'   - Worksheets: SvgSheet, SettingsSheet.
'   - Modules: CellValueEditForm, clipboard helpers, status-bar helpers.
'
' CROSS-PLATFORM NOTES:
'   - Clipboard operations are hidden on macOS.
'   - SVG editing and post-processing logic behave consistently across platforms.
'
' ERROR HANDLING:
'   - Localized checks ensure edit controls are only enabled for valid cells.
'   - Callback signatures follow IRibbonControl requirements.
'
' RELATED WIKI PAGES:
'   - SVG Worksheet
'   - Output, Publishing & Post-Processing
'   - Working with Replacement Strings
' =============================================================================

Option Explicit

' ===========================================================================
' Callbacks for svgPostprocess

'@Ignore ParameterNotUsed
Public Sub svgPostprocess_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_POST_PROCESS_SVG).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub svgPostprocess_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_POST_PROCESS_SVG)
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub svgHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSvgTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for svgEditCell

'@Ignore ParameterNotUsed
Public Sub svgEditCell_onAction(ByVal control As IRibbonControl)
    CellValueEditForm.show
End Sub

'@Ignore ParameterNotUsed
Public Sub svgEditCell_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    If ActiveSheet.name <> SvgSheet.name Then
        Enabled = False
    ElseIf Selection.Cells.count > 1 Then
        Enabled = False
    ElseIf ActiveCell.HasFormula Then
        Enabled = False
    Else
        Enabled = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub svgClipboard_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

'@Ignore ParameterNotUsed
Public Sub svgClipboard_onAction(ByVal control As IRibbonControl)
#If Not Mac Then
    
    If ClipBoard_SetData(ActiveCell.value) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySvgSuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySvgFailed"), 5
    End If
    
#End If
End Sub


