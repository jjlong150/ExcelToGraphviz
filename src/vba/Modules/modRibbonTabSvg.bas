Attribute VB_Name = "modRibbonTabSvg"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabSvg
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon Callbacks
'
' ROLE:
'   Ribbon callback module for the SVG Tab. Manages post-processing toggles,
'   cell-level SVG editing, clipboard operations, and help navigation. Provides
'   the UI surface for configuring SVG replacement strings, animation flags,
'   and worksheet-level SVG utilities.
'
' RESPONSIBILITIES:
'   - Manage SVG post-processing state (SETTINGS_POST_PROCESS_SVG) and update
'     Ribbon visibility for On/Off/Turn-On/Turn-Off controls.
'   - Display risk-warning prompts when enabling SVG post-processing.
'   - Launch the cell-value SVG editor (CellValueEditForm) for editing large
'     replacement strings.
'   - Support clipboard copy operations for SVG cell content (Windows only).
'   - Provide help-link navigation for SVG documentation via workbook-configured
'     URLs.
'   - Synchronize Ribbon state through getVisible/getEnabled callbacks and
'     targeted invalidation (SvgRefreshButtons).
'
' INTERACTIONS:
'   - Named Ranges: SETTINGS_POST_PROCESS_SVG, HelpURLSvgTab.
'   - Worksheets: SvgSheet, SettingsSheet.
'   - Modules: CellValueEditForm, clipboard helpers, status-bar helpers.
'
' CROSS-PLATFORM NOTES:
'   - Clipboard operations are hidden on macOS.
'   - Post-processing logic and SVG editing behave consistently across platforms.
'
' ERROR HANDLING:
'   - Validates active sheet, selection size, and formula presence before
'     enabling SVG cell editing.
'   - All callbacks follow IRibbonControl signature requirements.
'
' RELATED WIKI PAGES:
'   - SVG Worksheet
'   - Replacement Strings & Animation Flags
'   - Output, Publishing & Post-Processing
' =============================================================================

Option Explicit

Private Sub svgOn_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If GetSettingBoolean(SETTINGS_POST_PROCESS_SVG) Then visible = True
End Sub

Private Sub svgOff_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If Not GetSettingBoolean(SETTINGS_POST_PROCESS_SVG) Then visible = True
End Sub

Private Sub svgTurnOn_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If Not GetSettingBoolean(SETTINGS_POST_PROCESS_SVG) Then visible = True
End Sub

Private Sub svgTurnOff_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If GetSettingBoolean(SETTINGS_POST_PROCESS_SVG) Then visible = True
End Sub

Private Sub svgTurnOn_onAction(ByVal control As IRibbonControl)
    Dim response As VbMsgBoxResult
    response = MsgBox(GetSupertip("svgRiskWarning"), _
                      vbYesNo + vbExclamation + vbDefaultButton2, _
                      GetScreentip("svgRiskWarning"))
    
    If response = vbYes Then
        SettingsSheet.Range(SETTINGS_POST_PROCESS_SVG).value = TOGGLE_YES
        SvgRefreshButtons
    End If
End Sub

Private Sub svgTurnOff_onAction(ByVal control As IRibbonControl)
    SettingsSheet.Range(SETTINGS_POST_PROCESS_SVG).value = TOGGLE_NO
    SvgRefreshButtons
End Sub

Private Sub SvgRefreshButtons()
    InvalidateRibbonControl "svgOn"
    InvalidateRibbonControl "svgOff"
    InvalidateRibbonControl "svgTurnOn"
    InvalidateRibbonControl "svgTurnOff"
End Sub

' ===========================================================================
' Callbacks for Help

Private Sub svgHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSvgTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for svgEditCell

Private Sub svgEditCell_onAction(ByVal control As IRibbonControl)
    CellValueEditForm.show
End Sub

Private Sub svgEditCell_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    If ActiveSheet.name <> SvgSheet.name Then
        Enabled = False
    ElseIf Selection.Cells.Count > 1 Then
        Enabled = False
    ElseIf ActiveCell.HasFormula Then
        Enabled = False
    Else
        Enabled = True
    End If
End Sub

Private Sub svgClipboard_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

Private Sub svgClipboard_onAction(ByVal control As IRibbonControl)
#If Not Mac Then
    
    If ClipBoard_SetData(ActiveCell.value) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySvgSuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySvgFailed"), 5
    End If
    
#End If
End Sub


