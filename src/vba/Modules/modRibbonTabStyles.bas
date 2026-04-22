Attribute VB_Name = "modRibbonTabStyles"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabStyles
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Callback bridge for the "Styles" Ribbon Tab, providing worksheet-level
'   style preview actions, suffix configuration, and integration with the
'   full Style Designer. Acts as the UI surface for managing and inspecting
'   style definitions stored on the Styles worksheet.
'
' RESPONSIBILITIES:
'   - Dispatch IRibbonControl callbacks for all Styles tab controls.
'   - Trigger style previews (single row, all rows) and clear preview output.
'   - Persist style suffix settings (open/close markers) via named ranges.
'   - Enable "Edit Style" only when the active row represents a valid
'     style object (Node / Edge / Subgraph).
'   - Bridge to the Style Designer via RestoreStyleDesigner.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml (control IDs -> callbacks).
'   - Named Ranges:
'       StylesSuffixOpen, StylesSuffixClose,
'       SETTINGS_STYLES_COL_OBJECT_TYPE, HelpURLStylesTab.
'   - Worksheets: StylesSheet, SettingsSheet, DataSheet.
'   - Modules: ClearStylesPreview, PreviewStyleForCurrentRow,
'              GenerateStylesPreviewAll, RestoreStyleDesigner.
'   - Global State: internalMyRibbon (via Ribbon invalidation in parent tabs).
'
' CROSS-PLATFORM NOTES:
'   - Fully supported on Windows and macOS.
'   - All actions rely on worksheet operations and hyperlink navigation.
'
' ERROR HANDLING:
'   - Localized checks ensure "Edit Style" is only enabled for valid rows.
'   - Callback signatures follow IRibbonControl requirements.
'
' RELATED WIKI PAGES:
'   - Styles & the Style Gallery
'   - Style Designer Ribbon Tab
'   - Working with the Data Worksheet
'   - Worksheet Architecture & Named Ranges
' =============================================================================

Option Explicit

Private Sub stylesClear_onAction(ByVal control As IRibbonControl)
    ClearStylesPreview
End Sub

Private Sub stylesPreview_onAction(ByVal control As IRibbonControl)
    PreviewStyleForCurrentRow
End Sub

Private Sub stylesPreviewAll_onAction(ByVal control As IRibbonControl)
    StylesSheet.Activate
    ClearStylesPreview
    GenerateStylesPreviewAll
    ClearStatusBar
End Sub

' ===========================================================================
' Callbacks for stylesSuffixBegin

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixBegin_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range("StylesSuffixOpen").value = Text
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixBegin_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range("StylesSuffixOpen"))
End Sub

' ===========================================================================
' Callbacks for stylesSuffixEnd

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixEnd_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range("StylesSuffixClose").value = Text
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixEnd_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range("StylesSuffixClose"))
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Private Sub stylesHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLStylesTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for stylesEdit

'@Ignore ParameterNotUsed
Private Sub stylesEdit_onAction(ByVal control As IRibbonControl)
    RestoreStyleDesigner
End Sub

'@Ignore ParameterNotUsed
Private Sub stylesEdit_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False

    If ActiveSheet.name <> StylesSheet.name Then Exit Sub
    If Not TypeOf Selection Is Range Then Exit Sub
    If Selection.rows.count <> 1 Then Exit Sub

    Dim row As Long
    row = Selection.row

    Dim typeCol As Long
    typeCol = GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)

    Dim styleType As String
    styleType = StylesSheet.Cells(row, typeCol).value

    If Not (styleType = TYPE_NODE Or styleType = TYPE_EDGE Or styleType = TYPE_SUBGRAPH_OPEN) Then Exit Sub

    Enabled = True
End Sub
