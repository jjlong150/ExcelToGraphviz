Attribute VB_Name = "modRibbonTabData"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabData
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon Callbacks
'
' ROLE:
'   Callback hub for the Data Ribbon Tab. Manages label/tooltip visibility
'   modes, JSON preview and viewer selection, help navigation, and renderer
'   selection. Provides the UI glue between user actions and the workbook's
'   data-driven configuration model.
'
' RESPONSIBILITIES:
'   - Manage blank/omit modes for node labels, cluster labels, node tooltips,
'     cluster tooltips, and edge tooltips.
'   - Control JSON preview actions:
'       o Generate JSON for the current view
'       o Display token/character estimates
'       o Launch or reopen the JSON viewer
'       o Enforce radio-button behavior for viewer selection
'   - Handle "Publish to File" enablement logic for DOT, Graphviz, and
'     Knowledge Graph JSON output.
'   - Manage renderer selection (Cairo, GD, GDI+, Quartz) and enforce
'     mutual exclusivity across rendering engines.
'   - Provide help-link navigation for Data-tab documentation.
'   - Coordinate AutoDraw mode (Auto/Manual) and reflect state in Ribbon UI.
'
' INTERACTIONS:
'   - Named Ranges: SETTINGS_BLANK_NODE_LABELS, BlankClusterLabel,
'     BlankClusterTooltip, BlankNodeTooltip, BlankEdgeTooltip,
'     SETTINGS_JSON_VIEWER, SETTINGS_RUN_MODE, SETTINGS_RENDER_*,
'     SETTINGS_PUBLISH_*, HelpURLDataTab.
'   - Modules: modCreateGraph, modCreateDot, modCreateJson,
'     modUtilityRibbon, modUtilitySettings, workspace utilities.
'   - Sheets: SettingsSheet, StylesSheet.
'
' CROSS-PLATFORM NOTES:
'   - macOS hides clipboard-related and GVEdit controls.
'   - Renderer availability varies by image type and platform.
'
' ERROR HANDLING:
'   - Uses OptimizeCode_Begin/End to reduce flicker and improve responsiveness
'     during long-running operations.
'
' RELATED WIKI PAGES:
'   - Data Ribbon Tab
'   - Label & Tooltip Modes
'   - JSON Viewer Options
'   - Rendering Engines (Cairo/GD/GDI+/Quartz)
'   - Publish to File Workflow
' =============================================================================

Option Explicit

' ===========================================================================
' Callbacks for blankNodeLabels

Private Sub blankNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_NODE_LABELS).value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_DEFAULT
    AutoDraw
End Sub

Private Sub blankNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_NODE_LABELS, TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for blankClusterLabels

Sub blankClusterLabels_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankClusterLabel").value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl "blankClusterLabels"
    InvalidateRibbonControl "omitClusterLabels"
    AutoDraw
End Sub

Sub blankClusterLabels_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankClusterLabel", TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for omitClusterLabels

Sub omitClusterLabels_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankClusterLabel").value = TOGGLE_BLANK_USE_DEFAULT

    InvalidateRibbonControl "blankClusterLabels"
    InvalidateRibbonControl "omitClusterLabels"
    AutoDraw
End Sub

Sub omitClusterLabels_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankClusterLabel", TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for blankClusterTooltips

Sub blankClusterTooltips_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankClusterTooltip").value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl "blankClusterTooltips"
    InvalidateRibbonControl "omitClusterTooltips"
    AutoDraw
End Sub

Sub blankClusterTooltips_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankClusterTooltip", TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for omitClusterTooltips

Sub omitClusterTooltips_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankClusterTooltip").value = TOGGLE_BLANK_USE_DEFAULT

    InvalidateRibbonControl "blankClusterTooltips"
    InvalidateRibbonControl "omitClusterTooltips"
    AutoDraw
End Sub

Sub omitClusterTooltips_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankClusterTooltip", TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for blankNodeTooltips

Sub blankNodeTooltips_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankNodeTooltip").value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl "blankNodeTooltips"
    InvalidateRibbonControl "omitNodeTooltips"
    AutoDraw
End Sub

Sub blankNodeTooltips_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankNodeTooltip", TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for omitNodeTooltips

Sub omitNodeTooltips_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankNodeTooltip").value = TOGGLE_BLANK_USE_DEFAULT

    InvalidateRibbonControl "blankNodeTooltips"
    InvalidateRibbonControl "omitNodeTooltips"
    AutoDraw
End Sub

Sub omitNodeTooltips_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankNodeTooltip", TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for blankEdgeTooltips

Sub blankEdgeTooltips_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankEdgeTooltip").value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl "blankEdgeTooltips"
    InvalidateRibbonControl "omitEdgeTooltips"
    AutoDraw
End Sub

Sub blankEdgeTooltips_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankEdgeTooltip", TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for omitEdgeTooltips

Sub omitEdgeTooltips_onAction(control As IRibbonControl, pressed As Boolean)
    SettingsSheet.Range("BlankEdgeTooltip").value = TOGGLE_BLANK_USE_DEFAULT

    InvalidateRibbonControl "blankEdgeTooltips"
    InvalidateRibbonControl "omitEdgeTooltips"
    AutoDraw
End Sub

Sub omitEdgeTooltips_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = getPressed(SettingsSheet.name, "BlankEdgeTooltip", TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for showJson

Sub showJson_onAction(control As IRibbonControl)
    ' Save the current cursor
    Dim saveCursor As Long
    saveCursor = Application.Cursor
    
    ' Show the hourglass wait cursor
    Application.Cursor = xlWait
    DoEvents
    
    ' Generate the knowledge graph
    Dim jsonText As String
    jsonText = JsonCurrentViewToString(-1)
    
    ' Display sizes in the status bar
    Dim statusBarText As String
    statusBarText = GetMessage("statusbarKnowledgeGraphEstTokens")
    statusBarText = replace(statusBarText, "{characters}", Len(jsonText), 1, 1, vbTextCompare)
    statusBarText = replace(statusBarText, "{tokens}", GetTokenEstimate(jsonText), 1, 1, vbTextCompare)
    UpdateStatusBarForNSeconds statusBarText, 10
    
    ' Show the json in the browser
    ShowKnowledgeGraph jsonText
    
    ' Restore cursor to value we started with
    Application.Cursor = saveCursor
End Sub

Private Sub showJson_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = True
End Sub

' Callback for toggle buttons getPressed (Checks the active option)
Private Sub showJsonOptions_getToggleState(control As IRibbonControl, ByRef returnedVal)
    Dim selectedMethod As String
    selectedMethod = SettingsSheet.Range(SETTINGS_JSON_VIEWER).Value2
    
    If LCase$(Mid$(control.id, Len("jsonViewer") + 1)) = selectedMethod Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

' Callback for toggle buttons onAction (Enforces radio button behavior)
Private Sub showJsonOptions_onToggleChange(control As IRibbonControl, pressed As Boolean)
    ' Update the selection tracking variable
    Dim selectedMethod As String
    selectedMethod = LCase$(Mid$(control.id, Len("jsonViewer") + 1))
    SettingsSheet.Range(SETTINGS_JSON_VIEWER).Value2 = selectedMethod
    
    ' Force the ribbon to refresh and update the checkmarks visually
    InvalidateRibbonControl "jsonViewerNew"
    InvalidateRibbonControl "jsonViewerShared"
    InvalidateRibbonControl "jsonReopen"
End Sub

Private Sub jsonReopen_onAction(control As IRibbonControl)
    ReopenGraphViewer
End Sub

Private Sub jsonReopen_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
    If SettingsSheet.Range(SETTINGS_JSON_VIEWER).Value2 = "shared" And SharedBrowserWasOpened() Then
        Enabled = True
    End If
End Sub

Private Sub publishDot_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub publishDot_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(control.id).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    InvalidateRibbonControl "graphToFile"
    InvalidateRibbonControl "graphAllViewsToFile"
End Sub

Private Sub publishGraphviz_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub publishGraphviz_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(control.id).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    InvalidateRibbonControl "graphToFile"
    InvalidateRibbonControl "graphAllViewsToFile"
End Sub

Private Sub publishGraphviz_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetLabel(control.id) & " (." & SettingsSheet.Range(SETTINGS_FILE_FORMAT).value & ")"
End Sub

Private Sub publishKnowledge_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub publishKnowledge_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(control.id).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    InvalidateRibbonControl "graphToFile"
    InvalidateRibbonControl "graphAllViewsToFile"
End Sub

' ===========================================================================
' Callbacks for graphToFile

' ==========================================================================
' CALLBACK: graphToFile_onAction
'
' PURPOSE:
'   Triggers the "Publish to File" workflow for the currently selected view.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN RESOLUTION: Identifies the active View column using
'      'GetSettingColNum' to define the render boundaries.
'   2. UI FEEDBACK: Sets the 'xlWait' cursor and executes 'DoEvents' to
'      ensure the UI remains responsive during the initial handshake.
'   3. EXECUTION:
'      - Wraps the call in 'OptimizeCode_Begin/End' to maximize performance.
'      - Invokes 'CreateGraphFile' to handle DOT generation and binary execution.
'   4. STATE RESTORATION: Reverts the cursor to 'xlDefault' once the
'      file has been successfully written to disk.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Graph to File button.
'   - Layer: UI / Orchestration.
' ==========================================================================
Private Sub graphToFile_onAction(ByVal control As IRibbonControl)
    Dim firstColumn As Long
    Dim lastColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    lastColumn = firstColumn
    
    ' Save the current cursor
    Dim saveCursor As Long
    saveCursor = Application.Cursor
    
    ' Show the hourglass wait cursor
    Application.Cursor = xlWait
    DoEvents
       
    ' Disable screen updating and events
    OptimizeCode_Begin
    
    ' Publish graph diagram
    If GetSettingBoolean(SETTINGS_PUBLISH_GRAPHVIZ) Then
        CreateGraphFile firstColumn, lastColumn
    End If
    
    ' Publish DOT source code
    If GetSettingBoolean(SETTINGS_PUBLISH_DOT) Then
        CreateDotFiles firstColumn, lastColumn
    End If
    
    ' Publish Knowledge Graph JSON
    If GetSettingBoolean(SETTINGS_PUBLISH_KNOWLEDGE) Then
        ' Get JSON formatting options
        Dim indent As Long
        indent = -1
        If Not GetSettingBoolean("KnowledgeMinify") Then
            indent = CLng(SettingsSheet.Range("KnowledgeIndent").value)
        End If
        
        ' Create the Knowledge Graph files
        CreateJsonFiles firstColumn, lastColumn, indent
    End If
    
    ' Enable screen updating and events
    OptimizeCode_End
    
    ' Restore cursor to value we started with
    Application.Cursor = saveCursor
End Sub

' ==========================================================================
' CALLBACK: graphToFile_getEnabled
'
' PURPOSE:
'   Determines if the "Graph to File" button should be active based on
'   whether a valid data View has been selected in the Style Gallery.
'
' TECHNICAL WORKFLOW:
'   1. VALIDATION: Invokes 'IsAViewSpecified' to verify that the user has
'      chosen a specific view (column) for rendering.
'   2. LOGICAL RETURN: Sets the 'pressed' (Enabled) state to TRUE only if
'       a view is active.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Tab Activation.
'   - UX Strategy: Prevents execution errors by disabling file export
'     functionality when no view context exists.
' ==========================================================================
Private Sub graphToFile_getEnabled(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If Not GetSettingBoolean(SETTINGS_PUBLISH_DOT) And _
       Not GetSettingBoolean(SETTINGS_PUBLISH_GRAPHVIZ) And _
       Not GetSettingBoolean(SETTINGS_PUBLISH_KNOWLEDGE) Then
       pressed = False
       Exit Sub
    End If
    
    pressed = Not (IsAViewSpecified() = False)
End Sub

' ===========================================================================
' Callbacks for graphAllViewsToFile

' ==========================================================================
' CALLBACK: graphAllViewsToFile_onAction
'
' PURPOSE:
'   The batch-processing entry point. Iterates through all defined "Views"
'   in the Style Gallery and exports each as a separate file.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Identifies the 'firstColumn' of the View gallery
'      using the 'SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW' setting.
'   2. BOUNDARY CALCULATION: Scans the header row of the 'Styles' sheet
'      to count non-empty View names, determining the 'lastColumn' index.
'   3. UI FEEDBACK: Activates the 'xlWait' cursor and executes 'DoEvents'
'      to maintain responsiveness during the initial calculation.
'   4. BATCH EXECUTION:
'      - Invokes 'OptimizeCode_Begin' to suppress UI updates.
'      - Calls 'CreateGraphFile' with the resolved column range.
'      - Invokes 'OptimizeCode_End' and restores the default cursor.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Publish All Views button.
'   - Strategy: Automates the production of multiple graph perspectives
'     (e.g., Logical, Physical, Security) in a single operation.
' ==========================================================================
Private Sub graphAllViewsToFile_onAction(ByVal control As IRibbonControl)

    Dim nonEmptyCellCount As Long
    Dim row As Long
    Dim col As Long
    Dim columnName As String
    Dim firstColumn As Long
    Dim lastColumn As Long
    
    row = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    nonEmptyCellCount = 0
    
    ' Get the configured location of the first view name column
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    
    ' Count the non-empty cells beginning at the first view column
    For col = firstColumn To GetLastColumn(StylesSheet.name, row)
        columnName = StylesSheet.Cells.item(row, col)
        If columnName <> vbNullString Then
            nonEmptyCellCount = nonEmptyCellCount + 1
        End If
    Next col

    ' Calaculate the absolute column number of the last view column
    lastColumn = firstColumn + nonEmptyCellCount - 1
    
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents
    
    ' Graph all the views
    OptimizeCode_Begin
    
    If GetSettingBoolean(SETTINGS_PUBLISH_GRAPHVIZ) Then
        CreateGraphFile firstColumn, lastColumn
    End If
    
    If GetSettingBoolean(SETTINGS_PUBLISH_DOT) Then
        CreateDotFiles firstColumn, lastColumn
    End If
    
    If GetSettingBoolean(SETTINGS_PUBLISH_KNOWLEDGE) Then
        CreateJsonFiles firstColumn, lastColumn, 2
    End If
    
    OptimizeCode_End
    
    DoEvents
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
End Sub

' ===========================================================================
' Callbacks for graphToWorksheet

' ==========================================================================
' CALLBACK: graphToWorksheet_onAction
'
' PURPOSE:
'   UI entry point that triggers the standard in-workbook graph rendering
'   process.
'
' TECHNICAL WORKFLOW:
'   1. REDIRECTION: Hands off execution to 'CreateGraphWorksheetQuickly'.
'   2. UI CONTEXT: Inherits the performance optimizations and wait-cursor
'      feedback defined in the target procedure.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Render Graph button.
'   - Strategy: Decouples the Ribbon callback from the core rendering
'     logic to allow for shared use by hotkeys.
' ==========================================================================
Private Sub graphToWorksheet_onAction(ByVal control As IRibbonControl)
    CreateGraphWorksheetQuickly
End Sub

' ==========================================================================
' CALLBACK: graphToWorksheet_getEnabled
'
' PURPOSE:
'   Controls the availability of the primary "Render Graph" button on the
'   Ribbon based on the workbook's current configuration state.
'
' TECHNICAL WORKFLOW:
'   1. VALIDATION: Calls 'IsAViewSpecified' to check if a specific view
'      column in the Style Gallery has been selected from the dropdown list.
'   2. UI FEEDBACK: Sets 'Enabled' to TRUE only if a view context exists,
'      preventing the user from attempting a render without a defined style set.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Tab Activation.
'   - UX Strategy: Enforces the "View-First" workflow required for
'     successful Graphviz source generation.
' ==========================================================================
Private Sub graphToWorksheet_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = IsAViewSpecified()
End Sub

' ===========================================================================
' Callbacks for graphAuto

' ==========================================================================
' CALLBACK: graphAuto_onAction
'
' PURPOSE:
'   Toggles the global execution mode between 'Auto' (Live Preview) and
'   'Manual' via the Ribbon interface.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_RUN_MODE' named range using
'      the 'Toggle' helper to map the Ribbon's boolean state to project
'      constants.
'   2. REACTIVITY: Immediately invokes 'AutoDraw' so that if the user
'      enables Auto mode, the graph refreshes to reflect current data.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> AutoDraw Toggle.
'   - Impact: Determines if 'Worksheet_Change' events trigger the renderer.
' ==========================================================================
Private Sub graphAuto_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = Toggle(pressed, TOGGLE_AUTO, TOGGLE_MANUAL)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: graphAuto_getPressed
'
' PURPOSE:
'   Ensures the Ribbon's 'AutoDraw' toggle visually reflects the current
'   workbook setting.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Evaluates the 'SETTINGS_RUN_MODE' named range.
'   2. UI FEEDBACK: Returns TRUE if the setting matches 'TOGGLE_AUTO'.
' ==========================================================================
Private Sub graphAuto_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_AUTO
End Sub

' ===========================================================================
' Callbacks for graphAuto

Private Sub openAfterPublish_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(control.id).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

Private Sub openAfterPublish_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(control.id).value = TOGGLE_YES
End Sub

' ==========================================================================
' CALLBACK: dataHelp_onAction
'
' PURPOSE:
'   Redirects the user to the official Graphviz documentation or a
'   project-specific help page for the Graphviz Tab.
'
' TECHNICAL WORKFLOW:
'   1. URL RESOLUTION: Retrieves the target URL from the "HelpURLDataTab"
'      named range on the Settings worksheet.
'   2. NAVIGATION: Invokes 'ActiveWorkbook.FollowHyperlink' to launch the
'      link in the user's default web browser.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Help button (Last group).
'   - Strategy: Centralizes documentation links within the workbook's
'     Settings sheet to allow for URL updates without code modification.
' ==========================================================================
Private Sub dataHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLDataTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for renderCairo

Private Sub renderCairo_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub renderCairo_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RENDER_CAIRO).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    
    SettingsSheet.Range(SETTINGS_RENDER_GD).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_GD
    
    SettingsSheet.Range(SETTINGS_RENDER_GDIPLUS).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_GDIPLUS
    
    SettingsSheet.Range(SETTINGS_RENDER_QUARTZ).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_QUARTZ
End Sub

Private Sub renderCairo_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim imageFileType As String
    imageFileType = LCase$(GetSettingsForGraph.imageTypeFile)
    
    Select Case imageFileType
        Case "bmp", "gif", "jpg", "pdf", "png", "ps", "svg", "tiff"
            returnedVal = True
        Case Else
            returnedVal = False
    End Select
End Sub

' ===========================================================================
' Callbacks for renderGD

Private Sub renderGD_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub renderGD_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RENDER_GD).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    
    SettingsSheet.Range(SETTINGS_RENDER_CAIRO).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_CAIRO
    
    SettingsSheet.Range(SETTINGS_RENDER_GDIPLUS).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_GDIPLUS
    
    SettingsSheet.Range(SETTINGS_RENDER_QUARTZ).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_QUARTZ
End Sub

Private Sub renderGD_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim imageFileType As String
    imageFileType = LCase$(GetSettingsForGraph.imageTypeFile)
    
    Select Case imageFileType
        Case "gif", "jpg", "png"
            returnedVal = True
        Case Else
            returnedVal = False
    End Select
End Sub

' ===========================================================================
' Callbacks for renderGDIPlus

Private Sub renderGDIPlus_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub renderGDIPlus_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RENDER_GDIPLUS).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    
    SettingsSheet.Range(SETTINGS_RENDER_CAIRO).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_CAIRO
    
    SettingsSheet.Range(SETTINGS_RENDER_GD).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_GD
    
    SettingsSheet.Range(SETTINGS_RENDER_QUARTZ).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_QUARTZ
End Sub

Private Sub renderGDIPlus_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim imageFileType As String
    imageFileType = LCase$(GetSettingsForGraph.imageTypeFile)
    
    Select Case imageFileType
        Case "bmp", "gif", "jpg", "png", "tiff"
            returnedVal = True
        Case Else
            returnedVal = False
    End Select
End Sub

' ===========================================================================
' Callbacks for renderQuartz

Private Sub renderQuartz_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Private Sub renderQuartz_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RENDER_QUARTZ).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    
    SettingsSheet.Range(SETTINGS_RENDER_CAIRO).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_CAIRO
    
    SettingsSheet.Range(SETTINGS_RENDER_GD).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_GD
    
    SettingsSheet.Range(SETTINGS_RENDER_GDIPLUS).value = False
    InvalidateRibbonControl RIBBON_CTL_RENDER_GDIPLUS
End Sub

Private Sub renderQuartz_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim imageFileType As String
    imageFileType = LCase$(GetSettingsForGraph.imageTypeFile)
    
    Select Case imageFileType
        Case "bmp", "gif", "jpg", "pdf", "png", "tiff"
            returnedVal = True
        Case Else
            returnedVal = False
    End Select
End Sub

'=============================================================================
' GetRenderer
'
' Determines the active Graphviz rendering engine based on user-selected
' settings. Returns the renderer keyword used by downstream graph-generation
' routines.
'
' BEHAVIOR
'   o Checks renderer-selection flags in priority order:
'         Cairo -> GD -> GDI+ -> Quartz
'   o Returns the corresponding renderer name as a lowercase string.
'   o Returns an empty string if no renderer option is enabled.
'
' RETURNS
'   o "cairo", "gd", "gdiplus", or "quartz"
'   o Empty string if no renderer setting is active.
'
'=============================================================================
Public Function GetRenderer() As String
    If GetSettingBoolean(SETTINGS_RENDER_CAIRO) Then
        GetRenderer = "cairo"
    ElseIf GetSettingBoolean(SETTINGS_RENDER_GD) Then
        GetRenderer = "gd"
    ElseIf GetSettingBoolean(SETTINGS_RENDER_GDIPLUS) Then
        GetRenderer = "gdiplus"
    ElseIf GetSettingBoolean(SETTINGS_RENDER_QUARTZ) Then
        GetRenderer = "quartz"
    Else
        GetRenderer = ""
    End If
End Function
       
'=============================================================================
' JsonCurrentViewToString
'
' Generates a JSON string for the currently selected view, using the active
' runtime settings, enabled styles, and graph options. The JSON is returned
' directly rather than written to disk.
'
' PARAMETERS
'   indent   - Long
'       Indentation level for JSON output. Use a negative value for compact
'       (no-whitespace) formatting.
'
' BEHAVIOR
'   o Determines the active view column from settings.
'   o Validates the presence of the data worksheet.
'   o Retrieves the current view name and exposes it for graph generation.
'   o Refreshes graph options that may reference the view name.
'   o Caches enabled styles for the current view.
'   o Generates JSON via GetGraphJson.
'
' SIDE EFFECTS
'   o Updates SettingsSheet("ViewNameLabel") during processing.
'
' RETURNS
'   o String containing the JSON representation of the current view.
'
'=============================================================================
Private Function JsonCurrentViewToString(indent As Long) As String
    Dim viewColumn As Long
    viewColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    ' Clear the status bar
    ClearStatusBar
    
    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(GetDataWorksheetName())

    If Not WorksheetExists(ini.data.worksheetName) Then
        EmitMessage GetMessage("msgboxNoDataToGraph")
        Exit Function
    End If

    ' Get the name of the view
    Dim viewName As String
    viewName = StylesSheet.Cells.item(ini.styles.headingRow, viewColumn).value
        
    ' Expose the view name so it can be used as data in the graph
    SettingsSheet.Range("ViewNameLabel").value = viewName
        
    ' View name might be referenced in the graph options, so refresh the value
    ini.graph.options = Trim$(SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).value)
        
    ' Cache the names included in the current view
    Dim viewStyles As Dictionary
    Set viewStyles = CacheEnabledStyles(ini, viewColumn)

    ' Generate the JSON for the current view
    Dim graphJson As String
    graphJson = GetGraphJson(ini, viewName, viewStyles, indent)
        
    ' Sync up settings with dropdown choice
    SettingsSheet.Range("ViewNameLabel").value = SettingsSheet.Range("ViewName").value

    JsonCurrentViewToString = graphJson
End Function


