Attribute VB_Name = "modRibbonTabGraphviz"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabGraphviz
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Primary callback bridge for the "Graphviz" Ribbon Tab, coordinating layout
'   engines, graph type, spline routing, workspace operations, and rendering
'   settings.
'
' RESPONSIBILITIES:
'   - Manage layout engine selection and synchronize Ribbon state.
'   - Control graph orientation (Directed/Undirected).
'   - Manage spline routing, rankdir, output order, and engine-specific options.
'   - Handle workspace operations (Clear Data, Clear Graphs, Clear Messages).
'   - Manage column visibility for the Data worksheet.
'   - Trigger AutoDraw for real-time reactivity.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Named Ranges: SETTINGS_GRAPHVIZ_ENGINE, SETTINGS_GRAPH_TYPE, SETTINGS_SPLINES, etc.
'   - Modules: modCreateGraph, modUtilityRibbon, workspace utilities.
'
' CROSS-PLATFORM NOTES:
'   - macOS hides certain controls (clipboard, GVEdit).
'
' ERROR HANDLING:
'   - Uses OptimizeCode blocks to reduce flicker.
'
' RELATED WIKI PAGES:
'   - Graphviz Ribbon Tab
'   - Event-Driven Architecture & AutoDraw
' =============================================================================

Option Explicit

Private Const MAX_ZOOM As Long = 150
Private Const MIN_ZOOM As Long = 5
Private Const ZOOM_STEP As Long = 5

' ===========================================================================
' Callbacks for Show/Hide Labels

' ==========================================================================
' CALLBACK: showColumn_onAction
'
' PURPOSE:
'   Toggles the visibility of a specific data column in the active worksheet
'   via Ribbon checkbox controls.
'
' TECHNICAL WORKFLOW:
'   1. UI RESET: Clears existing graphs to avoid visual artifacts during
'      column width shifts.
'   2. STATE PERSISTENCE: Updates the corresponding setting in 'SettingsSheet'
'      based on the 'pressed' state of the Ribbon control.
'   3. EXECUTION: Invokes 'ShowHideDataColumn' to perform the physical
'      hiding/unhiding logic.
'   4. REACTIVITY: Triggers 'AutoDraw' to refresh the graph layout.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showColumn_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    ClearWorksheetGraphs
    SettingsSheet.Range(control.id).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    ShowHideDataColumn (control.id)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: showColumn_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon checkbox state with the actual visibility of
'   worksheet columns during UI initialization or invalidation.
'
' TECHNICAL WORKFLOW:
'   1. SYNC: Ensures the worksheet column state matches the stored setting.
'   2. UI FEEDBACK: Returns TRUE if the associated setting is "Show".
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showColumn_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ShowHideDataColumn (control.id)
    returnedVal = GetSettingBoolean(control.id)
End Sub

' ==========================================================================
' PROCEDURE: ShowHideDataColumn
'
' PURPOSE:
'   The engine behind column management. Translates Ribbon control IDs into
'   physical Excel column operations.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN MAPPING: Uses a 'Select Case' to map Ribbon constants (e.g.,
'      'RIBBON_CTL_SHOW_COMMENT') to specific column letters stored in Settings.
'   2. WORKSHEET ACTIVATION: Switches focus to the target data sheet.
'   3. VISIBILITY TOGGLE: Sets the 'Hidden' property of the 'EntireColumn'
'      based on the saved boolean setting.
'   4. FOCUS MANAGEMENT: Returns the selection to the header row of the
'      affected column for a consistent user experience.
'
' TECHNICAL NOTES:
'   - Layer: UI / Workspace Management.
'   - Strategy: Centralizes column mapping to ensure the Ribbon can manage
'     any arbitrary column layout defined in Settings.
' ==========================================================================
Public Sub ShowHideDataColumn(ByVal columnId As String)
    Dim ShowColumn As Boolean
    Dim columnRange As String
    Dim col As String
    
    OptimizeCode_Begin
    
    ' Map the menu item to the column name
    Select Case columnId
        Case RIBBON_CTL_SHOW_COMMENT
            col = SettingsSheet.Range(SETTINGS_DATA_COL_COMMENT).value
        Case RIBBON_CTL_SHOW_ITEM
            col = SettingsSheet.Range(SETTINGS_DATA_COL_ITEM).value
        Case RIBBON_CTL_SHOW_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL).value
        Case RIBBON_CTL_SHOW_OUTSIDE_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL_X).value
        Case RIBBON_CTL_SHOW_TAIL_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL_TAIL).value
        Case RIBBON_CTL_SHOW_HEAD_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL_HEAD).value
        Case RIBBON_CTL_SHOW_TOOLTIP
            col = SettingsSheet.Range(SETTINGS_DATA_COL_TOOLTIP).value
        Case RIBBON_CTL_SHOW_IS_RELATED_TO_ITEM
            col = SettingsSheet.Range(SETTINGS_DATA_COL_IS_RELATED_TO).value
        Case RIBBON_CTL_SHOW_STYLE
            col = SettingsSheet.Range(SETTINGS_DATA_COL_STYLE).value
        Case RIBBON_CTL_SHOW_EXTRA_STYLE_ATTRIBUTES
            col = SettingsSheet.Range(SETTINGS_DATA_COL_EXTRA_ATTRIBUTES).value
        Case RIBBON_CTL_SHOW_MESSAGES
            col = SettingsSheet.Range(SETTINGS_DATA_COL_ERROR_MESSAGES).value
    End Select
    
    ' Activate the "data" worksheet
    ActiveWorkbook.Sheets.[_Default](GetDataWorksheetName()).Activate
    
    ' Compose the column range to show/hide
    columnRange = col & ":" & col
    
    ' Show/Hide column according the saved value that corresponds to the check mark in the dropdown list
    ActiveSheet.columns(columnRange).Select
    ShowColumn = GetSettingBoolean(columnId)
    Selection.EntireColumn.Hidden = Not ShowColumn
    
    ' Put the focus on the heading
    ActiveSheet.Range(col & CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))).Select

    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for data worksheet

' ==========================================================================
' CALLBACK: clearData_onAction
'
' PURPOSE:
'   Performs a "Hard Reset" of the active workspace, purging all data rows,
'   rendered imagery, and diagnostic source code.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT RESOLUTION: Identifies the target worksheet via
'      'GetDataWorksheetName' to ensure system sheets aren't cleared.
'   2. UI NORMALIZATION: Restores standard row heights and activates the sheet.
'   3. MULTI-LAYER PURGE:
'      - 'ClearDataWorksheet': Wipes the logical data rows.
'      - 'ClearWorksheetGraphs': Deletes OLE picture objects/shapes.
'      - 'ClearSourceWorksheet' / 'ClearSourceForm': Resets DOT source buffers.
'   4. UX CLEANUP: Clears the Excel StatusBar and exits optimization mode.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Clear Data button.
'   - Layer: UI / Workspace Management.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub clearData_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    
    Dim worksheetName As String
    worksheetName = GetDataWorksheetName()
    
    ActiveWorkbook.Sheets.[_Default](worksheetName).Activate
    ActiveWorkbook.Sheets.[_Default](worksheetName).rows.UseStandardHeight = True

    ClearDataWorksheet (worksheetName)
    ClearWorksheetGraphs
    ClearSourceWorksheet
    ClearSourceForm
    Application.StatusBar = False
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for includeImagePath

'@Ignore ParameterNotUsed
Public Sub includeImagePath_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_INCLUDE_IMAGE_PATH).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub includeImagePath_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_INCLUDE_IMAGE_PATH)
End Sub

' ===========================================================================
' Callbacks for addOptions

'@Ignore ParameterNotUsed
Public Sub addOptions_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_APPEND_OPTIONS).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub addOptions_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_APPEND_OPTIONS)
End Sub

' ===========================================================================
' Callbacks for addTimestamp

'@Ignore ParameterNotUsed
Public Sub addTimestamp_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_APPEND_TIMESTAMP).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub addTimestamp_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_APPEND_TIMESTAMP)
End Sub

' ===========================================================================
' Callbacks for blankNodeLabels

'@Ignore ParameterNotUsed
Public Sub blankNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_NODE_LABELS).value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub blankNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_NODE_LABELS, TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for defaultNodeLabels

'@Ignore ParameterNotUsed
Public Sub defaultNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_NODE_LABELS).value = TOGGLE_BLANK_USE_DEFAULT
    
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub defaultNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_NODE_LABELS, TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for blankEdgeLabels

'@Ignore ParameterNotUsed
Public Sub blankEdgeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_EDGE_LABELS).value = TOGGLE_BLANK_USE_BLANK
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub blankEdgeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_EDGE_LABELS, TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for defaultEdgeLabels

'@Ignore ParameterNotUsed
Public Sub defaultEdgeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_EDGE_LABELS).value = TOGGLE_BLANK_USE_DEFAULT
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub defaultEdgeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_EDGE_LABELS, TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for clearMessages

'@Ignore ParameterNotUsed
Public Sub clearMessages_onAction(ByVal control As IRibbonControl)
    ClearErrors
End Sub

' ===========================================================================
' Callbacks for clearWorksheetGraphs

'@Ignore ParameterNotUsed
Public Sub clearWorksheetGraphs_onAction(ByVal control As IRibbonControl)
    ClearWorksheetGraphs
End Sub

' ===========================================================================
' Callbacks for directed

' ==========================================================================
' CALLBACK: directed_onAction
'
' PURPOSE:
'   Toggles the graph's fundamental orientation between 'Directed' (digraph)
'   and 'Undirected' (graph) modes via the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE EVALUATION: Updates the 'SETTINGS_GRAPH_TYPE' named range based
'      on the 'pressed' state and the current existing value.
'   2. UI SYNCHRONIZATION: Explicitly triggers 'InvalidateRibbonControl' for
'      the 'Undirected' toggle to ensure the Ribbon buttons stay in sync.
'   3. REACTIVITY: Invokes 'AutoDraw' to immediately re-render the graph if
'      the user has "Auto" mode enabled.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Directed Toggle.
'   - Contract: Updates global Graphviz syntax (-> vs --).
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub directed_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        If SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED Then
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
        Else
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
        End If
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
    End If
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_UNDIRECTED
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: directed_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Directed' button state with the current
'   graph configuration.
'
' TECHNICAL WORKFLOW:
'   1. STATE COMPARISON: Polls the 'SETTINGS_GRAPH_TYPE' named range.
'   2. UI FEEDBACK: Returns TRUE to 'pressed' if the setting matches the
'      'TOGGLE_DIRECTED' constant, causing the button to appear active.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Tab Activation.
'   - Strategy: Ensures the visual UI consistently represents the
'     underlying DOT syntax (digraph vs. graph).
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub directed_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
End Sub

' ===========================================================================
' Callbacks for layout

' ==========================================================================
' PROCEDURE: RefreshLayoutGroup
'
' PURPOSE:
'   Force-refreshes the visual state of all layout engine controls on the
'   Ribbon to ensure the "Pressed" indicator matches the current settings.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Sequentially triggers 'InvalidateRibbonControl' for
'      the entire Layout group and every individual engine toggle (dot,
'      neato, sfdp, etc.).
'   2. CALLBACK TRIGGER: Forces Excel to re-execute the 'getPressed'
'      callbacks for these IDs, ensuring only the active engine is
'      highlighted in the UI.
'
' TECHNICAL NOTES:
'   - Layer: UI / Ribbon Synchronization.
'   - Usage: Typically called after a layout engine change or a project
'     load to align the Ribbon with 'SETTINGS_GRAPHVIZ_ENGINE'.
' ==========================================================================
Private Sub RefreshLayoutGroup()
    InvalidateRibbonControl RIBBON_CTL_GROUP_LAYOUT
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_CIRCO
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_DOT
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_FDP
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_NEATO
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_OSAGE
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_PATCHWORK
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_SFDP
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_TWOPI
End Sub

' ==========================================================================
' PROCEDURE: RefreshGraphTypeGroup
'
' PURPOSE:
'   Synchronizes the Ribbon UI for the "Graph Type" category, ensuring the
'   'Directed' and 'Undirected' toggles visually reflect the current state.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Triggers 'InvalidateRibbonControl' for the parent
'      group and both the 'Directed' and 'Undirected' button IDs.
'   2. CALLBACK REFRESH: Forces the Ribbon engine to re-query the
'      'getPressed' state from the 'SETTINGS_GRAPH_TYPE' named range.
'
' TECHNICAL NOTES:
'   - Layer: UI / Ribbon Synchronization.
'   - Strategy: Prevents UI "ghosting" where the Ribbon appears out of
'     sync with the actual worksheet settings.
' ==========================================================================
Private Sub RefreshGraphTypeGroup()
    InvalidateRibbonControl RIBBON_CTL_GROUP_GRAPH_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_DIRECTED
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_UNDIRECTED
End Sub

' ==========================================================================
' PROCEDURE: RefreshLayoutParametersGroup
'
' PURPOSE:
'   Synchronizes the Ribbon's engine-specific parameter controls (Dimension,
'   Smoothing, Overlap, etc.) to match the current Graphviz layout engine context.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Triggers a comprehensive refresh across the
'      Algorithm group, including dynamic drop-downs, separators,
'      and boolean toggles (Newrank, Compound).
'   2. CONTEXT REFRESH: Forces the Ribbon to re-evaluate the 'getVisible'
'      and 'getEnabled' states, ensuring parameters irrelevant to the
'      current layout (e.g., 'Overlap' for 'Dot') are appropriately managed.
'
' TECHNICAL NOTES:
'   - Layer: UI / Ribbon Synchronization.
'   - Strategy: Essential for maintaining a clean UI by hiding or
'     disabling parameters that do not apply to the selected layout engine.
' ==========================================================================
Private Sub RefreshLayoutParametersGroup()
    InvalidateRibbonControl RIBBON_CTL_ALGORITHM_GROUP
    InvalidateRibbonControl RIBBON_CTL_GRAPH_MODE
    InvalidateRibbonControl RIBBON_CTL_GRAPH_MODEL
    InvalidateRibbonControl RIBBON_CTL_GRAPH_SMOOTHING
    InvalidateRibbonControl RIBBON_CTL_ALGORITHM_GROUP_SEPARATOR1
    InvalidateRibbonControl RIBBON_CTL_ALGORITHM_GROUP_SEPARATOR2
    InvalidateRibbonControl RIBBON_CTL_GRAPH_DIM
    InvalidateRibbonControl RIBBON_CTL_GRAPH_DIMEN
    InvalidateRibbonControl RIBBON_CTL_GRAPH_CLUSTER_RANK
    InvalidateRibbonControl RIBBON_CTL_NEWRANK
    InvalidateRibbonControl RIBBON_CTL_COMPOUND
    InvalidateRibbonControl RIBBON_CTL_OVERLAP
End Sub

' ==========================================================================
' CALLBACK: layout_onAction
'
' PURPOSE:
'   Handles the selection of a Graphviz layout engine (dot, neato, circo, etc.)
'   via the Ribbon and orchestrates a full UI state synchronization.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE:
'      - If pressed: Extracts the engine name from the control ID (e.g.,
'        stripping "layout" from "layoutNeato") and updates 'SETTINGS_GRAPHVIZ_ENGINE'.
'      - If unpressed: Reverts to the system default 'LAYOUT_DOT'.
'   2. UI REFRESH CASCADE: Invokes a series of 'Refresh...' procedures to
'      update the Ribbon. This ensures that parameters irrelevant to the
'      new engine (e.g., 'rankdir' for 'neato') are hidden or disabled.
'   3. PERFORMANCE: Wraps UI updates in 'OptimizeCode' blocks to minimize
'      Ribbon flickering.
'   4. REACTIVITY: Triggers 'AutoDraw' to re-render the graph immediately
'      using the newly selected engine.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Layout Engine Gallery/Buttons.
'   - Logic: Implements "Exclusive Selection" behavior for engine types.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub layout_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LCase$(Mid$(control.id, Len("layout") + 1))
    Else
        SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LAYOUT_DOT
    End If
    
    OptimizeCode_Begin
    RefreshLayoutGroup
    RefreshSplinesGroup
    RefreshGraphTypeGroup
    RefreshOutputorderGroup
    RefreshRankdirGroup
    RefreshOrderingGroup
    RefreshLayoutParametersGroup
    OptimizeCode_End
    
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: layout_getPressed
'
' PURPOSE:
'   Determines which layout engine button should appear "active" or
'   pressed on the Ribbon based on the current workbook settings.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Retrieves the active engine name from the
'      'SETTINGS_GRAPHVIZ_ENGINE' named range.
'   2. ALIAS MAPPING: Implements a 'Select Case' to handle backward
'      compatibility by mapping legacy layout aliases (compound, splines)
'      to the current button naming convention.
'   3. ID COMPARISON: Compares the resolved layout string against the
'      control's ID (stripping the "layout" prefix).
'   4. UI FEEDBACK: Returns TRUE to 'pressed' if the strings match,
'      visually anchoring the Ribbon state to the underlying configuration.
'
' TECHNICAL NOTES:
'   - Layer: UI / Ribbon State.
'   - Pattern: Dynamic Property Callback.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub layout_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    Dim layout As String
    layout = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
    
    ' Backward compatibility. Map layout aliases to buttons provided
    Select Case layout
        Case "compound": layout = "polyline"
        Case "splines": layout = "true"
        Case "line": layout = "false"
    End Select
    
    pressed = layout = LCase$(Mid$(control.id, Len("layout") + 1))
End Sub

' ==========================================================================
' CALLBACK: layoutOptions_getLabel
'
' PURPOSE:
'   Dynamically updates the text label of the Layout Options group or
'   control on the Ribbon to display the currently active engine.
'
' TECHNICAL WORKFLOW:
'   1. SETTINGS RETRIEVAL: Accesses the 'SETTINGS_GRAPHVIZ_ENGINE' named
'      range to fetch the active layout value (e.g., "dot", "neato").
'   2. STRING COMPOSITION: Prepends the "layout=" prefix to the engine
'      name to create an informative, contextual label.
'   3. UI UPDATE: Returns the concatenated string to the Ribbon's 'label'
'      property for immediate display.
'
' TECHNICAL NOTES:
'   - Trigger: Occurs during Ribbon invalidation or Tab activation.
'   - Strategy: Provides at-a-glance confirmation of the rendering
'     engine without needing to open a menu or gallery.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub layoutOptions_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    label = "layout=" & SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
End Sub

' ===========================================================================
' Callbacks for undirected

' ==========================================================================
' CALLBACK: undirected_onAction
'
' PURPOSE:
'   Toggles the graph's fundamental orientation to 'Undirected' (graph)
'   mode via the Ribbon interface.
'
' TECHNICAL WORKFLOW:
'   1. STATE EVALUATION: Updates the 'SETTINGS_GRAPH_TYPE' named range.
'      - If currently Undirected, it flips to Directed.
'      - If currently Directed, it flips to Undirected.
'   2. UI SYNCHRONIZATION: Triggers 'InvalidateRibbonControl' for the
'      'Directed' toggle to keep the visual state consistent.
'   3. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram if the
'      workbook is in Auto-run mode.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Undirected Toggle.
'   - Syntax Impact: Changes the DOT operator from '->' to '--'.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub undirected_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        If SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED Then
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
        Else
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
        End If
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
    End If
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_DIRECTED
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: undirected_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Undirected' button state with the underlying
'   workbook configuration.
'
' TECHNICAL WORKFLOW:
'   1. STATE COMPARISON: Queries the 'SETTINGS_GRAPH_TYPE' named range.
'   2. BOOLEAN EVALUATION: Returns TRUE to 'pressed' if the value matches
'      the 'TOGGLE_UNDIRECTED' constant.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Tab Activation.
'   - Strategy: Ensures the UI correctly highlights the active graph type.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub undirected_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
End Sub

' ==========================================================================
' CALLBACK: directed_getVisible
'
' PURPOSE:
'   Dynamically controls the visibility of the 'Directed' graph toggle
'   based on the currently selected Graphviz layout engine.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Evaluates the active engine in 'SETTINGS_GRAPHVIZ_ENGINE'.
'   2. ENGINE-SPECIFIC LOGIC:
'      - Forces 'visible = False' if using 'PATCHWORK', as tree-map layouts
'        do not support edge directionality.
'      - Defaults to 'visible = True' for all other engines (dot, neato, etc.).
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Layout Engine change.
'   - UX Strategy: Prevents user confusion by hiding controls that are
'     incompatible with specific layout algorithms.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub directed_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_PATCHWORK
            visible = False
        Case Else
            visible = True
    End Select
End Sub

' ===========================================================================
' Callbacks for splines

' ==========================================================================
' PROCEDURE: RefreshSplinesGroup
'
' PURPOSE:
'   Synchronizes the "Splines" (Edge Routing) group on the Ribbon to ensure
'   the visual "Pressed" state matches the current Graphviz configuration.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Sequentially triggers 'InvalidateRibbonControl' for
'      every supported spline routing algorithm (Ortho, Curved, Polyline, etc.).
'   2. CALLBACK REFRESH: Forces the Ribbon to re-execute the 'getPressed'
'      logic for these controls, ensuring the UI accurately reflects
'      the 'SETTINGS_SPLINES' named range value.
'
' TECHNICAL NOTES:
'   - Layer: UI / Ribbon Synchronization.
'   - Strategy: Ensures that toggling edge routing via one control
'     correctly deselects the others within the same logical group.
' ==========================================================================
Private Sub RefreshSplinesGroup()
    InvalidateRibbonControl RIBBON_CTL_SPLINES_COMPOUND
    InvalidateRibbonControl RIBBON_CTL_SPLINES_CURVED
    InvalidateRibbonControl RIBBON_CTL_SPLINES_LINE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_NONE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_ORTHO
    InvalidateRibbonControl RIBBON_CTL_SPLINES_POLYLINE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_SPLINE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_TRUE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_FALSE
End Sub

' ==========================================================================
' CALLBACK: splines_onAction
'
' PURPOSE:
'   Configures the edge-routing algorithm (splines) via the Ribbon and
'   synchronizes the UI and rendering state.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE:
'      - If pressed: Extracts the spline type from the control ID (e.g.,
'        "splineOrtho" becomes "ortho") and updates 'SETTINGS_SPLINES'.
'      - If unpressed: Clears the setting to use Graphviz defaults.
'   2. UI REFRESH: Invokes 'RefreshSplinesGroup' to update the visual
'      toggles across the Ribbon gallery.
'   3. REACTIVITY: Triggers 'AutoDraw' to immediately re-render the graph
'      with the new edge routing logic.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Splines Gallery/Buttons.
'   - Impact: Directly controls the 'splines=' attribute in the DOT source.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub splines_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_SPLINES).value = LCase$(Mid$(control.id, Len("spline") + 1))
    Else
        SettingsSheet.Range(SETTINGS_SPLINES).value = vbNullString
    End If
    RefreshSplinesGroup
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: splines_getVisible
'
' PURPOSE:
'   Dynamically manages the visibility of the Spline (Edge Routing) controls
'   on the Ribbon based on the compatibility of the active layout engine.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT EVALUATION: Polls 'SETTINGS_GRAPHVIZ_ENGINE' to identify the
'      active rendering algorithm.
'   2. COMPATIBILITY CHECK:
'      - Hides the controls (visible = False) for 'PATCHWORK' layouts, as
'        squarified treemaps do not utilize edge routing.
'      - Defaults to True for all other engines (dot, neato, etc.).
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Layout Engine change.
'   - UX Strategy: Reduces "UI Noise" by hiding specialized routing
'     parameters when they are functionally irrelevant.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub splines_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_PATCHWORK
            visible = False
        Case Else
            visible = True
    End Select
End Sub

' ==========================================================================
' CALLBACK: splines_getPressed
'
' PURPOSE:
'   Determines which edge-routing (spline) button should appear active on
'   the Ribbon, with a specific fallback for the "Off" state.
'
' TECHNICAL WORKFLOW:
'   1. DEFAULT HANDLING: If 'SETTINGS_SPLINES' is empty, it forces the
'      'splineFalse' button (Off/None) to appear pressed.
'   2. ID COMPARISON: For all other states, it compares the active setting
'      against the control's ID (stripping the "spline" prefix).
'   3. UI FEEDBACK: Returns TRUE if the setting matches the button context,
'      ensuring a single visual toggle is highlighted.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Spline selection.
'   - Strategy: Normalizes the relationship between a null setting and
'     the "False/Off" UI representation.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub splines_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If SettingsSheet.Range(SETTINGS_SPLINES).value = vbNullString And control.id = "splineFalse" Then
        pressed = True
    Else
        pressed = SettingsSheet.Range(SETTINGS_SPLINES).value = LCase$(Mid$(control.id, Len("spline") + 1))
    End If
End Sub

' ===========================================================================
' Callbacks for dirName

' ==========================================================================
' CALLBACK: getDir_getVisible
'
' PURPOSE:
'   Acts as a "Self-Healing" validator for the output directory whenever
'   the Ribbon is refreshed or the tab is activated.
'
' TECHNICAL WORKFLOW:
'   1. UI INITIALIZATION: Sets 'visible = True' to ensure the directory
'      selection control is always accessible.
'   2. PATH VALIDATION: Retrieves the current 'SETTINGS_OUTPUT_DIRECTORY'
'      and checks its physical existence using 'DirectoryExists'.
'   3. AUTOMATIC RESET: If a previously configured directory has been
'      deleted or moved, the function programmatically clears the
'      named range to prevent downstream file I/O errors.
'
' TECHNICAL NOTES:
'   - Strategy: Prevents "Broken Path" states by validating the file
'     system context before the user attempts a batch export.
'   - Trigger: Ribbon Invalidation or Tab Activation.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub getDir_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = True
    
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    If dirName = vbNullString Then Exit Sub

    ' Validate that the output directory exists
    If Not DirectoryExists(dirName) Then
        SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY).value = vbNullString
    End If

End Sub

' ==========================================================================
' CALLBACK: getDir_getLabel
'
' PURPOSE:
'   Dynamically updates the Ribbon label for the directory selection control
'   to provide visual feedback on the current configuration state.
'
' TECHNICAL WORKFLOW:
'   1. PATH INSPECTION: Retrieves and trims the value from the
'      'SETTINGS_OUTPUT_DIRECTORY' named range.
'   2. CONDITIONAL LABELING:
'      - If the path is empty: Returns a localized default label (e.g.,
'        "Select Directory...") via the 'GetLabel' helper.
'      - If a path exists: Returns a null string, typically allowing the
'        Ribbon to collapse the label or display the path in an adjacent edit box.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Project Load.
'   - UX Strategy: Prompts the user for action when the output destination
'     is not yet established.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub getDir_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    If dirName = vbNullString Then
        label = GetLabel("getDir")
    Else
        label = vbNullString
    End If
End Sub

' ===========================================================================
' Callbacks for getDirLabel

' ==========================================================================
' CALLBACK: getDirLabel_getLabel
'
' PURPOSE:
'   Dynamically updates a Ribbon label to display a truncated, human-readable
'   version of the current output directory.
'
' TECHNICAL WORKFLOW:
'   1. PATH RETRIEVAL: Accesses the 'SETTINGS_OUTPUT_DIRECTORY' named range
'      to fetch the full absolute file system path.
'   2. STRING OPTIMIZATION: Invokes 'ShortenToLastTwoFolders' to condense
'      long, complex paths into a "..\Folder\Subfolder" format.
'   3. UI UPDATE: Returns the shortened string to the Ribbon to provide
'      contextual awareness without occupying excessive horizontal space.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Directory selection.
'   - UX Strategy: Balances information density with UI aesthetics.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub getDirLabel_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim folder As String
    folder = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    label = ShortenToLastTwoFolders(folder)
End Sub

' ==========================================================================
' FUNCTION: ShortenToLastTwoFolders
'
' PURPOSE:
'   Condenses an absolute file path into a compact version suitable for
'   Ribbon display, focusing on the immediate parent and target directories.
'
' TECHNICAL WORKFLOW:
'   1. PRE-FLIGHT: Returns a null string if the input 'fullPath' is empty.
'   2. PARSING: Identifies the OS-specific 'pathSeparator' and splits the
'      string into an array of directory segments.
'   3. DEPTH-BASED FORMATTING:
'      - 1 Level (n=0): Returns the root/single directory.
'      - 2 Levels (n=1): Returns the full path (e.g., C:\Folder).
'      - 3+ Levels (n>1): Truncates the leading path with ellipses ("...")
'        and returns only the final two directories (e.g., ...\Parent\Leaf).
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (Uses 'Application.pathSeparator').
'   - Strategy: Optimizes limited Ribbon real estate while maintaining
'     navigational context for the user.
' ==========================================================================
Private Function ShortenToLastTwoFolders(ByVal fullPath As String) As String
    Dim parts() As String
    Dim n As Long
    Dim sep As String

    ' Ensure a path was passed in
    If Trim$(fullPath) = vbNullString Then
        ShortenToLastTwoFolders = vbNullString
        Exit Function
    End If
    
    ' We have a non-null string, parse it.
    sep = Application.pathSeparator
    parts = split(fullPath, sep)
    n = UBound(parts)

    Select Case n
        Case 0
            ShortenToLastTwoFolders = parts(0)

        Case 1
            ShortenToLastTwoFolders = parts(0) & sep & parts(1)

        Case Else
            ShortenToLastTwoFolders = "..." & sep & parts(n - 1) & sep & parts(n)
    End Select
End Function

' ===========================================================================
' Callbacks for fileFormat

' ==========================================================================
' CALLBACK: fileFormat_onAction
'
' PURPOSE:
'   Sets the target image file extension (e.g., png, svg, pdf) for batch
'   export operations via the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the format string from the 'controlId'
'      suffix (e.g., "ff_png" becomes "png") and updates 'SETTINGS_FILE_FORMAT'.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub fileFormat_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_FILE_FORMAT).value = Mid$(controlId, Len("ff_") + 1)
End Sub

' ==========================================================================
' CALLBACK: fileFormat_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon's file format gallery with the current
'   workbook configuration.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "ff_" with the value from 'SETTINGS_FILE_FORMAT'
'      to highlight the active selection in the gallery.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub fileFormat_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "ff_" & SettingsSheet.Range(SETTINGS_FILE_FORMAT).value
End Sub

' ===========================================================================
' Callbacks for filePrefix

' ==========================================================================
' CALLBACK: filePrefix_onChange
'
' PURPOSE:
'   Updates the global filename prefix used for exported graph files.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Writes the user-entered 'Text' directly to the
'      'SETTINGS_FILE_NAME' named range.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub filePrefix_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range(SETTINGS_FILE_NAME).value = Text
End Sub

' ==========================================================================
' CALLBACK: filePrefix_getText
'
' PURPOSE:
'   Populates the Ribbon's edit box with the current filename prefix setting.
'
' TECHNICAL WORKFLOW:
'   1. UI INITIALIZATION: Retrieves and trims the value from
'      'SETTINGS_FILE_NAME' to display in the UI control.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub filePrefix_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range(SETTINGS_FILE_NAME))
End Sub

' ===========================================================================
' Callbacks for getDir

' ==========================================================================
' CALLBACK: getDir_onAction
'
' PURPOSE:
'   Triggers an OS-native directory picker to set the target folder for
'   batch graph exports.
'
' TECHNICAL WORKFLOW:
'   1. DIRECTORY PICKER: Invokes 'SelectDirectoryToCell', which launches the
'      File Dialog (Folder Picker) and writes the resulting path to the
'      'SETTINGS_OUTPUT_DIRECTORY' named range.
'   2. UI SYNCHRONIZATION: Calls 'RefreshRibbon' to update the directory
'      labels and truncation logic across the UI.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Select Directory button.
'   - Strategy: Ensures valid absolute paths are captured via the standard
'     OS interface rather than manual text entry.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub getDir_onAction(ByVal control As IRibbonControl)
    SelectDirectoryToCell SettingsSheet.name, SETTINGS_OUTPUT_DIRECTORY
    RefreshRibbon
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
'@Ignore ParameterNotUsed
Public Sub graphToFile_onAction(ByVal control As IRibbonControl)
    Dim firstColumn As Long
    Dim lastColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    lastColumn = firstColumn
    
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents
    
    OptimizeCode_Begin
    CreateGraphFile firstColumn, lastColumn
    OptimizeCode_End
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
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
'@Ignore ParameterNotUsed
Public Sub graphToFile_getEnabled(ByVal control As IRibbonControl, ByRef pressed As Variant)
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
'@Ignore ParameterNotUsed
Public Sub graphAllViewsToFile_onAction(ByVal control As IRibbonControl)

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
    CreateGraphFile firstColumn, lastColumn
    OptimizeCode_End
    
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
'@Ignore ParameterNotUsed
Public Sub graphToWorksheet_onAction(ByVal control As IRibbonControl)
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
'@Ignore ParameterNotUsed
Public Sub graphToWorksheet_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
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
'@Ignore ParameterNotUsed
Public Sub graphAuto_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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
'@Ignore ParameterNotUsed
Public Sub graphAuto_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_AUTO
End Sub

' ===========================================================================
' Callbacks for graphWorksheet

' ==========================================================================
' CALLBACK: graphWorksheet_onAction
'
' PURPOSE:
'   Sets the target destination for rendered images (either the Data sheet
'   or the dedicated Graph sheet) via a Ribbon dropdown/gallery.
'
' TECHNICAL WORKFLOW:
'   1. SELECTION MAPPING: Maps index 0 to the "data" sheet and index 1
'      to the "graph" sheet.
'   2. STATE PERSISTENCE: Updates the 'SETTINGS_IMAGE_WORKSHEET' named range.
'   3. REACTIVITY: Invokes 'AutoDraw' to immediately move or re-render
'      the image to the new destination.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphWorksheet_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value = "data"
    Else
        SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value = "graph"
    End If
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: graphWorksheet_getItemLabel
'
' PURPOSE:
'   Provides localized labels for the "Target Worksheet" selection items.
'
' TECHNICAL WORKFLOW:
'   1. LOCALIZATION: Retrieves language-localized "Data" or "Graph" labels
'      via the 'GetLabel' helper based on the requested item index.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphWorksheet_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef itemLabel As Variant)
    If index = 0 Then
        itemLabel = GetLabel("worksheetDataName")
    Else
        itemLabel = GetLabel("worksheetGraphName")
    End If
End Sub

' ==========================================================================
' CALLBACK: graphWorksheet_getItemCount
'
' PURPOSE:
'   Defines the fixed number of options (2) available in the target
'   worksheet selection control. A dynamic dropdown is used as it allows
'   for language-localized values in the dropdown list as opposed to static
'   values if the dropdown were definex in CustomUI.xml.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphWorksheet_getItemCount(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = 2
End Sub

' ==========================================================================
' CALLBACK: graphWorksheet_getSelectedItemIndex
'
' PURPOSE:
'   Synchronizes the Ribbon dropdown selection with the current
'   'SETTINGS_IMAGE_WORKSHEET' value.
'
' TECHNICAL WORKFLOW:
'   1. STATE EVALUATION: Returns 0 if the setting is "data", otherwise 1.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphWorksheet_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef itemIndex As Variant)
    If SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value = "data" Then
        itemIndex = 0
    Else
        itemIndex = 1
    End If
End Sub

' ===========================================================================
' Callbacks for imageFormat

' ==========================================================================
' CALLBACK: imageFormat_onAction
'
' PURPOSE:
'   Sets the image format (SVG, PNG, JPG) for in-workbook rendering and
'   triggers a visual refresh.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the format string (e.g., "svg") from the
'      'controlId' suffix and updates the 'SETTINGS_IMAGE_TYPE' named range.
'   2. REACTIVITY: Invokes 'AutoDraw' to immediately re-render the graph
'      on the target worksheet using the new file format.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Image Format Gallery.
'   - Strategy: Allows users to switch between vector (SVG) and raster
'     (PNG/JPG) rendering modes instantly.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub imageFormat_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value = Mid$(controlId, Len("img_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: imageFormat_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon gallery's selection indicator with the
'   workbook's active image format setting.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "img_" with the value from 'SETTINGS_IMAGE_TYPE'
'      to resolve the matching control ID in the Ribbon XML.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub imageFormat_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "img_" & SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value
End Sub

' ===========================================================================
' Callbacks for includeOrphanEdges

' ==========================================================================
' CALLBACK: includeOrphanEdges_onAction
'
' PURPOSE:
'   Toggles the suppression of "Orphan Edges"—relationships that refer to
'   nodes not currently defined or styled in the active view.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates 'SETTINGS_RELATIONSHIPS_WITHOUT_NODES'
'      using the 'Toggle' helper to map the boolean state to project constants.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram, triggering the
'      connectivity audit in the parsing engine.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub includeOrphanEdges_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RELATIONSHIPS_WITHOUT_NODES).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: includeOrphanEdges_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Orphan Edges' toggle with the workbook
'   configuration.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns the Boolean value of the associated
'      named range to set the visual 'pressed' state.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub includeOrphanEdges_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_RELATIONSHIPS_WITHOUT_NODES)
End Sub

' ===========================================================================
' Callbacks for includeOrphanNodes

' ==========================================================================
' CALLBACK: includeOrphanNodes_onAction
'
' PURPOSE:
'   Toggles the suppression of "Orphan Nodes"—nodes that do not have
'   any active relationships (edges) defined within the current view.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates 'SETTINGS_NODES_WITHOUT_RELATIONSHIPS'
'      using the 'Toggle' helper to map the boolean state to project constants.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram, which triggers
'      the orphan detection logic in the core transformation engine.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub includeOrphanNodes_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_NODES_WITHOUT_RELATIONSHIPS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: includeOrphanNodes_getPressed
'
' PURPOSE:
'   Ensures the Ribbon's 'Orphan Nodes' toggle correctly reflects the
'   current setting in the workbook.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns the Boolean value of the associated
'      named range to set the visual 'pressed' state.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub includeOrphanNodes_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_NODES_WITHOUT_RELATIONSHIPS)
End Sub

' ===========================================================================
' Callbacks for keepGvFile

' ==========================================================================
' CALLBACK: keepGvFile_onAction
'
' PURPOSE:
'   Controls the persistence of the intermediate Graphviz source file (.gv)
'   after a rendering operation is completed.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_FILE_DISPOSITION' named range
'      using the 'Toggle' helper to map the Ribbon state to 'Keep' or 'Delete'.
'
' TECHNICAL NOTES:
'   - Usage: Enabling this is essential for power users who wish to manually
'     inspect or modify the generated DOT source in an external editor.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub keepGvFile_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_FILE_DISPOSITION).value = Toggle(pressed, TOGGLE_KEEP, TOGGLE_DELETE)
End Sub

' ==========================================================================
' CALLBACK: keepGvFile_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Keep .gv File' toggle with the workbook
'   configuration.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns TRUE if 'SETTINGS_FILE_DISPOSITION' is set
'      to the 'TOGGLE_KEEP' constant.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub keepGvFile_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_FILE_DISPOSITION).value = TOGGLE_KEEP
End Sub

' ===========================================================================
' Callbacks for rankdir

' ==========================================================================
' CALLBACK: rankdir_getVisible
'
' PURPOSE:
'   Controls the visibility of the Rank Direction (layout flow) controls.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Queries 'SETTINGS_GRAPHVIZ_ENGINE'.
'   2. ENGINE GATE: Returns TRUE only for the 'DOT' engine, as hierarchical
'      flow (TB, LR, etc.) is specific to the Sugiyama-style layout.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub rankdir_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LAYOUT_DOT
End Sub

' ==========================================================================
' CALLBACK: rankdir_onAction
'
' PURPOSE:
'   Sets the graph layout direction (Top-to-Bottom, Left-to-Right, etc.).
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the direction code (TB, BT, LR, RL)
'      from the control ID and updates 'SETTINGS_RANKDIR'.
'   2. UI SYNC: Invokes 'RefreshRankdirGroup' to update toggle states.
'   3. REACTIVITY: Triggers 'AutoDraw' for immediate visual feedback.
' ==========================================================================
Public Sub rankdir_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_RANKDIR).value = Mid$(control.id, Len("rankdir") + 1)
    Else
        SettingsSheet.Range(SETTINGS_RANKDIR).value = vbNullString
    End If
    RefreshRankdirGroup
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: rankdir_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's direction toggles with the active setting.
'
' TECHNICAL WORKFLOW:
'   1. DEFAULT FALLBACK: If the setting is empty, 'rankdirTB' (Top-Bottom)
'      is forced to the pressed state to reflect the Graphviz default.
'   2. ID MATCH: Otherwise, compares the control ID suffix to the setting.
' ==========================================================================
Public Sub rankdir_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If SettingsSheet.Range(SETTINGS_RANKDIR).value = vbNullString And control.id = "rankdirTB" Then
        pressed = True
    Else
        pressed = SettingsSheet.Range(SETTINGS_RANKDIR).value = Mid$(control.id, Len("rankdir") + 1)
    End If
End Sub

' ==========================================================================
' PROCEDURE: RefreshRankdirGroup
'
' PURPOSE:
'   Forces a visual refresh of the entire Rank Direction Ribbon group.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Triggers 'InvalidateRibbonControl' for the group
'      header and all direction/spacer buttons to ensure toggle sync.
' ==========================================================================
Private Sub RefreshRankdirGroup()
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_GROUP
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_TB
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_BT
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_LR
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_RL
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_DUMMY1
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_DUMMY2
End Sub

' ===========================================================================
' Callbacks for showNodeLabels

' ==========================================================================
' CALLBACK: showNodeLabels_onAction
'
' PURPOSE:
'   Toggles the global visibility of labels on all graph nodes via the
'   Ribbon interface.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_NODE_LABELS' named range
'      using the 'Toggle' helper to persist the 'Include/Exclude' preference.
'   2. REACTIVITY: Invokes 'AutoDraw' to immediately re-generate the DOT
'      source, determining whether 'label' attributes are emitted for nodes.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_NODE_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: showNodeLabels_getPressed
'
' PURPOSE:
'   Ensures the Ribbon's 'Node Labels' toggle accurately reflects the
'   current project setting.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns the Boolean result of 'GetSettingBoolean' for
'      the node label visibility range to set the visual 'pressed' state.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_NODE_LABELS)
End Sub

' ===========================================================================
' Callbacks for showNodeXLabels

' ==========================================================================
' CALLBACK: showNodeXLabels_onAction
'
' PURPOSE:
'   Toggles the global visibility of external node labels (xLabels)
'   via the Ribbon interface.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates 'SETTINGS_NODE_XLABELS' using the 'Toggle'
'      helper to map the Ribbon state to project constants.
'   2. REACTIVITY: Invokes 'AutoDraw' to re-render the graph, determining
'      if 'xlabel' attributes are included in the DOT output.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showNodeXLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_NODE_XLABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: showNodeXLabels_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Node xLabels' toggle with the workbook
'   configuration.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns the Boolean value of the associated
'      named range to set the visual 'pressed' state.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showNodeXLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_NODE_XLABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeLabels

' ==========================================================================
' CALLBACK: showEdgeLabels_onAction
'
' PURPOSE:
'   Toggles the global visibility of primary labels on graph edges via
'   the Ribbon interface.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_EDGE_LABELS' named range
'      using the 'Toggle' helper to persist the 'Include/Exclude' state.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram, determining
'      if 'label' attributes are generated for edge relationships.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showEdgeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: showEdgeLabels_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Edge Labels' toggle with the workbook settings.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns the Boolean result of 'GetSettingBoolean' for
'      the edge label visibility range to set the visual 'pressed' state.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub showEdgeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_LABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeXLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeXLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_XLABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeXLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_XLABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeHeadLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeHeadLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_HEAD_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeHeadLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_HEAD_LABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeTailLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeTailLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_TAIL_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeTailLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_TAIL_LABELS)
End Sub

' ===========================================================================
' Callbacks for showPorts

'@Ignore ParameterNotUsed
Public Sub showPorts_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_PORTS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showPorts_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_PORTS)
End Sub

' ===========================================================================
' Callbacks for strict

'@Ignore ParameterNotUsed
Public Sub strict_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_STRICT).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub strict_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_STRICT)
End Sub

' ===========================================================================
' Callbacks for transparent

'@Ignore ParameterNotUsed
Public Sub transparent_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_TRANSPARENT).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub transparent_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_TRANSPARENT)
End Sub

' ===========================================================================
' Callbacks for center

'@Ignore ParameterNotUsed
Public Sub center_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_CENTER).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub center_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_CENTER)
End Sub

' ===========================================================================
' Callbacks for compound

' ==========================================================================
' CALLBACK: compound_onAction
'
' PURPOSE:
'   Toggles the Graphviz 'compound' attribute, allowing edges to clip to
'   subgraph/cluster boundaries rather than just internal nodes.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_GRAPH_COMPOUND' named range
'      using the 'Toggle' helper to map boolean state to 'Yes/No' constants.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram layout.
' =========================================================================='@Ignore ParameterNotUsed
Public Sub compound_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_COMPOUND).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: compound_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Compound' toggle with the workbook settings.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Retrieves the current value via 'GetSettingBoolean'
'      to set the visual 'pressed' state.
' =========================================================================='@Ignore ParameterNotUsed
Public Sub compound_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_COMPOUND)
End Sub

' ==========================================================================
' CALLBACK: compound_getVisible
'
' PURPOSE:
'   Manages the visibility of the Compound toggle based on engine compatibility.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Returns TRUE only for the 'DOT' engine, as compound
'      routing is a specific feature of hierarchical layouts.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub compound_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for concentrate

'@Ignore ParameterNotUsed
Public Sub concentrate_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_CONCENTRATE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub concentrate_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_CONCENTRATE)
End Sub

' ===========================================================================
' Callbacks for forceLabels

'@Ignore ParameterNotUsed
Public Sub forceLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_FORCE_LABELS).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub forceLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_FORCE_LABELS)
End Sub

' ===========================================================================
' Callbacks for newrank

' ==========================================================================
' CALLBACK: newrank_onAction
'
' PURPOSE:
'   Toggles the 'newrank' attribute for the DOT engine, enabling a more
'   flexible ranking algorithm that handles nested clusters more efficiently.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Maps the Ribbon's 'pressed' state to the
'      'SETTINGS_GRAPH_NEWRANK' named range using the 'Toggle' helper.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram layout.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub newrank_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_NEWRANK).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: newrank_getPressed
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Newrank' toggle with the workbook settings.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOOKUP: Returns the Boolean value of 'SETTINGS_GRAPH_NEWRANK'
'      to set the visual 'pressed' state.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub newrank_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_NEWRANK)
End Sub

' ==========================================================================
' CALLBACK: newrank_getVisible
'
' PURPOSE:
'   Manages the visibility of the Newrank toggle based on engine context.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Displays the control (visible = True) only for the
'      'DOT' layout engine, as this attribute is specific to hierarchical
'      Sugiyama-style rendering.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub newrank_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for rotate

'@Ignore ParameterNotUsed
Public Sub rotate_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_ORIENTATION).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub rotate_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_ORIENTATION)
End Sub

Public Function getPressed(ByVal worksheetName As String, ByVal keyword As String, ByVal matchValue As String) As Boolean
    getPressed = UCase$(GetCellString(worksheetName, keyword)) = UCase$(matchValue)
End Function

' ===========================================================================
' Callbacks for overlap

' ==========================================================================
' CALLBACK: overlap_getVisible
'
' PURPOSE:
'   Dynamically manages the visibility of the 'Overlap' parameter based on
'   layout engine compatibility.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT EVALUATION: Polls the 'SETTINGS_GRAPHVIZ_ENGINE' named range.
'   2. ALGORITHM BRANCHING:
'      - Forces 'visible = True' for force-directed engines (FDP, NEATO, SFDP)
'        where node overlap management is a key layout constraint.
'      - Forces 'visible = False' for hierarchical (DOT), radial (TWOPI),
'        or deterministic (PATCHWORK) engines where 'overlap' is not an
'        exposed or applicable attribute.
'
' TECHNICAL NOTES:
'   - UX Strategy: Reduces "Attribute Bloat" by presenting only parameters
'     relevant to the active Graphviz algorithm.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub overlap_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = False
        Case LAYOUT_DOT
            visible = False
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_OSAGE
            visible = False
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = True
        Case LAYOUT_TWOPI
            visible = False
        Case Else
            visible = False
    End Select
End Sub

'@Ignore ParameterNotUsed
Public Sub overlap_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value = Mid$(controlId, Len("overlap_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub overlap_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value = control.id
End Sub


'@Ignore ParameterNotUsed
Public Sub overlap_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "overlap_" & SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value
End Sub

' ===========================================================================
' Callbacks for toggleDebugLabels

'@Ignore ParameterNotUsed
Public Sub toggleDebugLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_DEBUG).value = Toggle(pressed, TOGGLE_ON, TOGGLE_OFF)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleDebugLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_DEBUG)
End Sub

' ===========================================================================
' Callbacks for toggleLogToConsole

'@Ignore ParameterNotUsed
Public Sub toggleDebugLogToConsole_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_LOG_TO_CONSOLE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleDebugLogToConsole_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub toggleDebugLogToConsole_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = enableConsole()
End Sub

' ===========================================================================
' Callbacks for toggleGraphvizVerbose

'@Ignore ParameterNotUsed
Public Sub toggleGraphvizVerbose_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPHVIZ_VERBOSE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleGraphvizVerbose_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPHVIZ_VERBOSE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub toggleGraphvizVerbose_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = enableConsole()
End Sub

' ===========================================================================
' Callbacks for useDefinedStyles

'@Ignore ParameterNotUsed
Public Sub useDefinedStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_INCLUDE_STYLE_FORMAT).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub useDefinedStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_INCLUDE_STYLE_FORMAT)
End Sub

' ===========================================================================
' Callbacks for useExtraStyles

'@Ignore ParameterNotUsed
Public Sub useExtraStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_INCLUDE_EXTRA_ATTRIBUTES).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub useExtraStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_INCLUDE_EXTRA_ATTRIBUTES)
End Sub

' ===========================================================================
' Callbacks for yesView

' ==========================================================================
' CALLBACK: yesNoView_onAction
'
' PURPOSE:
'   Sets the active View (Style Gallery column) based on the user's
'   selection from a Ribbon gallery or dropdown.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN CALCULATION: Determines the target column index by adding
'      the selected 'index' to the project's 'SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW'
'      base offset.
'   2. ADDRESS CONVERSION: Converts the numeric index into standard Excel
'      letters (e.g., "D") via 'ConvertColumnNumberToLetters'.
'   3. STATE PERSISTENCE: Updates the 'SETTINGS_YES_NO_SWITCH_COLUMN'
'      named range to inform the parsing engine which data to render.
'   4. REACTIVITY: Invokes 'AutoDraw' to immediately refresh the graph
'      using the newly selected View.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Styles/Graphviz Tab -> View Selection.
'   - Layer: UI / Settings Management.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub yesNoView_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    Dim columnName As String
    columnName = ConvertColumnNumberToLetters(index + GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW))
    SettingsSheet.Range(SETTINGS_YES_NO_SWITCH_COLUMN).value = columnName
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: yesNoView_getItemCount
'
' PURPOSE:
'   Dynamically calculates the number of available "Views" (columns) in the
'   Style Gallery to populate Ribbon dropdowns and galleries.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Iterates through the 'Styles' sheet header row,
'      starting from the first View column, counting non-empty headers.
'   2. SELF-HEALING LOGIC: Detects if the currently selected View column
'      was deleted from the worksheet.
'   3. AUTOMATIC RECOVERY: If the active selection index is now out of
'      bounds, it programmatically shifts the selection to the new
'      'lastCol' and triggers a full 'RefreshRibbon'.
'   4. UI FEEDBACK: Returns the final 'itemCount' to the Ribbon engine.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon Invalidation or Tab Activation.
'   - Strategy: Ensures the UI remains synchronized even if the user
'     manually modifies the 'Styles' worksheet structure.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub yesNoView_getItemCount(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim itemCount As Long
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim columnName As String
    
    row = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    itemCount = 0
    
    ' Count the non-empty cells beginning at the first view column
    For col = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW) To GetLastColumn(StylesSheet.name, row)
        columnName = StylesSheet.Cells.item(row, col)
        If columnName <> vbNullString Then
            itemCount = itemCount + 1
        End If
    Next col

    ' If the last view column is the currently selected column, and the user deletes the column then it
    ' is necessary to change the selection to the last column which will be present after the deletion occurs.
    lastCol = itemCount + GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW) - 1
    
    If lastCol < GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE) Then
        SettingsSheet.Range(SETTINGS_YES_NO_SWITCH_COLUMN).value = ConvertColumnNumberToLetters(lastCol)
        RefreshRibbon
    End If
    
    returnedVal = itemCount
End Sub

' ==========================================================================
' CALLBACK: yesNoView_getItemLabel
'
' PURPOSE:
'   Dynamically retrieves the display name for a specific "View" item within
'   a Ribbon gallery or dropdown.
'
' TECHNICAL WORKFLOW:
'   1. ADDRESS RESOLUTION: Calculates the target column by adding the
'      provided 'index' to the starting View column index.
'   2. LABEL RETRIEVAL: Pulls the header text directly from the 'StylesSheet'
'      using the row index defined in 'SETTINGS_STYLES_ROW_HEADING'.
'   3. UI FEEDBACK: Returns the header string (e.g., "Logical", "Security")
'      to the Ribbon's item label property.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub yesNoView_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef itemLabel As Variant)
    itemLabel = StylesSheet.Cells.item(CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING)), _
                            index + GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW))
End Sub

' ==========================================================================
' CALLBACK: yesNoView_getSelectedItemIndex
'
' PURPOSE:
'   Synchronizes the Ribbon's selection indicator with the currently
'   active View configuration.
'
' TECHNICAL WORKFLOW:
'   1. OFFSET CALCULATION: Subtracts the starting View column index from
'      the currently active view index ('SETTINGS_STYLES_COL_SHOW_STYLE').
'   2. UI SYNC: Returns the zero-based 'itemIndex', ensuring the correct
'      View is highlighted in the dropdown.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub yesNoView_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef itemIndex As Variant)
    Dim indx As Long
    indx = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE) - GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    itemIndex = indx
End Sub

' Utility routines

' ==========================================================================
' FUNCTION: IsAViewSpecified
'
' PURPOSE:
'   Validates if the user has selected a valid "View" column from the
'   Style Gallery to act as the rendering filter.
'
' TECHNICAL WORKFLOW:
'   1. STATE EVALUATION: Queries the 'SETTINGS_VIEW_NAME' named range.
'   2. LOGICAL CHECK: Returns TRUE if the value is anything other than "0"
'      (the system's null/default state).
'
' TECHNICAL NOTES:
'   - Layer: Logic / UI State.
'   - Usage: Acts as a critical "Safety Gate" for the 'getEnabled'
'     callbacks of the Render and Publish buttons.
' ==========================================================================
Public Function IsAViewSpecified() As Boolean
    IsAViewSpecified = Not (SettingsSheet.Range(SETTINGS_VIEW_NAME).value = "0")
End Function

' ==========================================================================
' CALLBACK: sql_getVisible
'
' PURPOSE:
'   Enforces the "Windows-Only" restriction for SQL-related Ribbon controls.
'
' TECHNICAL WORKFLOW:
'   1. PLATFORM BRANCHING: Uses the '#If Mac' compiler directive.
'   2. UI SUPPRESSION:
'      - Returns FALSE on macOS to hide controls dependent on ADO.
'      - Returns TRUE on Windows to enable the full SQL subsystem.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Prevents user interaction with the ADO-based SQL
'     engine on non-supported platforms.
' ==========================================================================
'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sql_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub

' ==========================================================================
' CALLBACK: mac_getVisible
'
' PURPOSE:
'   Controls the visibility of macOS-specific Ribbon elements or alerts.
'
' TECHNICAL WORKFLOW:
'   1. PLATFORM BRANCHING: Uses the '#If Mac' compiler directive.
'   2. UI INVERSION:
'      - Returns TRUE on macOS to show platform-specific guidance.
'      - Returns FALSE on Windows to keep the UI clean.
' ==========================================================================
'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub mac_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = True
#Else
    returnedVal = False
#End If
End Sub

' ===========================================================================
' Callbacks for graphZoomLevel

' ==========================================================================
' CALLBACK: graphZoomLevel_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon's zoom gallery selection with the current
'   workbook scale settings.
'
' TECHNICAL WORKFLOW:
'   1. STATE RETRIEVAL: Invokes 'GetCurrentZoom' to fetch the numeric scale.
'   2. ID COMPOSITION: Concatenates the control ID with the zoom value
'      (e.g., "zoom100") to highlight the correct item in the gallery.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphZoomLevel_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = control.id & GetCurrentZoom()
End Sub

' ==========================================================================
' CALLBACK: graphZoomLevel_onAction
'
' PURPOSE:
'   Updates the visual scaling of the rendered graph via a Ribbon selection.
'
' TECHNICAL WORKFLOW:
'   1. VALUE EXTRACTION: Parses the numeric zoom percentage from the
'      'controlId' string.
'   2. STATE UPDATE: Invokes 'UpdateZoom' to persist the new scale and
'      immediately trigger a visual resize of the graph.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphZoomLevel_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    Dim zoomLevel As String
    zoomLevel = Mid$(controlId, Len(control.id) + 1)
    UpdateZoom CLng(zoomLevel)
End Sub

' ==========================================================================
' CALLBACK: graphZoomLevel_getLabel
'
' PURPOSE:
'   Provides live feedback of the current zoom percentage on the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. UI UPDATE: Returns the value from 'GetCurrentZoom' formatted with
'      a percentage symbol (e.g., "150%") to the control's label property.
' ==========================================================================
Public Sub graphZoomLevel_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    label = GetCurrentZoom() & "%"
End Sub

' ===========================================================================
' Callbacks for graphZoomOut

'@Ignore ParameterNotUsed
Public Sub graphZoomOut_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = GetCurrentZoom() > MIN_ZOOM
End Sub

'@Ignore ParameterNotUsed
Public Sub GraphZoomOut_OnAction(ByVal control As IRibbonControl)
    Dim zoom As Long
    zoom = SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value - ZOOM_STEP
    
    UpdateZoom zoom
End Sub

' ===========================================================================
' Callbacks for graphZoomIn

' ==========================================================================
' CALLBACK: graphZoomOut_getEnabled
'
' PURPOSE:
'   Enforces the lower boundary for graph scaling by enabling or disabling
'   the "Zoom Out" control.
'
' TECHNICAL WORKFLOW:
'   1. THRESHOLD CHECK: Compares the current zoom level against the
'      'MIN_ZOOM' constant.
'   2. UI FEEDBACK: Disables the button if the minimum scale is reached
'      to prevent invalid scaling values.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphZoomIn_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = GetCurrentZoom() < MAX_ZOOM
End Sub

' ==========================================================================
' CALLBACK: GraphZoomOut_OnAction
'
' PURPOSE:
'   Decrements the graph's visual scale by a predefined step value.
'
' TECHNICAL WORKFLOW:
'   1. CALCULATION: Subtracts the 'ZOOM_STEP' from the current
'      'SETTINGS_SCALE_IMAGE' setting.
'   2. STATE UPDATE: Invokes 'UpdateZoom' to apply the new scale and
'      trigger a refresh of the Ribbon and graph image.
' ==========================================================================
Public Sub GraphZoomIn_OnAction(ByVal control As IRibbonControl)
    Dim zoom As Long
    zoom = SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value + ZOOM_STEP
    
    UpdateZoom zoom
End Sub

' ==========================================================================
' PROCEDURE: UpdateZoom
'
' PURPOSE:
'   The central scaling orchestrator. Manages the visual magnification of
'   the graph while maintaining boundary integrity and UI synchronization.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY CLAMPING: Enforces project-specific 'MIN_ZOOM' and
'      'MAX_ZOOM' constraints to prevent rendering errors.
'   2. STATE PERSISTENCE: Writes the sanitized value to the
'      'SETTINGS_SCALE_IMAGE' named range.
'   3. UI SYNCHRONIZATION: Explicitly invalidates the specific Ribbon
'      controls (In, Out, and Gallery) to refresh their labels and 'Enabled'
'      states.
'   4. REFRESH: Invokes 'CreateGraphWorksheetQuickly' to re-render the
'      image at the new scale.
' ==========================================================================
Private Sub UpdateZoom(zoom As Long)
    ' Clamp zoom value within bounds
    If zoom < MIN_ZOOM Then zoom = MIN_ZOOM
    If zoom > MAX_ZOOM Then zoom = MAX_ZOOM
    
    ' Update zoom value in settings
    SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value = zoom
    
    ' Invalidate ribbon controls
    InvalidateRibbonControl "graphZoomIn"
    InvalidateRibbonControl "graphZoomOut"
    InvalidateRibbonControl "graphZoomLevel"
    
    ' Refresh graph
    CreateGraphWorksheetQuickly
End Sub

' ==========================================================================
' FUNCTION: GetCurrentZoom
'
' PURPOSE:
'   Standardized getter for retrieving the current image scale from the
'   workbook's global settings.
'
' TECHNICAL WORKFLOW:
'   1. SETTINGS LOOKUP: Returns the Long value from the
'      'SETTINGS_SCALE_IMAGE' named range on the SettingsSheet.
' ==========================================================================
Private Function GetCurrentZoom() As Long
    GetCurrentZoom = SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value
End Function

' ===========================================================================
' Callbacks for dim

' ==========================================================================
' CALLBACK: dim_onAction
'
' PURPOSE:
'   Sets the 'dim' attribute (dimensionality for internal layout calculation)
'   via the Ribbon dropdown.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the dimension value (e.g., 2, 3) from the
'      'controlId' suffix and updates 'SETTINGS_GRAPH_DIM'.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the rendering engine with
'      the new dimensionality constraint.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub dim_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_DIM).value = Mid$(controlId, Len("dim_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: dim_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon's dimensionality selection with the underlying
'   workbook configuration.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates the "dim_" prefix with the value from
'      'SETTINGS_GRAPH_DIM' to highlight the active item in the gallery.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub dim_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "dim_" & SettingsSheet.Range(SETTINGS_GRAPH_DIM).value
End Sub

' ==========================================================================
' CALLBACK: dim_getVisible
'
' PURPOSE:
'   Manages the visibility of the 'dim' parameter based on layout engine context.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Displays the control (visible = True) only for
'      force-directed layouts (FDP, NEATO, SFDP) where dimensionality is a
'      valid layout parameter.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub dim_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for dimen

' ==========================================================================
' CALLBACK: dimen_onAction
'
' PURPOSE:
'   Sets the 'dimen' attribute (dimensionality for external rendering)
'   via the Ribbon dropdown.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the numeric value from the 'controlId'
'      suffix (e.g., "dimen_2") and updates 'SETTINGS_GRAPH_DIMEN'.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the graph with the
'      updated rendering dimensions.
' =========================================================================='@Ignore ParameterNotUsed
Public Sub dimen_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_DIMEN).value = Mid$(controlId, Len("dimen_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: dimen_GetSelectedItemID
'
' PURPOSE:
'   Ensures the Ribbon gallery correctly highlights the currently
'   configured 'dimen' value.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "dimen_" with the value from
'      'SETTINGS_GRAPH_DIMEN' to resolve the matching control ID.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub dimen_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "dimen_" & SettingsSheet.Range(SETTINGS_GRAPH_DIMEN).value
End Sub

' ==========================================================================
' CALLBACK: dimen_getVisible
'
' PURPOSE:
'   Restricts the visibility of the 'dimen' parameter to compatible engines.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Returns TRUE only for force-directed engines
'      (FDP, NEATO, SFDP), hiding it for all other layout algorithms.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub dimen_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for mode

' ==========================================================================
' CALLBACK: mode_onAction
'
' PURPOSE:
'   Sets the 'mode' attribute (heuristic used for layout optimization) via
'   the Ribbon dropdown.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the mode value from the 'controlId'
'      suffix and updates the 'SETTINGS_GRAPH_MODE' named range.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the graph with the new
'      optimization algorithm.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub mode_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_MODE).value = Mid$(controlId, Len("mode_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: mode_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Mode' gallery selection with the current
'   workbook configuration.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "mode_" with the value from 'SETTINGS_GRAPH_MODE'
'      to highlight the active item in the Ribbon gallery.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub mode_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "mode_" & SettingsSheet.Range(SETTINGS_GRAPH_MODE).value
End Sub

' ==========================================================================
' CALLBACK: mode_getVisible
'
' PURPOSE:
'   Controls the visibility of the 'mode' parameter based on layout engine.
'
' TECHNICAL WORKFLOW:
'   1. ENGINE GATE: Returns TRUE only for 'NEATO' and 'SFDP' engines, as the
'      mode attribute (e.g., major, KK) is specific to these force-directed
'      and spring-model algorithms.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub mode_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for model

' ==========================================================================
' CALLBACK: model_onAction
'
' PURPOSE:
'   Sets the 'model' attribute (distance metric used for layout) via
'   the Ribbon dropdown.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the model type (e.g., shortpath, circuit)
'      from the 'controlId' suffix and updates 'SETTINGS_GRAPH_MODEL'.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the graph using the new
'      mathematical model for node positioning.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub model_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_MODEL).value = Mid$(controlId, Len("model_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: model_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Model' selection with the current
'   workbook configuration.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "model_" with the value from 'SETTINGS_GRAPH_MODEL'
'      to highlight the active item in the gallery.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub model_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "model_" & SettingsSheet.Range(SETTINGS_GRAPH_MODEL).value
End Sub

' ==========================================================================
' CALLBACK: model_getVisible
'
' PURPOSE:
'   Manages the visibility of the 'model' parameter based on engine context.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Displays the control (visible = True) only for the
'      'NEATO' engine, as this attribute specifically dictates how neato
'      interprets edge lengths.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub model_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_NEATO
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for smoothing

' ==========================================================================
' CALLBACK: smoothing_onAction
'
' PURPOSE:
'   Sets the 'smoothing' attribute (velocity damping/layout refinement)
'   via the Ribbon dropdown.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the smoothing algorithm (e.g., avg_dist,
'      rng, spring) from the 'controlId' and updates 'SETTINGS_GRAPH_SMOOTHING'.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the graph with the
'      specified layout refinement logic.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub smoothing_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_SMOOTHING).value = Mid$(controlId, Len("smoothing_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: smoothing_GetSelectedItemID
'
' PURPOSE:
'   Synchronizes the Ribbon's 'Smoothing' selection with the current
'   workbook configuration.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "smoothing_" with the value from
'      'SETTINGS_GRAPH_SMOOTHING' to highlight the active gallery item.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub smoothing_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "smoothing_" & SettingsSheet.Range(SETTINGS_GRAPH_SMOOTHING).value
End Sub

' ==========================================================================
' CALLBACK: smoothing_getVisible
'
' PURPOSE:
'   Manages the visibility of the 'smoothing' parameter based on engine context.
'
' TECHNICAL WORKFLOW:
'   1. ENGINE GATE: Returns TRUE only for the 'SFDP' engine, as smoothing
'      algorithms are specialized for large-scale force-directed layouts.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub smoothing_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for clusterrank

' ==========================================================================
' CALLBACK: clusterrank_onAction
'
' PURPOSE:
'   Sets the 'clusterrank' attribute via the Ribbon, determining how
'   subgraphs/clusters are treated during the ranking phase.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the rank mode (e.g., local, global, none)
'      from the 'controlId' and updates 'SETTINGS_GRAPH_CLUSTER_RANK'.
'   2. REACTIVITY: Invokes 'AutoDraw' to refresh the diagram with the
'      new cluster ranking logic.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub clusterrank_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_CLUSTER_RANK).value = Mid$(controlId, Len("clusterrank_") + 1)
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: clusterrank_GetSelectedItemID
'
' PURPOSE:
'   Ensures the Ribbon gallery correctly highlights the currently
'   active cluster ranking mode.
'
' TECHNICAL WORKFLOW:
'   1. UI SYNC: Concatenates "clusterrank_" with the value from
'      'SETTINGS_GRAPH_CLUSTER_RANK' to resolve the matching control ID.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub clusterrank_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "clusterrank_" & SettingsSheet.Range(SETTINGS_GRAPH_CLUSTER_RANK).value
End Sub

' ==========================================================================
' CALLBACK: clusterrank_getVisible
'
' PURPOSE:
'   Restricts the visibility of the 'clusterrank' parameter to compatible engines.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Displays the control (visible = True) only for the
'      'DOT' layout engine, as this attribute dictates hierarchical
'      nesting logic specific to Sugiyama-style rendering.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub clusterrank_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for ordering

' ==========================================================================
' PROCEDURE: RefreshOrderingGroup
'
' PURPOSE:
'   Synchronizes the "Ordering" Ribbon group to ensure the visual state
'   reflects the current edge-sorting constraints (in/out).
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Triggers 'InvalidateRibbonControl' for the parent
'      group and the specific ordering toggles (In, Out, and Spacers).
' ==========================================================================
Private Sub RefreshOrderingGroup()
    InvalidateRibbonControl RIBBON_CTL_ORDERING_GROUP
    InvalidateRibbonControl RIBBON_CTL_ORDERING_IN
    InvalidateRibbonControl RIBBON_CTL_ORDERING_OUT
    InvalidateRibbonControl RIBBON_CTL_ORDERING_DUMMY1
End Sub

' ==========================================================================
' CALLBACK: ordering_getVisible
'
' PURPOSE:
'   Restricts the visibility of edge-ordering controls to compatible engines.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Displays the group (visible = True) only for the
'      'DOT' layout engine, as edge port ordering is specific to
'      hierarchical ranking.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub ordering_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ==========================================================================
' CALLBACK: ordering_onAction
'
' PURPOSE:
'   Configures the 'ordering' attribute (sorts edges by in-degree or
'   out-degree) via the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Extracts the sorting type from the control ID
'      suffix and updates 'SETTINGS_GRAPH_ORDERING'.
'   2. UI SYNC: Invokes 'RefreshOrderingGroup' to update toggle indicators.
'   3. REACTIVITY: Triggers 'AutoDraw' to refresh the diagram.
' ==========================================================================
Public Sub ordering_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value = LCase$(Mid$(control.id, Len("ordering") + 1))
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value = vbNullString
    End If
    RefreshOrderingGroup
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: ordering_getPressed
'
' PURPOSE:
'   Determines which edge-ordering button appears active on the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE COMPARISON: Compares the normalized value of
'      'SETTINGS_GRAPH_ORDERING' against the control's ID suffix.
' ==========================================================================
Public Sub ordering_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If LCase$(SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value) = LCase$(Mid$(control.id, Len("ordering") + 1)) Then
        pressed = True
    Else
        pressed = False
    End If
End Sub

' ===========================================================================
' Callbacks for outputorder

' ==========================================================================
' PROCEDURE: RefreshOutputorderGroup
'
' PURPOSE:
'   Synchronizes the "Output Order" Ribbon group to ensure only the active
'   rendering sequence is visually selected.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Triggers 'InvalidateRibbonControl' for the parent
'      group and all sequence toggles (NodesFirst, EdgesFirst, BreadthFirst).
' ==========================================================================
Private Sub RefreshOutputorderGroup()
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_GROUP
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_NODES_FIRST
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_EDGES_FIRST
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_BREADTH_FIRST
End Sub

' ==========================================================================
' CALLBACK: outputorderBreadthFirst_getPressed
'
' PURPOSE:
'   Determines which Output Order button appears active on the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. DEFAULT STATE: If 'SETTINGS_GRAPH_OUTPUT_ORDER' is empty,
'      'BreadthFirst' is forced to the pressed state as the Graphviz default.
'   2. STATE MATCH: Returns TRUE if the setting matches the button's
'      associated value (breadthfirst).
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorderBreadthFirst_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString Then
        pressed = True
    Else
        pressed = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "breadthfirst"
    End If
End Sub

' ==========================================================================
' CALLBACK: outputorderBreadthFirst_onAction
'
' PURPOSE:
'   Updates the 'outputorder' Graphviz attribute via the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_GRAPH_OUTPUT_ORDER' named
'      range based on the selected sequence.
'   2. UI REFRESH: Invokes 'RefreshOutputorderGroup' to update toggle states.
'   3. REACTIVITY: Triggers 'AutoDraw' for immediate visual feedback.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorderBreadthFirst_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "breadthfirst"
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString
    End If
    RefreshOutputorderGroup
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: outputorderEdgesFirst_getPressed
'
' PURPOSE:
'   Determines which Output Order button appears active on the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. DEFAULT STATE: If 'SETTINGS_GRAPH_OUTPUT_ORDER' is empty,
'      'BreadthFirst' is forced to the pressed state as the Graphviz default.
'   2. STATE MATCH: Returns TRUE if the setting matches the button's
'      associated value (edgesfirst).
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorderEdgesFirst_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "edgesfirst"
End Sub

' ==========================================================================
' CALLBACK: outputorderEdgesFirst_onAction
'
' PURPOSE:
'   Updates the 'outputorder' Graphviz attribute via the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_GRAPH_OUTPUT_ORDER' named
'      range based on the selected sequence.
'   2. UI REFRESH: Invokes 'RefreshOutputorderGroup' to update toggle states.
'   3. REACTIVITY: Triggers 'AutoDraw' for immediate visual feedback.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorderEdgesFirst_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "edgesfirst"
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString
    End If
    RefreshOutputorderGroup
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: outputorderNodesFirst_getPressed
'
' PURPOSE:
'   Determines which Output Order button appears active on the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. DEFAULT STATE: If 'SETTINGS_GRAPH_OUTPUT_ORDER' is empty,
'      'BreadthFirst' is forced to the pressed state as the Graphviz default.
'   2. STATE MATCH: Returns TRUE if the setting matches the button's
'      associated value (nodesfirst).
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorderNodesFirst_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "nodesfirst"
End Sub

' ==========================================================================
' CALLBACK: outputorderNodesFirst_onAction
'
' PURPOSE:
'   Updates the 'outputorder' Graphviz attribute via the Ribbon.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Updates the 'SETTINGS_GRAPH_OUTPUT_ORDER' named
'      range based on the selected sequence.
'   2. UI REFRESH: Invokes 'RefreshOutputorderGroup' to update toggle states.
'   3. REACTIVITY: Triggers 'AutoDraw' for immediate visual feedback.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorderNodesFirst_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "nodesfirst"
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString
    End If
    RefreshOutputorderGroup
    AutoDraw
End Sub

' ==========================================================================
' CALLBACK: outputorder_getVisible
'
' PURPOSE:
'   Controls the visibility of the Output Order parameter based on layout
'   engine compatibility.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CHECK: Evaluates the active engine in 'SETTINGS_GRAPHVIZ_ENGINE'.
'   2. ENGINE GATE: Returns TRUE for most engines, but forces FALSE for
'      'PATCHWORK', as rendering order is fixed in treemap layouts.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub outputorder_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = True
        Case LAYOUT_DOT
            visible = True
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_OSAGE
            visible = True
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = True
        Case LAYOUT_TWOPI
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ==========================================================================
' CALLBACK: algsep1_getVisible
'
' PURPOSE:
'   Manages the visibility of Ribbon separators within the Algorithm
'   group to maintain a clean UI layout.
'
' TECHNICAL WORKFLOW:
'   1. UI LOGIC: Logic-driven visibility based on the active engine.
'   2. CONTEXT: Reserved exclusively for 'DOT' layouts to group relevant
'      parameters.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub algsep1_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = False
        Case LAYOUT_DOT
            visible = True
        Case LAYOUT_FDP
            visible = False
        Case LAYOUT_NEATO
            visible = False
        Case LAYOUT_OSAGE
            visible = False
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = False
        Case LAYOUT_TWOPI
            visible = False
        Case Else
            visible = False
    End Select
End Sub

' ==========================================================================
' CALLBACK: algsep2_getVisible
'
' PURPOSE:
'   Manages the visibility of Ribbon separators within the Algorithm
'   group to maintain a clean UI layout.
'
' TECHNICAL WORKFLOW:
'   1. UI LOGIC: Logic-driven visibility based on the active engine.
'   2. CONTEXT: Appears for 'DOT', 'NEATO', and 'SFDP' to group relevant
'      parameters.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub algsep2_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = False
        Case LAYOUT_DOT
            visible = True
        Case LAYOUT_FDP
            visible = False
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_OSAGE
            visible = False
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = True
        Case LAYOUT_TWOPI
            visible = False
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for Help

' ==========================================================================
' CALLBACK: graphvizHelp_onAction
'
' PURPOSE:
'   Redirects the user to the official Graphviz documentation or a
'   project-specific help page for the Graphviz Tab.
'
' TECHNICAL WORKFLOW:
'   1. URL RESOLUTION: Retrieves the target URL from the "HelpURLGraphvizTab"
'      named range on the Settings worksheet.
'   2. NAVIGATION: Invokes 'ActiveWorkbook.FollowHyperlink' to launch the
'      link in the user's default web browser.
'
' TECHNICAL NOTES:
'   - Trigger: Ribbon -> Graphviz Tab -> Help button (Last group).
'   - Strategy: Centralizes documentation links within the workbook's
'     Settings sheet to allow for URL updates without code modification.
' ==========================================================================
'@Ignore ParameterNotUsed
Public Sub graphvizHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLGraphvizTab").value, NewWindow:=True
End Sub

