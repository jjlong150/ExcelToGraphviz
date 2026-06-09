VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: SqlSheet
' COPYRIGHT:        Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / SQL
'
' ROLE:
'   Interactive SQL development surface supporting authoring, testing, and
'   execution of workbook-driven SQL statements. Provides floating UI
'   buttons, contextual Ribbon activation, and row-level validation to ensure
'   safe and predictable execution against the SQL engine.
'
' DESCRIPTION:
'   This worksheet module governs the behavior of the SQL authoring sheet.
'   When activated, it switches the Ribbon to the SQL tab, refreshes
'   SQL-specific controls, and prepares the workspace for editing. As the
'   user moves through the sheet, a "Follow-the-Cursor" floating UI system
'   injects contextual Run, Edit, and Status buttons next to the active row.
'   These controls are validated against row type and SQL semantics to
'   prevent illegal operations. The module also manages cleanup, Ribbon
'   invalidation, and platform-safe activation.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE RIBBON ACTIVATION:
'       o Activates the SQL Ribbon tab on entry
'       o Invalidates Edit Cell and Clipboard controls to sync with cursor
'
'   - FLOATING UI ENGINE:
'       o Uses Worksheet_SelectionChange to purge and recreate floating
'         Run/Edit/Status buttons
'       o Positions buttons using ButtonConfig metadata (HOffset/VOffset)
'       o Uses Unicode iconography (ChrW) for high-resolution, zero-dependency
'         button visuals
'
'   - ROW-LEVEL VALIDATION:
'       o Binds ValidationFunc (e.g., IsSqlRowActive) to Run buttons
'       o Prevents execution on comment rows, blank rows, or invalid SQL
'
'   - PLATFORM-SAFE ACTIVATION:
'       o Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       o Performs immediate activation on Windows for seamless transitions
'
'   - UI HYGIENE & STATE MANAGEMENT:
'       o Invalidates Ribbon controls on selection changes and deactivation
'       o Ensures SQL-specific UI does not leak into other worksheets
'
' USAGE:
'   - Fires automatically when navigating to or interacting with the SQL sheet
'   - Provides the primary environment for SQL authoring, testing, and execution
'
' RELATED WIKI PAGES:
'   - SQL Worksheet Architecture
'   - Floating UI Engine
'   - SQL Execution Pipeline & ADO Provider Notes
'   - SQL Edit Form & Clipboard Tools
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI and initializes the SQL developer workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabSql' to surface database-specific
'      utilities (Run SQL, Connection Pool, Batch Run).
'   2. STATE REFRESH: Specifically invalidates the 'Edit Cell' and 'Clipboard'
'      ribbon controls to ensure they sync with the starting cursor position.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSql"
#Else
    ActivateTabSql
#End If
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Deactivate
' PURPOSE:
'   Synchronizes the Ribbon UI state upon exiting the SQL workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Explicitly refreshes the 'Edit Cell' and 'Clipboard'
'      Ribbon controls.
'   2. STATE TRANSITION: Ensures that the Ribbon "forgets" the specific
'      SQL-row context to avoid illegal operation attempts on standard
'      worksheet ranges.
'
' USAGE:
'   - Fires automatically whenever the user clicks away from the SQL sheet.
'   - Maintains UI hygiene by resetting specialized developer tools.
' ==========================================================================
Private Sub Worksheet_Deactivate()
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_SelectionChange
' PURPOSE:
'   Maintains the "Follow-the-Cursor" UI for SQL execution.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Forces a Ribbon refresh to update the 'Active Row'
'      status and toggle button availability (e.g., Run vs. Comment).
'   2. FLOATING UI MANAGEMENT:
'      - PURGE: Removes existing floating buttons to prevent duplicates.
'      - INJECTION: Calls 'UpdateFloatingButtonsOnSheet' to recreate
'        interactive "Run" and "Copy" buttons next to the selected row.
' ==========================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
    RefreshRibbon
    On Error GoTo 0

    ' Get configs once -> reuse for both creation and cleanup
    Dim configs() As ButtonConfig
    configs = GetFloatingButtonConfigs()

    ' Cleanup using the shared utility
    RemoveFloatingButtons Me, GetButtonNamesFromConfigs(configs)

    ' Create buttons using the shared utility
    UpdateFloatingButtonsOnSheet Me, Target, configs
End Sub

' ==========================================================================
' FUNCTION: GetFloatingButtonConfigs
' PURPOSE:
'   Defines the behavior, appearance, and logic for on-sheet SQL controls.
'
' TECHNICAL WORKFLOW:
'   1. REGISTRATION: Initializes an array of 'ButtonConfig' UDTs to define
'      the interactive command suite.
'   2. POSITIONING: Sets specific horizontal (HOffset) and vertical (VOffset)
'      coordinates to stack buttons neatly next to cell boundaries.
'   3. ICONOGRAPHY: Uses Unicode characters (ChrW) for lightweight,
'      high-resolution visual cues:
'      - Edit: Triggers the large-format SQL editor.
'      - Run: Executes the specific query row.
'      - Status: Opens the editor focused on the result/error status.
'   4. LOGIC BINDING:
'      - 'ActionMacro': Links the button to its corresponding VBA procedure.
'      - 'ValidationFunc': specifically uses 'IsSqlRowActive' for the Run
'        button to ensure only valid SELECT queries show the "Play" icon.
' ==========================================================================
Private Function GetFloatingButtonConfigs() As ButtonConfig()
    Dim configs() As ButtonConfig
    ReDim configs(0 To 2)   ' change upper bound when adding more

    With configs(0)
        .ColumnNumber = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
        .ButtonName = "SqlEditButton"
        .ActionMacro = "ShowSQLEditForm"
        .IconText = ChrW(&H270E)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = ""
    End With

    With configs(1)
        .ColumnNumber = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
        .ButtonName = "SqlRunButton"
        .ActionMacro = "RunOneSqlStatement"
        .IconText = ChrW(&H25B6)
        .VOffset = 22
        .HOffset = 5
        .ValidationFunc = "IsSqlRowActive"
    End With

    With configs(2)
        .ColumnNumber = GetSettingColNum(SETTINGS_SQL_COL_STATUS)
        .ButtonName = "SqlStatusButton"
        .ActionMacro = "ShowSQLEditForm"
        .IconText = ChrW(&H2315)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = ""
    End With

    GetFloatingButtonConfigs = configs
End Function


    
