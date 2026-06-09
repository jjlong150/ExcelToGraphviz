VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: DataSheet
' COPYRIGHT:        Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Data
'
' ROLE:
'   Primary data-entry surface for graph construction. Provides reactive
'   rendering, context-aware Ribbon activation, and intelligent style
'   filtering to support a seamless "edit -> preview" workflow.
'
' DESCRIPTION:
'   This worksheet module manages all interactive behavior on the "data" sheet.
'   It activates the Graphviz Ribbon tab when selected, triggers AutoDraw
'   whenever data changes, and dynamically filters the Style Name dropdown
'   based on the logical identity of the current row (Node, Edge, Cluster
'   Open/Close, Native DOT, etc.). The module serves as the UI bridge between
'   user edits and the graph-generation engine.
'
' KEY RESPONSIBILITIES:
'   - REACTIVE RENDERING:
'       o Calls AutoDraw on Worksheet_Change to provide instant visual feedback
'
'   - CONTEXT-AWARE NAVIGATION:
'       o Activates the Graphviz Ribbon tab on entry
'       o Clears the StatusBar on entry and exit for UI hygiene
'
'   - INTELLIGENT DATA ENTRY:
'       o Dynamically filters Style Name dropdowns based on row type
'       o Populates the ChoicesSheet with matching styles
'       o Resizes and sorts RANGE_STYLE_NAMES in real time
'
'   - STRUCTURAL INTERPRETATION:
'       o getRowType identifies Nodes, Edges, Cluster boundaries, Native DOT,
'         and Graph-level directives
'
' USAGE:
'   - Fires automatically during data entry, selection changes, and sheet
'     activation/deactivation
'   - Provides the core "edit -> preview" workflow for graph construction
'
' RELATED WIKI PAGES:
'   - Data Worksheet Architecture
'   - AutoDraw Pipeline
'   - Style Filtering & Dropdown Logic
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the Data worksheet's specific tools.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabGraphviz' to switch the Ribbon
'      focus to the primary graphing controls.
'   2. PLATFORM ADAPTATION:
'      - macOS: Uses 'Application.OnTime' to introduce a one-second delay,
'        preventing UI race conditions in the Mac sandbox.
'      - Windows: Executes the activation immediately for a snappier feel.
'   3. STATUS CLEANUP: Resets 'Application.StatusBar' to 'False' to clear
'      any lingering messages from previous SQL or Graphing operations.
'
' USAGE:
'   - Fires automatically whenever the user clicks the "Data" tab.
'   - Part of the "Context-Sensitive" UI strategy.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabGraphviz"
#Else
    ActivateTabGraphviz
#End If
    Application.StatusBar = False
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Change
'
' PURPOSE:
'   Responds to user edits on the 'data' worksheet and conditionally
'   triggers a graph refresh based on the current run mode.
'
' TECHNICAL WORKFLOW:
'   1. SAFETY CHECK:
'        - Exits immediately if events are disabled, preventing recursion
'          during programmatic updates or internal redraw operations.
'
'   2. MODE GATE:
'        - Reads the 'SETTINGS_RUN_MODE' named range on the Settings sheet.
'        - Only invokes AutoDraw when the workbook is in "Automatic" mode.
'
'   3. RENDER DISPATCH:
'        - Delegates to AutoDraw, which handles event suspension, screen
'          suppression, and the full graph-generation pipeline.
'
' USAGE:
'   - Provides the live-update behavior for users editing graph data.
'   - Integrates seamlessly with manual/auto toggling.
' ==========================================================================
'@Ignore ParameterNotUsed
Private Sub Worksheet_Change(ByVal Target As Range)
    If Application.enableEvents = False Then Exit Sub
    If SettingsSheet.Range(SETTINGS_RUN_MODE).value <> TOGGLE_AUTO Then Exit Sub
    AutoDraw
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Deactivate
' PURPOSE:
'   Performs UI housekeeping when the user switches away from the Data sheet.
'
' TECHNICAL WORKFLOW:
'   1. STATUS RESET: Clears the 'Application.StatusBar' to ensure that
'      context-specific messages (like "Ready to Draw") do not persist
'      on other worksheets.
'
' USAGE:
'   - Fires automatically whenever the user clicks another tab.
'   - Maintains a clean, professional workspace by removing stale feedback.
' ==========================================================================
Private Sub Worksheet_Deactivate()
    Application.StatusBar = False
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_SelectionChange
' PURPOSE:
'   Orchestrates the context-sensitive "Style Name" dropdown list.
'
' TECHNICAL WORKFLOW:
'   1. TARGET VALIDATION: Exit early if multiple cells are selected or if
'      the cursor is outside the 'Style Name' column.
'   2. CONTEXT ANALYSIS: Calls 'getRowType' to determine the row's logical
'      role (Node, Edge, Subgraph Open/Close, etc.).
'   3. STYLE FILTERING: Retrieves a filtered Dictionary of style names
'      via 'getMatchingStyles' relevant only to that row type.
'   4. CACHE SYNCHRONIZATION:
'      - Clears and repopulates the 'Choices' worksheet with matching keys.
'      - Resizes the 'StyleNames' Named Range to fit the new result set.
'      - Performs an 'xlAscending' sort to ensure a clean user interface.
'   5. RESOURCE HYGIENE: Explicitly destroys objects to prevent memory leaks.
'
' USAGE:
'   - Ensures that Edge styles never appear for Nodes, and vice versa.
'   - Provides the "Intelligent Data Entry" experience.
' ==========================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Exit if multiple cells selected
    If Target.Cells.CountLarge > 1 Then Exit Sub
    
    ' Exit if the cell is outside the Style Name column
    If Target.Column <> GetSettingColNum(SETTINGS_DATA_COL_STYLE) Then Exit Sub
    
    ' Clear the named range which is associated with the Style Name dropdown list
    ChoicesSheet.Range(RANGE_STYLE_NAMES).ClearContents
    
    ' Exit if heading row
    If Target.row = 1 Then Exit Sub
    
    ' Examine the row and determine how it is being used
    Dim rowType As String
    rowType = getRowType(ActiveSheet.name, Target.row)
    
    ' Use a dictionary to load the unique style names for this row type
    Dim styleNames As Dictionary
    Set styleNames = getMatchingStyles(rowType)
    
    If styleNames.count > 0 Then
        ' For rows which can be styled
        If rowType <> TYPE_BLANK_ROW And rowType <> TYPE_GRAPH Then
            ' Paste the dictionary keys into the named range associated with the
            ' Style Name dropdown list
            ChoicesSheet.Cells.item(1, 1).Resize(styleNames.count, 1).Value2 = Application.WorksheetFunction.Transpose(styleNames.keys)
            
            ' Adjust the size of the named range to match the number of values
            Dim namedRange As name
            Set namedRange = ActiveWorkbook.names.[_Default](RANGE_STYLE_NAMES)
            
            With namedRange
                '@Ignore ImplicitDefaultMemberAccess
                .RefersTo = .RefersToRange.Resize(styleNames.count, 1)
            End With
            
            ' Sort the values within the named range
            Dim columnRange As String
            columnRange = "A1:A" & styleNames.count
            
            ChoicesSheet.sort.SortFields.Clear
            ChoicesSheet.sort.SortFields.Add key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ChoicesSheet.sort
                .SetRange Range(columnRange)
                .Header = xlNo
                .MatchCase = False
                .orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    End If
    
    ' Don't leak memory
    Set namedRange = Nothing
    Set styleNames = Nothing
End Sub

