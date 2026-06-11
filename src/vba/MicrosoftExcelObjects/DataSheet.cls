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
' SUB: Worksheet_SelectionChange
' PURPOSE:
'   Dynamically updates the Style Name dropdown list (data validation)
'   whenever the user selects a cell in the Style column.
'
' BEHAVIOR:
'   - Clears the previous list and rebuilds it based on the current row's type.
'   - Only processes rows that support styling (skips blank rows and graphs).
'   - Relies on getMatchingStyles() which returns a pre-sorted, cached list.
'
' PERFORMANCE DESIGN:
'   - Multiple early-exit conditions to minimize unnecessary work.
'   - Temporarily disables ScreenUpdating, Calculation, and Events.
'   - Uses pre-sorted data from getMatchingStyles (no redundant sorting here).
'
' WORKFLOW:
'   1. Early exits for irrelevant selections.
'   2. Clear existing named range contents.
'   3. Determine rowType using getRowType().
'   4. Retrieve (cached) style list via getMatchingStyles().
'   5. Write styles to ChoicesSheet and resize the named range.
'
' DEPENDENCIES:
'   - RANGE_STYLE_NAMES (workbook named range)
'   - getMatchingStyles() function (provides cached + sorted results)
'   - getRowType(), GetSettingColNum()
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
    
    ' Skip rows that don't use styles
    If rowType = TYPE_BLANK_ROW Or rowType = TYPE_GRAPH Then Exit Sub
        
    ' Adjust Excel Performance Settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.enableEvents = False
    
    On Error GoTo CleanExit
    
    ' Use a dictionary to load the unique style names for this row type
    Dim styleNames As Dictionary
    Set styleNames = getMatchingStyles(rowType)
    
    If styleNames.count = 0 Then GoTo CleanExit
    
    ' Populate Dropdown List
    With ChoicesSheet
        Dim count As Long
        count = styleNames.count
        
        ' Write the (already sorted) styles
        .Cells(1, 1).Resize(count, 1).Value2 = _
            Application.WorksheetFunction.Transpose(styleNames.keys)
        
        ' Update named range size
        Dim namedRange As name
        Set namedRange = ActiveWorkbook.names(RANGE_STYLE_NAMES)
        
        With namedRange
            .RefersTo = .RefersToRange.Resize(count, 1)
        End With
    End With
    
CleanExit:
    ' Restore Excel settings
    Application.enableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Clean up objects
    Set namedRange = Nothing
    Set styleNames = Nothing
End Sub

