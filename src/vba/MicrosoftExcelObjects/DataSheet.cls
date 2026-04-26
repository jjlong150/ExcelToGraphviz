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
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
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
'       • Calls AutoDraw on Worksheet_Change to provide instant visual feedback
'
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Graphviz Ribbon tab on entry
'       • Clears the StatusBar on entry and exit for UI hygiene
'
'   - INTELLIGENT DATA ENTRY:
'       • Dynamically filters Style Name dropdowns based on row type
'       • Populates the ChoicesSheet with matching styles
'       • Resizes and sorts RANGE_STYLE_NAMES in real time
'
'   - STRUCTURAL INTERPRETATION:
'       • getRowType identifies Nodes, Edges, Cluster boundaries, Native DOT,
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
    If Application.EnableEvents = False Then Exit Sub
    If SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_AUTO Then AutoDraw
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
    ' Exit if multiple cells selected?
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
            ChoicesSheet.Cells.item(1, 1).Resize(styleNames.count, 1).Value2 = Application.WorksheetFunction.Transpose(styleNames.Keys)
            
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

' ==========================================================================
' FUNCTION: getRowType
' PURPOSE:
'   Determines the logical identity of a worksheet row based on its 'Item'
'   and 'Is Related To' values.
'
' TECHNICAL WORKFLOW:
'   1. PRIMARY DETECTION: Scans the 'Item' column for structural markers:
'      - '{' -> TYPE_SUBGRAPH_OPEN (Cluster start)
'      - '}' -> TYPE_SUBGRAPH_CLOSE (Cluster end)
'      - '>' -> TYPE_NATIVE (Raw DOT injection)
'   2. RELATIONSHIP ANALYSIS: If no structural marker is found, it checks
'      the 'Is Related To' column:
'      - If populated -> TYPE_EDGE (A connection between nodes)
'   3. KEYWORD EVALUATION: If still unresolved, it checks for specific
'      Reserved Keywords (node, edge, graph) to assign global defaults.
'   4. FALLBACK: Defaults to 'TYPE_NODE' if the cell contains data but
'      no other patterns match.
'
' USAGE:
'   - Crucial for 'Worksheet_SelectionChange' to filter the Style dropdown.
'   - Used by the rendering engine to translate rows into DOT syntax.
' ==========================================================================
Private Function getRowType(ByVal worksheetName As String, ByVal row As Long) As String

    Dim rowType As String
    '@Ignore AssignmentNotUsed
    rowType = TYPE_BLANK_ROW
    
    Dim dataItem As String
    dataItem = UCase$(GetCell(worksheetName, row, GetSettingColNum(SETTINGS_DATA_COL_ITEM)))

    If dataItem <> vbNullString Then
        If EndsWith(dataItem, OPEN_BRACE) Then
            rowType = TYPE_SUBGRAPH_OPEN
        
        ElseIf dataItem = CLOSE_BRACE Then
            rowType = TYPE_SUBGRAPH_CLOSE
        
        ElseIf dataItem = GREATER_THAN Then
            rowType = TYPE_NATIVE
        
        Else
            Dim dataIsRelatedtoItem As String
            dataIsRelatedtoItem = GetCell(worksheetName, row, GetSettingColNum(SETTINGS_DATA_COL_IS_RELATED_TO))
            
            If dataIsRelatedtoItem = vbNullString Then
                If dataItem = KEYWORD_NODE Then
                    rowType = TYPE_NODE
                ElseIf dataItem = KEYWORD_EDGE Then
                    rowType = TYPE_EDGE
                ElseIf dataItem = KEYWORD_GRAPH Then
                    rowType = TYPE_GRAPH
                Else
                    rowType = TYPE_NODE
                End If
            Else
                rowType = TYPE_EDGE
            End If
        End If
    End If

    getRowType = rowType
    
End Function

' ==========================================================================
' FUNCTION: getMatchingStyles
' PURPOSE:
'   Retrieves a unique list of style names from the 'Styles' sheet that
'   correspond to a specific object type.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA RESOLUTION: Fetches the 'Styles' worksheet layout (columns for
'      Flag, Type, and Name) using 'GetSettingsForStylesWorksheet'.
'   2. FILTERED SCAN: Iterates through the style gallery:
'      - Skips 'Comment' rows (rows with a '#' flag).
'      - Matches the 'Style Type' column against the provided 'rowType'.
'   3. DEDUPLICATION: Uses a 'Dictionary' to collect unique style names,
'      automatically ignoring duplicate entries.
'   4. SYSTEM CLEANUP: Removes internal structural keywords (like 'node',
'      'edge', 'subgraph_open') from the result set to ensure only
'      user-defined, selectable styles appear in the dropdown.
'
' USAGE:
'   - Powers the dynamic dropdown menus on the 'Data' worksheet.
' ==========================================================================
Private Function getMatchingStyles(ByVal rowType As String) As Dictionary

    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    ' Dictionary to hold the key and associated values
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    ' Loop through the specified range
    Dim row As Long
    Dim styleName As String
    
    For row = styles.firstRow To styles.lastRow
        '@Ignore EmptyIfBlock
        If StylesSheet.Cells.item(row, styles.flagColumn).value = FLAG_COMMENT Then
            ' Comment row, ignore it
        ElseIf StylesSheet.Cells.item(row, styles.typeColumn).value = rowType Then
            ' Retrieve the style name
            styleName = StylesSheet.Cells.item(row, styles.nameColumn).value

            If styleName <> vbNullString Then    ' a style name is present
                ' Add style name to the dictionary if not already present
                If Not dictionaryObj.Exists(styleName) Then ' ignore duplicate style names
                    dictionaryObj.Add styleName, "1"    ' value of name/value pair is not used
                End If
            End If
        End If
    Next row
    
    If dictionaryObj.Exists(TYPE_NODE) Then dictionaryObj.Remove (TYPE_NODE)
    If dictionaryObj.Exists(TYPE_EDGE) Then dictionaryObj.Remove (TYPE_EDGE)
    If dictionaryObj.Exists(TYPE_SUBGRAPH_OPEN) Then dictionaryObj.Remove (TYPE_SUBGRAPH_OPEN)
    If dictionaryObj.Exists(TYPE_SUBGRAPH_CLOSE) Then dictionaryObj.Remove (TYPE_SUBGRAPH_CLOSE)
    If dictionaryObj.Exists(TYPE_KEYWORD) Then dictionaryObj.Remove (TYPE_KEYWORD)
    If dictionaryObj.Exists(TYPE_NATIVE) Then dictionaryObj.Remove (TYPE_NATIVE)

    Set getMatchingStyles = dictionaryObj
    
End Function


