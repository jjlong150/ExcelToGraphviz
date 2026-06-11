Attribute VB_Name = "modWorksheetData"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetData
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Data
'
' ROLE:
'   Manage lifecycle operations for Data-model worksheets, including safe
'   clearing of record rows, schema-aware range targeting, and multi-sheet
'   compatibility. Ensures all operations honor the Named Range API contract
'   and the dataWorksheet UDT.
'
' RESPONSIBILITIES:
'   - Data clearing:
'       o ClearDataWorksheet: purge record rows while preserving header rows,
'         formatting, and structural integrity
'
'   - Schema adaptation:
'       o Resolve firstRow and headingRow via GetSettingsForDataWorksheet
'       o Use GetLastColumn to dynamically determine horizontal boundaries
'         regardless of column reordering
'
'   - Multi-target support:
'       o Operates on any worksheet conforming to the Data-model schema
'         (Data, SQL, Imports, etc.)
'
' ARCHITECTURAL NOTES:
'   - Uses dataWorksheet UDT to remain independent of physical cell locations.
'   - Ensures header rows are never cleared, even when UsedRange reports
'     minimal content.
'   - Dynamic range construction ensures compatibility with user-driven
'     column movement and schema evolution.
'   - Consumed by Import, SQL Refresh, and initialization workflows.
'
' USAGE:
'   - Ideal for resetting Data-model sheets before imports, SQL refreshes,
'     or batch graph-generation cycles.
'
' RELATED WIKI PAGES:
'   - Data Worksheet Contract (Named Range API)
'   - Schema-Aware Clearing Logic
'   - Data Lifecycle & Refresh Pipeline
' =============================================================================

Option Explicit

Private styleCache As Dictionary
Private cacheIsValid As Boolean

' ==========================================================================
' PROCEDURE: ClearDataWorksheet
'
' PURPOSE:
'   Purges all record data from a specified Data worksheet while preserving
'   the integrity of the header row and worksheet structure.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Invokes 'GetSettingsForDataWorksheet' to retrieve
'      the 'dataWorksheet' UDT, identifying the 'firstRow' and 'headingRow'
'      per the Named Range API "Contract."
'   2. BOUNDARY CALCULATION:
'      - Vertical: Identifies the 'lastRow' using 'UsedRange'. Includes
'        safety logic to ensure the header row is never targeted.
'      - Horizontal: Resolves the 'lastColumn' via 'GetLastColumn' to
'        capture all dynamic attribute fields.
'   3. BULK PURGE: Constructs a dynamic 'cellRange' and executes
'      '.ClearContents' for high-performance data removal.
'
' USAGE:
'   - Essential for resetting the workspace before a new "Import" or "SQL
'     Refresh" operation.
'   - Supports multi-sheet targets by passing the 'worksheetName' parameter.
' ==========================================================================
Public Sub ClearDataWorksheet(ByVal worksheetName As String)
    Dim lastColumn As Long
    Dim cellRange As String
    Dim lastRow As Long
    Dim dataLayout As dataWorksheet
    
    ' Get the layout of the 'data' worksheet
    dataLayout = GetSettingsForDataWorksheet(worksheetName)

    ' Determine the range of the cells which need to be cleared
    With ActiveWorkbook.worksheets.[_Default](worksheetName).UsedRange
        lastRow = .Cells(.Cells.count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < dataLayout.firstRow Then
        lastRow = dataLayout.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(worksheetName, dataLayout.headingRow)

    ' Remove any existing content
    cellRange = "A" & dataLayout.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    ActiveWorkbook.worksheets.[_Default](worksheetName).Range(cellRange).ClearContents
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
Public Function getRowType(ByVal worksheetName As String, ByVal row As Long) As String

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
' SUB: InvalidateStyleCache
' PURPOSE:
'   Clears the cache so that getMatchingStyles will rebuild fresh data
'   the next time it is called.
'   Call this whenever styles are added, modified, or deleted on the Styles sheet.
' ==========================================================================
Public Sub InvalidateStyleCache()
    cacheIsValid = False
    
    If Not styleCache Is Nothing Then
        styleCache.RemoveAll
    End If
End Sub

' ==========================================================================
' FUNCTION: getMatchingStyles
' PURPOSE:
'   Retrieves a unique, sorted list of style names from the 'Styles' sheet
'   that correspond to a specific object type (rowType).
'
' PERFORMANCE FEATURES:
'   - Array-based reading of the Styles sheet (much faster than cell-by-cell access)
'   - Dictionary caching by rowType to minimize repeated processing
'   - Returns pre-sorted results for consistent dropdown ordering
'
' TECHNICAL WORKFLOW:
'   1. CACHE CHECK: Returns cached (already sorted) result if available and valid.
'   2. SCHEMA RESOLUTION: Calls GetSettingsForStylesWorksheet() to get column
'      positions and data range.
'   3. BULK READ: Loads the entire relevant portion of the Styles sheet into
'      a Variant array for high performance.
'   4. FILTERING & DEDUPLICATION:
'      - Skips rows marked as comments (FLAG_COMMENT)
'      - Matches rows based on the Style Type column
'      - Uses Dictionary to ensure unique style names
'   5. CLEANUP: Removes internal structural types (node, edge, subgraph_*, etc.)
'   6. SORTING: Sorts the style names using GetSortedKeys (QuickSort)
'   7. CACHING: Stores the final sorted dictionary for future calls
'
' USAGE:
'   - Called from Worksheet_SelectionChange to populate the dynamic
'     Style Name dropdown list.
'   - Cache is invalidated externally when the Styles sheet is modified.
'
' DEPENDENCIES:
'   - Module-level variables: styleCache (Dictionary), cacheIsValid (Boolean)
'   - Helper function: GetSortedKeys()
'   - Constants: FLAG_COMMENT, TYPE_NODE, TYPE_EDGE, etc.
' ==========================================================================
Public Function getMatchingStyles(ByVal rowType As String) As Dictionary
    
    ' Skip special row types
    If rowType = TYPE_BLANK_ROW Or rowType = TYPE_GRAPH Then
        Set getMatchingStyles = New Dictionary
        Exit Function
    End If
    
    ' Initialize cache
    If styleCache Is Nothing Then
        Set styleCache = New Dictionary
        cacheIsValid = False
    End If
    
    ' Return from cache if possible
    If cacheIsValid And styleCache.Exists(rowType) Then
        Set getMatchingStyles = styleCache(rowType)
        Exit Function
    End If
    
    ' === Build dictionary for this rowType ===
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    Dim data As Variant
    Dim lastCol As Long
    lastCol = Application.max(styles.nameColumn, styles.typeColumn, styles.flagColumn)
    
    data = StylesSheet.Cells(styles.firstRow, 1) _
                     .Resize(styles.lastRow - styles.firstRow + 1, lastCol).Value2
    
    Dim relativeRow As Long
    Dim styleName As String
    
    For relativeRow = 1 To UBound(data, 1)
        If data(relativeRow, styles.flagColumn) = FLAG_COMMENT Then GoTo NextRow
        
        If data(relativeRow, styles.typeColumn) = rowType Then
            styleName = Trim$(data(relativeRow, styles.nameColumn) & "")
            If styleName <> vbNullString Then
                If Not dict.Exists(styleName) Then
                    dict.Add styleName, vbNullString
                End If
            End If
        End If
NextRow:
    Next relativeRow
    
    ' Remove special types
    Dim itemsToRemove As Variant
    itemsToRemove = Array(TYPE_NODE, TYPE_EDGE, TYPE_SUBGRAPH_OPEN, _
                         TYPE_SUBGRAPH_CLOSE, TYPE_KEYWORD, TYPE_NATIVE)
    
    Dim i As Long
    For i = LBound(itemsToRemove) To UBound(itemsToRemove)
        If dict.Exists(itemsToRemove(i)) Then dict.Remove itemsToRemove(i)
    Next i
    
    ' === Create Sorted Dictionary ===
    Dim sortedDict As Dictionary
    Set sortedDict = New Dictionary
    
    If dict.count > 0 Then
        Dim sortedKeys As Variant
        sortedKeys = GetSortedKeys(dict)
        
        For i = LBound(sortedKeys) To UBound(sortedKeys)
            sortedDict.Add sortedKeys(i), vbNullString
        Next i
    End If
    
    ' Cache the sorted dictionary
    If styleCache.Exists(rowType) Then
        Set styleCache(rowType) = sortedDict
    Else
        styleCache.Add rowType, sortedDict
    End If
    
    cacheIsValid = True
    Set getMatchingStyles = sortedDict
    
End Function

' ==========================================================================
' FUNCTION: GetSortedKeys
'
' PURPOSE:
'   Produces a reliably sorted array of Dictionary keys, ensuring stable,
'   alphabetical ordering for downstream consumers such as dropdown lists,
'   style galleries, and cache-driven lookup tables.
'
' PERFORMANCE FEATURES:
'   - Zero-overhead extraction: Reads keys directly from the Dictionary
'     without intermediate collections.
'   - Array-based sorting: Converts keys to a fixed-size String array for
'     optimal QuickSort performance.
'   - Lean branching: Immediately returns an empty array when the Dictionary
'     contains no entries.
'
' TECHNICAL WORKFLOW:
'   1. EMPTY CHECK:
'        - If the Dictionary has no entries, returns an empty Variant array.
'
'   2. KEY EXTRACTION:
'        - Allocates a String array sized exactly to dict.Count.
'        - Iterates the Dictionary’s Keys collection, copying each key into
'          the array in its native String form.
'
'   3. SORTING:
'        - Delegates ordering to QuickSort(), ensuring deterministic,
'          case-sensitive alphabetical sorting.
'
'   4. RETURN VALUE:
'        - Outputs the sorted String array for use by callers that require
'          predictable ordering (e.g., dropdown population, cache hydration).
'
' USAGE:
'   - Called by getMatchingStyles() to produce a sorted list of style names.
'   - Suitable for any module requiring stable, alphabetical ordering of
'     Dictionary keys.
'
' DEPENDENCIES:
'   - QuickSort(): In-module or shared sorting routine implementing a
'     comparison-based quicksort algorithm.
'
' ==========================================================================
Private Function GetSortedKeys(ByVal dict As Dictionary) As Variant
    Dim keys() As String
    Dim i As Long
    
    If dict.count = 0 Then
        GetSortedKeys = Array()
        Exit Function
    End If
    
    ReDim keys(0 To dict.count - 1)
    
    i = 0
    Dim key As Variant
    For Each key In dict.keys
        keys(i) = CStr(key)
        i = i + 1
    Next key
    
    ' Sort using reliable method
    Call QuickSort(keys, 0, UBound(keys))
    
    GetSortedKeys = keys
    
End Function

' ==========================================================================
' PROCEDURE: QuickSort
'
' PURPOSE:
'   Performs an in-place quicksort on a zero-based String array, producing a
'   deterministic ascending alphabetical ordering. Used as the core sorting
'   engine for style lists, key collections, and other lookup structures
'   requiring high-performance ordering.
'
' PERFORMANCE FEATURES:
'   - In-place partitioning: Eliminates the need for temporary arrays,
'     minimizing memory churn during recursive calls.
'   - Median-pivot strategy: Selects the midpoint element as the pivot to
'     reduce worst-case behavior on already-sorted or reverse-sorted data.
'   - Tight inner loops: Uses compact comparison loops for minimal overhead
'     during partition scanning.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY CHECK:
'        - Exits immediately when the current segment (low..high) contains
'          fewer than two elements.
'
'   2. PIVOT SELECTION:
'        - Chooses the midpoint element as the pivot to balance partitions.
'
'   3. PARTITIONING:
'        - Moves two indices (i, j) inward:
'            • i advances while arr(i) < pivot
'            • j retreats while arr(j) > pivot
'        - Swaps elements when i <= j to maintain correct ordering.
'
'   4. RECURSION:
'        - Recursively sorts the left partition (low..j).
'        - Recursively sorts the right partition (i..high).
'
'   5. RESULT:
'        - The input array is fully sorted in ascending order upon return.
'
' USAGE:
'   - Called by GetSortedKeys() to alphabetize Dictionary key arrays.
'   - Suitable for any module requiring fast, in-place String sorting.
'
' DEPENDENCIES:
'   - None. Operates solely on the provided String array.
'
' ==========================================================================
Private Sub QuickSort(ByRef arr() As String, ByVal low As Long, ByVal high As Long)
    Dim pivot As String
    Dim i As Long, j As Long
    Dim temp As String
    
    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        
        Do While i <= j
            Do While arr(i) < pivot And i < high: i = i + 1: Loop
            Do While arr(j) > pivot And j > low: j = j - 1: Loop
            
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        Call QuickSort(arr, low, j)
        Call QuickSort(arr, i, high)
    End If
End Sub
