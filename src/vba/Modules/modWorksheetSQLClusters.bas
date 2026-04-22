Attribute VB_Name = "modWorksheetSQLClusters"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetSQLClusters
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / SQL
'
' ROLE:
'   The flagship engine for N-level hierarchical cluster nesting. Performs
'   multi-level cluster discovery, sorting, scope management, and token-driven
'   metadata injection for deeply nested subgraph structures.
'
' RESPONSIBILITIES:
'   - Multi-level detection:
'       • DetectMultiLevel and DetectMaxLevels probe CLUSTER1…CLUSTERn fields
'         and determine the active hierarchy depth.
'
'   - Hierarchy orchestration:
'       • ProcessMultiLevelRecordset performs delta-driven open/close logic,
'         maintains per-level counters, and ensures structural continuity.
'
'   - Cluster emission:
'       • EmitClusterOpen / EmitClusterClose write Graphviz subgraph braces,
'         labels, styles, attributes, and tooltips with suffix-aware formatting.
'
'   - Token substitution:
'       • ProcessClusterProperty applies {cluster}, {subcluster}, and {level}
'         placeholders for dynamic naming and styling.
'
'   - Data emission:
'       • EmitRows / EmitOneRow map SQL records into the Data worksheet while
'         filtering structural fields and applying enumeration and wrapping rules.
'
' ARCHITECTURAL NOTES:
'   - Built as an extension of the SQL subsystem's modern clustering model.
'   - Uses ADO recordset sorting to guarantee stable hierarchical ordering.
'   - Integrates with sqlContext, dataWorksheet, and sqlFieldName UDTs.
'   - Backward-compatible with legacy CLUSTER/SUBCLUSTER logic via dispatcher.
'
' VERSION NOTES:
'   - v10.3.0 (Apr 3, 2026):
'       • Introduced full N-level clustering (CLUSTER1, CLUSTER2, …)
'       • Added per-level label/style/attribute/tooltip fields
'       • Added {label} placeholder support for cluster label formatting
'       • Added revised format-string parsing for HTML-like syntax
'
' USAGE:
'   - Automatically invoked by RunSQL when CLUSTER1 is detected.
'   - Enables unlimited hierarchical depth with zero configuration.
'
' RELATED WIKI PAGES:
'   - SQL Engine & Multi-Level Clustering
'   - Hierarchical Graph Construction
'   - Token-Driven Label and Style Formatting
' =============================================================================

Option Explicit

' ==========================================================================
' FUNCTION: DetectMultiLevel
' PURPOSE:
'   Determines if the SQL results require the modern N-level clustering engine.
'   Detect if multi-level clustering is present (prefers over old "CLUSTER").
'
' TECHNICAL WORKFLOW:
'   1. SIGNATURE CHECK: Scans the recordset for the specific field 'CLUSTER1'.
'   2. PRECEDENCE LOGIC: If 'CLUSTER1' is found, this function returns True,
'      signaling the 'MapResultsToDataWorksheet' dispatcher to prioritize
'      Multi-Level processing over the legacy 'CLUSTER/SUBCLUSTER' logic.
'
' USAGE:
'   - The primary decision gate for modern hierarchical rendering.
'   - Allows for backward compatibility while enabling infinite nesting.
' ==========================================================================
Public Function DetectMultiLevel(ByVal rs As Object, prefix As String) As Boolean
    DetectMultiLevel = HasField(rs, prefix & "1")
End Function

' ==========================================================================
' PROCEDURE: ProcessMultiLevelRecordset
' PURPOSE:
'   The flagship engine for generating deeply nested graph hierarchies.
'
' TECHNICAL WORKFLOW:
'   1. LEVEL DISCOVERY: Calls 'DetectMaxLevels' to identify the depth of the
'      hierarchy (e.g., discovering CLUSTER1 through CLUSTER5).
'   2. IN-MEMORY SORTING: Dynamically builds an ADO Sort string to group
'      records by level (ASC) to ensure structural continuity during emission.
'   3. STATE TRACKING: Initializes parallel arrays to monitor 'currentClusters'
'      and 'clusterCounters' for every depth level.
'   4. DELTA ANALYSIS: For every record, 'FindChangeLevel' identifies the
'      exact tier where the hierarchy shifts (e.g., moving from Dept A to Dept B).
'   5. SYMMETRICAL EMISSION:
'      - Closes all "inner" levels that have ended.
'      - Opens "outer" levels that have just begun, maintaining both
'        Absolute and Relative (per-level) cluster counts.
'   6. ATOMIC MAPPING: Calls 'EmitRows' to populate the node/edge data
'      within the currently active nested scope.
'   7. FINAL SEALING: Closes all remaining open levels after the recordset
'      is exhausted to ensure valid DOT syntax.
' ==========================================================================
Public Sub ProcessMultiLevelRecordset( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Exit early if invalid or empty
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Build simple sort string (no expressions)
    Dim maxLevels As Long
    maxLevels = DetectMaxLevels(rs, ctx.fields.Cluster, ctx.fields.clusterLevelLimit)
    
    Dim sortStr As String
    sortStr = ""
    
    Dim i As Long
    For i = 1 To maxLevels
        Dim fieldName As String
        fieldName = "[" & ctx.fields.Cluster & i & "]"
        
        If Len(sortStr) > 0 Then sortStr = sortStr & ", "
        sortStr = sortStr & fieldName & " ASC"
    Next i

    ' Attempt sort with error handling
    On Error GoTo SortFailed
    rs.sort = sortStr
    On Error GoTo 0

    ' Ensure we're at the beginning after sort
    rs.MoveFirst

    ' Proceed with clustering logic
    Dim currentClusters() As String
    ReDim currentClusters(1 To maxLevels)
    
    ' Per-level cluster counters
    Dim clusterCounters() As Long
    ReDim clusterCounters(1 To maxLevels)
    For i = 1 To maxLevels
        clusterCounters(i) = 0
    Next i
    
    Dim openLevels As Long: openLevels = 0
    Dim currentOpenLevel As Long:  currentOpenLevel = 0   ' 0 = no cluster open yet (flat nodes)
    Dim absoluteClusterCount As Long: absoluteClusterCount = 0
    Dim relativeClusterCount As Long: relativeClusterCount = 0   ' fallback if no clusters ever opened
    
    Do While Not rs.EOF
        Dim thisClusters() As String
        ReDim thisClusters(1 To maxLevels)
        For i = 1 To maxLevels
            thisClusters(i) = SafeFieldValue(rs, ctx.fields.Cluster & i)
        Next i

        Dim changeLevel As Long
        changeLevel = FindChangeLevel(currentClusters, thisClusters, maxLevels)

        ' Close inner levels
        For i = openLevels To changeLevel Step -1
            EmitClusterClose ctx, rs, row, i, absoluteClusterCount, clusterCounters(i)
            openLevels = openLevels - 1
        Next i

        ' After closing, update currentOpenLevel to the new deepest open level
        currentOpenLevel = openLevels
        
        ' Open new clusters
        For i = changeLevel To maxLevels
            If thisClusters(i) <> "" Then
                ' Increment counter for **this level only**
                clusterCounters(i) = clusterCounters(i) + 1
                
                ' Remember this as the most recently opened relative count
                relativeClusterCount = clusterCounters(i)
                
                ' Increment absolute cluster count
                absoluteClusterCount = absoluteClusterCount + 1
                
                EmitClusterOpen ctx, rs, row, i, absoluteClusterCount, clusterCounters(i)
                
                openLevels = openLevels + 1
                currentOpenLevel = i
                currentClusters(i) = thisClusters(i)
                
                ' Reset deeper levels' counters
                Dim deeper As Long
                For deeper = i + 1 To maxLevels
                    currentClusters(deeper) = ""
                    clusterCounters(deeper) = 0
                Next deeper
            End If
        Next i

        ' Increment the rs record count
        recordCnt = recordCnt + 1
        
        ' Emit row
        EmitRows ctx, rs, row, maxLevels, recordCnt, currentOpenLevel, absoluteClusterCount, relativeClusterCount

        rs.MoveNext
    Loop

    ' Close remaining
    rs.MoveFirst
    For i = 1 To openLevels
        EmitClusterClose ctx, rs, row, i, absoluteClusterCount, clusterCounters(i)
    Next i

    Exit Sub

SortFailed:
    ' Continue without sorting (data will still be processed, just not ordered)
    On Error GoTo 0
    rs.MoveFirst
    Resume Next

End Sub

' ==========================================================================
' FUNCTION: DetectMaxLevels
' PURPOSE:
'   Determines the total number of sequential clustering tiers in the data.
'
' TECHNICAL WORKFLOW:
'   1. LINEAR PROBING: Starts at index 1 and incrementaly checks for the
'      existence of fields matching the pattern [Prefix] + [Index]
'      (e.g., CLUSTER1, CLUSTER2).
'   2. GOVERNOR COMPLIANCE: Respects the 'levelLimit' setting to ensure
'      the probe does not exceed user-defined or system boundaries.
'   3. TERMINATION: Stops as soon as a break in the sequence is found
'      (e.g., if CLUSTER1 and CLUSTER2 exist, but CLUSTER3 is missing).
'   4. COORDINATE CALCULATION: Returns the final successful index,
'      providing the 'maxLevels' value used to drive the clustering loop.
'
' USAGE:
'   - Called at the start of 'ProcessMultiLevelRecordset'.
'   - Enables "Zero-Configuration" hierarchies—just add columns to your
'     SQL and the engine adapts.
' ==========================================================================
Private Function DetectMaxLevels(ByVal rs As Object, prefix As String, levelLimit As Long) As Long
    Dim i As Long
    i = 1
    While i <= levelLimit And HasField(rs, prefix & i)
        i = i + 1
    Wend
    DetectMaxLevels = i - 1
End Function

' ==========================================================================
' FUNCTION: FindChangeLevel
' PURPOSE:
'   Compares the current row's hierarchy against the previous row to
'   identify the highest-level grouping shift.
'
' TECHNICAL WORKFLOW:
'   1. LINEAR COMPARISON: Iterates through the cluster levels (1 to maxLevels)
'      comparing the cached 'current' array with the 'thisOne' array.
'   2. BREAKPOINT IDENTIFICATION: Returns the index of the first level
'      where the values differ (e.g., if Level 1 is the same but Level 2
'      changes, it returns 2).
'   3. NO-CHANGE SIGNAL: If all levels match exactly, it returns 'maxLevels + 1',
'      signaling the engine to continue emitting nodes within the existing scope.
'
' USAGE:
'   - The core decision logic for 'ProcessMultiLevelRecordset'.
'   - Determines the "Symmetry Point" for closing and opening Graphviz braces.
' ==========================================================================
Private Function FindChangeLevel(ByRef current() As String, ByRef thisOne() As String, ByVal maxLevels As Long) As Long
    Dim i As Long
    For i = 1 To maxLevels
        If current(i) <> thisOne(i) Then
            FindChangeLevel = i
            Exit Function
        End If
    Next i
    FindChangeLevel = maxLevels + 1  ' No change
End Function

' ==========================================================================
' PROCEDURE: EmitClusterOpen (Multi-Level)
' PURPOSE:
'   Writes the structural 'Open' row for a specific tier in an N-level hierarchy.
'
' TECHNICAL WORKFLOW:
'   1. SCOPE INITIATION: Places the mandatory 'OPEN_BRACE' ({) in the
'      Data worksheet's 'Item' column to start the Graphviz subgraph.
'   2. PROPERTY DELEGATION: Calls 'ProcessClusterProperty' for each
'      visual attribute (Label, Style, Attributes, Tooltip).
'   3. TOKEN CONTEXT: Passes both 'absoluteClusterCount' (global) and
'      'relativeClusterCount' (per-level) to allow for sophisticated
'      style and label formatting.
'   4. STYLE SYNCHRONIZATION: Injects the global 'Suffix Open' constant
'      when processing the StyleName column to ensure proper registry lookup.
'   5. ROW MANAGEMENT: Automatically increments the worksheet 'row' index
'      after the write is complete.
' ==========================================================================
Private Sub EmitClusterOpen( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long)

    With DataSheet
        ' Mandatory opening brace
        .Cells(row, ctx.dataLayout.itemColumn).value = OPEN_BRACE

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterLabel, ctx.dataLayout.labelColumn

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterStyleName, ctx.dataLayout.styleNameColumn, _
            suffixFromSetting:=SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_OPEN).value

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterAttributes, ctx.dataLayout.extraAttributesColumn

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterTooltip, ctx.dataLayout.tooltipColumn
    End With
    
    row = row + 1
End Sub

' ==========================================================================
' PROCEDURE: EmitClusterClose (Multi-Level)
' PURPOSE:
'   Writes the structural 'Close' row for a specific tier in an N-level hierarchy.
'
' TECHNICAL WORKFLOW:
'   1. SCOPE SEALING: Places the mandatory 'CLOSE_BRACE' (}) in the
'      Data worksheet's 'Item' column to end the Graphviz subgraph.
'   2. STYLE RE-ALIGNMENT:
'      - Re-evaluates the style name for the current level.
'      - Appends the global 'Suffix Close' (e.g., "_CLOSE") to enable
'        visual differentiation between the start and end of a container.
'   3. TOKEN SUBSTITUTION: Passes the Absolute and Relative counters to
'      'ProcessClusterProperty' to maintain consistent naming/ID references.
'   4. ROW MANAGEMENT: Increments the worksheet 'row' index to prepare for
'      the next set of data or the next closing brace.
'
' USAGE:
'   - Called by 'ProcessMultiLevelRecordset' whenever a hierarchy depth
'     shift is detected or the recordset ends.
' ==========================================================================
Private Sub EmitClusterClose( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long)

    With DataSheet
        ' Mandatory: closing brace
        .Cells(row, ctx.dataLayout.itemColumn).value = CLOSE_BRACE

        ' Only style name is emitted on close (for stylesheet toggle support)
        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterStyleName, ctx.dataLayout.styleNameColumn, _
            suffixFromSetting:=SafeStr(SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_CLOSE).value)
    End With

    row = row + 1
End Sub

' ==========================================================================
' PROCEDURE: ProcessClusterProperty
' PURPOSE:
'   Extracts, transforms, and writes specific cluster attributes (Labels,
'   Styles, etc.) for a specific level of the N-tier hierarchy.
'
' TECHNICAL WORKFLOW:
'   1. FIELD RESOLUTION: Dynamically constructs the source field name by
'      appending the current 'levelNumber' to the base template (e.g.,
'      transforming "CLUSTER_LABEL" into "CLUSTER_LABEL2").
'   2. SUFFIX APPLICATION: Appends the 'Suffix Open' or 'Suffix Close'
'      to style names to maintain alignment with the Styles worksheet.
'   3. TRIPLE-TOKEN SUBSTITUTION: Resolves dynamic placeholders within
'      the value:
'      - {cluster}: Replaced by the Absolute Count (global index).
'      - {subcluster}: Replaced by the Relative Count (per-level index).
'      - {level}: Replaced by the current depth (1, 2, 3...).
'   4. DATA EMISSION: Writes the finalized, substituted string into the
'      target 'Data' worksheet column.
' ==========================================================================
Private Sub ProcessClusterProperty( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef wsRow As Long, _
    ByVal levelNumber As Long, _
    ByVal absCount As Long, _
    ByVal relCount As Long, _
    ByVal templateField As String, _
    ByVal targetColumn As Long, _
    Optional ByVal suffixFromSetting As String = vbNullString)

    Dim clusterPrefix As String
    clusterPrefix = ctx.fields.Cluster & levelNumber

    Dim fieldName As String
    fieldName = replace(templateField, ctx.fields.Cluster, clusterPrefix, , , vbTextCompare)

    If Not HasField(rs, fieldName) Then Exit Sub

    Dim value As String
    value = SafeFieldValue(rs, fieldName)

    ' Apply suffix if provided (different for open vs close)
    If Len(suffixFromSetting) > 0 Then
        value = value & suffixFromSetting
    End If

    ' Apply placeholders (same logic for all properties)
    value = replace(value, ctx.fields.clusterPlaceholder, SafeStr(absCount), , , vbTextCompare)
    value = replace(value, ctx.fields.subclusterPlaceholder, SafeStr(relCount), , , vbTextCompare)
    value = replace(value, ctx.fields.clusterLevelPlaceholder, SafeStr(levelNumber), , , vbTextCompare)

     ' Write result
    DataSheet.Cells(wsRow, targetColumn).value = value
End Sub

' ==========================================================================
' PROCEDURE: EmitRows (Multi-Level)
' PURPOSE:
'   Translates ADO records into worksheet rows while preserving the active
'   hierarchical state (Depth, Absolute ID, and Relative ID).
'
' TECHNICAL WORKFLOW:
'   1. SAFETY VALIDATION: Prevents infinite loops by verifying 'stepBy'
'      values and ensuring mathematical directionality (start vs stop).
'   2. LOOP ENUMERATION: Supports "Enumeration Mode" within a cluster,
'      allowing a single SQL record to generate a sequenced range of nodes.
'   3. STATE INJECTION: Passes the current clustering context—including
'      'levelNumber', 'absoluteClusterCount', and 'relativeClusterCount'—
'      down to the atomic 'EmitOneRow' writer.
'   4. GOVERNOR COMPLIANCE: Tracks the total 'ctx.loop.count' against
'      system limits to ensure stability in high-density graphs.
'   5. COORDINATE MANAGEMENT: Increments the worksheet 'row' index after
'      each successful write to maintain the vertical pipeline.
' ==========================================================================
Private Sub EmitRows( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal maxLevels As Long, _
    ByVal recordCount As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long)
    
    Dim i As Long

    ' Safety: prevent infinite loop
    If ctx.loop.stepBy = 0 Then Exit Sub

    ' Safety: prevent direction mismatch infinite loop
    If ctx.loop.stepBy > 0 Then
        If ctx.loop.startAt > ctx.loop.stopAt Then Exit Sub
    Else
        If ctx.loop.startAt < ctx.loop.stopAt Then Exit Sub
    End If

    For i = ctx.loop.startAt To ctx.loop.stopAt Step ctx.loop.stepBy
        ctx.loop.count = ctx.loop.count + 1
        If ctx.loop.count > ctx.loop.max Then Exit For
        
        EmitOneRow ctx, rs, row, maxLevels, recordCount, levelNumber, absoluteClusterCount, relativeClusterCount, i
        row = row + 1
    Next i

End Sub

' ==========================================================================
' PROCEDURE: EmitOneRow (Multi-Level)
' PURPOSE:
'   Maps an ADO record to the 'Data' worksheet while applying hierarchical tokens.
'
' TECHNICAL WORKFLOW:
'   1. FIELD FILTERING: Uses 'IsClusterRelatedField' to ignore structural
'      columns (like CLUSTER1) so they aren't mistakenly written as node data.
'   2. QUAD-TOKEN SUBSTITUTION: Resolves dynamic placeholders in every field:
'      - {record}: Sequential record count.
'      - {cluster}: The global Absolute Cluster ID.
'      - {subcluster}: The local Relative Cluster ID.
'      - {level}: The current depth of the active cluster.
'   3. ENUMERATION SUPPORT: If a loop is active, injects the {i} placeholder
'      using the 'enumStep' value.
'   4. AUTOMATIC TEXT WRAPPING: Identifies 'Label' and 'xLabel' fields and
'      applies 'SplitMultilineText' if a 'SPLIT_LENGTH' is detected.
'   5. DYNAMIC COLUMN MAPPING: Routes sanitized values to the correct
'      worksheet column based on the language-agnostic 'ctx.headings' map.
' ==========================================================================
Private Sub EmitOneRow( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal maxLevels As Long, _
    ByVal recordCount As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long, _
    ByVal enumStep As Long)
    
    ' Process all rs fields and write to 'data' worksheet
    With DataSheet
        Dim fld As Object
        Dim v As String
        Dim targetCol As Long
        
        For Each fld In rs.fields
            If Not IsClusterRelatedField(fld.name, maxLevels) Then

                ' Common transformation: null -> "", placeholder replacement
                v = SafeStr(fld.value)
                
                If Len(v) > 0 Then
                    v = replace(v, ctx.fields.recordsetPlaceholder, CStr(recordCount), , , vbTextCompare)
                    v = replace(v, ctx.fields.clusterPlaceholder, CStr(absoluteClusterCount), , , vbTextCompare)
                    v = replace(v, ctx.fields.subclusterPlaceholder, CStr(relativeClusterCount), , , vbTextCompare)
                    v = replace(v, ctx.fields.clusterLevelPlaceholder, CStr(levelNumber), , , vbTextCompare)

                    If ctx.loop.Enabled Then
                        v = replace(v, ctx.fields.enumeratePlaceholder, SafeStr(enumStep), , , vbTextCompare)
                    End If
                End If
                
                Select Case LCase$(fld.name)
                    Case ctx.headings.flag
                        .Cells(row, ctx.dataLayout.flagColumn).value = v
                    
                    Case ctx.headings.item
                        .Cells(row, ctx.dataLayout.itemColumn).value = v
                    
                    Case ctx.headings.label, ctx.headings.xLabel
                        targetCol = IIf(LCase$(fld.name) = ctx.headings.label, _
                                        ctx.dataLayout.labelColumn, _
                                        ctx.dataLayout.xLabelColumn)
                                        
                        ' Apply multiline splitting only when requested & meaningful
                        Dim splitLength As Long
                        splitLength = GetSplitLength(rs, ctx.fields.splitLength)
                        If splitLength > 0 Then
                            Dim lineEnding  As String
                            lineEnding = GetLineEnding(rs, ctx.fields.lineEnding, NEWLINE)
                            v = SplitMultilineText(v, splitLength, lineEnding)
                        End If
                        
                        .Cells(row, targetCol).value = v
                    
                    Case ctx.headings.tailLabel
                        .Cells(row, ctx.dataLayout.tailLabelColumn).value = v
                    
                    Case ctx.headings.headLabel
                        .Cells(row, ctx.dataLayout.headLabelColumn).value = v
                    
                    Case ctx.headings.Tooltip
                        .Cells(row, ctx.dataLayout.tooltipColumn).value = v
                    
                    Case ctx.headings.isRelatedToItem
                        .Cells(row, ctx.dataLayout.isRelatedToItemColumn).value = v
                    
                    Case ctx.headings.styleName
                        .Cells(row, ctx.dataLayout.styleNameColumn).value = v
                    
                    Case ctx.headings.extraAttributes
                        .Cells(row, ctx.dataLayout.extraAttributesColumn).value = v
                    
                    Case ctx.headings.errorMessage
                        .Cells(row, ctx.dataLayout.errorMessageColumn).value = v
                    
                    ' Case Else: ignore unknown columns (intentional, general-purpose)
                End Select
            End If
        Next fld
    End With
End Sub

' ==========================================================================
' FUNCTION: IsClusterRelatedField
' PURPOSE:
'   Determines if a field is a "structural" clustering column to prevent
'   it from being written as standard node/edge data.
'
' TECHNICAL WORKFLOW:
'   1. NORMALIZATION: Converts the 'fieldName' to lowercase for
'      case-insensitive matching.
'   2. PATTERN SCAN: Iterates from Level 1 up to 'maxLevels' to build
'      expected metadata names:
'      - Base Cluster ID: "cluster1", "cluster2"...
'      - Visual Attributes: "cluster1 label", "cluster1 style name",
'        "cluster1 attributes", "cluster1 tooltip".
'   3. LOGIC GATE: Returns True if the field matches any of these structural
'      patterns, signaling the emission engine to skip this column.
'
' USAGE:
'   - Called by 'EmitOneRow' (Multi-Level) to filter the Recordset fields.
'   - Keeps the 'Data' worksheet clean by isolating "How to group" data
'     from "What to show" data.
' ==========================================================================
Private Function IsClusterRelatedField(ByVal fieldName As String, ByVal maxLevels As Long) As Boolean
    Dim i As Long
    Dim lcName As String
    lcName = LCase$(fieldName)
    
    Dim prefix As String
    For i = 1 To maxLevels
        prefix = LCase$("CLUSTER" & i)
        If lcName = prefix Or _
           lcName = prefix & " label" Or _
           lcName = prefix & " style name" Or _
           lcName = prefix & " attributes" Or _
           lcName = prefix & " tooltip" Then
            IsClusterRelatedField = True
            Exit Function
        End If
    Next i
End Function

