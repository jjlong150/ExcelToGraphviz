Attribute VB_Name = "modCreateJson"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modCreateJson
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Logic / Knowledge Graph Synthesis Pipeline
'
' ROLE:
'   The Knowledge Graph orchestration engine. Transforms structured worksheet
'   data into a fully normalized, style-aware, inheritance-driven JSON document
'   suitable for the Relationship Visualizer's interactive viewer. Coordinates
'   the complete JSON-emission lifecycle, including context initialization,
'   row processing, semantic-property merging, cluster-scope management, and
'   final serialization.
'
' RESPONSIBILITIES:
'   o Knowledge Graph Construction:
'       - Initialize and finalize the KnowledgeGraphContext.
'       - Maintain node dictionaries, edge collections, cluster scopes, and
'         semantic-property layers.
'       - Track viewStyles and usedStyles for style-aware JSON emission.
'
'   o Row Processing Pipeline:
'       - Parse worksheet rows into dataRow UDTs (GetDataRow).
'       - Normalize style names and resolve style metadata (PrepareDataRow).
'       - Apply label inheritance from graph/node/edge keyword scopes.
'       - Dispatch row-type logic for nodes, edges, clusters, and keywords.
'
'   o Semantic Property & Inheritance Model:
'       - Merge graph-level, node-level, edge-level, and row-level properties
'         following DeepWiki precedence rules (graph < node < edge < row).
'       - Maintain DotScope stack for cluster-level inheritance of labels,
'         tooltips, and keyword metadata.
'
'   o JSON Tree Assembly:
'       - Build metadata, styles, graph, nodes, and edges sections.
'       - Normalize HTML-like labels and construct label dictionaries.
'       - Emit only styles actually referenced in the selected view.
'
'   o JSON Serialization:
'       - Convert the assembled dictionary tree into compact or indented JSON.
'       - Ensure byte-consistent output across platforms via EscapeNonAscii.
'
' INTERACTIONS:
'   o modCreateCommon:
'       - Shared parsing, normalization, style caching, label overrides,
'         placeholder expansion, HTML-label detection, and filesystem helpers.
'
'   o modDataTypes:
'       - settings, dataRow, style, DotScope, and worksheet-contract UDTs.
'
'   o Utility Modules:
'       - modUtilityString: normalization, token substitution, affix stripping.
'       - modUtilityFileSystem: temp directories, environment variables.
'
'   o ViewerAssets:
'       - Consumes JSON output for interactive rendering.
'
' CROSS-PLATFORM NOTES:
'   o JSON emission is platform-neutral; all filesystem interactions are routed
'     through shared helpers that normalize path separators and sandbox behavior.
'   o EscapeNonAscii ensures identical byte output on Windows and macOS.
'
' ERROR HANDLING:
'   o Defensive checks for cluster-stack integrity.
'   o Graceful fallback when styles or metadata are missing.
'   o No external binary execution; all errors remain within the VBA layer.
'
' RELATED WIKI PAGES:
'   o Knowledge Graph Pipeline Architecture
'   o Semantic Properties & Inheritance
'   o Clusters & DotScope
'   o JSON Output Format & Serialization Rules
'   o Styles & View Definitions
' =============================================================================
Option Explicit

' ==========================================================================
' TYPE: KnowledgeGraphContext
'
' PURPOSE:
'   Serves as the central working context for Knowledge Graph synthesis.
'   Maintains all node/edge dictionaries, cluster scopes, semantic-property
'   layers, and style metadata required during parsing, inheritance, and
'   label/attribute resolution. Acts as the shared state passed through all
'   processing routines.
'
' FIELDS:
'   nodes As Dictionary
'       - Canonical dictionary of all nodes keyed by uppercase node ID.
'       - Each entry is a Dictionary containing node attributes, labels,
'         semantic properties, and style metadata.
'
'   edges As Collection
'       - Ordered list of synthesized edges.
'       - Each entry is a Dictionary containing edge attributes, labels,
'         semantic properties, and style metadata.
'
'   clusters As Stack
'       - Stack of DotScope objects representing nested cluster scopes.
'       - Used for inheritance of graph/node/edge keyword metadata and
'         cluster-level label/tooltip overrides.
'
'   clusterCount As Long
'       - Running count of clusters created, used for generating unique
'         cluster identifiers when needed.
'
'   properties As Dictionary
'       - Reserved for future expansion; currently unused in the pipeline.
'
'   viewStyles As Dictionary
'       - All styles associated with the selected view.
'       - Populated during initialization and used by PrepareDataRow to
'         resolve styleType and styleFormat.
'
'   usedStyles As Dictionary
'       - Tracks styles actually referenced by nodes/edges during synthesis.
'       - Enables style pruning and view-specific optimization.
'
'   propertiesGraph As Dictionary
'       - Semantic properties defined at the graph keyword level.
'       - Serves as the base layer for graph/node/edge property merging.
'
'   propertiesNode As Dictionary
'       - Semantic properties defined at the node keyword level.
'       - Merged with graph-level properties and row-level properties during
'         node creation/update.
'
'   propertiesEdge As Dictionary
'       - Semantic properties defined at the edge keyword level.
'       - Merged with graph-level properties and row-level properties during
'         edge synthesis.
'
' TECHNICAL NOTES:
'   - The context is mutated throughout the pipeline and passed by reference
'     to all processing routines.
'   - Cluster scopes (DotScope) provide inheritance surfaces for labels,
'     tooltips, and keyword metadata.
'   - Semantic-property precedence follows DeepWiki rules:
'         graph < node < edge < row.
'   - DeepWiki Context: Implements the shared-state architecture described in
'     the "Pipeline Architecture", "Clusters", and "Semantic Properties"
'     documentation.
' ==========================================================================
Private Type KnowledgeGraphContext
    nodes As Dictionary
    edges As Collection
    
    clusters As Stack
    clusterCount As Long
    
    properties As Dictionary
    
    viewStyles As Dictionary    ' All styles associated to a chosen view
    usedStyles As Dictionary    ' Styles used in the graph
    
    propertiesGraph As Dictionary
    propertiesNode As Dictionary
    propertiesEdge As Dictionary
End Type

' ==========================================================================
' ROUTINE: GetGraphJson
'
' PURPOSE:
'   Synthesizes the complete Knowledge Graph JSON document for a selected
'   view. Initializes the working context, processes all input rows, assembles
'   metadata, styles, graph-level attributes, nodes, and edges into a unified
'   dictionary tree, and serializes the result to JSON.
'
' FUNCTIONAL WORKFLOW:
'   1. CONTEXT INITIALIZATION:
'        - Creates and initializes a KnowledgeGraphContext via ContextInitialize.
'        - Injects viewStyles and computes usedStyles for the selected view.
'        - Enables tooltip support for Knowledge Graph emission.
'
'   2. ROW PROCESSING:
'        - Invokes ProcessAllRows to parse, normalize, inherit, override,
'          and synthesize all nodes, edges, clusters, and metadata.
'
'   3. JSON TREE CONSTRUCTION:
'        - Creates the root dictionary (body) that will become the JSON tree.
'
'        o METADATA SECTION:
'             - Populates global metadata (title, view name, etc.) via
'               ProcessMetadata.
'
'        o STYLES SECTION:
'             - Adds ctx.usedStyles when at least one style is referenced.
'
'        o GRAPH SECTION:
'             - Retrieves the active DotScope from the cluster stack.
'             - Adds graph-level label and tooltip (if present).
'             - Adds debuglabel when debugging is enabled.
'
'        o NODES SECTION:
'             - Converts ctx.nodes (Dictionary) into a Collection of node
'               dictionaries and adds it under "nodes".
'
'        o EDGES SECTION:
'             - Adds ctx.edges directly as the "edges" array.
'
'   4. JSON SERIALIZATION:
'        - Converts the assembled dictionary tree to a JSON string.
'        - Uses compact formatting when indent < 0; otherwise applies the
'          caller-specified whitespace indentation.
'
'   5. CLEANUP:
'        - Finalizes the context via ContextFinalize.
'        - Releases temporary objects.
'
' TECHNICAL NOTES:
'   - The cluster stack must contain at least one scope at the end of row
'     processing; a defensive check prevents emitting graph metadata when
'     the stack has been over-popped.
'   - Label dictionaries are created using CreateLabelDictionary to ensure
'     consistent formatting across nodes, edges, and graph-level metadata.
'   - DeepWiki Context: Implements the JSON-emission rules described in the
'     "Output Format", "Metadata", and "Serialization" documentation.
' ==========================================================================
Public Function GetGraphJson(ByRef ini As settings, _
                                  ByVal viewName As String, _
                                  ByRef viewStyles As Dictionary, _
                                  ByVal indent As Long _
                                  ) As String
    
    ' Establish a context for passing values around
    Dim ctx As KnowledgeGraphContext
    ContextInitialize ctx
    
    ' Get the information on styles in use
    Set ctx.viewStyles = viewStyles
    Set ctx.usedStyles = GetStyles(ini, viewStyles)
    
    ' Process all the rows
    ProcessAllRows ini, ctx
    
    ' Construct the tree
    Dim body As Dictionary
    Set body = New Dictionary
    
    ' Metadata section
    ProcessMetadata ini, viewName, body
    
    ' styles{} section
    If ctx.usedStyles.Count > 0 Then
        body.Add "styles", ctx.usedStyles
    End If
    
    ' graph{} section
    Dim graphDict As Dictionary
    Set graphDict = New Dictionary
    
    Dim scope As DotScope
    Set scope = ctx.clusters.Peek
    
    ' Defensive check protecting against too many stack pops
    If Not scope Is Nothing Then
        ' Graph label
        If Len(scope.graphLabel) > 0 Then
            Dim graphLabelDict As Dictionary
            Set graphLabelDict = CreateLabelDictionary(scope.graphLabel)
            graphDict.Add "label", graphLabelDict
        End If
        
        ' Debug Label
        If ini.graph.debug And scope.graphRow > 0 Then
            Dim graphDebugDict As Dictionary
            Set graphDebugDict = CreateLabelDictionary("Row: " & scope.graphRow)
            graphDict.Add "debuglabel", graphDebugDict
        End If
    End If
    
    If graphDict.Count > 0 Then
        body.Add "graph", graphDict
    End If
    
    ' nodes[] section
    If ctx.nodes.Count > 0 Then
        Dim nodeArray As Collection
        Set nodeArray = DictionaryValuesToCollection(ctx.nodes)
        body.Add "nodes", nodeArray
    End If
    
    ' edges[] section
    If ctx.edges.Count > 0 Then
        body.Add "edges", ctx.edges
    End If
    
    ' Convert the tree to a JSON string
    If indent < 0 Then
        GetGraphJson = ConvertToJson(body)
    Else
        GetGraphJson = ConvertToJson(body, whitespace:=indent)
    End If
    
    ' Release all allocated objects
    ContextFinalize ctx
    Set body = Nothing
End Function

' ==========================================================================
' ROUTINE: CreateLabelDictionary
'
' PURPOSE:
'   Constructs a standardized label dictionary containing the resolved label
'   value and its detected type ("text" or "html"). Removes Graphviz-style
'   HTML-like affixes when present, ensuring that only the true label content
'   is serialized into the JSON output.
'
' FUNCTIONAL WORKFLOW:
'   1. LABEL-TYPE DETECTION:
'        - Determines whether the label should be treated as plain text or
'          HTML-like using GetLabelType.
'
'   2. HTML-LIKE STRIPPING:
'        - When IsLabelHTMLLike returns True:
'             o Removes the outer "<" and ">" affixes using StripAffix.
'             o Preserves the inner content exactly as written.
'        - Otherwise:
'             o Uses the label value as-is.
'
'   3. DICTIONARY CONSTRUCTION:
'        - Creates a new Dictionary containing:
'             o "value" -> the cleaned label text.
'             o "type"  -> the detected label type.
'
'   4. OUTPUT:
'        - Returns the constructed dictionary for inclusion in node, edge,
'          or graph-level JSON sections.
'
' TECHNICAL NOTES:
'   - HTML-like detection is intentionally conservative: only labels wrapped
'     in literal "<...>" are treated as HTML-like.
'   - This routine does not perform normalization; callers may apply
'     NormalizeLabel or ApplyLabelOverrides beforehand.
'   - DeepWiki Context: Implements the label-dictionary rules described in
'     the "Labels", "Formatting", and "Serialization" documentation.
' ==========================================================================
Function CreateLabelDictionary(ByVal label As String) As Dictionary
    ' text or html?
    Dim lblType As String
    lblType = GetLabelType(label)

    ' Remove graphviz html-like label signals as they are not part of the actual label
    Dim lblValue As String
    If IsLabelHTMLLike(label) Then
        lblValue = StripAffix(label, "<", ">")
    Else
        lblValue = label
    End If
        
    Dim lblDict As Dictionary
    Set lblDict = New Dictionary
    
    lblDict.Add "value", lblValue
    lblDict.Add "type", lblType
    
    Set CreateLabelDictionary = lblDict
End Function

' ==========================================================================
' ROUTINE: ContextInitialize
'
' PURPOSE:
'   Initializes a new KnowledgeGraphContext with all required dictionaries,
'   collections, and cluster scaffolding. Establishes the root DotScope and
'   prepares the context for node/edge synthesis, label inheritance, keyword
'   processing, and JSON emission.
'
' FUNCTIONAL WORKFLOW:
'   1. CORE COLLECTION INITIALIZATION:
'        - Creates empty containers for:
'             o ctx.nodes        (Dictionary of node dictionaries)
'             o ctx.edges        (Collection of edge dictionaries)
'
'   2. CLUSTER STACK SETUP:
'        - Creates a new Stack for cluster scopes.
'        - Instantiates a root DotScope representing the top-level graph.
'        - Pushes the root scope onto ctx.clusters.
'        - Ensures all subsequent cluster operations have a valid base scope.
'
'   3. STYLE METADATA INITIALIZATION:
'        - Creates empty dictionaries for:
'             o ctx.viewStyles   (styles available for the selected view)
'             o ctx.usedStyles   (styles actually referenced during synthesis)
'
'   4. SEMANTIC-PROPERTY LAYERS:
'        - Initializes property dictionaries for:
'             o ctx.properties        (reserved/general)
'             o ctx.propertiesGraph   (graph-level keyword properties)
'             o ctx.propertiesNode    (node-level keyword properties)
'             o ctx.propertiesEdge    (edge-level keyword properties)
'        - These dictionaries form the base layers for DeepWiki's property
'          precedence model: graph < node < edge < row.
'
' TECHNICAL NOTES:
'   - ContextInitialize must be called before any row-processing routines.
'   - The root DotScope ensures that label inheritance and keyword metadata
'     always have a valid scope, even when no clusters are explicitly opened.
'   - All objects are created fresh; ContextFinalize is responsible for cleanup.
'   - DeepWiki Context: Implements the initialization rules described in the
'     "Pipeline Architecture", "Clusters", and "Semantic Properties" sections.
' ==========================================================================
Private Sub ContextInitialize(ctx As KnowledgeGraphContext)
    Set ctx.nodes = New Dictionary
    Set ctx.edges = New Collection
    
    Set ctx.clusters = New Stack
    Dim scope As DotScope
    Set scope = New DotScope    ' Root graph
    ctx.clusters.Push scope
    
    Set ctx.usedStyles = New Dictionary
    Set ctx.viewStyles = New Dictionary
    
    Set ctx.properties = New Dictionary
    Set ctx.propertiesGraph = New Dictionary
    Set ctx.propertiesNode = New Dictionary
    Set ctx.propertiesEdge = New Dictionary
End Sub

' ==========================================================================
' ROUTINE: ContextFinalize
'
' PURPOSE:
'   Releases all objects associated with a KnowledgeGraphContext. Performs a
'   deterministic teardown of node/edge collections, cluster scopes, style
'   dictionaries, and semantic-property layers. Ensures that no lingering
'   references remain after JSON emission is complete.
'
' FUNCTIONAL WORKFLOW:
'   1. CORE COLLECTION RELEASE:
'        - Clears references to:
'             o ctx.nodes        (Dictionary of node dictionaries)
'             o ctx.edges        (Collection of edge dictionaries)
'
'   2. CLUSTER STACK RELEASE:
'        - Clears the cluster stack (ctx.clusters), including the root
'          DotScope created during ContextInitialize.
'
'   3. STYLE METADATA RELEASE:
'        - Clears references to:
'             o ctx.usedStyles
'             o ctx.viewStyles
'          ensuring no style dictionaries persist across graph builds.
'
'   4. SEMANTIC-PROPERTY LAYER RELEASE:
'        - Clears references to:
'             o ctx.properties
'             o ctx.propertiesGraph
'             o ctx.propertiesNode
'             o ctx.propertiesEdge
'          removing all keyword-level property dictionaries.
'
' TECHNICAL NOTES:
'   - ContextFinalize must be called exactly once per context lifecycle.
'   - This routine does not attempt to clear nested dictionaries or collections;
'     VB6 reference semantics ensure that setting the parent object to Nothing
'     releases all child objects.
'   - Complements ContextInitialize, forming a clean initialization/finalization
'     pair for the Knowledge Graph pipeline.
'   - DeepWiki Context: Implements the teardown rules described in the
'     "Pipeline Architecture", "Memory Management", and "Lifecycle" sections.
' ==========================================================================
Private Sub ContextFinalize(ctx As KnowledgeGraphContext)
    Set ctx.nodes = Nothing
    Set ctx.edges = Nothing
    
    Set ctx.clusters = Nothing
    
    Set ctx.usedStyles = Nothing
    Set ctx.viewStyles = Nothing
    
    Set ctx.properties = Nothing
    Set ctx.propertiesGraph = Nothing
    Set ctx.propertiesNode = Nothing
    Set ctx.propertiesEdge = Nothing
End Sub

' ==========================================================================
' ROUTINE: GetStyles
'
' PURPOSE:
'   Produces a filtered dictionary of styles actually used in the graph for
'   the selected view. Reduces JSON output size by excluding unused styles
'   and limits emission to node-, edge-, and cluster-open styles. Normalizes
'   cluster style names and constructs JSON-ready dictionaries for each style.
'
' FUNCTIONAL WORKFLOW:
'   1. DETERMINE STYLES IN USE:
'        - Invokes DetermineStylesInUse to compute the subset of viewStyles
'          referenced by nodes, edges, or cluster scopes.
'        - Reduces downstream JSON size and AI token usage.
'
'   2. ITERATE OVER USED STYLES:
'        - For each style key in stylesInUse:
'             o Retrieves the corresponding style object from viewStyles.
'
'   3. STYLE-TYPE FILTERING:
'        - Only emits styles whose styleType is one of:
'             o "node"
'             o "edge"
'             o "subgraph-open"
'        - All other style types are skipped.
'
'   4. CLUSTER-STYLE NORMALIZATION:
'        - When styleType = "subgraph-open":
'             o Converts type to "cluster".
'             o Strips the configured suffix (ini.styles.affixOpen) from
'               the style name using StripSuffix.
'
'   5. DICTIONARY CONSTRUCTION:
'        - Builds a JSON-ready dictionary containing:
'             o JSON_STYLES_DESCRIPTION -> trimmed style description (if present)
'             o JSON_STYLES_TYPE        -> normalized style type
'        - Adds the dictionary under the normalized style name only when at
'          least one field is populated.
'
'   6. OUTPUT:
'        - Returns a dictionary mapping style names -> style dictionaries.
'
' TECHNICAL NOTES:
'   - Style dictionaries are intentionally minimal: only description and type
'     are emitted to reduce JSON size.
'   - Name normalization for cluster styles ensures consistent naming across
'     inheritance, overrides, and JSON emission.
'   - DeepWiki Context: Implements the style-emission rules described in the
'     "Styles", "View Definitions", and "Serialization" documentation.
' ==========================================================================
Private Function GetStyles(ByRef ini As settings, viewStyles As Dictionary) As Dictionary

    ' Only return information on styles actually being used (helps reduce AI token use).
    Dim stylesInUse As Dictionary
    Set stylesInUse = DetermineStylesInUse(ini, viewStyles)

    Dim items As New Dictionary
    Dim dictKey As Variant
    Dim styleObj As Object
    Dim dictObj As Dictionary

    For Each dictKey In stylesInUse.keys
        Set styleObj = viewStyles(dictKey)

        ' Filter: only include node, edge, subgraph-open
        Select Case LCase$(Trim$(styleObj.styleType))
            Case "node", "edge", "subgraph-open"
                ' Allowed -> continue
            Case Else
                GoTo NextStyle
        End Select

        ' Build dictionary only if at least one field is populated
        Set dictObj = New Dictionary

        With styleObj
            Dim name As String: name = Trim$(.styleName)
            Dim desc As String: desc = Trim$(.styleDescription)
            Dim typ As String:  typ = Trim$(.styleType)

            If typ = "subgraph-open" Then
                typ = "cluster"
                name = Trim$(StripSuffix(name, ini.styles.affixOpen))
            End If
            If desc <> vbNullString Then dictObj.Add JSON_STYLES_DESCRIPTION, desc
            If typ <> vbNullString Then dictObj.Add JSON_STYLES_TYPE, typ
        End With

        If dictObj.Count > 0 Then
            items.Add name, dictObj
        End If

NextStyle:
    Next dictKey

    Set GetStyles = items
End Function

' ==========================================================================
' ROUTINE: ProcessMetadata
'
' PURPOSE:
'   Populates the metadata section of the Knowledge Graph JSON document.
'   Emits format identifiers, versioning, graph directionality, workbook
'   attribution, view name, and a local export timestamp. Establishes the
'   top-level metadata fields consumed by downstream renderers and tools.
'
' FUNCTIONAL WORKFLOW:
'   1. FORMAT IDENTIFICATION:
'        - Adds the canonical Knowledge Graph Format identifier:
'             o "format" -> "RV-KGF"
'        - Adds the schema version:
'             o "version" -> "1.0"
'
'   2. GRAPH DIRECTIONALITY:
'        - Emits a Boolean "directed" flag based on ini.graph.graphType.
'        - Ensures downstream consumers know whether edges should be treated
'          as directed or undirected.
'
'   3. WORKBOOK ATTRIBUTION:
'        - Records the name of the workbook that generated the export:
'             o "source_workbook" -> ThisWorkbook.Name
'        - Provides traceability for multi-workbook environments.
'
'   4. VIEW IDENTIFICATION:
'        - Emits the selected view name under:
'             o "view" -> viewName
'        - Enables multi-view Knowledge Graphs to be distinguished cleanly.
'
'   5. EXPORT TIMESTAMP:
'        - Emits a local timestamp (no timezone offset) using:
'             o "export_datetime" -> yyyy-mm-ddThh:nn:ss
'        - Timestamp reflects the machine-local time of export.
'
' TECHNICAL NOTES:
'   - Metadata is always emitted, regardless of node/edge count.
'   - Timestamp intentionally omits timezone offset to preserve compatibility
'     with legacy consumers.
'   - DeepWiki Context: Implements the metadata rules described in the
'     "Output Format", "Metadata", and "Serialization" documentation.
' ==========================================================================
Private Sub ProcessMetadata(ini As settings, viewName As String, ByRef body As Dictionary)
    body.Add "format", "RV-KGF" ' Relationship Visualizer - Knowledge Graph Format
    body.Add "version", "1.0"
    body.Add "directed", IIf(ini.graph.graphType = "directed", True, False)
    body.Add "source_workbook", ThisWorkbook.name
    body.Add "view", viewName
    'export_datetime is local time on the machine that generated the export; no timezone offset is included
    body.Add "export_datetime", Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
End Sub

' ==========================================================================
' ROUTINE: DetermineStylesInUse
'
' PURPOSE:
'   Scans the data worksheet to identify which styles are actually referenced
'   by rows in the selected view. Produces a dictionary of unique style names
'   that appear in non-comment rows and exist in the viewStyles dictionary.
'   Used to reduce JSON output size and eliminate unused style definitions.
'
' FUNCTIONAL WORKFLOW:
'   1. WORKSHEET RESOLUTION:
'        - Retrieves the worksheet specified by ini.data.worksheetName.
'        - Ensures all row reads occur against a cached Worksheet reference.
'
'   2. ROW ITERATION:
'        - Iterates from ini.data.firstRow to ini.data.lastRow.
'        - Skips comment rows based on FLAG_COMMENT in the flagColumn.
'
'   3. STYLE-NAME NORMALIZATION:
'        - Reads the style name from styleNameColumn.
'        - When empty, computes a default style name via DetermineStyleName.
'        - Trims and uppercases the style name for dictionary lookup.
'
'   4. STYLE VALIDATION:
'        - Skips rows with empty style names.
'        - Skips rows whose style names do not exist in viewStyles.
'
'   5. UNIQUE-STYLE COLLECTION:
'        - Adds each validated style name to styleIds (Dictionary).
'        - Ensures each style appears only once, regardless of row count.
'
'   6. OUTPUT:
'        - Returns styleIds, a dictionary of unique style names actually used
'          in the worksheet for the selected view.
'
' TECHNICAL NOTES:
'   - This routine performs no style-type filtering; that occurs later in
'     GetStyles.
'   - Style-name normalization ensures consistent behavior across worksheets
'     with mixed or missing style-name entries.
'   - DeepWiki Context: Implements the style-usage detection rules described
'     in the "Styles", "View Definitions", and "Row Normalization" sections.
' ==========================================================================
Private Function DetermineStylesInUse(ByRef ini As settings, _
                                      ByVal viewStyles As Dictionary) As Dictionary
    Dim row As Long
    Dim styleName As String
    Dim ws As Worksheet

    ' Cache worksheet reference
    Set ws = worksheets(ini.data.worksheetName)

    Dim styleIds As New Dictionary

    For row = ini.data.firstRow To ini.data.lastRow

        ' Skip comment rows
        If ws.Cells(row, ini.data.flagColumn).value = FLAG_COMMENT Then GoTo NextRow

        ' Normalize style name
        styleName = Trim$(ws.Cells(row, ini.data.styleNameColumn).value)

        If styleName = vbNullString Then
            styleName = DetermineStyleName(ini, row)
        End If

        styleName = UCase$(Trim$(styleName))

        ' Skip empty or unknown styles
        If styleName = vbNullString Then GoTo NextRow
        If Not viewStyles.Exists(styleName) Then GoTo NextRow

        ' Add to dictionary of unique styles
        If Not styleIds.Exists(styleName) Then
            styleIds.Add styleName, True
        End If

NextRow:
    Next row

    Set DetermineStylesInUse = styleIds
End Function

' ==========================================================================
' ROUTINE: ProcessAllRows
'
' PURPOSE:
'   Iterates through all data rows in the worksheet and performs the complete
'   Knowledge Graph synthesis pipeline. Normalizes each row, applies style
'   metadata, performs label inheritance, and dispatches row-type logic for
'   nodes, edges, clusters, and keyword metadata. Forms the backbone of the
'   graph-building process.
'
' FUNCTIONAL WORKFLOW:
'   1. INITIALIZATION:
'        - Resets ctx.clusterCount to zero.
'        - Ensures cluster counting begins fresh for this synthesis pass.
'
'   2. ROW ITERATION:
'        - Loops from ini.data.firstRow to ini.data.lastRow.
'        - Retrieves each row's parsed data via GetDataRow.
'
'   3. ROW SKIP LOGIC:
'        - Skips rows that are blank, commented, or otherwise excluded by
'          RowShouldBeSkipped.
'
'   4. ROW PREPARATION:
'        - Normalizes style name and resolves style metadata via PrepareDataRow.
'        - Assigns styleType and optional styleFormat.
'
'   5. LABEL INHERITANCE:
'        - Applies graph/node/edge label inheritance via ApplyLabelInheritance.
'        - Ensures missing labels are filled from cluster-scope defaults
'          before override resolution.
'
'   6. ROW DISPATCH:
'        - Executes row-type logic via DispatchRow:
'             o Node synthesis
'             o Edge synthesis
'             o Cluster open/close
'             o Keyword metadata processing
'        - Updates ctx.nodes, ctx.edges, ctx.clusters, and semantic properties.
'
'   7. LOOP CONTINUATION:
'        - Advances to the next row until the worksheet range is exhausted.
'
' TECHNICAL NOTES:
'   - ProcessAllRows is the central driver of the Knowledge Graph pipeline.
'   - All major subsystems-style resolution, inheritance, overrides, keyword
'     processing, and node/edge synthesis-are invoked here.
'   - Cluster count increments occur inside DispatchRow when cluster-open
'     rows are encountered.
'   - DeepWiki Context: Implements the row-processing rules described in the
'     "Pipeline Architecture", "Row Types", and "Cluster Semantics" sections.
' ==========================================================================
Private Sub ProcessAllRows(ByRef ini As settings, _
                               ByRef ctx As KnowledgeGraphContext)

    ' Count each cluster as encountered.
    ctx.clusterCount = 0
    
    Dim row As Long
    Dim data As dataRow

    For row = ini.data.firstRow To ini.data.lastRow
        data = GetDataRow(ini, ini.data.worksheetName, row)

        ' Skip rows that are blank or commented out
        If RowShouldBeSkipped(data) Then GoTo NextRow

        ' Prepare the row
        data = PrepareDataRow(ini, ctx.viewStyles, data, row)
        
        ' Fill missing labels from node/edge/graph defaults first
        ApplyLabelInheritance ini, ctx, data

        ' Perform the row-type logic
        DispatchRow ini, ctx, data

NextRow:
    Next row

End Sub

' ==========================================================================
' ROUTINE: RowShouldBeSkipped
'
' PURPOSE:
'   Determines whether a parsed data-row should be excluded from the Knowledge
'   Graph synthesis pipeline. Filters out comment rows and structurally blank
'   rows before any style resolution, inheritance, or dispatch logic occurs.
'
' FUNCTIONAL WORKFLOW:
'   1. COMMENT ROW DETECTION:
'        - If data.comment equals FLAG_COMMENT:
'             o Marks the row as skipped.
'             o Prevents processing of user-commented or disabled rows.
'
'   2. BLANK-ITEM FILTER:
'        - Trims data.item and checks for an empty value.
'        - Skips rows that contain no actionable item keyword (node, edge,
'          graph, cluster-open, etc.).
'
'   3. OUTPUT:
'        - Returns True when the row should be skipped.
'        - Returns False otherwise.
'
' TECHNICAL NOTES:
'   - This routine performs minimal validation; deeper semantic checks occur
'     later in PrepareDataRow and DispatchRow.
'   - Early skipping improves performance and reduces unnecessary style and
'     inheritance operations.
'   - DeepWiki Context: Implements the row-filtering rules described in the
'     "Row Normalization", "Comments", and "Pipeline Architecture" sections.
' ==========================================================================
Private Function RowShouldBeSkipped(data As dataRow) As Boolean
    If data.comment = FLAG_COMMENT Then
        RowShouldBeSkipped = True
        Exit Function
    End If
    
    If Len(Trim$(data.item)) = 0 Then
        RowShouldBeSkipped = True
        Exit Function
    End If
End Function

' ==========================================================================
' ROUTINE: PrepareDataRow
'
' PURPOSE:
'   Normalizes and enriches a dataRow record by resolving its style name,
'   applying cached style metadata, and enforcing keyword overrides for
'   Graphviz-reserved items ("node", "edge", "graph"). Produces a fully
'   prepared dataRow ready for downstream formatting and graph generation.
'
' FUNCTIONAL WORKFLOW:
'   1. STYLE NAME NORMALIZATION:
'        - If the incoming dataRow has no styleName, DetermineStyleName is
'          invoked to derive one from worksheet content.
'        - The resolved name is uppercased to form a dictionary key.
'
'   2. EMPTY-STYLE SHORT-CIRCUIT:
'        - If the normalized style key is empty, the row cannot participate
'          in style lookup; styleType is set to 0 and the function exits.
'
'   3. STYLE CACHE APPLICATION:
'        - Checks the enabled-style Dictionary for the normalized key.
'        - When present:
'             o Copies styleType from the cached style object.
'             o Copies styleFormat only when ini.graph.includeStyleFormat=True.
'               Otherwise assigns vbNullString.
'        - Uses a single dictionary lookup for performance.
'
'   4. KEYWORD OVERRIDE:
'        - Graphviz keywords ("node", "edge", "graph") always force
'          data.styleType = TYPE_KEYWORD, regardless of any cached style.
'        - Implemented via a compact Select Case on the lowercase item value.
'
'   5. OUTPUT:
'        - Returns the fully prepared dataRow with normalized styleName,
'          resolved styleType, optional styleFormat, and keyword enforcement.
'
' TECHNICAL NOTES:
'   - Strongly typed styleObj improves clarity and avoids late binding.
'   - Centralizes all row-normalization logic so the main generation loop
'     operates on fully prepared records.
'   - DeepWiki Context: Implements the row-preparation rules described in the
'     "Styles", "Generation Pipeline", and "Normalization" sections.
' ==========================================================================
Private Function PrepareDataRow(ByRef ini As settings, _
                                ByRef styles As Dictionary, _
                                ByRef data As dataRow, _
                                ByVal row As Long) As dataRow

    ' Normalize style name if missing
    If Len(data.styleName) = 0 Then
        data.styleName = DetermineStyleName(ini, row)
    End If

    Dim styleKey As String
    styleKey = UCase$(data.styleName)

    ' No style name ? no style type
    If Len(styleKey) = 0 Then
        data.styleType = 0
        PrepareDataRow = data
        Exit Function
    End If

    ' Apply cached style if present
    If styles.Exists(styleKey) Then
        Dim styleObj As style
        Set styleObj = styles(styleKey)

        data.styleType = styleObj.styleType

        If ini.graph.includeStyleFormat Then
            data.Format = styleObj.styleFormat
        Else
            data.Format = vbNullString
        End If
    End If

    ' Keyword override
    Select Case LCase$(data.item)
        Case "node", "edge", "graph"
            data.styleType = TYPE_KEYWORD
    End Select

    PrepareDataRow = data
End Function

' ==========================================================================
' ROUTINE: ApplyLabelInheritance
'
' PURPOSE:
'   Applies label inheritance from the active cluster scope to the current
'   data-row. Delegates to the appropriate inheritance routine based on the
'   row's styleType, ensuring that graph-, node-, and edge-level metadata
'   propagate correctly when row-level values are absent.
'
' FUNCTIONAL WORKFLOW:
'   1. CLUSTER-SCOPE VALIDATION:
'        - Ensures that a cluster scope is active before attempting inheritance.
'        - Emits a warning and exits when the cluster stack is empty, since
'          inheritance requires an active DotScope.
'
'   2. SCOPE RETRIEVAL:
'        - Retrieves the current DotScope from the top of the cluster stack.
'
'   3. ROW-TYPE DISPATCH:
'        - Delegates inheritance based on data.styleType:
'
'             o TYPE_SUBGRAPH_OPEN
'             o TYPE_GRAPH
'                   - Applies graph-level inheritance via ApplyLabelInheritanceGraph.
'
'             o TYPE_NODE
'                   - Applies node-level inheritance via ApplyLabelInheritanceNode.
'
'             o TYPE_EDGE
'                   - Applies edge-level inheritance via ApplyLabelInheritanceEdge.
'
'        - Each delegated routine performs fill-only inheritance, preserving
'          any row-level values already provided.
'
' TECHNICAL NOTES:
'   - Inheritance occurs *before* override resolution (ApplyLabelOverrides),
'     ensuring that templates and placeholders operate on the correct base
'     values.
'   - Graph/node/edge inheritance routines are intentionally separate to
'     maintain clarity and avoid cross-type contamination.
'   - DeepWiki Context: Implements the unified inheritance rules described in
'     the "Labels", "Clusters", and "Metadata Propagation" documentation.
' ==========================================================================
Private Sub ApplyLabelInheritance(ByRef ini As settings, _
                                  ByRef ctx As KnowledgeGraphContext, _
                                  ByRef data As dataRow)

    If ctx.clusters.Count <= 0 Then
        EmitMessageSilent GetMessage("errormsgALIClusterStackIsEmpty"), esWarning
        Exit Sub
    End If

    Dim scope As DotScope
    Set scope = ctx.clusters.Peek

    With data
        Select Case .styleType
            
            Case TYPE_SUBGRAPH_OPEN
                ApplyLabelInheritanceGraph scope, data
            
            Case TYPE_GRAPH
                ApplyLabelInheritanceGraph scope, data
            
            Case TYPE_NODE
                ApplyLabelInheritanceNode scope, data
                
            Case TYPE_EDGE
                ApplyLabelInheritanceEdge scope, data
        End Select
    End With

End Sub

' ==========================================================================
' ROUTINE: ApplyLabelInheritanceGraph
'
' PURPOSE:
'   Applies graph-level label inheritance from the active cluster scope to the
'   current data-row. Populates missing label fields (label/tooltip) using
'   scope-level defaults while preserving any values explicitly provided by
'   the row.
'
' FUNCTIONAL WORKFLOW:
'   1. INHERIT PRIMARY LABEL:
'        - When data.label is empty and scope.graphLabel is defined,
'          assigns the inherited graph label.
'
'   2. INHERIT TOOLTIP:
'        - When data.tooltip is empty and scope.graphTooltip is defined,
'          assigns the inherited graph tooltip.
'
' TECHNICAL NOTES:
'   - Inheritance is strictly *fill-only*: row-level values always override
'     scope-level defaults and are never replaced.
'   - Graph-level keyword metadata originates from ProcessKeywordGraph and is
'     stored on the DotScope for later use.
'   - Inheritance occurs before override resolution (ApplyLabelOverrides),
'     ensuring that templates and placeholders operate on the correct base
'     values.
'   - DeepWiki Context: Implements the graph-label inheritance rules described
'     in the "Labels", "Clusters", and "Graph Metadata" documentation.
' ==========================================================================
Private Sub ApplyLabelInheritanceGraph(ByRef scope As DotScope, ByRef data As dataRow)
    
    ' Label
    If Len(data.label) = 0 And Len(scope.graphLabel) > 0 Then
        data.label = scope.graphLabel
    End If
    
    ' Tooltip
    If Len(data.tooltip) = 0 And Len(scope.graphTooltip) > 0 Then
        data.tooltip = scope.graphTooltip
    End If

End Sub

' ==========================================================================
' ROUTINE: ApplyLabelInheritanceNode
'
' PURPOSE:
'   Applies node-level label inheritance from the active cluster scope to the
'   current data-row. Populates missing label fields (label/xlabel/tooltip)
'   using scope-level defaults while preserving any values explicitly provided
'   by the row.
'
' FUNCTIONAL WORKFLOW:
'   1. INHERIT PRIMARY LABEL:
'        - When data.label is empty and scope.nodeLabel is defined,
'          assigns the inherited node label.
'
'   2. INHERIT X-LABEL:
'        - When data.xlabel is empty and scope.nodeXLabel is defined,
'          assigns the inherited xlabel.
'
'   3. INHERIT TOOLTIP:
'        - When data.tooltip is empty and scope.nodeTooltip is defined,
'          assigns the inherited tooltip.
'
' TECHNICAL NOTES:
'   - Inheritance is strictly *fill-only*: row-level values always override
'     scope-level defaults and are never replaced.
'   - Node-level keyword metadata originates from ProcessKeywordNode and is
'     stored on the DotScope for later use.
'   - Inheritance occurs before override resolution (ApplyLabelOverrides),
'     ensuring that templates and placeholders operate on the correct base
'     values.
'   - DeepWiki Context: Implements the node-label inheritance rules described
'     in the "Labels", "Clusters", and "Node Metadata" documentation.
' ==========================================================================
Private Sub ApplyLabelInheritanceNode(ByRef scope As DotScope, ByRef data As dataRow)
    
    ' Label
    If Len(data.label) = 0 And Len(scope.nodeLabel) > 0 Then
        data.label = scope.nodeLabel
    End If
    
    ' XLabel
    If Len(data.xlabel) = 0 And Len(scope.nodeXLabel) > 0 Then
        data.xlabel = scope.nodeXLabel
    End If

    ' Tooltip
    If Len(data.tooltip) = 0 And Len(scope.nodeTooltip) > 0 Then
        data.tooltip = scope.nodeTooltip
    End If

End Sub

' ==========================================================================
' ROUTINE: ApplyLabelInheritanceEdge
'
' PURPOSE:
'   Applies edge-level label inheritance from the active cluster scope to the
'   current data-row. Populates missing label fields (label/xlabel/taillabel/
'   headlabel/tooltip) using scope-level defaults while preserving any values
'   explicitly provided by the row.
'
' FUNCTIONAL WORKFLOW:
'   1. INHERIT PRIMARY LABEL:
'        - When data.label is empty and scope.edgeLabel is defined,
'          assigns the scope-level label to the row.
'
'   2. INHERIT X-LABEL:
'        - When data.xlabel is empty and scope.edgeXLabel is defined,
'          assigns the inherited xlabel.
'
'   3. INHERIT TAIL LABEL:
'        - When data.taillabel is empty and scope.edgeTailLabel is defined,
'          assigns the inherited tail label.
'
'   4. INHERIT HEAD LABEL:
'        - When data.headlabel is empty and scope.edgeHeadLabel is defined,
'          assigns the inherited head label.
'
'   5. INHERIT TOOLTIP:
'        - When data.tooltip is empty and scope.edgeTooltip is defined,
'          assigns the inherited tooltip.
'
' TECHNICAL NOTES:
'   - Inheritance is strictly *fill-only*: row-level values always override
'     scope-level defaults and are never replaced.
'   - Edge-level keyword metadata originates from ProcessKeywordEdge and is
'     stored on the DotScope for later use.
'   - Inheritance occurs before override resolution (ApplyLabelOverrides),
'     ensuring that templates and placeholders operate on the correct base
'     values.
'   - DeepWiki Context: Implements the edge-label inheritance rules described
'     in the "Labels", "Clusters", and "Edge Metadata" documentation.
' ==========================================================================
Private Sub ApplyLabelInheritanceEdge(ByRef scope As DotScope, ByRef data As dataRow)
    
    ' Label
    If Len(data.label) = 0 And Len(scope.edgeLabel) > 0 Then
        data.label = scope.edgeLabel
    End If
    
    ' XLabel
    If Len(data.xlabel) = 0 And Len(scope.edgeXLabel) > 0 Then
        data.xlabel = scope.edgeXLabel
    End If
    
    ' Tail Label
    If Len(data.taillabel) = 0 And Len(scope.edgeTailLabel) > 0 Then
        data.taillabel = scope.edgeTailLabel
    End If
    
    ' Head Label
    If Len(data.headlabel) = 0 And Len(scope.edgeHeadLabel) > 0 Then
        data.headlabel = scope.edgeHeadLabel
    End If

    ' Tooltip
    If Len(data.tooltip) = 0 And Len(scope.edgeTooltip) > 0 Then
        data.tooltip = scope.edgeTooltip
    End If

End Sub

' ==========================================================================
' ROUTINE: ProcessKeywordNode
'
' PURPOSE:
'   Applies node-level keyword semantics to the current cluster scope.
'   Resolves all supported node-label variants (label/xlabel/tooltip)
'   using the full override pipeline and stores the results in the active
'   DotScope for inheritance by subsequent node rows.
'
' FUNCTIONAL WORKFLOW:
'   1. CLUSTER-SCOPE VALIDATION:
'        - Ensures that a cluster scope is active before applying node-level
'          keyword semantics.
'        - Emits a warning and exits when the cluster stack is empty, since
'          node-level metadata must attach to an active scope.
'
'   2. SCOPE RETRIEVAL:
'        - Retrieves the current DotScope from the top of the cluster stack.
'
'   3. LABEL OVERRIDE RESOLUTION:
'        - Invokes ApplyLabelOverrides for each supported node label type:
'             o "label"   -> scope.nodeLabel
'             o "xlabel"  -> scope.nodeXLabel
'             o "tooltip" -> scope.nodeTooltip
'        - Each override resolution incorporates:
'             o row-level values,
'             o style-format templates,
'             o placeholder expansion rules.
'
'   4. METADATA ASSIGNMENT:
'        - Stores the resolved label values directly on the DotScope, making
'          them available for inheritance by subsequent node rows.
'
' TECHNICAL NOTES:
'   - Node-level keyword semantics apply only within the active cluster scope.
'   - Override resolution follows the same pipeline used for nodes, edges,
'     clusters, and graph-level keywords, ensuring consistent behavior across
'     all attribute types.
'   - DeepWiki Context: Implements the node-keyword rules described in the
'     "Keywords", "Clusters", and "Node Metadata" documentation.
' ==========================================================================
Private Sub ProcessKeywordNode(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow)

    If ctx.clusters.Count <= 0 Then
        EmitMessageSilent GetMessage("errormsgPKNClusterStackIsEmpty"), esWarning
        Exit Sub
    End If

    Dim scope As DotScope
    Set scope = ctx.clusters.Peek

    Dim label As String
    ApplyLabelOverrides ini, data, "label", data.label, label
    scope.nodeLabel = label

    Dim xlabel As String
    ApplyLabelOverrides ini, data, "xlabel", data.xlabel, xlabel
    scope.nodeXLabel = xlabel

    Dim tooltip As String
    ApplyLabelOverrides ini, data, "tooltip", data.tooltip, tooltip
    scope.nodeTooltip = tooltip

End Sub

' ==========================================================================
' ROUTINE: ProcessKeywordEdge
'
' PURPOSE:
'   Applies edge-level keyword semantics to the current cluster scope.
'   Resolves all supported edge-label variants (label/xlabel/taillabel/
'   headlabel/tooltip) using the full override pipeline and stores the
'   results in the active DotScope for later inheritance by edge rows.
'
' FUNCTIONAL WORKFLOW:
'   1. CLUSTER-SCOPE VALIDATION:
'        - Ensures that a cluster scope is active before applying edge-level
'          keyword semantics.
'        - Emits a warning and exits when the cluster stack is empty, since
'          edge-level metadata must attach to an active scope.
'
'   2. SCOPE RETRIEVAL:
'        - Retrieves the current DotScope from the top of the cluster stack.
'
'   3. LABEL OVERRIDE RESOLUTION:
'        - Invokes ApplyLabelOverrides for each supported edge label type:
'             o "label"      -> scope.edgeLabel
'             o "xlabel"     -> scope.edgeXLabel
'             o "taillabel"  -> scope.edgeTailLabel
'             o "headlabel"  -> scope.edgeHeadLabel
'             o "tooltip"    -> scope.edgeTooltip
'        - Each override resolution incorporates:
'             o row-level values,
'             o style-format templates,
'             o placeholder expansion rules.
'
'   4. METADATA ASSIGNMENT:
'        - Stores the resolved label values directly on the DotScope, making
'          them available for inheritance by subsequent edge rows.
'
' TECHNICAL NOTES:
'   - Edge-level keyword semantics apply only within the active cluster scope.
'   - Override resolution follows the same pipeline used for nodes, edges,
'     clusters, and graph-level keywords, ensuring consistent behavior across
'     all attribute types.
'   - DeepWiki Context: Implements the edge-keyword rules described in the
'     "Keywords", "Clusters", and "Edge Metadata" documentation.
' ==========================================================================
Private Sub ProcessKeywordEdge(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow)

    If ctx.clusters.Count <= 0 Then
        EmitMessageSilent GetMessage("errormsgPKEClusterStackIsEmpty"), esWarning
        Exit Sub
    End If

    Dim scope As DotScope
    Set scope = ctx.clusters.Peek

    Dim label As String
    ApplyLabelOverrides ini, data, "label", data.label, label
    scope.edgeLabel = label

    Dim xlabel As String
    ApplyLabelOverrides ini, data, "xlabel", data.xlabel, xlabel
    scope.edgeXLabel = xlabel

    Dim taillabel As String
    ApplyLabelOverrides ini, data, "taillabel", data.taillabel, taillabel
    scope.edgeTailLabel = taillabel

    Dim headlabel As String
    ApplyLabelOverrides ini, data, "headlabel", data.headlabel, headlabel
    scope.edgeHeadLabel = headlabel

    Dim tooltip As String
    ApplyLabelOverrides ini, data, "tooltip", data.tooltip, tooltip
    scope.edgeTooltip = tooltip

End Sub

' ==========================================================================
' ROUTINE: ProcessKeywordGraph
'
' PURPOSE:
'   Applies graph-level keyword semantics to the current cluster scope.
'   Resolves label and tooltip overrides using the full label-pipeline
'   (including style-format rules and placeholder expansion) and records
'   the originating row for diagnostic and debug purposes.
'
' FUNCTIONAL WORKFLOW:
'   1. CLUSTER-SCOPE VALIDATION:
'        - Ensures that a cluster scope is active before applying graph-level
'          keyword semantics.
'        - Emits a warning and exits when the cluster stack is empty, since
'          graph-level metadata must attach to an active scope.
'
'   2. SCOPE RETRIEVAL:
'        - Retrieves the current DotScope from the top of the cluster stack.
'
'   3. LABEL OVERRIDE RESOLUTION:
'        - Invokes ApplyLabelOverrides to compute the final graph label using:
'             o row-level label,
'             o style-format templates,
'             o placeholder expansion rules.
'        - Stores the resolved label in scope.graphLabel.
'
'   4. TOOLTIP OVERRIDE RESOLUTION:
'        - Invokes ApplyLabelOverrides to compute the final graph tooltip.
'        - Stores the resolved tooltip in scope.graphTooltip.
'
'   5. ROW TRACKING:
'        - Records the originating row number in scope.graphRow for debugging,
'          traceability, and optional debug-label synthesis.
'
' TECHNICAL NOTES:
'   - Graph-level keyword semantics apply only within the active cluster scope.
'   - Label and tooltip resolution follow the same override pipeline used for
'     nodes, edges, and clusters, ensuring consistent behavior across all
'     attribute types.
'   - DeepWiki Context: Implements the graph-keyword rules described in the
'     "Keywords", "Clusters", and "Graph Metadata" documentation.
' ==========================================================================
Private Sub ProcessKeywordGraph(ByRef ini As settings, _
                                    ByRef ctx As KnowledgeGraphContext, _
                                    ByRef data As dataRow)

    If ctx.clusters.Count <= 0 Then
        EmitMessageSilent GetMessage("errormsgPKGClusterStackIsEmpty"), esWarning
        Exit Sub
    End If

    Dim scope As DotScope
    Set scope = ctx.clusters.Peek

    Dim label As String
    ApplyLabelOverrides ini, data, "label", data.label, label
    scope.graphLabel = label

    Dim tooltip As String
    ApplyLabelOverrides ini, data, "tooltip", data.tooltip, tooltip
    scope.graphTooltip = tooltip

    scope.graphRow = data.row
End Sub

' ==========================================================================
' FUNCTION: NormalizeLabel
'
' PURPOSE:
'   Normalizes a label string by removing Graphviz escape sequences and
'   replacing actual control characters with spaces. Produces a clean,
'   serialization-safe label suitable for downstream processing and
'   Graphviz emission.
'
' FUNCTIONAL WORKFLOW:
'   1. ESCAPE-SEQUENCE NORMALIZATION:
'        - Replaces blank labels specified as "" as a null string
'          and early terminates to avoid additional string tests.
'        - Replaces Graphviz escape tokens:
'             o "\n" -> " "
'             o "\r" -> " "
'             o "\l" -> " "
'          ensuring that escaped line/left-justify markers do not propagate
'          into final label output.
'
'   2. CONTROL-CHARACTER REMOVAL:
'        - Replaces actual control characters:
'             o vbCrLf -> " "
'             o vbCr   -> " "
'             o vbLf   -> " "
'          preventing multi-line or malformed labels from entering the
'          attribute dictionary.
'
'   3. OUTPUT:
'        - Returns the fully normalized string with all escape sequences and
'          control characters replaced by spaces.
'
' TECHNICAL NOTES:
'   - This routine does not trim leading/trailing whitespace; callers may
'     apply additional formatting as needed.
'   - Normalization is intentionally conservative: it avoids collapsing
'     multiple spaces or altering visible content beyond escape/control
'     cleanup.
'   - DeepWiki Context: Implements the label-normalization rules described
'     in the "Labels", "Formatting", and "Graphviz Compatibility" sections.
' ==========================================================================
Private Function NormalizeLabel(ByVal s As String) As String
    Dim t As String
    t = s

    ' Blank label "" passed as string, instead of empty string
    If t = Chr$(34) & Chr$(34) Then
        NormalizeLabel = vbNullString
        Exit Function
    End If
    
    ' Graphviz escapes
    t = replace(t, "\n", " ")
    t = replace(t, "\r", " ")
    t = replace(t, "\l", " ")

    ' Actual control characters
    t = replace(t, vbCrLf, " ")
    t = replace(t, vbCr, " ")
    t = replace(t, vbLf, " ")

    NormalizeLabel = t
End Function

' ==========================================================================
' ROUTINE: DispatchRow
'
' PURPOSE:
'   Routes a parsed data-row to the appropriate processing subsystem based on
'   its classified styleType. Acts as the central dispatcher for the Knowledge
'   Graph pipeline, ensuring that each row is handled by the correct semantic
'   processor (keyword, node, edge, cluster-open/close, native).
'
' FUNCTIONAL WORKFLOW:
'   1. ROW-TYPE CLASSIFICATION:
'        - Examines data.styleType to determine the semantic meaning of the row.
'        - Uses predefined TYPE_* constants to identify the correct handler.
'
'   2. SUBSYSTEM DISPATCH:
'        - TYPE_KEYWORD
'             o Invokes ProcessKeyword to configure graph/node/edge-level
'               keyword semantics and merge keyword-level properties.
'
'        - TYPE_NODE
'             o Invokes ProcessNode to synthesize or update node definitions.
'
'        - TYPE_EDGE
'             o Invokes ProcessEdge to synthesize edges, merge semantic
'               properties, and tag orphan nodes.
'
'        - TYPE_SUBGRAPH_OPEN
'             o Invokes ProcessClusterOpen to begin a new cluster scope.
'
'        - TYPE_SUBGRAPH_CLOSE
'             o Invokes ProcessClusterClose to unwind the cluster stack.
'
'        - TYPE_NATIVE
'             o Invokes ProcessNative to surface unsupported native-row usage.
'
'        - Unknown row type
'             o Emits a warning indicating the row type is not recognized.
'
'   3. ERROR HANDLING:
'        - All subsystem routines handle their own errors; DispatchRow simply
'          forwards the row to the correct processor.
'
' TECHNICAL NOTES:
'   - DispatchRow is the top-level control point for the entire row-processing
'     pipeline and is invoked once per parsed input row.
'   - Ensures strict separation of concerns: each subsystem handles only its
'     own semantics, while DispatchRow handles classification and routing.
'   - DeepWiki Context: Implements the row-dispatch rules described in the
'     "Input Processing", "Row Types", and "Pipeline Architecture" documentation.
' ==========================================================================
Private Sub DispatchRow(ByRef ini As settings, _
                                      ByRef ctx As KnowledgeGraphContext, _
                                      ByRef data As dataRow)
    Select Case data.styleType
        Case TYPE_KEYWORD
            ProcessKeyword ini, ctx, data
            
        Case TYPE_NODE
            ProcessNode ini, ctx, data
            
        Case TYPE_EDGE
            ProcessEdge ini, ctx, data
            
        Case TYPE_SUBGRAPH_OPEN
            ProcessClusterOpen ini, ctx, data
            
        Case TYPE_SUBGRAPH_CLOSE
            ProcessClusterClose ini, ctx, data
            
        Case TYPE_NATIVE
            ProcessNative ini, ctx, data
            
        Case Else
            EmitMessageSilent GetMessage("errormsgUnknownRowType") & data.styleType, esWarning
    End Select

End Sub

' ==========================================================================
' ROUTINE: ProcessKeyword
'
' PURPOSE:
'   Dispatches keyword-typed rows (graph/node/edge) to their respective
'   processing routines and merges any associated semantic properties into
'   the appropriate keyword-level property dictionaries. Ensures that global
'   keyword configuration is applied before node/edge synthesis occurs.
'
' FUNCTIONAL WORKFLOW:
'   1. PROPERTY PARSING:
'        - When the row contains a non-empty properties string:
'             o Parses the string into a dictionary of semantic properties.
'             o Leaves properties = Nothing when no properties are present.
'
'   2. KEYWORD DISPATCH:
'        - Normalizes the .item field to lowercase and dispatches based on
'          recognized keyword types:
'
'             o TYPE_GRAPH
'                   - Invokes ProcessKeywordGraph.
'                   - Merges parsed properties into ctx.propertiesGraph.
'
'             o TYPE_NODE
'                   - Invokes ProcessKeywordNode.
'                   - Merges parsed properties into ctx.propertiesNode.
'
'             o TYPE_EDGE
'                   - Invokes ProcessKeywordEdge.
'                   - Merges parsed properties into ctx.propertiesEdge.
'
'             o Unknown keyword
'                   - Emits a warning indicating the keyword is not supported.
'
'   3. PROPERTY MERGING:
'        - Uses MergeDictionaries to combine existing keyword-level properties
'          with newly parsed row-level properties.
'        - Ensures keyword-level configuration follows DeepWiki precedence:
'             graph < node < edge < row.
'
'   4. ERROR HANDLING:
'        - Surfaces any runtime errors with a descriptive message including
'          the originating keyword.
'        - Cleans up temporary property dictionaries to avoid memory leaks.
'
' TECHNICAL NOTES:
'   - Keyword rows configure global behavior for subsequent node/edge/cluster
'     synthesis and are processed before structural rows.
'   - Property dictionaries stored in ctx.propertiesGraph / Node / Edge
'     serve as base layers for semantic-property merging during node/edge
'     creation and update.
'   - DeepWiki Context: Implements the keyword-processing rules described in
'     the "Keywords", "Properties", and "Graph Configuration" documentation.
' ==========================================================================
Private Sub ProcessKeyword(ByRef ini As settings, _
                               ByRef ctx As KnowledgeGraphContext, _
                               ByRef data As dataRow)
    
    Dim properties As Dictionary
    
    On Error GoTo Cleanup

    ' Only parse if there are actual properties
    If Len(data.properties) > 0 Then
        Set properties = ParsePropertyString(data.properties)
    End If

    With data
        Select Case LCase$(.item)
            
            Case TYPE_GRAPH
                ProcessKeywordGraph ini, ctx, data
                If Not properties Is Nothing Then
                    Set ctx.propertiesGraph = MergeDictionaries(ctx.propertiesGraph, properties)
                End If
                
            Case TYPE_NODE
                ProcessKeywordNode ini, ctx, data
                If Not properties Is Nothing Then
                    Set ctx.propertiesNode = MergeDictionaries(ctx.propertiesNode, properties)
                End If
                
            Case TYPE_EDGE
                ProcessKeywordEdge ini, ctx, data
                If Not properties Is Nothing Then
                    Set ctx.propertiesEdge = MergeDictionaries(ctx.propertiesEdge, properties)
                End If
                
            Case Else
                EmitMessageSilent GetMessage("errormsgUnknownKeywordType") & .item, esWarning
                
        End Select
    End With

Cleanup:
    Set properties = Nothing
    
    If Err.number <> 0 Then
        ' Localize the full error message
        Dim fullMessage As String
        fullMessage = GetMessage("errormsgProcessKeywordError")
        fullMessage = replace(fullMessage, "{item}", data.item, 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.number}", CStr(Err.number), 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.description}", Err.Description, 1, 1, vbTextCompare)
        
        EmitMessageSilent fullMessage, esError
    End If

End Sub

' ==========================================================================
' ROUTINE: ProcessNode
'
' PURPOSE:
'   Synthesizes one or more node definitions from a single data-row entry.
'   Supports multi-ID expansion, cluster inheritance, node creation/update
'   dispatch, and safe error handling. Ensures that each referenced node is
'   fully defined or updated according to the current row context.
'
' FUNCTIONAL WORKFLOW:
'   1. NODE-LIST EXTRACTION:
'        - Retrieves and trims the item field.
'        - Exits early when the field is empty.
'        - Splits comma-delimited node identifiers into individual tokens.
'
'   2. CLUSTER INHERITANCE:
'        - If a cluster scope is active, retrieves the current cluster name
'          from the top of the cluster stack.
'
'   3. MULTI-NODE EXPANSION:
'        - Iterates over each node token:
'             o Trims whitespace.
'             o Removes any port suffix (node:port).
'             o Normalizes the node key to uppercase for deduplication.
'
'   4. CREATE/UPDATE DISPATCH:
'        - When the node does not exist:
'             o Invokes ProcessNodeCreate to build a new node dictionary.
'        - When the node already exists:
'             o Invokes ProcessNodeUpdate to enrich or finalize the node,
'               resolving orphan placeholders and applying row-level updates.
'
'   5. ERROR HANDLING:
'        - Surfaces any runtime errors with a descriptive message including
'          the originating item field.
'
' TECHNICAL NOTES:
'   - Node creation and update follow DeepWiki's precedence rules:
'        graph < node < row.
'   - Uppercase node keys ensure consistent deduplication across edges,
'     clusters, and node definitions.
'   - Port removal ensures that node identity is not conflated with port
'     syntax used by edges.
'   - DeepWiki Context: Implements the node-processing rules described in
'     the "Nodes", "Clusters", and "Edges" documentation.
' ==========================================================================
Private Sub ProcessNode(ByRef ini As settings, _
                            ByRef ctx As KnowledgeGraphContext, _
                            ByRef data As dataRow)

    Dim idList As String
    idList = Trim$(data.item)
    If Len(idList) = 0 Then Exit Sub

    Dim itemIds() As String
    itemIds = split(idList, COMMA)

    Dim clusterName As String
    clusterName = vbNullString
    
    If ctx.clusters.Count > 0 Then
        Dim scope As DotScope
        Set scope = ctx.clusters.Peek
        clusterName = scope.clusterName
    End If

    Dim nodeId As String
    Dim nodeKey As String
    Dim arrItem As Variant

    On Error GoTo Cleanup

    For Each arrItem In itemIds
        nodeId = RemovePort(Trim$(CStr(arrItem)))
        If Len(nodeId) = 0 Then GoTo NextItem

        ' Always use uppercase for deduplication
        nodeKey = UCase$(nodeId)
        
        If Not ctx.nodes.Exists(nodeKey) Then
            ProcessNodeCreate ini, ctx, data, nodeId, clusterName
        Else
            ProcessNodeUpdate ini, ctx, data, nodeId, clusterName
        End If

NextItem:
    Next arrItem

Cleanup:
    If Err.number <> 0 Then
        ' Localize the full error message
        Dim fullMessage As String
        fullMessage = GetMessage("errormsgProcessNodeError")
        fullMessage = replace(fullMessage, "{item}", data.item, 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.number}", CStr(Err.number), 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.description}", Err.Description, 1, 1, vbTextCompare)
        
        EmitMessageSilent fullMessage, esError
    End If

End Sub

' ==========================================================================
' ROUTINE: ProcessNodeCreate
'
' PURPOSE:
'   Creates a fully defined node dictionary for a new node encountered in the
'   Knowledge Graph pipeline. Assigns identity, cluster membership, style,
'   synthesized labels, semantic properties, and optional debug metadata
'   before registering the node in the context.
'
' FUNCTIONAL WORKFLOW:
'   1. NODE INITIALIZATION:
'        - Constructs a new dictionary for the node.
'        - Adds the canonical "id" field using the caller-supplied nodeId.
'
'   2. CLUSTER & STYLE ASSIGNMENT:
'        - Adds "cluster" when the node is created inside an active cluster.
'        - Adds "style" when the data row specifies a styleName.
'
'   3. LABEL SYNTHESIS:
'        - Applies ProcessLabel for each supported node label type:
'             o "label"   - primary node label.
'             o "xlabel"  - auxiliary node label.
'             o "tooltip" - hover text (when enabled).
'        - Each call applies override rules, placeholder expansion, and
'          normalization before merging into the node dictionary.
'
'   4. SEMANTIC-PROPERTY MERGING:
'        - Merges graph-level and node-level semantic properties with
'          row-level properties.
'        - Adds the resulting merged dictionary under the "properties" field.
'
'   5. CONTEXT REGISTRATION:
'        - Inserts the completed node dictionary into ctx.nodes using the
'          uppercase node key.
'
'   6. DEBUG AUGMENTATION:
'        - When debugging is enabled, appends a "debuglabel" attribute
'          containing the originating row number.
'
'   7. ERROR HANDLING:
'        - Surfaces any runtime errors with a descriptive message including
'          the node identifier.
'        - Cleans up temporary dictionaries to avoid memory leaks.
'
' TECHNICAL NOTES:
'   - Node creation is distinct from node update: this routine initializes
'     a new node, while ProcessNodeUpdate enriches or replaces orphan entries.
'   - Semantic-property precedence follows DeepWiki rules:
'        graph < node < row.
'   - Label synthesis uses the same pipeline as edges and clusters, ensuring
'     consistent override and placeholder behavior.
'   - DeepWiki Context: Implements the node-creation rules described in the
'     "Nodes", "Properties", and "Labels" documentation.
' ==========================================================================
Private Sub ProcessNodeCreate(ByRef ini As settings, _
                              ByRef ctx As KnowledgeGraphContext, _
                              ByRef data As dataRow, _
                              ByVal nodeId As String, _
                              ByVal currentCluster As String)
    Dim nodeKey As String
    nodeKey = UCase$(nodeId)

    Dim d As Dictionary
    Set d = New Dictionary
    
    d.Add "id", nodeId
    
    If Len(currentCluster) > 0 Then d.Add "cluster", currentCluster
    If Len(data.styleName) > 0 Then d.Add "style", data.styleName

    ' === Label Processing ===
    ProcessLabel ini, data, d, "label", data.label, ini.graph.includeNodeLabels
    ProcessLabel ini, data, d, "xlabel", data.xlabel, ini.graph.includeNodeXLabels
    ProcessLabel ini, data, d, "tooltip", data.tooltip, ini.graph.includeNodeTooltips

    ' === Semantic Properties ===
    If ctx.propertiesGraph.Count > 0 Or ctx.propertiesNode.Count > 0 Or Len(data.properties) > 0 Then
        Dim d1 As Dictionary, d2 As Dictionary, propertiesDict As Dictionary
        
        Set d1 = MergeDictionaries(ctx.propertiesGraph, ctx.propertiesNode)
        Set d2 = ParsePropertyString(data.properties)
        Set propertiesDict = MergeDictionaries(d1, d2)
        
        d.Add "properties", propertiesDict
    End If

    ctx.nodes.Add nodeKey, d

    ' --- Debug Label ---
    If ini.graph.debug Then MergeLabelField d("id"), d, "debuglabel", "Row: " & data.row

Cleanup:
    Set d1 = Nothing
    Set d2 = Nothing
    Set propertiesDict = Nothing
    
    If Err.number <> 0 Then
        ' Localize the full error message
        Dim fullMessage As String
        fullMessage = GetMessage("errormsgProcessNodeCreateError")
        fullMessage = replace(fullMessage, "{item}", nodeId, 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.number}", CStr(Err.number), 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.description}", Err.Description, 1, 1, vbTextCompare)
        
        EmitMessageSilent fullMessage, esError
    End If

End Sub

' ==========================================================================
' ROUTINE: ProcessNodeUpdate
'
' PURPOSE:
'   Updates an existing node dictionary with cluster membership, style
'   inheritance, label synthesis, semantic-property merging, and optional
'   debug augmentation. Converts orphan placeholders into fully defined
'   nodes and ensures all node attributes reflect the current data-row
'   context.
'
' FUNCTIONAL WORKFLOW:
'   1. NODE LOOKUP & ORPHAN RESOLUTION:
'        - Retrieves the existing node dictionary using the uppercase key.
'        - Removes the "defined" sentinel when present, marking the node as
'          explicitly defined rather than an orphan placeholder.
'
'   2. CLUSTER ASSIGNMENT:
'        - When a cluster is active and the node lacks a "cluster" field,
'          assigns the current cluster name to the node.
'
'   3. STYLE UPDATE:
'        - If the row specifies a styleName, updates or inserts the "style"
'          attribute accordingly.
'
'   4. LABEL SYNTHESIS:
'        - Applies ProcessLabel for each supported node label type:
'             o "label"   - primary node label.
'             o "xlabel"  - auxiliary node label.
'             o "tooltip" - hover text (when enabled).
'        - Each call applies override rules, placeholder expansion, and
'          normalization before merging into the node dictionary.
'
'   5. SEMANTIC-PROPERTY MERGING:
'        - Merges graph-level and node-level semantic properties with
'          row-level properties.
'        - If the node already contains a "properties" dictionary, merges
'          into it; otherwise, adds a new merged dictionary.
'
'   6. DEBUG AUGMENTATION:
'        - When debugging is enabled, appends a "debuglabel" attribute
'          containing the originating row number.
'
'   7. ERROR HANDLING:
'        - Surfaces any runtime errors with a descriptive message including
'          the node identifier.
'        - Cleans up temporary dictionaries to avoid memory leaks.
'
' TECHNICAL NOTES:
'   - Converts orphan nodes (created implicitly by edges) into fully defined
'     nodes once a row explicitly defines them.
'   - Semantic-property precedence follows DeepWiki rules:
'        graph < node < row.
'   - Label synthesis uses the same pipeline as edges and clusters, ensuring
'     consistent override and placeholder behavior.
'   - DeepWiki Context: Implements the node-update rules described in the
'     "Nodes", "Properties", and "Labels" documentation.
' ==========================================================================
Private Sub ProcessNodeUpdate(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow, _
                                   ByVal nodeId As String, _
                                   ByVal currentCluster As String)
    Dim nodeKey As String
    nodeKey = UCase$(nodeId)

    Dim d As Dictionary
    
    On Error GoTo Cleanup

    Set d = ctx.nodes(nodeKey)
    
    If d.Exists("defined") Then d.Remove "defined"

    If Len(currentCluster) > 0 And Not d.Exists("cluster") Then
        d.Add "cluster", currentCluster
    End If

    If Len(data.styleName) > 0 Then d("style") = data.styleName

    ' === Label Processing ===
    ProcessLabel ini, data, d, "label", data.label, ini.graph.includeNodeLabels
    ProcessLabel ini, data, d, "xlabel", data.xlabel, ini.graph.includeNodeXLabels
    ProcessLabel ini, data, d, "tooltip", data.tooltip, ini.graph.includeNodeTooltips

    ' Semantic Properties
    If ctx.propertiesGraph.Count > 0 Or ctx.propertiesNode.Count > 0 Or Len(data.properties) > 0 Then
        Dim baseProps As Dictionary, rowProps As Dictionary, mergedProps As Dictionary
        
        Set baseProps = MergeDictionaries(ctx.propertiesGraph, ctx.propertiesNode)
        Set rowProps = ParsePropertyString(data.properties)
        Set mergedProps = MergeDictionaries(baseProps, rowProps)
        
        If d.Exists("properties") Then
            Set d("properties") = MergeDictionaries(d("properties"), mergedProps)
        Else
            d.Add "properties", mergedProps
        End If
    End If

    ' --- Debug Label ---
    If ini.graph.debug Then MergeLabelField d("id"), d, "debuglabel", "Row: " & data.row

Cleanup:
    Set d = Nothing
    Set baseProps = Nothing
    Set rowProps = Nothing
    Set mergedProps = Nothing
    
    If Err.number <> 0 Then
        ' Localize the full error message
        Dim fullMessage As String
        fullMessage = GetMessage("errormsgProcessNodeUpdateError")
        fullMessage = replace(fullMessage, "{item}", nodeId, 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.number}", CStr(Err.number), 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.description}", Err.Description, 1, 1, vbTextCompare)
        
        EmitMessageSilent fullMessage, esError
    End If

End Sub

' ==========================================================================
' ROUTINE: ProcessLabel
'
' PURPOSE:
'   Synthesizes a single label field for a node, edge, or cluster by applying
'   feature-switch gating, override rules, placeholder expansion, normalization,
'   and structured-label merging. Produces the final label value that will be
'   stored in the attribute dictionary.
'
' FUNCTIONAL WORKFLOW:
'   1. FEATURE-SWITCH GATING:
'        - Immediately exits when includeField = False, allowing callers to
'          disable specific label types (label/xlabel/headlabel/taillabel/
'          tooltip) via configuration.
'
'   2. OVERRIDE RESOLUTION:
'        - Initializes finalValue with the caller-supplied fieldValue.
'        - Invokes ApplyLabelOverrides to incorporate:
'             o extra-attribute overrides,
'             o style-format template substitutions,
'             o placeholder expansion rules.
'
'   3. NORMALIZATION:
'        - When the resolved label is non-empty:
'             o Applies NormalizeLabel to ensure consistent whitespace,
'               quoting, and Graphviz-compatible formatting.
'
'   4. STRUCTURED-LABEL MERGING:
'        - Invokes MergeLabelField to insert or update the label field in the
'          attribute dictionary, preserving both "value" and "type" metadata.
'
' TECHNICAL NOTES:
'   - This routine does not perform HTML-like detection directly; normalization
'     and type classification occur inside MergeLabelField.
'   - Placeholder expansion is performed only when enabled by style-format
'     rules; no implicit expansion occurs.
'   - DeepWiki Context: Implements the label-synthesis rules described in the
'     "Labels", "Styles", and "Attribute Inheritance" documentation.
' ==========================================================================
Private Sub ProcessLabel(ByRef ini As settings, _
                         ByRef data As dataRow, _
                         ByRef d As Dictionary, _
                         ByVal fieldName As String, _
                         ByRef fieldValue As String, _
                         ByVal includeField As Boolean)

    ' Ensure field is not returned if disabled by user
    If Not includeField Then
        If d.Exists(fieldName) Then d.Remove fieldName
        Exit Sub
    End If

    Dim finalValue As String
    finalValue = fieldValue
    
    ' Resolve the actual label taking feature switches and inheritance into consideration
    ApplyLabelOverrides ini, data, fieldName, fieldValue, finalValue
    
    ' --- Normalize + Merge ---
    If Len(Trim$(finalValue)) > 0 Then
        finalValue = NormalizeLabel(finalValue)
        MergeLabelField d("id"), d, fieldName, finalValue
    End If
End Sub

' ==========================================================================
' ROUTINE: ProcessClusterOpen
'
' PURPOSE:
'   Opens a new Graphviz cluster scope and registers it in the context stack.
'   Assigns a unique cluster name, inherits style attributes when appropriate,
'   and emits the corresponding cluster node via ProcessClusterCreate.
'
' FUNCTIONAL WORKFLOW:
'   1. CURRENT SCOPE CAPTURE:
'        - Retrieves the current cluster count.
'        - If a cluster is already open, captures its name and scope object
'          for inheritance.
'
'   2. CLUSTER-NAME GENERATION:
'        - Increments the cluster counter and constructs a unique name
'          ("cluster_n") for the new scope.
'
'   3. SCOPE INHERITANCE & CREATION:
'        - If no clusters exist, creates a fresh DotScope.
'        - If a cluster is active, clones the existing scope to preserve
'          inherited attributes and styling rules.
'        - Assigns the new cluster name to the scope.
'
'   4. CONTEXT STACK REGISTRATION:
'        - Pushes the new cluster scope onto the cluster stack so that
'          subsequent nodes and edges resolve attributes within the correct
'          hierarchical context.
'
'   5. CLUSTER-NODE EMISSION:
'        - Invokes ProcessClusterCreate to generate the Graphviz node
'          representing the cluster boundary, passing both the new and
'          previous cluster names.
'
' TECHNICAL NOTES:
'   - Cluster scopes are hierarchical; cloning preserves inherited attributes
'     while allowing new scopes to override or extend them.
'   - The cluster stack ensures correct resolution of nested cluster rules
'     during node and edge synthesis.
'   - DeepWiki Context: Implements the cluster-opening rules described in the
'     "Clusters" and "Scope Inheritance" documentation.
' ==========================================================================
Private Sub ProcessClusterOpen(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow)

    Dim openClusterCount As Long
    openClusterCount = ctx.clusterCount
    
    Dim oldClusterName As String
    oldClusterName = vbNullString

    Dim oldCluster As DotScope
    If ctx.clusters.Count > 0 Then
        Set oldCluster = ctx.clusters.Peek
        oldClusterName = oldCluster.clusterName
    End If
    
    ctx.clusterCount = openClusterCount + 1
    Dim newClusterName As String
    newClusterName = "cluster_" & CStr(ctx.clusterCount)

    Dim newCluster As DotScope
    If ctx.clusters.Count = 0 Then
        Set newCluster = New DotScope
    Else
        Set newCluster = oldCluster.Clone
    End If
    
    newCluster.clusterName = newClusterName
    ctx.clusters.Push newCluster

    ' Add the cluster node
    ProcessClusterCreate ini, ctx, data, newClusterName, oldClusterName

End Sub

' ==========================================================================
' ROUTINE: ProcessClusterCreate
'
' PURPOSE:
'   Creates the Graphviz node representing an opened cluster scope. Assigns
'   identity and styling attributes, applies label and tooltip synthesis
'   rules, and registers the resulting node dictionary in the context's
'   node collection.
'
' FUNCTIONAL WORKFLOW:
'   1. STYLE-NAME RESOLUTION:
'        - Extracts the style name from the data row.
'        - Removes any cluster-opening suffix (styles.affixOpen) to obtain
'          the canonical style identifier.
'
'   2. CLUSTER-DICTIONARY CONSTRUCTION:
'        - Creates a new attribute dictionary for the cluster node.
'        - Adds required fields:
'             o "id"     - unique cluster identifier.
'             o "type"   - always "cluster".
'        - Adds optional fields:
'             o "cluster" - parent cluster name when nested.
'             o "style"   - resolved style name when present.
'
'   3. LABEL & TOOLTIP SYNTHESIS:
'        - Invokes ProcessLabel to apply override rules, placeholder
'          expansion, and conditional emission for:
'             o "label"   - cluster label.
'             o "tooltip" - cluster tooltip (when enabled).
'
'   4. CONTEXT REGISTRATION:
'        - Adds the completed cluster dictionary to ctx.nodes using the
'          cluster ID as the key.
'
'   5. DEBUG AUGMENTATION:
'        - When debugging is enabled, appends a "debuglabel" attribute
'          containing the originating row number.
'
' TECHNICAL NOTES:
'   - Cluster nodes represent Graphviz subgraph boundaries and participate
'     in hierarchical styling and inheritance rules.
'   - Label and tooltip synthesis follow the same pipeline used for nodes
'     and edges, ensuring consistent placeholder and override behavior.
'   - DeepWiki Context: Implements the cluster-node creation rules described
'     in the "Clusters" and "Scope Inheritance" documentation.
' ==========================================================================
Private Sub ProcessClusterCreate(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow, _
                                   ByVal id As String, _
                                   ByVal parentId As String)

    Dim styNam As String:  styNam = Trim$(data.styleName)

    If Len(styNam) > 0 Then
        styNam = Trim$(StripSuffix(styNam, ini.styles.affixOpen))
    End If

    Dim clusterDict As Dictionary
    Set clusterDict = New Dictionary

    With clusterDict
        .Add "id", id
        .Add "type", "cluster"
        
        If Len(parentId) > 0 Then .Add "cluster", parentId
        If Len(styNam) > 0 Then .Add "style", styNam

        ' Handle label inheritance and placeholder substitution
        ProcessLabel ini, data, clusterDict, "label", data.label, ini.graph.includeClusterLabels
        ProcessLabel ini, data, clusterDict, "tooltip", data.tooltip, ini.graph.includeClusterTooltips
    End With

    ctx.nodes.Add id, clusterDict

    ' --- Debug Label ---
    If ini.graph.debug Then MergeLabelField clusterDict("id"), clusterDict, "debuglabel", "Row: " & data.row

End Sub

' ==========================================================================
' ROUTINE: ProcessClusterClose
'
' PURPOSE:
'   Closes the most recently opened Graphviz cluster scope by popping the
'   top entry from the cluster stack. Ensures proper unwinding of nested
'   cluster contexts and surfaces malformed input when an unmatched closing
'   brace is encountered.
'
' FUNCTIONAL WORKFLOW:
'   1. STACK CHECK:
'        - If one or more cluster scopes are active, removes the top scope
'          from the stack, restoring the previous cluster context.
'
'   2. MALFORMED-INPUT HANDLING:
'        - If no cluster scopes are active, the closing brace represents an
'          unmatched or extraneous cluster terminator.
'        - Emits a non-fatal diagnostic message indicating the mismatch and
'          ignores the closing brace to preserve the root sentinel.
'
' TECHNICAL NOTES:
'   - Cluster scopes are managed as a LIFO stack; each ProcessClusterOpen
'     must be paired with a corresponding ProcessClusterClose.
'   - Unmatched closing braces do not interrupt processing; they are surfaced
'     as errors but safely ignored to prevent corruption of the cluster stack.
'   - DeepWiki Context: Implements the cluster-closing rules described in the
'     "Clusters" and "Scope Inheritance" documentation.
' ==========================================================================
Private Sub ProcessClusterClose(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow)
                                   
    If ctx.clusters.Count > 0 Then
        ctx.clusters.Pop
    Else
        ' Extra/unmatched '}' with no corresponding open - malformed input.
        ' Ignore it rather than popping the root sentinel away, and surface
        ' it so the malformed data doesn't fail silently.
        EmitMessageSilent GetMessage("errormsgUnmatchedClosingBrace"), esError
    End If
    
End Sub

' ==========================================================================
' ROUTINE: ProcessNative
'
' PURPOSE:
'   Handles rows marked as type "native" (>) within the Knowledge Graph
'   processing pipeline. Since native rows are not part of the supported
'   knowledge grapg synthesis model, the routine emits a warning and
'   performs no additional processing.
'
' FUNCTIONAL WORKFLOW:
'   1. INPUT CLASSIFICATION:
'        - Identifies rows of type "native" based on the data-row metadata.
'
'   2. WARNING EMISSION:
'        - Produces a non-fatal diagnostic message indicating that native
'          rows are unsupported and will be ignored.
'        - Surfaces the issue to prevent silent failure or confusion during
'          data preparation.
'
'   3. NO-OP PROCESSING:
'        - Does not modify the context, node set, cluster stack, or any
'          attribute dictionaries.
'        - Ensures that unsupported rows do not interfere with downstream
'          synthesis rules.
'
' TECHNICAL NOTES:
'   - Native rows are reserved for future extensions or external tooling
'     and are intentionally excluded from Knowledge Graph generation.
'   - Diagnostic messaging is non-localized pending integration with the
'     broader localization subsystem.
'   - DeepWiki Context: Documents the exclusion of native-row semantics
'     from the Knowledge Graph transformation pipeline.
' ==========================================================================
Private Sub ProcessNative(ByRef ini As settings, _
                                   ByRef ctx As KnowledgeGraphContext, _
                                   ByRef data As dataRow)
                                   
    EmitMessageSilent GetMessage("errormsgNativeTypeNotSupported"), esWarning
End Sub

' ==========================================================================
' ROUTINE: ProcessEdge
'
' PURPOSE:
'   Synthesizes one or more Graphviz edges from a single data-row definition.
'   Supports multi-source and multi-target expansion, port extraction,
'   semantic-property merging, edge-ID construction, full label cascading,
'   and orphan-node tagging within the Knowledge Graph pipeline.
'
' FUNCTIONAL WORKFLOW:
'   1. SOURCE/TARGET EXTRACTION:
'        - Retrieves and trims the source and target lists.
'        - If either list is empty, no edges are produced.
'        - Splits comma-delimited lists into individual source/target tokens.
'
'   2. SEMANTIC-PROPERTY MERGING:
'        - When row-level properties exist, merges:
'             o graph-level semantic properties,
'             o edge-level semantic properties,
'             o row-level semantic properties.
'        - Produces a per-edge cloned dictionary to avoid shared mutation.
'
'   3. MULTI-EDGE EXPANSION:
'        - Iterates over each source × target combination.
'        - Extracts node IDs and optional ports.
'        - Applies port-handling rules based on includeEdgePorts.
'
'   4. EDGE-ID CONSTRUCTION:
'        - Builds a unique edge identifier using:
'             o source ID and port,
'             o target ID and port,
'             o configured edge operator,
'             o port-inclusion rules.
'
'   5. EDGE CREATION:
'        - Invokes ProcessEdgeCreate to synthesize the full edge dictionary,
'          including label inheritance, placeholder expansion, and style rules.
'        - Attaches semantic properties when present.
'        - Registers the completed edge in ctx.edges.
'
'   6. ORPHAN-NODE TAGGING:
'        - Ensures both source and target nodes are marked as non-orphans
'          within the node dictionary.
'
'   7. ERROR HANDLING:
'        - Surfaces any runtime errors with a descriptive message including
'          the source and target identifiers.
'        - Cleans up temporary dictionaries to avoid memory leaks.
'
' TECHNICAL NOTES:
'   - Supports fan-out (one-to-many) and fan-in (many-to-one) edge expansion.
'   - Port extraction uses RemovePort and GetPort to maintain Graphviz syntax.
'   - Semantic properties follow DeepWiki's graph/edge/property precedence rules.
'   - DeepWiki Context: Implements the edge-synthesis rules described in the
'     "Edges", "Properties", and "Styles" documentation.
' ==========================================================================
Private Sub ProcessEdge(ByRef ini As settings, _
                            ByRef ctx As KnowledgeGraphContext, _
                            ByRef data As dataRow)

    Dim sourceList As String: sourceList = Trim$(data.item)
    Dim targetList As String: targetList = Trim$(data.relatedItem)

    If Len(sourceList) = 0 Or Len(targetList) = 0 Then Exit Sub

    Dim sources() As String
    Dim targets() As String
    sources = split(sourceList, COMMA)
    targets = split(targetList, COMMA)

    Dim propertiesDict As Dictionary   ' For semantic properties only
    
    On Error GoTo Cleanup

    ' Build semantic properties (Graph + Edge keywords + row)
    If Len(data.properties) > 0 Then
        Dim d1 As Dictionary, d2 As Dictionary
        Set d1 = MergeDictionaries(ctx.propertiesGraph, ctx.propertiesEdge)
        Set d2 = ParsePropertyString(data.properties)
        Set propertiesDict = MergeDictionaries(d1, d2)
        Set d1 = Nothing
        Set d2 = Nothing
    End If

    Dim src As Variant, tgt As Variant
    Dim srcId As String, tgtId As String
    Dim srcPort As String, tgtPort As String
    Dim edgeId As String

    For Each src In sources
        srcId = RemovePort(Trim$(CStr(src)))
        srcPort = IIf(ini.graph.includeEdgePorts, GetPort(Trim$(CStr(src))), "")
        If Len(srcId) = 0 Then GoTo NextSource

        For Each tgt In targets
            tgtId = RemovePort(Trim$(CStr(tgt)))
            tgtPort = IIf(ini.graph.includeEdgePorts, GetPort(Trim$(CStr(tgt))), "")
            If Len(tgtId) = 0 Then GoTo NextTarget

            edgeId = ProcessEdgeBuildId(srcId, srcPort, tgtId, tgtPort, _
                                 ini.graph.edgeOperator, ini.graph.includeEdgePorts)

            ' Create edge with full label cascading
            Dim edgeDict As Dictionary
            Set edgeDict = ProcessEdgeCreate(ini, ctx, data, edgeId, srcId, tgtId, srcPort, tgtPort)

            ' Add semantic properties (cloned per edge)
            If Not propertiesDict Is Nothing Then
                If propertiesDict.Count > 0 Then
                    edgeDict.Add "properties", CloneDictionary(propertiesDict)
                End If
            End If

            ctx.edges.Add edgeDict

            ProcessEdgeTagOrphanNodes ctx.nodes, srcId
            ProcessEdgeTagOrphanNodes ctx.nodes, tgtId

NextTarget:
        Next tgt
NextSource:
    Next src

Cleanup:
    Set propertiesDict = Nothing
    Set edgeDict = Nothing
    
    If Err.number <> 0 Then
        ' Localize the full error message
        Dim fullMessage As String
        fullMessage = GetMessage("errormsgProcessEdgeError")
        fullMessage = replace(fullMessage, "{item}", data.item, 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{relatedItem}", data.relatedItem, 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.number}", CStr(Err.number), 1, 1, vbTextCompare)
        fullMessage = replace(fullMessage, "{Err.description}", Err.Description, 1, 1, vbTextCompare)
        
        EmitMessageSilent fullMessage, esError
    End If
End Sub

' ==========================================================================
' FUNCTION: ProcessEdgeBuildId
'
' PURPOSE:
'   Constructs a canonical Graphviz edge identifier from source and target
'   node IDs, optionally including port information. Ensures consistent
'   formatting for downstream edge-creation and property-attachment routines.
'
' FUNCTIONAL WORKFLOW:
'   1. PORT-INCLUSIVE MODE:
'        - When includePorts = True:
'             o Appends the source port (srcId:srcPort) when present.
'             o Appends the configured edge operator.
'             o Appends the target port (tgtId:tgtPort) when present.
'
'   2. SIMPLE MODE:
'        - When includePorts = False:
'             o Constructs the identifier using only srcId, edgeOperator,
'               and tgtId.
'
'   3. OUTPUT:
'        - Returns the fully assembled edge identifier as a string.
'
' TECHNICAL NOTES:
'   - Port syntax follows Graphviz conventions: node:port.
'   - Edge operator is caller-supplied, allowing support for directed,
'     undirected, and custom syntactic forms.
'   - This routine performs no validation of node IDs or ports; callers
'     are responsible for upstream sanitization.
'   - DeepWiki Context: Implements the edge-identifier construction rules
'     described in the "Edges" and "Syntax" documentation.
' ==========================================================================
Private Function ProcessEdgeBuildId(srcId As String, _
                                        srcPort As String, _
                                        tgtId As String, _
                                        tgtPort As String, _
                                        edgeOperator As String, _
                                        includePorts As Boolean) As String
    If includePorts Then
        ProcessEdgeBuildId = srcId
        If Len(srcPort) > 0 Then ProcessEdgeBuildId = ProcessEdgeBuildId & ":" & srcPort
        ProcessEdgeBuildId = ProcessEdgeBuildId & edgeOperator & tgtId
        If Len(tgtPort) > 0 Then ProcessEdgeBuildId = ProcessEdgeBuildId & ":" & tgtPort
    Else
        ProcessEdgeBuildId = srcId & edgeOperator & tgtId
    End If

End Function

' ==========================================================================
' FUNCTION: ProcessEdgeCreate
'
' PURPOSE:
'   Constructs the complete attribute dictionary for a single Graphviz edge.
'   Applies port-handling rules, style inheritance, full label-synthesis
'   (label/xlabel/headlabel/taillabel/tooltip), and optional debug
'   augmentation before returning the finalized edge dictionary.
'
' FUNCTIONAL WORKFLOW:
'   1. BASE ATTRIBUTE ASSIGNMENT:
'        - Creates a new dictionary and assigns:
'             o "id"      - unique edge identifier.
'             o "source"  - source node ID.
'             o "target"  - target node ID.
'
'   2. PORT HANDLING:
'        - When includeEdgePorts = True:
'             o Adds "tailport" when a source port is present.
'             o Adds "headport" when a target port is present.
'
'   3. STYLE ASSIGNMENT:
'        - If the row specifies a styleName, adds it as the "style" attribute.
'
'   4. LABEL & TOOLTIP SYNTHESIS:
'        - Invokes ProcessLabel for each supported edge attribute:
'             o "label"      - primary edge label.
'             o "xlabel"     - auxiliary edge label.
'             o "headlabel"  - label near the target node.
'             o "taillabel"  - label near the source node.
'             o "tooltip"    - hover text (when enabled).
'        - Each call applies override rules, placeholder expansion, and
'          conditional emission based on pipeline configuration.
'
'   5. DEBUG AUGMENTATION:
'        - When debugging is enabled, appends a "debuglabel" attribute
'          containing the originating row number.
'
'   6. OUTPUT:
'        - Returns the fully assembled edge dictionary for registration
'          in ctx.edges.
'
' TECHNICAL NOTES:
'   - This routine centralizes all edge-attribute synthesis except semantic
'     properties, which are attached by the caller.
'   - Label processing follows the same dictionary-based pipeline used for
'     nodes and clusters, ensuring consistent override and placeholder rules.
'   - DeepWiki Context: Implements the edge-creation rules described in the
'     "Edges", "Labels", and "Styles" documentation.
' ==========================================================================
Private Function ProcessEdgeCreate(ByRef ini As settings, _
                                       ByRef ctx As KnowledgeGraphContext, _
                                       ByRef data As dataRow, _
                                       ByVal edgeId As String, _
                                       ByVal srcId As String, _
                                       ByVal tgtId As String, _
                                       ByVal srcPort As String, _
                                       ByVal tgtPort As String) As Dictionary
    Dim d As Dictionary
    Set d = New Dictionary

    With d
        .Add "id", edgeId
        .Add "source", srcId
        .Add "target", tgtId

        If ini.graph.includeEdgePorts Then
            If Len(srcPort) > 0 Then .Add "tailport", srcPort
            If Len(tgtPort) > 0 Then .Add "headport", tgtPort
        End If

        If Len(data.styleName) > 0 Then .Add "style", data.styleName

        ' === Label Processing - Extra Attributes have highest precedence ===
        ProcessLabel ini, data, d, "label", data.label, ini.graph.includeEdgeLabels
        ProcessLabel ini, data, d, "xlabel", data.xlabel, ini.graph.includeEdgeXLabels
        ProcessLabel ini, data, d, "headlabel", data.headlabel, ini.graph.includeEdgeHeadLabels
        ProcessLabel ini, data, d, "taillabel", data.taillabel, ini.graph.includeEdgeTailLabels
        ProcessLabel ini, data, d, "tooltip", data.tooltip, ini.graph.includeEdgeTooltips
    End With

    ' --- Debug Label ---
    If ini.graph.debug Then MergeLabelField d("id"), d, "debuglabel", "Row: " & data.row

    Set ProcessEdgeCreate = d

End Function

' ==========================================================================
' ROUTINE: ProcessEdgeTagOrphanNodes
'
' PURPOSE:
'   Ensures that any node referenced by an edge is represented in the node
'   dictionary, marking it as an orphan when no explicit node definition
'   exists. Prevents missing-node conditions and supports later resolution
'   when nodes are formally defined.
'
' FUNCTIONAL WORKFLOW:
'   1. NODE-ID NORMALIZATION:
'        - Converts the supplied nodeId to uppercase and trims whitespace
'          to produce a canonical dictionary key.
'
'   2. ORPHAN DETECTION:
'        - Checks whether the normalized node ID exists in the node
'          dictionary.
'
'   3. SENTINEL CREATION:
'        - When the node is not present:
'             o Creates a minimal dictionary containing:
'                   - "id"       - original node identifier.
'                   - "defined"  - False (indicating an orphan placeholder).
'             o Inserts the orphan entry under the normalized key.
'
'   4. DOWNSTREAM RESOLUTION:
'        - Explicit node definitions later replace or update the orphan
'          entry, ensuring correct attribute synthesis and preventing
'          dangling references.
'
' TECHNICAL NOTES:
'   - Orphan tagging is performed for both source and target nodes during
'     edge synthesis.
'   - The "defined" flag allows the pipeline to distinguish between nodes
'     created implicitly by edges and nodes defined explicitly in the input.
'   - DeepWiki Context: Implements the orphan-node handling rules described
'     in the "Nodes" and "Edges" documentation.
' ==========================================================================
Private Sub ProcessEdgeTagOrphanNodes(ByRef nodes As Dictionary, ByVal nodeId As String)

    Dim ucId As String
    ucId = Trim$(UCase$(nodeId))

    If Not nodes.Exists(ucId) Then
        Dim orphan As New Dictionary
        orphan.Add "id", nodeId
        orphan.Add "defined", False ' will be removed later if node is explicitly defined
        nodes.Add ucId, orphan
    End If

End Sub

' ==========================================================================
' ROUTINE: MergeLabelField
'
' PURPOSE:
'   Inserts or updates a structured label field within an attribute dictionary.
'   Normalizes the label value, detects its type (text or HTML-like), strips
'   Graphviz HTML affixes when necessary, and preserves both value and type
'   metadata for downstream rendering and serialization.
'
' FUNCTIONAL WORKFLOW:
'   1. EMPTY-VALUE CHECK:
'        - Ignores updates when newValue is empty, preventing accidental
'          overwrites with blank content.
'
'   2. LABEL-TYPE DETECTION:
'        - Determines whether the incoming label is plain text or HTML-like
'          using GetLabelType.
'
'   3. HTML-AFFIX STRIPPING:
'        - For HTML-like labels, removes the outer "<" and ">" affixes so the
'          stored value reflects the actual label content rather than the
'          Graphviz signal syntax.
'
'   4. FIELD CREATION:
'        - When the field does not exist:
'             o Creates a new dictionary containing:
'                   - "value" - normalized label text.
'                   - "type"  - "text" or "html".
'             o Adds the new field to jsonDict.
'
'   5. FIELD UPDATE:
'        - When the field already exists:
'             o Emits an informational diagnostic indicating that the field
'               is being updated after initial creation.
'             o Replaces both "value" and "type" in the existing dictionary.
'
' TECHNICAL NOTES:
'   - This routine stores labels in a structured form rather than as raw
'     strings, enabling consistent downstream formatting and serialization.
'   - HTML-like detection is based solely on Graphviz label syntax; callers
'     are responsible for ensuring semantic correctness.
'   - DeepWiki Context: Implements the structured-label rules described in
'     the "Labels", "HTML-Like Labels", and "Serialization" documentation.
' ==========================================================================
Private Sub MergeLabelField(ByVal id As String, _
                            ByRef jsonDict As Dictionary, _
                            ByVal fieldName As String, _
                            ByVal newValue As String)

    If Len(newValue) = 0 Then Exit Sub   ' ignore empty updates

    ' text or html?
    Dim lblType As String
    lblType = GetLabelType(newValue)

    ' Remove graphviz html-like label signals as they are not part of the actual label
    Dim lblValue As String
    If lblType = "html" Then
        lblValue = StripAffix(newValue, "<", ">")
    Else
        lblValue = newValue
    End If
        
    ' Field does not exist -> create new dictionary
    If Not jsonDict.Exists(fieldName) Then
        Dim lblDict As New Dictionary
        lblDict.Add "value", lblValue
        lblDict.Add "type", lblType
        jsonDict.Add fieldName, lblDict
        Exit Sub
    End If

    ' Field exists -> update dictionary contents
    Dim existing As Dictionary
    Set existing = jsonDict(fieldName)

    ' Expand the full localized error message
    Dim fullMessage As String
    fullMessage = GetMessage("errormsgNodeUpdatedAfterCreation")
    fullMessage = replace(fullMessage, "{id}", id, 1, 1, vbTextCompare)
    fullMessage = replace(fullMessage, "{field}", fieldName, 1, 1, vbTextCompare)

    EmitMessageSilent fullMessage, esInfo
    
    existing("value") = lblValue
    existing("type") = lblType
End Sub

' ==========================================================================
' FUNCTION: GetLabelType
'
' PURPOSE:
'   Classifies a label as either plain text or HTML-like based on Graphviz
'   label syntax. Provides a normalized type identifier used by downstream
'   routines that store structured label metadata.
'
' FUNCTIONAL WORKFLOW:
'   1. HTML-LIKE DETECTION:
'        - Uses IsLabelHTMLLike to determine whether the label contains
'          Graphviz-style HTML markers (e.g., <...>).
'
'   2. TYPE ASSIGNMENT:
'        - Returns "html" when HTML-like syntax is detected.
'        - Returns "text" otherwise.
'
' TECHNICAL NOTES:
'   - This routine performs no normalization or stripping; callers handle
'     HTML-affix removal when necessary.
'   - Classification is used by MergeLabelField and other structured-label
'     routines to preserve type metadata for serialization and rendering.
'   - DeepWiki Context: Supports the label-type rules described in the
'     "Labels" and "HTML-Like Labels" documentation.
' ==========================================================================
Private Function GetLabelType(label As String) As String
    If IsLabelHTMLLike(label) Then
        GetLabelType = "html"
    Else
        GetLabelType = "text"
    End If
End Function

' ==========================================================================
' FUNCTION: GetTokenEstimate
'
' PURPOSE:
'   Produces a conservative token estimate for a JSON string based on
'   character count, ceiling division, and an 8% safety buffer. Used by
'   upstream routines that need reliable preflight token sizing before
'   chunking or submitting payloads to LLM APIs.
'
' FUNCTIONAL WORKFLOW:
'   1. CHARACTER COUNT:
'        - Measures the total number of characters in the JSON text.
'
'   2. BASE TOKEN ESTIMATE:
'        - Applies a rule-of-thumb conversion (~4 characters per token).
'        - Uses ceiling division to avoid underestimating token usage.
'
'   3. SAFETY BUFFER:
'        - Adds an 8% buffer to account for BPE edge cases, punctuation
'          density, and JSON structural irregularities.
'
' TECHNICAL NOTES:
'   - This routine intentionally errs on the safe side to prevent API
'     truncation or incomplete responses caused by underestimated token
'     counts.
'   - The 4:1 heuristic aligns with common JSON tokenization patterns but
'     is not model-specific; callers should treat the result as a planning
'     estimate rather than an exact tokenizer output.
'   - DeepWiki Context: Supports the token-budgeting guidance described in
'     "Chunking Strategy" and "LLM Payload Preparation".
' ==========================================================================
Public Function GetTokenEstimate(jsonText As String) As Long
    Dim charCount As Long
    Dim estimatedTokens As Long
    Dim safetyBuffer As Long
    
    charCount = Len(jsonText)
    
    ' Rule of thumb: ~4 characters per token (common for JSON)
    ' Round up so we never underestimate
    estimatedTokens = -Int(-charCount / 4)   ' Ceiling division
    
    ' Add a small safety buffer (~5-10%) for BPE edge cases
    safetyBuffer = CLng(estimatedTokens * 0.08)
    
    GetTokenEstimate = estimatedTokens + safetyBuffer
End Function

'=============================================================================
' CreateJsonFiles
'
' Generates one JSON file per view column, using the current runtime settings,
' enabled styles, and graph options. Each JSON file is written to the configured
' output directory (or the workbook path if none is specified). Optionally opens
' the published file after creation.
'
' PARAMETERS
'   firstViewColumn  - Long
'       The first column containing a view definition to process.
'
'   lastViewColumn   - Long
'       The last column containing a view definition to process.
'
'   whitespace       - Long
'       Indentation level for JSON output. Use a negative value for compact
'       (no-whitespace) formatting.
'
' BEHAVIOR
'   o Validates data worksheet and file-output settings.
'   o Builds sanitized filenames for each view.
'   o Generates JSON via GetGraphJson for each view.
'   o Writes UTF-8 (Windows) or text (Mac) output.
'   o Optionally opens the published file based on user settings.
'   o Updates the status bar with progress information.
'
' SIDE EFFECTS
'   o Updates SettingsSheet("ViewNameLabel") during processing.
'   o Writes viewer-ready JSON files to disk.
'
'=============================================================================
Public Sub CreateJsonFiles(ByVal firstViewColumn As Long, ByVal lastViewColumn As Long, whitespace As Long)
    ' Clear the status bar
    ClearStatusBar
    
    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(GetDataWorksheetName())

    If Not WorksheetExists(ini.data.worksheetName) Then
        EmitMessage GetMessage("msgboxNoDataToGraph")
        Exit Sub
    End If

    ' Get file output settings
    Dim output As FileOutput
    output = GetSettingsForFileOutput()

    ' Determine output directory, and build file names
    If output.directory = vbNullString Then
        output.directory = ActiveWorkbook.path
    End If

    ' Validate filename info
    If Not FileLocationProvided(output) Then
        Exit Sub
    End If

    Dim jsonFilename As String
    Dim fullJsonFilename As String
    
    Dim viewColumn As Long
    Dim viewName As String
    
    For viewColumn = firstViewColumn To lastViewColumn
        ' Get the name of the view
        viewName = StylesSheet.Cells.item(ini.styles.headingRow, viewColumn).value
        
        ' Expose the view name so it can be used as data in the graph
        SettingsSheet.Range("ViewNameLabel").value = viewName
        
        ' View name might be referenced in the graph options, so refresh the value
        ini.graph.options = Trim$(SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).value)
        
        ' Build the JSON file name
        jsonFilename = SanitizeFilename(GetFilenameBase(ini, viewColumn) & ".json")
        
        If output.directory = vbNullString Then
            fullJsonFilename = jsonFilename
        Else
            fullJsonFilename = output.directory & Application.pathSeparator & jsonFilename
        End If
        
        ' Cache the names included in the current view
        Dim viewStyles As Dictionary
        Set viewStyles = CacheEnabledStyles(ini, viewColumn)

        ' Generate the JSON for the current view
        Dim graphJson As String
        graphJson = GetGraphJson(ini, viewName, viewStyles, whitespace)
        
        ' Write the JSON to a file
#If Mac Then
        WriteTextToFile graphJson, fullJsonFilename
#Else
        WriteTextToUTF8FileFileWithoutBOM graphJson, fullJsonFilename
#End If
        
        ' Display the published graph?
        If FileExists(fullJsonFilename) Then
            If SettingsSheet.Range("openAfterPublish").value = TOGGLE_YES Then
                SafeFollowHyperlink fullJsonFilename
            End If
            UpdateStatusBarForNSeconds GetLabel("publishKnowledge") & " " & GetMessage("statusbarGraphFilenameIs") & " " & fullJsonFilename, 10
        End If
        
        ' Cleanup
        Set viewStyles = Nothing
    Next viewColumn

    ' Sync up settings with dropdown choice
    SettingsSheet.Range("ViewNameLabel").value = SettingsSheet.Range("ViewName").value

End Sub


