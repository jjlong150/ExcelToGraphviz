Attribute VB_Name = "modWorksheetSettings"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetSettings
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Settings
'
' ROLE:
'   Central configuration hub and UI state manager. Encapsulates the contract
'   between the Settings worksheet and the VBA engine via Named Ranges and
'   strongly-typed UDTs, and drives the simulated tabbed Settings UI.
'
' RESPONSIBILITIES:
'   - Global settings aggregation:
'       • GetSettings: assemble the master settings UDT from all subsystems
'         (Graph, Data, Source, SQL, SVG, Styles, Output, CommandLine, Console)
'
'   - Worksheet layout mapping:
'       • GetSettingsForDataWorksheet, GetSettingsForStylesWorksheet,
'         GetSettingsForSourceWorksheet, GetSettingsForSqlWorksheet,
'         GetSettingsForSvgWorksheet: map logical settings to physical
'         worksheet coordinates via Named Range API and layout globals
'       • GetSettingColNum: translate logical column identifiers into
'         numeric indices
'
'   - SQL integration configuration:
'       • GetSettingsForSqlFields: load SQL field names, placeholders,
'         limits, and advanced behaviors (concatenate, enumerate, iterate)
'
'   - File output configuration:
'       • GetSettingsForFileOutput: resolve output directory, filename
'         prefix, timestamp/options flags, and date/time snapshot
'
'   - Image and output paths:
'       • SelectImageDirectory / SelectOutputDirectory: drive folder pickers
'         and persist choices into SETTINGS_ ranges
'
'   - Typed retrieval helpers:
'       • GetSettingText, GetSettingLong: safe wrappers around Settings
'         ranges with error shielding and optional case normalization
'
' ARCHITECTURAL NOTES:
'   - Implements the SETTINGS_ Named Range API so the Settings UI can be
'     restructured without breaking VBA logic.
'   - Uses UDTs (settings, dataWorksheet, stylesWorksheet, sqlWorksheet,
'     sourceWorksheet, svgWorksheet, FileOutput, sqlFieldName, consoleOptions)
'     as cached snapshots for the rendering and SQL pipelines.
'   - Some functions rely on ActiveSheet context when resolving columns,
'     requiring the appropriate sheet to be active during layout discovery.
'   - Enforces cross-platform behavior (e.g., SQL as Windows-only) via
'     higher-level visibility and feature-gating logic elsewhere in the module.
'
' USAGE:
'   - Call GetSettings(dataWorksheetName) at the start of graph generation,
'     SQL workflows, or diagnostics to obtain a stable configuration snapshot.
'
' RELATED WIKI PAGES:
'   - Settings & Diagnostics Architecture
'   - Named Range API & UDT Mapping
'   - Tabbed Settings Worksheet Design
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: SelectImageDirectory
'
' PURPOSE:
'   Triggers a directory picker dialog to allow the user to define the
'   primary search path for image assets (icons, backgrounds) used in graphs.
'
' TECHNICAL WORKFLOW:
'   1. DIALOG INVOCATION: Calls 'ChooseDirectory' to launch the platform-
'      specific folder picker.
'   2. VALIDATION: Aborts execution if the user cancels the dialog (null string).
'   3. STATE PERSISTENCE: Updates the 'SETTINGS_IMAGE_PATH' named range on
'      the 'Settings' worksheet using the 'SetCellString' utility.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the 'SETTINGS_' Named Range API for persistence.
' ==========================================================================
Public Sub SelectImageDirectory()
    ' Let the user select a directory
    Dim directoryName As String
    directoryName = ChooseDirectory(vbNullString)
    
    If directoryName = vbNullString Then Exit Sub
    
    ' Update the cell with the directory name chosen
    SetCellString SettingsSheet.name, SETTINGS_IMAGE_PATH, directoryName
End Sub

' ==========================================================================
' PROCEDURE: SelectOutputDirectory
'
' PURPOSE:
'   Facilitates the selection of a physical file system path where rendered
'   Graphviz diagrams and exported data will be saved.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXTUAL START: Retrieves the existing path from 'SETTINGS_OUTPUT_DIRECTORY'.
'      If the path is invalid or missing, it resets to a null starting point.
'   2. DIALOG INVOCATION: Launches 'ChooseDirectory', passing the existing
'      path as the initial folder for user convenience.
'   3. VALIDATION: Terminates the procedure if the user cancels the picker.
'   4. STATE PERSISTENCE: Commits the new directory path back to the
'      Settings worksheet using the 'SetCellString' utility.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings / File System.
'   - Contract: Updates the 'SETTINGS_OUTPUT_DIRECTORY' Named Range.
' ==========================================================================
Public Sub SelectOutputDirectory()
    ' Let the user select a directory
    Dim directoryName As String
    
    ' Get the directory currently specified
    directoryName = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    
    If directoryName <> vbNullString Then
        ' Start at a directory which exists
        If Not DirectoryExists(directoryName) Then
            directoryName = vbNullString
        End If
    End If
    
    ' Bring up the directory picker
    directoryName = ChooseDirectory(directoryName)
    
    If directoryName = vbNullString Then Exit Sub   ' Cancel was chosen
        
    ' Update the cell with the directory name chosen
    SetCellString SettingsSheet.name, SETTINGS_OUTPUT_DIRECTORY, directoryName
End Sub

' ==========================================================================
' FUNCTION: GetSettings
'
' PURPOSE:
'   Constructs a comprehensive 'settings' UDT (User Defined Type) by
'   aggregating configuration data from all functional subsystems.
'
' TECHNICAL WORKFLOW:
'   1. UDT ASSEMBLY: Sequentially invokes specialized "GetSettingsFor..."
'      functions to populate each member of the global settings structure.
'   2. LAYER MAPPING: Integrates parameters for Graph rendering, Data/SQL
'      worksheets, SVG post-processing, Style galleries, and Output paths.
'   3. ENGINE CONFIGURATION: Includes low-level CLI and Console settings
'      required for the Graphviz process execution.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Implements the "GetSettings UDT" pattern noted in
'     the 'Settings & Diagnostics' architectural page.
'   - Strategy: Acts as a centralized "Snapshot" of the workbook's state,
'     serving as the primary contract for the rendering pipeline.
' ==========================================================================
Public Function GetSettings(ByVal dataWorksheet As String) As settings
    GetSettings.graph = GetSettingsForGraph()
    GetSettings.data = GetSettingsForDataWorksheet(dataWorksheet)
    GetSettings.source = GetSettingsForSourceWorksheet()
    GetSettings.sql = GetSettingsForSqlWorksheet()
    GetSettings.svg = GetSettingsForSvgWorksheet()
    GetSettings.styles = GetSettingsForStylesWorksheet()
    GetSettings.output = GetSettingsForFileOutput()
    GetSettings.CommandLine = GetSettingsForCommandLine()
    GetSettings.console = GetSettingsForConsole()
End Function

' ==========================================================================
' FUNCTION: GetSettingsForStylesWorksheet
'
' PURPOSE:
'   Initializes the 'stylesWorksheet' UDT by mapping sheet-specific
'   constants to the physical row and column indices on the 'Styles' sheet.
'
' TECHNICAL WORKFLOW:
'   1. ROW MAPPING: Resolves the heading and first data row indices.
'   2. DYNAMIC BOUNDARY: Retrieves the 'lastRow' setting; if set to 0,
'      it dynamically calculates the end of the sheet using 'UsedRange'.
'   3. COLUMN RESOLUTION: Maps logical identifiers (Comment, Style Name,
'      Format, Object Type) to physical columns via 'GetSettingColNum'.
'   4. METADATA CAPTURE: Records user-defined suffixes for subgraph
'      containers (Open/Close markers).
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Part of the "Worksheet Architecture" documentation,
'     ensuring the "Contract" between VBA and the Styles UI is maintained.
' ==========================================================================
Public Function GetSettingsForStylesWorksheet() As stylesWorksheet
    GetSettingsForStylesWorksheet.headingRow = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    GetSettingsForStylesWorksheet.firstRow = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_FIRST))
    
    GetSettingsForStylesWorksheet.lastRow = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_LAST))
    If GetSettingsForStylesWorksheet.lastRow = 0 Then
        With StylesSheet.UsedRange
            GetSettingsForStylesWorksheet.lastRow = .Cells.item(.Cells.count).row
        End With
    End If
    
    GetSettingsForStylesWorksheet.flagColumn = GetSettingColNum(SETTINGS_STYLES_COL_COMMENT)
    GetSettingsForStylesWorksheet.nameColumn = GetSettingColNum(SETTINGS_STYLES_COL_STYLE)
    GetSettingsForStylesWorksheet.formatColumn = GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)
    GetSettingsForStylesWorksheet.typeColumn = GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)
    GetSettingsForStylesWorksheet.firstYesNoColumn = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    GetSettingsForStylesWorksheet.selectedViewColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    GetSettingsForStylesWorksheet.suffixOpen = SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_OPEN).value
    GetSettingsForStylesWorksheet.suffixClose = SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_CLOSE).value
End Function

' ==========================================================================
' FUNCTION: GetSettingColNum
'
' PURPOSE:
'   Translates a logical setting name into a physical Excel column index.
'
' TECHNICAL WORKFLOW:
'   1. INDIRECTION LOOKUP: Reads the value of the 'namedRange' from the
'      'Settings' sheet (which typically contains an Excel column letter).
'   2. COORDINATE RESOLUTION: Appends "1" to the letter to create a valid
'      cell address (e.g., "C1").
'   3. INDEX EXTRACTION: Returns the '.Column' property of that resolved
'      address on the 'ActiveSheet'.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: This is a core component of the "Named Range API"
'     that allows the user to reorder columns without breaking VBA logic.
'   - Constraint: The function relies on the 'ActiveSheet' context, which
'     requires the target worksheet to be active during invocation.
' ==========================================================================
Public Function GetSettingColNum(ByVal namedRange As String) As Long
    GetSettingColNum = ActiveSheet.Range(SettingsSheet.Range(namedRange).value & 1).Column
End Function

' ==========================================================================
' FUNCTION: GetSettingsForDataWorksheet
'
' PURPOSE:
'   Maps the logical configuration of a Data worksheet into a 'dataWorksheet'
'   UDT, establishing the "Contract" for the core parsing engine.
'
' TECHNICAL WORKFLOW:
'   1. IDENTITY ASSIGNMENT: Binds the provided 'worksheetName' to the UDT.
'   2. ROW RESOLUTION: Retrieves header and data row anchors. If the
'      'lastRow' setting is 0, it dynamically calculates the boundary
'      using the worksheet's 'UsedRange'.
'   3. COLUMN MAPPING: Uses 'GetSettingColNum' to resolve physical indices
'      for core attributes (Item, Style, Label, Tooltip, etc.).
'   4. RELATIONSHIP MAPPING: Resolves the 'isRelatedToItemColumn' which
'      defines the 'Related Item' (target) of an edge.
'   5. ALPHA-NUMERIC CACHING: Stores both the numeric index and the
'      raw column letter (AsAlpha) for the 'Graph' display column to
'      facilitate diverse range operations.
'
' TECHNICAL NOTES:
'   - Terminology: Adheres to the 'Item' and 'Related Item' convention
'     specified in the architecture documentation.
'   - Strategy: Decouples the VBA parser from worksheet geometry.
' ==========================================================================
Public Function GetSettingsForDataWorksheet(ByVal worksheetName As String) As dataWorksheet
    GetSettingsForDataWorksheet.worksheetName = worksheetName
    
    GetSettingsForDataWorksheet.headingRow = CLng(SettingsSheet.Range(SETTINGS_DATA_ROW_HEADING))
    GetSettingsForDataWorksheet.firstRow = CLng(SettingsSheet.Range(SETTINGS_DATA_ROW_FIRST))
    GetSettingsForDataWorksheet.lastRow = CLng(SettingsSheet.Range(SETTINGS_DATA_ROW_LAST))
    If GetSettingsForDataWorksheet.lastRow = 0 Then
        With ActiveWorkbook.worksheets.[_Default](worksheetName).UsedRange
            GetSettingsForDataWorksheet.lastRow = .Cells(.Cells.count).row
        End With
    End If

    GetSettingsForDataWorksheet.flagColumn = GetSettingColNum(SETTINGS_DATA_COL_COMMENT)
    GetSettingsForDataWorksheet.styleNameColumn = GetSettingColNum(SETTINGS_DATA_COL_STYLE)
    GetSettingsForDataWorksheet.itemColumn = GetSettingColNum(SETTINGS_DATA_COL_ITEM)
    GetSettingsForDataWorksheet.labelColumn = GetSettingColNum(SETTINGS_DATA_COL_LABEL)
    GetSettingsForDataWorksheet.xLabelColumn = GetSettingColNum(SETTINGS_DATA_COL_LABEL_X)
    GetSettingsForDataWorksheet.tailLabelColumn = GetSettingColNum(SETTINGS_DATA_COL_LABEL_TAIL)
    GetSettingsForDataWorksheet.headLabelColumn = GetSettingColNum(SETTINGS_DATA_COL_LABEL_HEAD)
    GetSettingsForDataWorksheet.tooltipColumn = GetSettingColNum(SETTINGS_DATA_COL_TOOLTIP)
    GetSettingsForDataWorksheet.isRelatedToItemColumn = GetSettingColNum(SETTINGS_DATA_COL_IS_RELATED_TO)
    GetSettingsForDataWorksheet.extraAttributesColumn = GetSettingColNum(SETTINGS_DATA_COL_EXTRA_ATTRIBUTES)
    GetSettingsForDataWorksheet.errorMessageColumn = GetSettingColNum(SETTINGS_DATA_COL_ERROR_MESSAGES)
    GetSettingsForDataWorksheet.graphDisplayColumn = GetSettingColNum(SETTINGS_DATA_COL_GRAPH)
    GetSettingsForDataWorksheet.graphDisplayColumnAsAlpha = SettingsSheet.Range(SETTINGS_DATA_COL_GRAPH).value
End Function

' ==========================================================================
' FUNCTION: GetSettingsForSourceWorksheet
'
' PURPOSE:
'   Initializes the 'sourceWorksheet' UDT by mapping the logical settings
'   for the DOT source viewer to physical worksheet coordinates and parameters.
'
' TECHNICAL WORKFLOW:
'   1. ROW MAPPING: Resolves the 'headingRow' and 'firstRow' anchors from
'      the Settings worksheet.
'   2. COLUMN RESOLUTION: Maps the 'lineNumberColumn' and 'sourceColumn'
'      indices using 'GetSettingColNum'.
'   3. INDENTATION LOGIC: Retrieves the user-defined indentation level and
'      applies boundary clamping (0 to 8 spaces) to ensure valid DOT
'      formatting and UI readability.
'
' TECHNICAL NOTES:
'   - Layer: Settings / Source Viewer.
'   - Contract: Part of the "DOT Source Viewer & Console" architecture.
' ==========================================================================
Public Function GetSettingsForSourceWorksheet() As sourceWorksheet
    GetSettingsForSourceWorksheet.headingRow = CLng(SettingsSheet.Range(SETTINGS_SOURCE_ROW_HEADING))
    GetSettingsForSourceWorksheet.firstRow = CLng(SettingsSheet.Range(SETTINGS_SOURCE_ROW_FIRST))

    GetSettingsForSourceWorksheet.lineNumberColumn = GetSettingColNum(SETTINGS_SOURCE_COL_LINE_NUMBER)
    GetSettingsForSourceWorksheet.sourceColumn = GetSettingColNum(SETTINGS_SOURCE_COL_SOURCE)
    GetSettingsForSourceWorksheet.indent = CLng(SettingsSheet.Range(SETTINGS_SOURCE_INDENT))
    
    If GetSettingsForSourceWorksheet.indent < 0 Then
        GetSettingsForSourceWorksheet.indent = 0
    ElseIf GetSettingsForSourceWorksheet.indent > 8 Then
        GetSettingsForSourceWorksheet.indent = 8
    End If
End Function

' ==========================================================================
' FUNCTION: GetSettingsForSqlWorksheet
'
' PURPOSE:
'   Initializes the 'sqlWorksheet' UDT by mapping the logical SQL settings
'   to physical coordinates on the 'SQL' sheet.
'
' TECHNICAL WORKFLOW:
'   1. ROW RESOLUTION: Retrieves header and data row anchors from the
'      Settings sheet.
'   2. DYNAMIC BOUNDARY: Calculates the 'lastRow' by inspecting the
'      'SqlSheet.UsedRange' to capture all active query definitions.
'   3. COLUMN MAPPING: Resolves physical indices for critical SQL fields
'      (Comment, SQL Statement, Target Excel File, and Status) via
'      'GetSettingColNum'.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Foundational for the "SQL Data Integration"
'     and "SQL Engine" architectural pages.
'   - Platform: These settings drive Windows-only ADO functionality.
' ==========================================================================
Public Function GetSettingsForSqlWorksheet() As sqlWorksheet
    GetSettingsForSqlWorksheet.headingRow = CLng(SettingsSheet.Range(SETTINGS_SQL_ROW_HEADING))
    GetSettingsForSqlWorksheet.firstRow = CLng(SettingsSheet.Range(SETTINGS_SQL_ROW_FIRST))
    With SqlSheet.UsedRange
        GetSettingsForSqlWorksheet.lastRow = .Cells.item(.Cells.count).row
    End With
    GetSettingsForSqlWorksheet.flagColumn = GetSettingColNum(SETTINGS_SQL_COL_COMMENT)
    GetSettingsForSqlWorksheet.sqlStatementColumn = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
    GetSettingsForSqlWorksheet.excelFileColumn = GetSettingColNum(SETTINGS_SQL_COL_EXCEL_FILE)
    GetSettingsForSqlWorksheet.statusColumn = GetSettingColNum(SETTINGS_SQL_COL_STATUS)
End Function

' ==========================================================================
' FUNCTION: GetSettingText
'
' PURPOSE:
'   Retrieves a text value from a specific named range on the Settings
'   worksheet, providing robust error handling and optional case normalization.
'
' TECHNICAL WORKFLOW:
'   1. RANGE LOOKUP: Attempts to read the '.value' from the 'SettingsSheet'
'      using the provided 'settingName'.
'   2. DATA VALIDATION: Checks if the retrieved variant is an Error or
'      Empty; returns a zero-length string if either is true to prevent
'      VBA runtime crashes.
'   3. STRING NORMALIZATION:
'      - If 'makeLower' is TRUE: Trims and converts the string to lowercase.
'      - If 'makeLower' is FALSE: Returns the trimmed string in its
'        original case.
'
' TECHNICAL NOTES:
'   - Strategy: Acts as a "Safety Wrapper" for the Named Range API.
'   - Usage: Preferred over direct range access for core configuration
'     strings to ensure system stability.
' ==========================================================================
Private Function GetSettingText(ByVal settingName As String, Optional ByVal makeLower As Boolean = True) As String
    Dim v As Variant
    v = SettingsSheet.Range(settingName).value

    If IsError(v) Or IsEmpty(v) Then
        GetSettingText = ""
    ElseIf makeLower Then
        GetSettingText = Trim$(LCase$(CStr(v)))
    Else
        GetSettingText = Trim$(CStr(v))
    End If
End Function

' ==========================================================================
' FUNCTION: GetSettingLong
'
' PURPOSE:
'   Retrieves a numeric setting from a named range and safely casts it
'   to a Long integer type.
'
' TECHNICAL WORKFLOW:
'   1. DATA EXTRACTION: Polls the 'SettingsSheet' for the specified
'      'settingName' value.
'   2. ERROR SHIELDING: Validates against 'IsError' or 'IsEmpty' to
'      prevent type mismatch crashes during casting.
'   3. TYPE CONVERSION: Returns 0 if invalid; otherwise, applies 'CLng'
'      to the variant for strict numeric processing.
'
' TECHNICAL NOTES:
'   - Strategy: Provides a stable interface for integer-based configuration
'     constants (e.g., row heights, column indices).
'   - Contract: Part of the centralized Settings retrieval API.
' ==========================================================================
Private Function GetSettingLong(ByVal settingName As String) As String
    Dim v As Variant
    v = SettingsSheet.Range(settingName).value

    If IsError(v) Or IsEmpty(v) Then
        GetSettingLong = 0
    Else
        GetSettingLong = CLng(v)
    End If
End Function

' ==========================================================================
' FUNCTION: GetSettingsForSqlFields
'
' PURPOSE:
'   Initializes the 'sqlFieldName' UDT by mapping logical SQL integration
'   settings to the Project's internal processing logic.
'
' TECHNICAL WORKFLOW:
'   1. CLUSTER HIERARCHY: Retrieves field aliases for multi-level clustering
'      (Labels, Styles, Tooltips) used to infer graph structure from recordsets.
'   2. PLACEHOLDER RESOLUTION: Captures specialized tokens used for dynamic
'      query injection (e.g., %ID%, %LEVEL%).
'   3. PERFORMANCE TUNING:
'      - Applies 'clusterLevelLimit' to prevent stack overflow during recursion.
'      - Sets 'maxConnectionMinutes' to manage the ADO connection pool
'        freshness threshold.
'   4. RECURSIVE PATTERNS: Loads settings for 'TreeQuery' and 'Iterate' modes
'      to handle parent-child traversal and nested query loops.
'   5. DATA TRANSFORMATION: Maps "Concatenate" and "Enumerate" flags used
'      to post-process SQL results into valid DOT attributes.
'
' TECHNICAL NOTES:
'   - Platform: Windows-Only (Supports ADO-based data integration).
'   - DeepWiki Context: Foundational for 'Advanced SQL Patterns' and
'     'SQL Engine & Connection Pooling' architectural pages.
' ==========================================================================
Public Function GetSettingsForSqlFields(ByVal makeLCase As Boolean) As sqlFieldName

    With GetSettingsForSqlFields
        ' Cluster
        .Cluster = GetSettingText(SETTINGS_SQL_FIELD_NAME_CLUSTER, makeLCase)
        .clusterLabel = GetSettingText(SETTINGS_SQL_FIELD_NAME_CLUSTER_LABEL, makeLCase)
        .clusterStyleName = GetSettingText(SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME, makeLCase)
        .clusterAttributes = GetSettingText(SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES, makeLCase)
        .clusterTooltip = GetSettingText(SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP, makeLCase)

        ' Subcluster
        .subcluster = GetSettingText(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER, makeLCase)
        .subclusterLabel = GetSettingText(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_LABEL, makeLCase)
        .subclusterStyleName = GetSettingText(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME, makeLCase)
        .subclusterAttributes = GetSettingText(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES, makeLCase)
        .subclusterTooltip = GetSettingText(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP, makeLCase)

        ' Placeholders (case preserved)
        .clusterPlaceholder = GetSettingText(SETTINGS_SQL_COUNT_PLACEHOLDER_CLUSTER, False)
        .clusterLevelPlaceholder = GetSettingText(SETTINGS_SQL_COUNT_PLACEHOLDER_LEVEL, False)
        .subclusterPlaceholder = GetSettingText(SETTINGS_SQL_COUNT_PLACEHOLDER_SUBCLUSTER, False)
        .recordsetPlaceholder = GetSettingText(SETTINGS_SQL_COUNT_PLACEHOLDER_RECORDSET, False)

        ' Other settings (case preserved)
        .splitLength = GetSettingText(SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH, False)
        .lineEnding = GetSettingText(SETTINGS_SQL_FIELD_NAME_LINE_ENDING, False)
        .filterColumn = GetSettingText(SETTINGS_SQL_COL_FILTER, False)
        .filterValue = GetSettingText(SETTINGS_SQL_FILTER_VALUE, False)
        .treeQuery = GetSettingText(SETTINGS_SQL_FIELD_NAME_TREE_QUERY, False)
        .whereColumn = GetSettingText(SETTINGS_SQL_FIELD_NAME_WHERE_COLUMN, False)
        .whereValue = GetSettingText(SETTINGS_SQL_FIELD_NAME_WHERE_VALUE, False)
        .maxDepth = GetSettingText(SETTINGS_SQL_FIELD_NAME_MAX_DEPTH, False)

        ' Boolean
        .closeConnections = GetSettingBoolean(SETTINGS_SQL_CLOSE_CONNECTIONS)
        
        ' Long
        
        ' Empose a limit on how many nested clusters are allowed
        .clusterLevelLimit = GetSettingLong(SETTINGS_SQL_MAX_CLUSTER_LEVELS)
        If .clusterLevelLimit < 1 Then .clusterLevelLimit = MAX_CLUSTERS
        
        ' We can retry a failed query if the cause was not due to syntax error
        ' If retry is not desired, it can be changed in the settings.
        .retryLimit = GetSettingLong(SETTINGS_SQL_RETRY_LIMIT)
        If .retryLimit <= 1 Then .retryLimit = MAX_RS_OPEN_RETRIES

        ' If the connection is older than the configured freshness threshold (default: 5 minutes),
        ' we treat it as stale. 0 minutes causes new connection with every invocation.
        .maxConnectionMinutes = GetSettingLong(SETTINGS_SQL_MAX_CONNECTION_MINUTES)
        If .maxConnectionMinutes < 0 Then .maxConnectionMinutes = DEFAULT_MAX_CONN_AGE_MINUTES

        ' Flags (case preserved)
        .CreateEdges = GetSettingText(SETTINGS_SQL_FIELD_NAME_CREATE_EDGES, False)
        .CreateRank = GetSettingText(SETTINGS_SQL_FIELD_NAME_CREATE_RANK, False)

        ' Concatenation (case preserved)
        .concatenateSwitch = GetSettingText(SETTINGS_SQL_FIELD_NAME_CONCATENATE_SWITCH, False)
        .concatenateField = GetSettingText(SETTINGS_SQL_FIELD_NAME_CONCATENATE_FIELD, False)
        .concatenateMapTo = GetSettingText(SETTINGS_SQL_FIELD_NAME_CONCATENATE_MAP_TO, False)
        .concatenatePrefix = GetSettingText(SETTINGS_SQL_FIELD_NAME_CONCATENATE_PREFIX, False)
        .concatenateSuffix = GetSettingText(SETTINGS_SQL_FIELD_NAME_CONCATENATE_SUFFIX, False)
        .concatenateSeparator = GetSettingText(SETTINGS_SQL_FIELD_NAME_CONCATENATE_SEPARATOR, False)

        ' Enumeration (case preserved)
        .enumerateSwitch = GetSettingText(SETTINGS_SQL_FIELD_NAME_ENUMERATE_SWITCH, False)
        .enumerateStartAt = GetSettingText(SETTINGS_SQL_FIELD_NAME_ENUMERATE_START_AT, False)
        .enumerateStopAt = GetSettingText(SETTINGS_SQL_FIELD_NAME_ENUMERATE_STOP_AT, False)
        .enumerateStepBy = GetSettingText(SETTINGS_SQL_FIELD_NAME_ENUMERATE_STEP_BY, False)
        .enumeratePlaceholder = GetSettingText(SETTINGS_SQL_FIELD_NAME_ENUMERATE_PLACEHOLDER, False)
        .enumerateMax = GetSettingText(SETTINGS_SQL_FIELD_NAME_ENUMERATE_MAX, False)

        ' Iteration + queries (case preserved)
        .iterate = GetSettingText(SETTINGS_SQL_FIELD_NAME_ITERATE_SWITCH, False)
        .idQuery = GetSettingText(SETTINGS_SQL_FIELD_NAME_ITERATE_ID_QUERY, False)
        .dataQuery = GetSettingText(SETTINGS_SQL_FIELD_NAME_ITERATE_DATA_QUERY, False)
        .idPlaceholder = GetSettingText(SETTINGS_SQL_FIELD_NAME_ITERATE_PLACEHOLDER, False)
    End With
End Function

' ==========================================================================
' FUNCTION: GetSettingsForSvgWorksheet
'
' PURPOSE:
'   Initializes the 'svgWorksheet' UDT by mapping the logical SVG post-processing
'   settings to physical worksheet coordinates.
'
' TECHNICAL WORKFLOW:
'   1. ROW MAPPING: Resolves the 'headingRow' and 'firstRow' anchors using
'      the 'svgLayoutRow' global configuration.
'   2. DYNAMIC BOUNDARY: Determines the 'lastRow' by inspecting the
'      'SvgSheet.UsedRange' to capture all active Find/Replace rules.
'   3. COLUMN RESOLUTION: Maps the 'flagColumn', 'findColumn', and
'      'replaceColumn' indices using the 'svgLayoutColumn' global configuration.
'
' TECHNICAL NOTES:
'   - Layer: Settings / Post-Processing.
'   - DeepWiki Context: Foundational for the "SVG Post-Processing & Animation"
'     architecture, enabling XML-level manipulation of rendered graphs.
' ==========================================================================
Public Function GetSettingsForSvgWorksheet() As svgWorksheet
    GetSettingsForSvgWorksheet.headingRow = svgLayoutRow.headingRow
    GetSettingsForSvgWorksheet.firstRow = svgLayoutRow.firstDataRow
    With SvgSheet.UsedRange
        GetSettingsForSvgWorksheet.lastRow = .Cells.item(.Cells.count).row
    End With
    GetSettingsForSvgWorksheet.flagColumn = svgLayoutColumn.flagColumn
    GetSettingsForSvgWorksheet.findColumn = svgLayoutColumn.findColumn
    GetSettingsForSvgWorksheet.replaceColumn = svgLayoutColumn.replaceColumn
End Function

' ==========================================================================
' FUNCTION: GetSettingsForFileOutput
'
' PURPOSE:
'   Initializes the 'FileOutput' UDT to establish the naming convention and
'   target directory for exported diagrams and data files.
'
' TECHNICAL WORKFLOW:
'   1. PREFERENCE RETRIEVAL: Captures user toggles for appending Graphviz
'      engine options or timestamps to the final filename.
'   2. DIRECTORY RESOLUTION: Retrieves the 'SETTINGS_OUTPUT_DIRECTORY';
'      defaults to 'ActiveWorkbook.path' if no specific folder is configured.
'   3. FILENAME COMPOSITION: Retrieves the base prefix; defaults to the
'      current Workbook's name (minus extension) if the field is empty.
'   4. TEMPORAL STAMPING: Populates the UDT with fresh Date and Time
'      strings via 'GetDate' and 'GetTime' for potential use in the
'      'appendTimeStamp' logic.
'
' TECHNICAL NOTES:
'   - Layer: Settings / File System.
'   - Strategy: Ensures a non-null destination for the "Publishing" pipeline.
' ==========================================================================
Public Function GetSettingsForFileOutput() As FileOutput
    GetSettingsForFileOutput.appendOptions = GetSettingBoolean(SETTINGS_APPEND_OPTIONS)
    GetSettingsForFileOutput.appendTimeStamp = GetSettingBoolean(SETTINGS_APPEND_TIMESTAMP)
    
    GetSettingsForFileOutput.directory = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    If GetSettingsForFileOutput.directory = vbNullString Then
        GetSettingsForFileOutput.directory = ActiveWorkbook.path
    End If
    
    GetSettingsForFileOutput.fileNamePrefix = Trim$(SettingsSheet.Range(SETTINGS_FILE_NAME))
    If GetSettingsForFileOutput.fileNamePrefix = vbNullString Then
        GetSettingsForFileOutput.fileNamePrefix = Mid$(ActiveWorkbook.name, 1, InStr(1, ActiveWorkbook.name, ".") - 1)
    End If
    
    GetSettingsForFileOutput.date = GetDate()
    GetSettingsForFileOutput.time = GetTime()
End Function

' ==========================================================================
' FUNCTION: GetSettingsForConsole
'
' PURPOSE:
'   Initializes the 'consoleOptions' UDT by capturing the active debugging
'   and logging preferences from the Settings worksheet.
'
' TECHNICAL WORKFLOW:
'   1. LOGGING PREFERENCE: Retrieves the 'logToConsole' toggle to determine
'      if stdout/stderr should be captured during rendering.
'   2. RETENTION LOGIC: Retrieves 'appendConsole' to decide if new logs
'      should clear the 'Console' sheet or append to existing entries.
'   3. ENGINE VERBOSITY: Pulls 'graphvizVerbose' to control whether the
'      '-v' flag is passed to the dot engine CLI.
'
' TECHNICAL NOTES:
'   - Layer: Settings / Diagnostics.
'   - DeepWiki Context: Directly supports the "Console Architecture"
'     diagnostic pipeline.
' ==========================================================================
Public Function GetSettingsForConsole() As consoleOptions
    GetSettingsForConsole.logToConsole = GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
    GetSettingsForConsole.appendConsole = GetSettingBoolean(SETTINGS_APPEND_CONSOLE)
    GetSettingsForConsole.graphvizVerbose = GetSettingBoolean(SETTINGS_GRAPHVIZ_VERBOSE)
End Function

' ==========================================================================
' FUNCTION: GetSettingsForGraph
'
' PURPOSE:
'   Constructs a detailed 'graphOptions' UDT by harvesting engine-specific
'   attributes and rendering toggles from the Settings worksheet.
'
' TECHNICAL WORKFLOW:
'   1. ATTRIBUTE AGGREGATION: Retrieves global Graphviz attributes (Strict,
'      Compound, Concentrate, Splines, etc.) and Graph-theory constraints
'      (Rankdir, Overlap, Newrank).
'   2. DATA FILTERING LOGIC: Maps boolean toggles that control which
'      components (Orphan nodes/edges, XLabels, Ports) are included in the
'      final DOT source generation.
'   3. ENGINE RESOLUTION: Invokes 'GetGraphvizEngine' and 'GetImagePath'
'      to resolve the binary environment and asset lookup paths.
'   4. FORMAT VALIDATION: Sets default image types for both File and
'      Worksheet output if the corresponding settings are null.
'   5. SYNTAX CONFIGURATION:
'      - Sets 'command' (graph/digraph) and 'edgeOperator' (--/->) based
'        on the Directed vs. Undirected toggle.
'   6. FEATURE ENHANCEMENT: Automatically enables 'includeTooltip' if the
'      output format is SVG, ensuring interactive metadata is preserved.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: This function is the "Heart" of the Transformation
'     Pipeline, defining the specific Graphviz flavor for every render.
'   - Strategy: Acts as a comprehensive bridge between Excel's high-level
'     UI and the low-level DOT language specifications.
' ==========================================================================
Public Function GetSettingsForGraph() As graphOptions
    GetSettingsForGraph.addStrict = GetSettingBoolean(SETTINGS_GRAPH_STRICT)
    GetSettingsForGraph.blankEdgeLabels = GetSettingBoolean(SETTINGS_BLANK_EDGE_LABELS)
    GetSettingsForGraph.blankNodeLabels = GetSettingBoolean(SETTINGS_BLANK_NODE_LABELS)
    GetSettingsForGraph.center = GetSettingBoolean(SETTINGS_GRAPH_CENTER)
    GetSettingsForGraph.clusterrank = SettingsSheet.Range(SETTINGS_GRAPH_CLUSTER_RANK).value
    GetSettingsForGraph.compound = GetSettingBoolean(SETTINGS_GRAPH_COMPOUND)
    GetSettingsForGraph.concentrate = GetSettingBoolean(SETTINGS_GRAPH_CONCENTRATE)
    GetSettingsForGraph.debug = GetSettingBoolean(SETTINGS_DEBUG)
    GetSettingsForGraph.engine = GetGraphvizEngine()
    GetSettingsForGraph.fileDisposition = Trim$(SettingsSheet.Range(SETTINGS_FILE_DISPOSITION))
    GetSettingsForGraph.forceLabels = GetSettingBoolean(SETTINGS_GRAPH_FORCE_LABELS)
    GetSettingsForGraph.imagePath = GetImagePath()
    GetSettingsForGraph.includeGraphImagePath = GetSettingBoolean(SETTINGS_GRAPH_INCLUDE_IMAGE_PATH)
    GetSettingsForGraph.includeEdgeHeadLabels = GetSettingBoolean(SETTINGS_EDGE_HEAD_LABELS)
    GetSettingsForGraph.includeEdgeLabels = GetSettingBoolean(SETTINGS_EDGE_LABELS)
    GetSettingsForGraph.includeEdgePorts = GetSettingBoolean(SETTINGS_EDGE_PORTS)
    GetSettingsForGraph.includeEdgeTailLabels = GetSettingBoolean(SETTINGS_EDGE_TAIL_LABELS)
    GetSettingsForGraph.includeEdgeXLabels = GetSettingBoolean(SETTINGS_EDGE_XLABELS)
    GetSettingsForGraph.includeExtraAttributes = GetSettingBoolean(SETTINGS_INCLUDE_EXTRA_ATTRIBUTES)
    GetSettingsForGraph.includeNodeLabels = GetSettingBoolean(SETTINGS_NODE_LABELS)
    GetSettingsForGraph.includeNodeXLabels = GetSettingBoolean(SETTINGS_NODE_XLABELS)
    GetSettingsForGraph.includeOrphanEdges = GetSettingBoolean(SETTINGS_RELATIONSHIPS_WITHOUT_NODES)
    GetSettingsForGraph.includeOrphanNodes = GetSettingBoolean(SETTINGS_NODES_WITHOUT_RELATIONSHIPS)
    GetSettingsForGraph.includeStyleFormat = GetSettingBoolean(SETTINGS_INCLUDE_STYLE_FORMAT)
    GetSettingsForGraph.layout = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
    GetSettingsForGraph.layoutDim = SettingsSheet.Range(SETTINGS_GRAPH_DIM).value
    GetSettingsForGraph.layoutDimen = SettingsSheet.Range(SETTINGS_GRAPH_DIMEN).value
    GetSettingsForGraph.mode = SettingsSheet.Range(SETTINGS_GRAPH_MODE).value
    GetSettingsForGraph.model = SettingsSheet.Range(SETTINGS_GRAPH_MODEL).value
    GetSettingsForGraph.newrank = GetSettingBoolean(SETTINGS_GRAPH_NEWRANK)
    GetSettingsForGraph.options = Trim$(SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).value)
    GetSettingsForGraph.ordering = SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value
    GetSettingsForGraph.orientation = GetSettingBoolean(SETTINGS_GRAPH_ORIENTATION)
    GetSettingsForGraph.outputOrder = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value
    GetSettingsForGraph.overlap = SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value
    GetSettingsForGraph.pictureName = SettingsSheet.Range(SETTINGS_PICTURE_NAME).value
    GetSettingsForGraph.postProcessSVG = GetSettingBoolean(SETTINGS_POST_PROCESS_SVG)
    GetSettingsForGraph.rankdir = Trim$(SettingsSheet.Range(SETTINGS_RANKDIR).value)
    GetSettingsForGraph.scaleImage = CLng(SettingsSheet.Range(SETTINGS_SCALE_IMAGE))
    GetSettingsForGraph.smoothing = SettingsSheet.Range(SETTINGS_GRAPH_SMOOTHING).value
    GetSettingsForGraph.splines = SettingsSheet.Range(SETTINGS_SPLINES).value
    GetSettingsForGraph.transparentBackground = GetSettingBoolean(SETTINGS_GRAPH_TRANSPARENT)

    GetSettingsForGraph.imageTypeFile = SettingsSheet.Range(SETTINGS_FILE_FORMAT).value
    If Trim$(GetSettingsForGraph.imageTypeFile) = vbNullString Then
        GetSettingsForGraph.imageTypeFile = SETTINGS_DEFAULT_TO_FILE_TYPE
    End If
    
    GetSettingsForGraph.imageTypeWorksheet = SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value
    If Trim$(GetSettingsForGraph.imageTypeWorksheet) = vbNullString Then
        GetSettingsForGraph.imageTypeWorksheet = GraphSheet.name
    End If
    
    GetSettingsForGraph.imageWorksheet = SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value
    If Trim$(GetSettingsForGraph.imageWorksheet) = vbNullString Then
        GetSettingsForGraph.imageWorksheet = SETTINGS_DEFAULT_TO_WORKSHEET_TYPE
    End If
    
    GetSettingsForGraph.graphType = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value
    If GetSettingsForGraph.graphType = TOGGLE_UNDIRECTED Then
        GetSettingsForGraph.command = "graph"
        GetSettingsForGraph.edgeOperator = "--"
    ElseIf GetSettingsForGraph.graphType = TOGGLE_DIRECTED Then
        GetSettingsForGraph.command = "digraph"
        GetSettingsForGraph.edgeOperator = "->"
    Else
        GetSettingsForGraph.command = "graph"
        GetSettingsForGraph.edgeOperator = "--"
    End If
    
    If LCase$(GetSettingsForGraph.imageTypeWorksheet) = FILETYPE_SVG Then
        GetSettingsForGraph.includeTooltip = True
    End If
    
    If LCase$(GetSettingsForGraph.imageTypeFile) = FILETYPE_SVG Then
        GetSettingsForGraph.includeTooltip = True
    End If
    
End Function

' ==========================================================================
' FUNCTION: GetGraphvizEngine
'
' PURPOSE:
'   Retrieves the identifier for the active Graphviz layout engine.
'
' TECHNICAL WORKFLOW:
'   1. STATIC RETURN: Currently returns the 'SETTINGS_DEFAULT_GRAPHVIZ_ENGINE'
'      constant.
'
' TECHNICAL NOTES:
'   - Strategy: Serves as a centralized hook for engine resolution.
'   - DeepWiki Context: Underpins the "Graphviz Ribbon Tab" logic where
'     engines like dot, neato, and circo are selected.
' ==========================================================================
Public Function GetGraphvizEngine() As String
    GetGraphvizEngine = SETTINGS_DEFAULT_GRAPHVIZ_ENGINE
End Function

' ==========================================================================
' FUNCTION: GetSettingsForCommandLine
'
' PURPOSE:
'   Initializes the 'CommandLine' UDT with the low-level execution parameters
'   required to invoke the external Graphviz binary.
'
' TECHNICAL WORKFLOW:
'   1. PARAMETER CAPTURE: Retrieves user-defined CLI flags (e.g., custom
'      scaling or output overrides) from the 'SETTINGS_COMMAND_LINE_PARAMETERS'
'      named range.
'   2. PATH RESOLUTION: Retrieves the absolute file system path to the
'      Graphviz executable (dot.exe) from 'SETTINGS_GV_PATH'.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Foundational for the "Graphviz Class & Process Execution"
'     page, defining the location and behavior of the external engine.
' ==========================================================================
Public Function GetSettingsForCommandLine() As CommandLine
    GetSettingsForCommandLine.parameters = SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).value
    GetSettingsForCommandLine.GraphvizPath = SettingsSheet.Range(SETTINGS_GV_PATH).value
End Function

' ==========================================================================
' FUNCTION: GetExchangeOptions
'
' PURPOSE:
'   Initializes the 'ExchangeOptions' UDT to define the scope and behavior
'   of the JSON-based data import/export (E2GXF) process.
'
' TECHNICAL WORKFLOW:
'   1. SUBSYSTEM MAPPING: Iteratively configures exchange parameters for
'      Data, SQL, SVG, and Styles worksheets.
'   2. GRANULAR CONTROL: Captures metadata toggles for specific row
'      attributes (Row Numbers, Heights, and Visibility status) to ensure
'      UI fidelity during transport.
'   3. IMPORT STRATEGY: Retrieves the 'action' setting (e.g., Append vs.
'      Overwrite) for each functional area to manage data collisions.
'   4. GLOBAL CONFIGURATION: Includes high-level toggles for worksheet
'      layouts, project metadata, and global graph settings to enable full
'      workbook portability.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Foundational for the "Data Exchange (JSON Import/Export)"
'     architecture page.
'   - Strategy: Facilitates Git-friendly text-based version control of
'     entire ExcelToGraphviz projects.
' ==========================================================================
Public Function GetExchangeOptions() As ExchangeOptions
    GetExchangeOptions.data.include = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET)
    GetExchangeOptions.data.action = Trim$(SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_IMPORT_ACTION))
    GetExchangeOptions.data.row.number = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_ROW)
    GetExchangeOptions.data.row.height = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_HEIGHT)
    GetExchangeOptions.data.row.visible = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_VISIBLE)
    
    GetExchangeOptions.sql.include = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SQL_WORKSHEET)
    GetExchangeOptions.sql.action = Trim$(SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_IMPORT_ACTION))
    GetExchangeOptions.sql.row.number = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_ROW)
    GetExchangeOptions.sql.row.height = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_HEIGHT)
    GetExchangeOptions.sql.row.visible = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_VISIBLE)
    
    GetExchangeOptions.svg.include = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SVG_WORKSHEET)
    GetExchangeOptions.svg.action = Trim$(SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_IMPORT_ACTION))
    GetExchangeOptions.svg.row.number = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_ROW)
    GetExchangeOptions.svg.row.height = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_HEIGHT)
    GetExchangeOptions.svg.row.visible = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_VISIBLE)
    
    GetExchangeOptions.styles.include = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_STYLES_WORKSHEET)
    GetExchangeOptions.styles.action = Trim$(SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION))
    GetExchangeOptions.styles.row.number = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_ROW)
    GetExchangeOptions.styles.row.height = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_HEIGHT)
    GetExchangeOptions.styles.row.visible = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_VISIBLE)
    
    GetExchangeOptions.includeLayouts = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS)
    GetExchangeOptions.includeMetadata = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_METADATA)
    GetExchangeOptions.includeSettings = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_GRAPH_OPTIONS)
End Function

' ==========================================================================
' FUNCTION: GetSettingBoolean
'
' PURPOSE:
'   Normalizes a wide variety of human-readable Excel status strings into
'   a standard VBA Boolean (True/False).
'
' TECHNICAL WORKFLOW:
'   1. DATA EXTRACTION: Retrieves and trims the value from the 'SettingsSheet'
'      using the provided 'cellName' (Named Range API).
'   2. HEURISTIC MAPPING: Uses a 'Select Case' statement to evaluate the
'      value against common "Positive" keywords (YES, TRUE, ON, AUTO, etc.).
'   3. DEFAULTING: Returns 'False' for any unrecognized string or explicit
'      negative values (OFF, NO, HIDE), ensuring system stability.
'
' TECHNICAL NOTES:
'   - Strategy: Centralizes the truth-evaluation logic for the entire
'     project, allowing the UI to use flexible terminology without
'     impacting code logic.
'   - Layer: Settings / Logic.
' ==========================================================================
Public Function GetSettingBoolean(ByVal cellName As String) As Boolean
    
    GetSettingBoolean = False
    
    Select Case UCase$(Trim$(SettingsSheet.Range(cellName).value))
        Case "ON", "YES", "TRUE", "AUTO", "SHOW", "INCLUDE", "DEFAULT"
            GetSettingBoolean = True
        Case Else
            GetSettingBoolean = False
    End Select
    
End Function

' ==========================================================================
' PROCEDURE: DisplayTabRows
'
' PURPOSE:
'   Dynamically controls the visibility of a specified row range on the
'   Settings worksheet to simulate a "Tabbed" user interface.
'
' TECHNICAL WORKFLOW:
'   1. ITERATION: Loops through the worksheet rows starting from 'rowFrom'
'      through to 'rowTo'.
'   2. STATE APPLICATION: Sets the '.Hidden' property of each row to the
'      inverse of the 'isVisible' parameter.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - DeepWiki Context: Implements the "Settings & Diagnostics" UI logic
'     that allows users to navigate complex configurations via a
'     simulated tabbed view.
' ==========================================================================
Public Sub DisplayTabRows(ByVal isVisible As Boolean, ByVal rowFrom As Long, ByVal rowTo As Long)
    Dim row As Long
    For row = rowFrom To rowTo
        SettingsSheet.rows.item(row).Hidden = Not isVisible
    Next row
End Sub

' ==========================================================================
' PROCEDURE: DisplayGraphOptions
'
' PURPOSE:
'   Toggles the visibility of the "Graph Options" configuration section
'   and its associated UI tab indicators.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Dynamically calculates the row range using
'      the 'SETTINGS_IMAGE_PATH' and 'SETTINGS_PICTURE_NAME' named ranges
'      as anchors (Contract API).
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to show or hide the
'      configuration fields.
'   3. TAB STATE MANAGEMENT: Swaps the visibility of the "Enabled" and
'      "Disabled" graphical shapes to provide visual feedback of the
'      active tab.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the Named Range API to identify start/end rows.
' ==========================================================================
Public Sub DisplayGraphOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_IMAGE_PATH).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_PICTURE_NAME).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabGraphOptions").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabGraphOptions").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplayCmdLineOptions
'
' PURPOSE:
'   Toggles the visibility of the "Command Line Options" section and its
'   associated UI tab indicators on the Settings worksheet.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using the
'      'SETTINGS_COMMAND_LINE_PARAMETERS' and 'SETTINGS_GV_PATH'
'      named ranges as anchors (Contract API).
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to expand or collapse
'      the configuration fields.
'   3. TAB STATE MANAGEMENT: Toggles the visibility of the "Enabled"
'      and "Disabled" graphical shapes to reflect the active tab state.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the Named Range API to identify the target
'     CLI configuration block.
' ==========================================================================
Public Sub DisplayCmdLineOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_GV_PATH).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabCmdLineOptions").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabCmdLineOptions").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplayStylesOptions
'
' PURPOSE:
'   Toggles the visibility of the "Styles Worksheet" configuration section
'   and its associated UI tab indicators on the Settings sheet.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using the
'      'SETTINGS_STYLES_COL_COMMENT' and 'SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW'
'      named ranges as anchors (Contract API).
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to show or hide the
'      Style Gallery configuration fields.
'   3. TAB STATE MANAGEMENT: Swaps the visibility of the "Enabled" and
'      "Disabled" graphical shapes to provide visual feedback for the
'      active UI tab.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the Named Range API to maintain a "tabbed"
'     interface within a standard worksheet.
' ==========================================================================
Public Sub DisplayStylesOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_STYLES_COL_COMMENT).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabStylesWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabStylesWorksheet").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplayDataOptions
'
' PURPOSE:
'   Toggles the visibility of the "Data Worksheet" configuration section
'   and its associated UI tab indicators on the Settings sheet.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using the
'      'SETTINGS_DATA_COL_COMMENT' and 'SETTINGS_DATA_COL_GRAPH'
'      named ranges as anchors (Contract API).
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to show or hide the
'      Data sheet structural configuration fields.
'   3. TAB STATE MANAGEMENT: Swaps the visibility of the "Enabled" and
'      "Disabled" graphical shapes to provide visual feedback for the
'      active UI tab.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the Named Range API to maintain the decoupling
'     between worksheet layout and UI logic.
' ==========================================================================
Public Sub DisplayDataOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_DATA_COL_COMMENT).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_DATA_COL_GRAPH).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabDataWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabDataWorksheet").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplaySourceOptions
'
' PURPOSE:
'   Toggles the visibility of the "Source Worksheet" configuration section
'   and its associated UI tab indicators on the Settings sheet.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using the
'      'SETTINGS_SOURCE_ROW_HEADING' and 'SETTINGS_SOURCE_INDENT'
'      named ranges as anchors (Contract API).
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to show or hide the
'      Source Viewer configuration fields.
'   3. TAB STATE MANAGEMENT: Swaps the visibility of the "Enabled" and
'      "Disabled" graphical shapes to provide visual feedback for the
'      active UI tab.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Adheres to the Named Range API for UI state management.
' ==========================================================================
Public Sub DisplaySourceOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_SOURCE_ROW_HEADING).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_SOURCE_INDENT).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabSourceWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabSourceWorksheet").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplaySqlOptions
'
' PURPOSE:
'   Toggles the visibility of the "SQL Worksheet" configuration section,
'   implementing platform-specific restrictions for macOS.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using the
'      'SETTINGS_SQL_COL_COMMENT' and 'SETTINGS_SQL_FIELD_NAME_CONCATENATE_SEPARATOR'
'      named ranges as anchors (Contract API).
'   2. MAC RESTRICTION (#If Mac): Forces the entire section and its tab
'      indicators to be hidden, as SQL features (ADO) are Windows-only.
'   3. WINDOWS EXECUTION (#Else): Invokes 'DisplayTabRows' and toggles the
'      visibility of the "Enabled"/"Disabled" graphical tab shapes.
'
' TECHNICAL NOTES:
'   - Platform: Windows-Only feature. Explicitly hidden on macOS to prevent
'     user confusion regarding ADO availability.
'   - DeepWiki Context: Reflects the "Windows-Only Restriction" noted in
'     the SQL Data Integration architectural page.
' ==========================================================================
Public Sub DisplaySqlOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_SQL_COL_COMMENT).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_CONCATENATE_SEPARATOR).row + 1
#If Mac Then
    DisplayTabRows False, rowFrom, rowTo
    SettingsSheet.Shapes.Range("enabledTabSqlWorksheet").visible = False
    SettingsSheet.Shapes.Range("disabledTabSqlWorksheet").visible = False
#Else
    DisplayTabRows isVisible, rowFrom, rowTo
    SettingsSheet.Shapes.Range("enabledTabSqlWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabSqlWorksheet").visible = Not isVisible
#End If
End Sub

' ==========================================================================
' PROCEDURE: DisplayGraphvizTab
'
' PURPOSE:
'   Toggles the visibility of the "Graphviz Tab" configuration section
'   on the Settings worksheet, maintaining the simulated tabbed UI.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using
'      'SETTINGS_TAB_GRAPHVIZ' as the top anchor and
'      'SETTINGS_TAB_SOURCE' as the bottom boundary (minus buffer).
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to expand or collapse
'      the configuration fields.
'   3. TAB STATE MANAGEMENT: Toggles the visibility of the "Enabled"
'      and "Disabled" graphical shapes to reflect the active tab state.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the Named Range API to identify the target
'     ribbon-tab configuration block.
' ==========================================================================
Public Sub DisplayGraphvizTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_TAB_GRAPHVIZ).row
    rowTo = SettingsSheet.Range(SETTINGS_TAB_SOURCE).row - 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabGraphvizTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabGraphvizTab").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplaySourceTab
'
' PURPOSE:
'   Toggles the visibility of the configuration rows associated with the
'   "Source" Ribbon tab interface on the Settings worksheet.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Dynamically calculates the row range using
'      the 'SETTINGS_TAB_SOURCE' and 'SETTINGS_EXT_TAB_NAME' named ranges
'      as the vertical anchors.
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to show or hide the
'      specific Ribbon-state configuration fields.
'   3. TAB STATE MANAGEMENT: Updates the visibility of the "Enabled" and
'      "Disabled" UI shapes to provide clear visual state feedback.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Adheres to the Named Range API "Contract" to allow for
'     flexible worksheet restructuring.
' ==========================================================================
Public Sub DisplaySourceTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_TAB_SOURCE).row
    rowTo = SettingsSheet.Range(SETTINGS_EXT_TAB_NAME).row - 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabSourceTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabSourceTab").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplayExtensionsTab
'
' PURPOSE:
'   Toggles the visibility of the "Extensions Tab" configuration rows and
'   associated graphical UI indicators on the Settings worksheet.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the dynamic row range using
'      'SETTINGS_EXT_TAB_NAME' and 'SETTINGS_TAB_EXCHANGE' as anchors.
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to expand or collapse the
'      relevant configuration fields.
'   3. TAB STATE MANAGEMENT: Swaps the visibility of the "Enabled" and
'      "Disabled" graphical shapes to reflect the active tab selection.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Relies on the Named Range API to maintain the simulated
'     tabbed interface logic.
' ==========================================================================
Public Sub DisplayExtensionsTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
  
    rowFrom = SettingsSheet.Range(SETTINGS_EXT_TAB_NAME).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_TAB_EXCHANGE).row - 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabExtensionsTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabExtensionsTab").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: DisplayExchangeTab
'
' PURPOSE:
'   Toggles the visibility of the "Exchange Tab" (JSON Import/Export)
'   configuration section and its associated graphical UI indicators.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the row range using the
'      "SettingsExchangeTab" anchor and the 'SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS'
'      named range as the vertical boundaries.
'   2. UI TRANSITION: Invokes 'DisplayTabRows' to expand or collapse the
'      data exchange configuration fields.
'   3. TAB STATE MANAGEMENT: Updates the visibility of the "Enabled" and
'      "Disabled" shapes to provide visual feedback for the active tab.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - DeepWiki Context: Controls the visibility of the configuration engine
'     for the "Data Exchange (JSON Import/Export)" subsystem.
' ==========================================================================
Public Sub DisplayExchangeTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range("SettingsExchangeTab").row - 1
    rowTo = SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabExchangeTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabExchangeTab").visible = Not isVisible
End Sub

' ==========================================================================
' PROCEDURE: TabSelectGraphOptions
'
' PURPOSE:
'   Activates the "Graph Options" view within the Settings worksheet's
'   simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to hide the bulk row
'      hiding/unhiding process from the user.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayGraphOptions' to TRUE.
'      - Explicitly sets all other tab display procedures to FALSE.
'   3. FOCUS MANAGEMENT: Selects the 'SETTINGS_IMAGE_PATH' range to orient
'      the user to the start of the configuration block.
'   4. REFRESH: Re-enables 'ScreenUpdating' to reveal the newly focused tab.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Centralizes the "Exclusive Visibility" logic for the
'     simulated tab system.
' ==========================================================================
Public Sub TabSelectGraphOptions()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions True
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_IMAGE_PATH).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectCmdLineOptions
'
' PURPOSE:
'   Activates the "Command Line Options" view within the Settings worksheet's
'   simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to ensure a smooth visual
'      transition during bulk row state changes.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayCmdLineOptions' to TRUE.
'      - Explicitly hides all other configuration blocks by calling their
'        respective display procedures with FALSE.
'   3. FOCUS MANAGEMENT: Shifts the selection to 'SETTINGS_COMMAND_LINE_PARAMETERS'
'      to land the user at the primary input field.
'   4. REFRESH: Re-enables 'ScreenUpdating' to commit the visual state.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Contract: Adheres to the "Exclusive Visibility" pattern for
'     simulated tab management.
' ==========================================================================
Public Sub TabSelectCmdLineOptions()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions True
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectStylesWorksheet
'
' PURPOSE:
'   Activates the "Styles Worksheet" configuration view within the
'   Settings worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform bulk row
'      manipulation without visual flickering.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayStylesOptions' to TRUE.
'      - Forces all other configuration blocks (Graph, SQL, Data, etc.)
'        to FALSE to maintain exclusive visibility.
'   3. FOCUS MANAGEMENT: Shifts active selection to 'SETTINGS_STYLES_COL_COMMENT'
'      to anchor the user's view to the Style Gallery schema settings.
'   4. REFRESH: Restores 'ScreenUpdating' to finalize the UI transition.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Implements the "Exclusive Tab Selection" pattern for the
'     Settings UI.
' ==========================================================================
Public Sub TabSelectStylesWorksheet()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions True
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_STYLES_COL_COMMENT).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectDataWorksheet
'
' PURPOSE:
'   Activates the "Data Worksheet" configuration view within the Settings
'   worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform the multi-row
'      visibility transition seamlessly.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayDataOptions' to TRUE.
'      - Invokes all other tab display procedures with FALSE to ensure
'        exclusive visibility of the Data configuration block.
'   3. FOCUS MANAGEMENT: Shifts active selection to 'SETTINGS_DATA_COL_COMMENT'
'      to align the user's view with the Data sheet structural settings.
'   4. REFRESH: Restores 'ScreenUpdating' to finalize the layout change.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Centralizes the "Exclusive Visibility" logic for managing
'     complex sheet-based configuration UI.
' ==========================================================================
Public Sub TabSelectDataWorksheet()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions True
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_DATA_COL_COMMENT).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectSourceWorksheet
'
' PURPOSE:
'   Activates the "Source Worksheet" configuration view within the Settings
'   worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform the multi-row
'      visibility transition without visual flicker.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplaySourceOptions' to TRUE.
'      - Invokes all other configuration display procedures with FALSE to
'        ensure exclusive visibility of the Source Viewer settings.
'   3. FOCUS MANAGEMENT: Selects 'SETTINGS_SOURCE_COL_LINE_NUMBER' to
'      orient the user to the start of the Source sheet structural settings.
'   4. REFRESH: Restores 'ScreenUpdating' to finalize the UI layout.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Implements the "Exclusive Visibility" pattern for the
'     simulated tab system on the Settings worksheet.
' ==========================================================================
Public Sub TabSelectSourceWorksheet()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions True
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_SOURCE_COL_LINE_NUMBER).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectSqlWorksheet
'
' PURPOSE:
'   Activates the "SQL Worksheet" configuration view within the Settings
'   worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform the multi-row
'      visibility transition without visual flicker.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplaySqlOptions' to TRUE (Note: This procedure includes
'        internal logic to force FALSE on macOS).
'      - Invokes all other configuration display procedures with FALSE to
'        ensure exclusive visibility of the SQL settings.
'   3. FOCUS MANAGEMENT: Selects 'SETTINGS_SQL_COL_COMMENT' to orient the
'      user to the start of the SQL-specific configuration block.
'   4. REFRESH: Restores 'ScreenUpdating' to finalize the UI layout.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Implements the "Exclusive Visibility" pattern while
'     respecting the Windows-only availability of SQL features.
' ==========================================================================
Public Sub TabSelectSqlWorksheet()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions True
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_SQL_COL_COMMENT).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectGraphvizTab
'
' PURPOSE:
'   Activates the "Graphviz Tab" configuration view within the Settings
'   worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform the multi-row
'      visibility transition without visual flicker.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayGraphvizTab' to TRUE to reveal Ribbon-state settings.
'      - Invokes all other configuration display procedures with FALSE to
'        ensure exclusive visibility.
'   3. FOCUS MANAGEMENT: Selects 'SETTINGS_OUTPUT_DIRECTORY' to orient the
'      user to the primary configuration field for this section.
'   4. REFRESH: Restores 'ScreenUpdating' to finalize the UI layout.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Part of the "Exclusive Visibility" pattern used to manage
'     the extensive configuration options on a single worksheet.
' ==========================================================================
Public Sub TabSelectGraphvizTab()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab True
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectSourceTab
'
' PURPOSE:
'   Activates the "Source Tab" configuration view within the Settings
'   worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform the multi-row
'      visibility transition without visual flicker.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplaySourceTab' to TRUE.
'      - Hides all other configuration blocks by calling their respective
'        display procedures with FALSE.
'   3. FOCUS MANAGEMENT: Selects the "SourceWeb1Text" range to anchor
'      the user's view to the Source Ribbon configuration settings.
'   4. REFRESH: Restores 'ScreenUpdating' to commit the visual layout.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Implements the "Exclusive Visibility" pattern for
'     navigating complex Ribbon-state settings on the Settings sheet.
' ==========================================================================
Public Sub TabSelectSourceTab()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab True
    DisplayExtensionsTab False
    DisplayExchangeTab False
    
    SettingsSheet.Range("SourceWeb1Text").Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectExtensionsTab
'
' PURPOSE:
'   Activates the "Extensions Tab" configuration view within the Settings
'   worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to perform the multi-row
'      visibility transition without visual flicker.
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayExtensionsTab' to TRUE.
'      - Hides all other configuration blocks by calling their respective
'        display procedures with FALSE.
'   3. FOCUS MANAGEMENT: Selects the 'SETTINGS_EXT_TAB_NAME' range to anchor
'      the user's view to the Extensions Ribbon configuration settings.
'   4. REFRESH: Restores 'ScreenUpdating' to commit the visual layout.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Implements the "Exclusive Visibility" pattern for
'     navigating extension-specific Ribbon settings.
' ==========================================================================
Public Sub TabSelectExtensionsTab()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab True
    DisplayExchangeTab False
    
    SettingsSheet.Range(SETTINGS_EXT_TAB_NAME).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: TabSelectExchangeTab
'
' PURPOSE:
'   Activates the "Exchange Tab" (JSON Import/Export) configuration view
'   within the Settings worksheet's simulated tabbed interface.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to hide the bulk row
'      manipulation required to swap "tabs."
'   2. TAB ORCHESTRATION:
'      - Sets 'DisplayExchangeTab' to TRUE.
'      - Invokes all other configuration block display procedures with
'        FALSE to ensure exclusive visibility.
'   3. FOCUS MANAGEMENT: Selects 'SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET'
'      to align the user's view with the start of the Exchange configuration.
'   4. REFRESH: Restores 'ScreenUpdating' to finalize the UI layout.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings.
'   - Strategy: Centralizes the "Exclusive Visibility" logic for the
'     JSON E2GXF data exchange configuration subsystem.
' ==========================================================================
Public Sub TabSelectExchangeTab()
    Application.ScreenUpdating = False
    
    DisplayGraphOptions False
    DisplayCmdLineOptions False
    DisplayStylesOptions False
    DisplayDataOptions False
    DisplaySourceOptions False
    DisplaySqlOptions False
    DisplayGraphvizTab False
    DisplaySourceTab False
    DisplayExtensionsTab False
    DisplayExchangeTab True
    
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET).Select
    
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: ShowOrHideWorksheets
'
' PURPOSE:
'   Synchronizes the visibility of all auxiliary and diagnostic worksheets
'   with the toggles defined on the Settings sheet.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Invokes 'OptimizeCode_Begin' to suppress screen
'      flicker during bulk visibility changes.
'   2. STATE EVALUATION: Iterates through all system worksheets (Console,
'      Diagnostics, Help, Locales, etc.), comparing their dedicated named
'      range settings against the 'TOGGLE_SHOW' constant.
'   3. MAC RESTRICTION (#If Mac): Forces the 'SQL' worksheet to remain hidden
'      regardless of setting, as ADO-based SQL features are Windows-only.
'   4. REFRESH: Invokes 'OptimizeCode_End' to restore normal Excel operation.
'
' TECHNICAL NOTES:
'   - Layer: UI / Settings Management.
'   - Contract: Adheres to the "Windows-Only Restriction" for SQL features
'     specified in the architectural documentation.
' ==========================================================================
Public Sub ShowOrHideWorksheets()
    OptimizeCode_Begin
    
    AboutSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_ABOUT).value = TOGGLE_SHOW
    ConsoleSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_CONSOLE).value = TOGGLE_SHOW
    DiagnosticsSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS).value = TOGGLE_SHOW
    HelpAttributesSheet.visible = SettingsSheet.Range(SETTINGS_HELP_ATTRIBUTES).value = TOGGLE_SHOW
    HelpColorsSheet.visible = SettingsSheet.Range(SETTINGS_HELP_COLORS).value = TOGGLE_SHOW
    HelpShapesSheet.visible = SettingsSheet.Range(SETTINGS_HELP_SHAPES).value = TOGGLE_SHOW
    ListsSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LISTS).value = TOGGLE_SHOW
    LocaleDeDeSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_DE_DE).value = TOGGLE_SHOW
    LocaleEnGbSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_EN_GB).value = TOGGLE_SHOW
    LocaleEnUsSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_EN_US).value = TOGGLE_SHOW
    LocaleFrFrSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_FR_FR).value = TOGGLE_SHOW
    LocaleItItSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_IT_IT).value = TOGGLE_SHOW
    LocalePlPlSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_PL_PL).value = TOGGLE_SHOW
    SettingsSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).value = TOGGLE_SHOW
    SourceSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SOURCE).value = TOGGLE_SHOW
#If Mac Then
    SqlSheet.visible = False
#Else
    SqlSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SQL).value = TOGGLE_SHOW
#End If
    StyleDesignerSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).value = TOGGLE_SHOW
    StylesSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLES).value = TOGGLE_SHOW
    SvgSheet.visible = SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SVG).value = TOGGLE_SHOW
    
    OptimizeCode_End
End Sub

