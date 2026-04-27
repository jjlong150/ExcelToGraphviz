Attribute VB_Name = "modWorksheetSQL"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetSQL
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / SQL
'
' ROLE:
'   Orchestrate the full SQL-driven data-generation pipeline. Translates
'   relational queries, pseudo-SQL commands, and automation directives into
'   hierarchical Graphviz-ready structures. Acts as the execution engine for
'   all SQL workflows, including clustering, enumeration, recursion, and
'   batch publishing.
'
' RESPONSIBILITIES:
'   - SQL execution lifecycle:
'       • RunSQL: primary dispatcher for SQL commands, SELECT queries,
'         placeholder substitution, environment setup, and cleanup
'       • PassesFilter / GetExcelFilePath: determine execution eligibility
'         and resolve multi-source file paths
'
'   - Algorithmic engines:
'       • Iterative Search (ID -> Data)
'       • Recursive Search (parent/child traversal with cycle detection)
'       • Sequential Edges (A -> B -> C)
'       • Enumeration Mode (mathematical range expansion)
'       • N-Level Clustering (CLUSTER1…CLUSTERn) with dynamic sorting
'
'   - Publishing & automation:
'       • Publish, PublishViews, PublishAllViews, PublishAsDirectedGraph,
'         PublishAsUndirectedGraph, PublishAllViewsAsDirectedGraph,
'         PublishAllViewsAsUndirectedGraph
'       • SQL-driven headless rendering and file-naming overrides
'
'   - Worksheet integration:
'       • Populate Data worksheet rows using dataWorksheet UDT
'       • Map SQL worksheet geometry using sqlWorksheet UDT
'       • Load SQL field names, placeholders, and limits via sqlFieldName UDT
'
'   - Stability & resilience:
'       • Connection pooling, retry logic, and stale-connection detection
'       • Null-safe field accessors and semantic error classification
'       • SleepMilliseconds throttling to avoid COM collisions
'
' ARCHITECTURAL NOTES:
'   - SQL subsystem is Windows-only due to ADO dependency; macOS execution
'     paths are gated at higher layers.
'   - Uses Named Range API to remain independent of worksheet geometry.
'   - Integrates with Style Designer (views), Data sheet (record placement),
'     and Graphviz rendering pipeline for automated publishing.
'   - Supports pseudo-SQL commands for automation, environment control,
'     placeholder injection, and batch rendering.
'
' VERSION NOTES:
'   - v6.0.00 (May 14, 2023):
'       • Added SQL clustering and subclustering support (CLUSTER / SUBCLUSTER fields)
'       • Added full pseudo-SQL automation suite:
'           RESET, PREVIEW, PREVIEW AS DIRECTED/UNDIRECTED,
'           PUBLISH, PUBLISH AS DIRECTED/UNDIRECTED,
'           PUBLISH ALL VIEWS, PUBLISH ALL VIEWS AS DIRECTED/UNDIRECTED
'       • Added SQL filtering support (filter columns to control execution)
'       • Added SQL label-splitting logic for multi-line labels with alignment options
'       • Added new SQL sample workbooks demonstrating clustering and automation
'
'   - v7.0.0 (Dec 4, 2024):
'       • Added SQL pop-up editor for large SQL statements
'       • Added Copy-to-Clipboard for SQL statements
'       • Added SET DATA FILE pseudo-SQL command
'       • Added support for CLUSTER LABEL and SUBCLUSTER LABEL in SQL
'       • Added RunSQLAsExtension utility for running SQL from the Extension tab
'
'   - v7.2.0 (Mar 14, 2025):
'       • Added support for recursive SQL queries (base case + recursive member)
'       • Added four new SQL-related settings for recursive keyword customization
'
'   - v8.0.0 (Aug 27, 2025):
'       • Added SQL connection pooling to mitigate slow ADO connections
'       • Added default data directory + default workbook support
'       • Added CREATE EDGES syntax for automatic A->B->C edge generation
'       • Added CREATE RANK syntax for rank-group subgraphs
'       • Added SQL editor access button (`[...]`)
'
'   - v10.0.0 (Jan 23, 2026):
'       • Added support for Microsoft Access (.accdb/.mdb) SQL data sources
'       • Added SQL enumeration mode (range-based result generation)
'       • Added iterative query-set execution with dynamic placeholder substitution
'       • Added SQL error logging to external log file
'       • Added ADO hardening and reliability improvements
'       • Added environment documentation to SQL log-to-file feature
'       • Fixed breaking issue: clustering now groups by CLUSTER (not CLUSTER LABEL)
'
'   - v10.1.0 (Feb 9, 2026):
'       • Added Concatenation Mode for iterative SQL queries
'       • Added floating action buttons for SQL rows (edit/run/status)
'
'   - v10.2.0 (Feb 27, 2026):
'       • Added SQL placeholder substitution via SET PLACEHOLDER name = value
'       • Added filename sanitization for SQL-driven exports
'
'   - v10.3.0 (Apr 3, 2026):
'       • Added support for n-level SQL clustering (CLUSTER1, CLUSTER2, …)
'       • Added {label} placeholder for cluster label formatting
'       • Updated format-string parsing for HTML-like syntax
'
' USAGE:
'   - Invoke RunSQL to execute all SQL rows or a specific row.
'   - Use SQL_PUBLISH*, SQL_PREVIEW*, and SQL_SET_* commands to drive
'     automation, publishing, and environment configuration.
'
' RELATED WIKI PAGES:
'   - SQL Engine & Connection Pooling
'   - Advanced SQL Patterns (Iterate, TreeQuery, Enumeration)
'   - N-Level Clustering Architecture
'   - SQL -> Graphviz Transformation Pipeline
' =============================================================================

Option Explicit

' ==========================================================================
' SECTION: DATA STRUCTURES & CONFIGURATION
' ==========================================================================

' Loop safety and timing constants
Private Const RETRY_DELAY_MS As Long = 100
Private Const SQL_DELAY_MS As Long = 20
Private Const DEFAULT_MAX_RECURSION_DEPTH As Long = 100
Private Const LOOP_MAX_STEPS As Long = 10000 ' Put an upper limit on DO loop to prevent infinite loops

''
' ENUMERATION PARAMETERS: Configures range-based result generation.
' Supports the 'Enumeration Mode' where results are generated mathematically
' rather than from an underlying table.
'
Private Type EnumerateParameters
    Enabled As Boolean
    startAt As Long
    stopAt As Long
    stepBy As Long
    max As Long
    count As Long
End Type

''
' SQL CONTEXT: The primary state-container for a SQL execution cycle.
' 1. dataLayout: Maps the destination 'Data' worksheet columns.
' 2. fields: Maps the expected SQL field names.
' 3. headings: Stores worksheet header text for validation.
' 4. sqlLayout: Maps the source 'SQL' worksheet columns.
' 5. loop: Stores active enumeration settings.
'
Public Type sqlContext
    dataLayout As dataWorksheet
    fields As sqlFieldName
    headings As DataWorksheetHeadings
    sqlLayout As sqlWorksheet
    loop As EnumerateParameters
End Type

''
' CONCAT SETTINGS: Controls the collapsing of multiple detail rows.
' Used to merge related data points into a single multi-line Node or Edge label.
'
Private Type ConcatSettings
    Enabled     As Boolean
    ConcatField As String
    TargetField As String
    prefix      As String
    suffix      As String
    separator   As String
End Type

' ==========================================================================
' SECTION: BUTTON ACTIONS
' ==========================================================================

' ==========================================================================
' PROCEDURE: RunSQL
' PURPOSE:
'   The primary entry point for executing SQL workflows. Iterates through
'   the 'SQL' worksheet to process queries, commands, and automation tasks.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT INITIALIZATION: Loads UDTs for Data/SQL layouts, headings,
'      and field mappings to ensure consistent data placement.
'   2. ENVIRONMENT LOCK: Toggles 'Run Mode' to Manual to prevent redundant
'      graph redraws during batch data population.
'   3. COMMAND DISPATCHER (The ElseIf Ladder):
'      - Environmental: SET_DATA_FILE, ENABLE_LOGGING, RESET.
'      - Substitution: SET_PLACEHOLDER (manages dynamic token replacement).
'      - Automation: SQL_PUBLISH, SQL_PREVIEW (triggers "headless" rendering).
'   4. QUERY EXECUTION:
'      - Retrieves pooled connections via 'getConnection'.
'      - Applies Placeholder substitutions to the raw SQL string.
'      - Routes SELECT statements to 'executeSQL' for processing.
'   5. STABILITY & CLEANUP:
'      - Implements 'SleepMilliseconds' delays to prevent COM collisions.
'      - Finalizes connection pooling and restores 'Run Mode' state.
'
' PARAMETERS:
'   - row [Optional Long]: If 0, processes the entire sheet; if specified,
'     runs only that specific row (allowing for targeted execution).
' ==========================================================================
Public Sub RunSQL(Optional ByVal row As Long = 0)
    ' Disable logging from prior run
    SetLoggingEnabled False

    ' Get the column layout of the 'data' worksheet
    Dim ctx As sqlContext
    ctx.dataLayout = GetSettingsForDataWorksheet(DataSheet.name)

    ' Get the heading values of the 'data' worksheet columns.
    ctx.headings = GetSQLWorksheetHeadings(ctx.dataLayout)

    ' Get the column layout of the 'sql' worksheet
    ctx.sqlLayout = GetSettingsForSqlWorksheet()

    ' Get the list of special field names used for determining clusters and subclusters.
    ctx.fields = GetSettingsForSqlFields(True)
    
    ' Create a dictionary to hold substitution placeholder values
    Dim placeholders As Dictionary
    Set placeholders = New Dictionary
    placeholders.CompareMode = TextCompare

    ' Establish the loop constraints. A row of 0 passed in means run all SQL statements
    Dim firstRow As Long
    Dim lastRow  As Long
    
    Dim sqlCol As Long
    sqlCol = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
    
    If row = 0 Then
        firstRow = ctx.sqlLayout.firstRow
        lastRow = GetLastRowInColumn(SqlSheet, sqlCol)
        
        If lastRow < firstRow Then
            Exit Sub
        End If
    Else
        firstRow = row
        lastRow = row
    End If

    ' Disable automatic graph rendering as cells change.
    Dim runMode As String
    runMode = SafeStr(SettingsSheet.Range(SETTINGS_RUN_MODE).value)
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_MANUAL

    ' Clear out the info from previous run
    ClearSQLStatus
    ClearDataWorksheet DataSheet.name

    Dim dataRow As Long
    dataRow = ctx.dataLayout.firstRow

    ' The column used to filter which SQL statements should be run
    Dim filterColumn As Long
    If Len(SafeStr(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value)) = 0 Then
        filterColumn = 0
    Else
        filterColumn = GetSettingColNum(SETTINGS_SQL_COL_FILTER)
    End If

    ' Loop through the data rows of SQL statements
    Dim sqlStatement As String
    Dim sqlUCase As String
    Dim sqlRow As Long
    Dim dataFile As String
    Dim message As String
    Dim filePath As String
    Dim connectionObject As Object  ' Connection

    For sqlRow = firstRow To lastRow
        Application.StatusBar = False
        
        ' Skip initializations if the SQL row is commented out
        If SafeStr(SqlSheet.Cells.item(sqlRow, ctx.sqlLayout.flagColumn).value) <> FLAG_COMMENT Then

            ' Establish the full path to the Excel file containing the data
            filePath = GetExcelFilePath(sqlRow, ctx.sqlLayout, dataFile)

            ' Get SQL statement, and convert to upper case
            sqlStatement = Trim$(SafeStr(SqlSheet.Cells.item(sqlRow, ctx.sqlLayout.sqlStatementColumn).value))
            sqlUCase = UCase$(sqlStatement)

            ' Get default SUCCESS message
            message = GetMessage("msgboxSqlStatusSuccess")
        End If

        If SafeStr(SqlSheet.Cells.item(sqlRow, ctx.sqlLayout.flagColumn).value) = FLAG_COMMENT Then
            sqlStatement = vbNullString
            sqlUCase = vbNullString
            message = GetMessage("msgboxSqlStatusSkipped")

        ElseIf Len(sqlStatement) = 0 Then
            message = vbNullString

        ElseIf Not PassesFilter(sqlRow, filterColumn) Then
            message = GetMessage("msgboxSqlStatusFiltered")

        ElseIf sqlUCase = SQL_SET_DATA_FILE Then
            dataFile = SafeStr(SqlSheet.Cells.item(sqlRow, ctx.sqlLayout.excelFileColumn).value)

        ElseIf sqlUCase = SQL_RESET Then
            ClearDataWorksheet DataSheet.name

        ElseIf sqlUCase = SQL_ENABLE_LOGGING Then
            SetLoggingEnabled True

        ElseIf sqlUCase = SQL_DISABLE_LOGGING Then
            SetLoggingEnabled False

        ElseIf sqlUCase = SQL_LOG_ENVIRONMENT Then
            Dim loggingEnabled As Boolean
            loggingEnabled = IsLoggingEnabled()
            SetLoggingEnabled True
            LogDiagnostic "ENVIRONMENT", includeFingerprint:=True
            SetLoggingEnabled loggingEnabled

        ElseIf StartsWith(sqlUCase, SQL_SET_PLACEHOLDER) Then
            ParsePlaceholderLine placeholders, sqlStatement
            
        ElseIf StartsWith(sqlUCase, SQL_SET_CLUSTER_LEVEL_LIMIT) Then
            ParseClusterLevelLimitLine sqlStatement, ctx.fields.clusterLevelLimit
            
        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_ALL_VIEWS_AS_DIRECTED_GRAPH) Then
            Application.StatusBar = sqlStatement
            PublishAllViewsAsDirectedGraph (sqlStatement)

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_ALL_VIEWS_AS_UNDIRECTED_GRAPH) Then
            Application.StatusBar = sqlStatement
            PublishAllViewsAsUndirectedGraph sqlStatement

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_ALL_VIEWS) Then
            Application.StatusBar = sqlStatement
            PublishAllViews sqlStatement, SQL_PUBLISH_ALL_VIEWS

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_AS_DIRECTED_GRAPH) Then
            Application.StatusBar = sqlStatement
            PublishAsDirectedGraph sqlStatement

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_AS_UNDIRECTED_GRAPH) Then
            Application.StatusBar = sqlStatement
            PublishAsUndirectedGraph sqlStatement

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH) Then
            Application.StatusBar = sqlStatement
            Publish sqlStatement, SQL_PUBLISH

        ElseIf StartsWith(sqlUCase, SQL_PREVIEW_AS_DIRECTED_GRAPH) Then
            Application.StatusBar = sqlStatement
            PreviewAs TOGGLE_DIRECTED

        ElseIf StartsWith(sqlUCase, SQL_PREVIEW_AS_UNDIRECTED_GRAPH) Then
            Application.StatusBar = sqlStatement
            PreviewAs TOGGLE_UNDIRECTED

        ElseIf StartsWith(sqlUCase, SQL_PREVIEW) Then
            Application.StatusBar = sqlStatement
            CreateGraphWorksheet

        ElseIf Not StartsWith(sqlUCase, SQL_SELECT) Then
            message = GetMessage("msgboxSqlStatusSkipped") & " - " & GetMessage("msgboxSqlMustBeginWithSelect")

        Else
            ' SELECT branch
            Set connectionObject = Nothing

            ' Get connection to data source
            On Error Resume Next
            Set connectionObject = getConnection(filePath, ctx.fields.maxConnectionMinutes)
            Dim connErrDescription As String: connErrDescription = err.Description
            Dim connErrNumber As Long: connErrNumber = err.number
            On Error GoTo 0

            ' Execute the SQL query
            If connectionObject Is Nothing Then
                LogDiagnostic connErrDescription, errorNumber:=connErrNumber, errorCategory:="Data / Connection"
                message = GetMessage("msgboxSqlStatusFailure") & " - " & connErrDescription
            Else
                ' Apply placeholder substitutions before executing SQL
                ApplyPlaceholders sqlStatement, placeholders
                
                err.Clear
                message = executeSQL(ctx, filePath, connectionObject, sqlStatement, dataRow)
            End If
        End If

        ' Display the status of the SQL query
        SqlSheet.Cells.item(sqlRow, ctx.sqlLayout.statusColumn).value = message

        ' Breathe. a small delay before each SQL execution can reduce COM collisions on slower machines
        DoEvents
        SleepMilliseconds SQL_DELAY_MS
    Next sqlRow

    ' Clean up connection pool if using narrow-scoped pooling
    If GetSettingBoolean(SETTINGS_SQL_CLOSE_CONNECTIONS) Then CleanupConnectionPool

    ' Clean up placeholder dictionary
    CleanupPlaceholders placeholders
    
    ' Restore the run mode setting
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = runMode
End Sub

' ==========================================================================
' PROCEDURE: PreviewAs
' PURPOSE:
'   Temporary state-manager for triggering on-demand graph previews.
'
' TECHNICAL WORKFLOW:
'   1. CACHE STATE: Stores the user's current 'Graph Type' setting.
'   2. OVERRIDE: Programmatically sets the graph to 'directed' or 'undirected'.
'   3. EXECUTION: Calls 'CreateGraphWorksheet' to render the diagram.
'   4. RESTORE: Reverts the 'Graph Type' setting back to the original value.
'
' USAGE:
'   - Triggered by SQL commands like 'PREVIEW AS DIRECTED GRAPH'.
'   - Allows SQL-driven automation to dictate the visual style of a preview.
' ==========================================================================
Private Sub PreviewAs(ByVal graphType As String)
    Dim originalGraphType As String
    originalGraphType = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = graphType
    CreateGraphWorksheet
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = originalGraphType
End Sub

' ==========================================================================
' PROCEDURE: PublishAsDirectedGraph
' PURPOSE:
'   Orchestrates the "headless" publication of a directed graph file.
'
' TECHNICAL WORKFLOW:
'   1. COMMAND CAPTURE: Receives the full 'SQL_PUBLISH...' string from the
'      RunSQL loop.
'   2. MODE ENFORCEMENT: Passes the 'TOGGLE_DIRECTED' constant to the
'      underlying 'PublishAs' router.
'
' USAGE:
'   - Called when the SQL engine encounters 'PUBLISH AS DIRECTED GRAPH'.
'   - Bridges the gap between a SQL script and the physical file export engine.
' ==========================================================================
Private Sub PublishAsDirectedGraph(ByRef commandStatement As String)
    PublishAs commandStatement, SQL_PUBLISH_AS_DIRECTED_GRAPH, TOGGLE_DIRECTED
End Sub

' ==========================================================================
' PROCEDURE: PublishAsUndirectedGraph
' PURPOSE:
'   Orchestrates the "headless" publication of an undirected graph file.
'
' TECHNICAL WORKFLOW:
'   1. COMMAND CAPTURE: Receives the full 'SQL_PUBLISH...' string from the
'      RunSQL loop.
'   2. MODE ENFORCEMENT: Passes the 'TOGGLE_UNDIRECTED' constant to the
'      underlying 'PublishAs' router.
'
' USAGE:
'   - Called when the SQL engine encounters 'PUBLISH AS UNDIRECTED GRAPH'.
'   - Ensures that relationships are rendered as undirected edges (--) during
'     automated publishing.
' ==========================================================================
Private Sub PublishAsUndirectedGraph(ByRef commandStatement As String)
    PublishAs commandStatement, SQL_PUBLISH_AS_UNDIRECTED_GRAPH, TOGGLE_UNDIRECTED
End Sub

' ==========================================================================
' PROCEDURE: PublishAs
' PURPOSE:
'   A state-managed router for automated file publication.
'
' TECHNICAL WORKFLOW:
'   1. STATE CACHING: Backs up the current 'Filename Prefix' and 'Graph Type'
'      from the Settings worksheet.
'   2. INJECTION: Updates the global settings with the custom filename
'      parsed from the SQL command and the specified graphType.
'   3. EXECUTION: Triggers the 'Publish' routine to render and save the file.
'   4. RESTORATION: Reverts the worksheet settings to their original state
'      to ensure a seamless user experience after the automation finishes.
'
' USAGE:
'   - The common backend for 'PublishAsDirectedGraph' and 'PublishAsUndirectedGraph'.
' ==========================================================================
Private Sub PublishAs(ByRef commandStatement As String, ByRef phrase As String, ByVal graphType As String)
    ' Backup current values
    Dim originalPrefix As String
    Dim originalGraphType As String
    
    originalPrefix = SettingsSheet.Range(SETTINGS_FILE_NAME).value
    originalGraphType = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value
    
    ' Override current values
    SetPrefix commandStatement, phrase
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = graphType
    
    ' Create the graph
    Publish commandStatement, phrase
    
    ' Restore values from backup
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = originalGraphType
    SettingsSheet.Range(SETTINGS_FILE_NAME).value = originalPrefix
End Sub

' ==========================================================================
' PROCEDURE: Publish
' PURPOSE:
'   Orchestrates the automated export of the primary graph view to a file.
'
' TECHNICAL WORKFLOW:
'   1. SCOPE RESOLUTION: Identifies the 'Show Style' column index (the
'      primary View toggle) from the workbook settings.
'   2. DISPATCH: Calls 'PublishViews' with identical start and end column
'      indices to force a single-view execution.
'
' USAGE:
'   - Triggered by the 'SQL_PUBLISH' command.
'   - Acts as the baseline automation for file generation via SQL.
' ==========================================================================
Private Sub Publish(ByRef commandStatement As String, ByRef phrase As String)
    Dim firstColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    Dim lastColumn As Long
    lastColumn = firstColumn

    PublishViews commandStatement, phrase, firstColumn, lastColumn
End Sub

' ==========================================================================
' PROCEDURE: PublishViews
' PURPOSE:
'   The low-level coordinator for automated (SQL-driven) file exports.
'
' TECHNICAL WORKFLOW:
'   1. PATH PRESERVATION: Backs up the current 'Filename Prefix' to ensure
'      the user's manual settings are not lost.
'   2. COMMAND PARSING: Calls 'SetPrefix' to extract the desired filename
'      from the SQL command string (e.g., extracting "MyGraph" from "PUBLISH MyGraph").
'   3. ENGINE INVOCATION: Triggers 'CreateGraphFile' (in modCreateGraph.bas)
'      to perform the actual DOT generation and image rendering.
'   4. STATE RESTORATION: Reverts the 'Filename Prefix' back to its
'      original value once the export is complete.
'
' USAGE:
'   - Common backend for both single-view (PUBLISH) and multi-view
'     (PUBLISH ALL VIEWS) automation commands.
' ==========================================================================
Private Sub PublishViews(ByRef commandStatement As String, ByRef phrase As String, ByVal firstColumn As Long, ByVal lastColumn As Long)
    ' Backup current values
    Dim originalPrefix As String
    originalPrefix = SettingsSheet.Range(SETTINGS_FILE_NAME).value
    
    ' Override current values
    SetPrefix commandStatement, phrase
    
    ' Create the graph
    CreateGraphFile firstColumn, lastColumn
    
    ' Restore values from backup
    SettingsSheet.Range(SETTINGS_FILE_NAME).value = originalPrefix
End Sub

' ==========================================================================
' SECTION: BATCH PUBLISHING AUTOMATION (SQL-DRIVEN)
' ==========================================================================

''
' BATCH HANDLER: Exports all defined views as directed graphs (->).
' Triggered by: 'PUBLISH ALL VIEWS AS DIRECTED GRAPH'
' Logic: Forces the 'TOGGLE_DIRECTED' state before iterating through
' every view column defined in the Styles worksheet.
'
Private Sub PublishAllViewsAsDirectedGraph(ByRef commandStatement As String)
    PublishAllViewsAs commandStatement, SQL_PUBLISH_ALL_VIEWS_AS_DIRECTED_GRAPH, TOGGLE_DIRECTED
End Sub

''
' BATCH HANDLER: Exports all defined views as undirected graphs (--).
' Triggered by: 'PUBLISH ALL VIEWS AS UNDIRECTED_GRAPH'
' Logic: Forces the 'TOGGLE_UNDIRECTED' state to ensure relationships
' are rendered as simple connections without arrows.
'
Private Sub PublishAllViewsAsUndirectedGraph(ByRef commandStatement As String)
    PublishAllViewsAs commandStatement, SQL_PUBLISH_ALL_VIEWS_AS_UNDIRECTED_GRAPH, TOGGLE_UNDIRECTED
End Sub

' ==========================================================================
' PROCEDURE: PublishAllViewsAs
' PURPOSE:
'   A state-managed router for automated batch (multi-view) file publication.
'
' TECHNICAL WORKFLOW:
'   1. STATE PERSISTENCE: Backs up the current 'Filename Prefix' and
'      'Graph Type' from the Settings worksheet.
'   2. CONFIGURATION OVERRIDE: Applies the custom filename from the SQL
'      command and enforces the specified graphType (Directed/Undirected).
'   3. EXECUTION: Triggers 'PublishAllViews' to iterate through and render
'      every view column.
'   4. RESTORATION: Reverts all settings to their pre-automation state to
'      maintain user workspace consistency.
'
' USAGE:
'   - Common backend for 'PublishAllViewsAsDirectedGraph' and
'     'PublishAllViewsAsUndirectedGraph'.
' ==========================================================================
Private Sub PublishAllViewsAs(ByRef commandStatement As String, ByRef phrase As String, ByVal graphType As String)
    ' Backup current values
    Dim originalPrefix As String
    Dim originalGraphType As String
    
    originalPrefix = SettingsSheet.Range(SETTINGS_FILE_NAME).value
    originalGraphType = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value
    
    ' Override current values
    SetPrefix commandStatement, phrase
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = graphType
    
    ' Create the graph
    PublishAllViews commandStatement, phrase
    
    ' Restore values from backup
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = originalGraphType
    SettingsSheet.Range(SETTINGS_FILE_NAME).value = originalPrefix
End Sub

' ==========================================================================
' PROCEDURE: PublishAllViews
' PURPOSE:
'   Calculates the full range of View columns for automated batch export.
'
' TECHNICAL WORKFLOW:
'   1. STARTING POINT: Identifies the first "Yes/No" View column using the
'      'SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW' named range.
'   2. DYNAMIC DISCOVERY: Calls 'GetLastViewColumn' to find the end of the
'      View matrix, allowing the tool to adapt as users add new columns.
'   3. EXECUTION: Passes the resolved column range to 'PublishViews' to
'      trigger the multi-file rendering loop.
'
' USAGE:
'   - Triggered by the 'SQL_PUBLISH_ALL_VIEWS' pseudo-SQL command.
'   - Essential for large-scale documentation where every "View" must be
'     exported as an independent asset.
' ==========================================================================
Private Sub PublishAllViews(ByRef commandStatement As String, ByRef phrase As String)
    Dim firstColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    
    Dim lastColumn As Long
    lastColumn = GetLastViewColumn(firstColumn)

    PublishViews commandStatement, phrase, firstColumn, lastColumn
End Sub

' ==========================================================================
' FUNCTION: GetLastViewColumn
' PURPOSE:
'   Identifies the final column in the Styles worksheet's "View" matrix.
'
' TECHNICAL WORKFLOW:
'   1. HEADING RESOLUTION: Locates the specific row index where View titles
'      reside (via SETTINGS_STYLES_ROW_HEADING).
'   2. CONTIGUOUS SCAN: Iterates horizontally from the first View column,
'      counting non-empty cells to find the edge of the defined Views.
'   3. COORDINATE CALCULATION: Computes the absolute Excel column index
'      to provide an accurate "Stop" point for batch rendering loops.
'
' USAGE:
'   - Used by 'PublishAllViews' to define the iteration range for mass
'     diagram generation.
'   - Ensures new View columns are automatically detected without code changes.
' ==========================================================================
Private Function GetLastViewColumn(ByVal firstColumn As Long) As Long
    Dim nonEmptyCellCount As Long
    Dim row As Long
    Dim col As Long
    row = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    
    ' Count the non-empty cells beginning at the first view column
    nonEmptyCellCount = 0
    For col = firstColumn To GetLastColumn(StylesSheet.name, row)
        If StylesSheet.Cells.item(row, col) <> vbNullString Then
            nonEmptyCellCount = nonEmptyCellCount + 1
        End If
    Next col

    ' Calaculate the absolute column number of the last view column
    GetLastViewColumn = firstColumn + nonEmptyCellCount - 1
End Function

' ==========================================================================
' PROCEDURE: SetPrefix
' PURPOSE:
'   Extracts and applies a custom filename from a SQL publishing command.
'
' TECHNICAL WORKFLOW:
'   1. STRING PARSING: Evaluates the 'commandStatement' to find text
'      appearing after the command keyword (e.g., after "PUBLISH ").
'   2. SETTINGS INJECTION: Trims the extracted string and writes it
'      directly into the SETTINGS_FILE_NAME named range on the Settings sheet.
'
' USAGE:
'   - Called by 'PublishViews' and 'PublishAllViewsAs' to set the
'     destination filename for automated exports.
'   - Enables syntax like: "PUBLISH NetworkDiagram_v1"
' ==========================================================================
Private Sub SetPrefix(ByRef commandStatement As String, ByRef phrase As String)
    ' Override the filename prefix with the value provided after the PUBLISH phrase
    If Len(commandStatement) > Len(phrase) Then
        Dim prefix As String
        prefix = Mid$(commandStatement, Len(phrase) + 1)
        SettingsSheet.Range(SETTINGS_FILE_NAME).value = Trim$(prefix)
    End If
End Sub

' ==========================================================================
' FUNCTION: GetExcelFilePath
' PURPOSE:
'   Determines the target database file path for a SQL query based on a
'   four-tier order of precedence.
'
' TECHNICAL WORKFLOW:
'   1. SCENARIO 1 (High Priority): Checks the 'Excel File' column on the
'      specific SQL row. Supports both absolute and relative paths.
'   2. SCENARIO 2 (State-Driven): Uses the value set by the previous
'      'SET DATA FILE' command if one was issued.
'   3. SCENARIO 3 (UI-Driven): Falls back to the global Data Source
'      selected in the Ribbon/Settings tab.
'   4. SCENARIO 4 (Default): If no other source is identified, targets the
'      active workbook itself.
'
' USAGE:
'   - Called by 'RunSQL' to identify the target for the ADO 'getConnection' call.
'   - Enables "Multi-Source" scripts that query multiple files in one pass.
' ==========================================================================
Private Function GetExcelFilePath(ByVal sqlRow As Long, ByRef sqlLayout As sqlWorksheet, ByVal dataFile As String) As String
    ' Order of precedence is
    ' 1) filename from current SQL row
    ' 2) SET DATA FILE filename passed in the dataFile parameter
    ' 3) Data file from the ribbon values
    ' 4) Current workbook
    
    ' Get the file from the SQL row
    Dim filePath As String
    
    ' Scenario 1 - filename from the current SQL row
    filePath = Trim$(SqlSheet.Cells.item(sqlRow, sqlLayout.excelFileColumn).value)
    If filePath <> vbNullString Then
        If InStr(filePath, Application.pathSeparator) Then
            GetExcelFilePath = filePath
        Else
            GetExcelFilePath = ActiveWorkbook.path & Application.pathSeparator & filePath
        End If
        Exit Function
    End If
    
    ' Scenario 2 - SET DATA FILE filename passed in the dataFile parameter
    If Trim$(dataFile) <> vbNullString Then
        If InStr(dataFile, Application.pathSeparator) Then
            GetExcelFilePath = dataFile
        Else
            GetExcelFilePath = ActiveWorkbook.path & Application.pathSeparator & Trim$(dataFile)
        End If
        Exit Function
    End If
    
    ' Scenario 3 - Get data file from the ribbon
    Dim dirName As String
    dirName = SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY)
    
    Dim fileName As String
    fileName = SettingsSheet.Range(SETTINGS_DATASOURCE_FILE)
    
    If dirName <> vbNullString And fileName <> vbNullString Then
        GetExcelFilePath = dirName & Application.pathSeparator & fileName
        Exit Function
    End If
    
    ' Scenario 4 - Current workbook
    GetExcelFilePath = ActiveWorkbook.FullName
End Function

' ==========================================================================
' FUNCTION: PassesFilter
' PURPOSE:
'   Determines if a specific SQL row should be executed based on the
'   workbook's active filter criteria.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN VALIDATION: If no filter column is defined, the check is
'      bypassed (returns True).
'   2. VALUE RETRIEVAL: Performs a null-safe capture of both the row's
'      tag and the global 'SQL Filter Value' from Settings.
'   3. COMPARISON: Uses 'vbTextCompare' to perform a case-insensitive
'      match between the row tag and the global filter.
'
' USAGE:
'   - Called by 'RunSQL' to decide whether to process or skip a query.
'   - Enables "Batch Tagging" for managing large SQL libraries.
' ==========================================================================
Private Function PassesFilter(ByVal sqlRow As Long, ByVal filterColumn As Long) As Boolean
    ' Default: does NOT pass
    PassesFilter = False

    ' No filter column selected
    If filterColumn <= 0 Then
        PassesFilter = True
        Exit Function
    End If

    ' Null-safe retrieval of the cell value
    Dim cellValue As String
    cellValue = SafeStr(SqlSheet.Cells.item(sqlRow, filterColumn).value)

    ' Null-safe retrieval of the filter value
    Dim filterValue As String
    filterValue = SafeStr(SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value)

    ' Apply filter
    If StrComp(Trim$(cellValue), filterValue, vbTextCompare) = 0 Then
        PassesFilter = True
    End If
End Function

' ==========================================================================
' PROCEDURE: ClearSQLStatus
' PURPOSE:
'   Resets the 'Status' column on the SQL worksheet to a blank state.
'
' TECHNICAL WORKFLOW:
'   1. LAYOUT DISCOVERY: Retrieves the current SQL worksheet column mapping
'      (sqlLayout) to identify the 'Status' column index.
'   2. RANGE CALCULATION: Determines the extent of the data using 'UsedRange'
'      and converts the column index into Excel A1 notation.
'   3. CLEARANCE: Calls 'ClearContents' on the specific range to wipe previous
'      execution feedback (e.g., Success, Failed, or Error messages).
'
' USAGE:
'   - Called by 'RunSQL' as part of the initialization phase.
'   - Ensures the 'Status' column accurately reflects only the current session.
' ==========================================================================
Public Sub ClearSQLStatus()
    Dim cellRange As String
    Dim sqlLayout As sqlWorksheet
    
    ' Get the layout of the 'sql' worksheet
    sqlLayout = GetSettingsForSqlWorksheet()
    
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    ' Format the range to clear
    cellRange = ConvertColumnNumberToLetters(sqlLayout.statusColumn) & sqlLayout.firstRow & ":" & ConvertColumnNumberToLetters(sqlLayout.statusColumn) & lastRow
    SqlSheet.Range(cellRange).ClearContents
End Sub

' ==========================================================================
' SECTION: SQL PROCESSING
' https://technet.microsoft.com/en-us/library/ee692882.aspx
' ==========================================================================

' ==========================================================================
' FUNCTION: executeSQL
' PURPOSE:
'   The core engine for ADO Recordset execution and result transformation.
'   Manages query execution, search algorithms, and data worksheet mapping.
'
' TECHNICAL WORKFLOW:
'   1. RECORDSET INITIALIZATION: Creates a late-bound ADODB.Recordset
'      to ensure cross-version Excel compatibility.
'   2. RESILIENCE LOOP: Implements a retry-loop (ctx.fields.retryLimit)
'      to handle transient locking while identifying fatal 'User Errors'
'      (syntax, missing tables) via 'IsUserSQLError'.
'   3. ALGORITHM DISPATCHER:
'      - IterativeSearch: Performs parameterized data retrieval.
'      - RecursiveSearch: Handles parent-child tree traversals.
'      - MergeRecordsets: Unifies anchor and recursive results into a single set.
'   4. RESULT MAPPING: Determines the output strategy based on field signatures:
'      - CreateEdges: Generates a sequential chain of edges.
'      - CreateRank: Groups nodes into specific Graphviz rank levels.
'      - ProcessMultiLevelRecordset: Handles CLUSTERn hierarchical nesting.
'      - MapResultsToDataWorksheet: The standard mapping for flat results.
'   5. FORENSIC LOGGING: Captures query failures and environmental context
'      via 'LogDiagnostic' for troubleshooting.
' ==========================================================================
Private Function executeSQL( _
    ByRef ctx As sqlContext, _
    ByVal filePath As String, _
    ByRef connectionObject As Object, _
    ByVal sqlStatement As String, _
    ByRef row As Long) As String

#If Win32 Or Win64 Then

    On Error GoTo executeSQLError
    
    Dim rs As Object
    Dim rsRecursion As Object
    Dim rsMerged As Object
    Dim recordCnt As Long: recordCnt = 0
    Dim attempts As Long: attempts = 0
    Dim userError As Boolean: userError = False
    Dim errNumber As Long: errNumber = 0
    Dim errDescription As String: errDescription = vbNullString
    
    ' A Microsoft bug is causing it to take 12 seconds to get a connection, so provide feedback
    Application.StatusBar = replace(GetMessage("statusbarSqlEstablishingConnection"), "{filePath}", filePath)
    
    ' Define a recordset for a SQL SELECT statement using late binding
    ' as we do not know which version of Excel this spreadsheet
    ' will be running on
    Set rs = CreateObject("ADODB.Recordset")
    
    err.Clear
    For attempts = 1 To ctx.fields.retryLimit
        On Error Resume Next
        
        ' Execute the SQL SELECT query
        rs.Open source:=sqlStatement, ActiveConnection:=connectionObject, CursorType:=CursorTypeEnum.adOpenStatic, LockType:=LockTypeEnum.adLockReadOnly, options:=CommandTypeEnum.adCmdText
        
        ' Immediately save error state, as processing an error could trigger it being cleared
        errNumber = err.number
        errDescription = err.Description
        
        ' Break from retry loop if query succeeded
        If errNumber = 0 Then Exit For
        
        ' Check for non-retryable SQL errors (bad sheet, bad column, syntax, etc.)
        userError = IsUserSQLError(errDescription)
        If userError Then
            Exit For
        End If

        LogDiagnostic "executeSQL(): rs.Open - " & errDescription, errorNumber:=errNumber, attempt:=attempts, sql:=sqlStatement, errorCategory:=ClassifyError(err.Description)
        err.Clear
        SleepMilliseconds RETRY_DELAY_MS
    Next attempts
    
    ' Reset status bar
    Application.StatusBar = False
    err.Clear
    
    ' If userError, then the SQL is bad. Stop processing.
    If userError Then GoTo executeSQLError
    If errNumber <> 0 Then GoTo executeSQLError
    
    ' If the recordset failed to open but didn't trigger userError, rs.State might be 0.
    
    If rs Is Nothing Or rs.State <> ObjectStateEnum.adStateOpen Then
        err.Raise vbObjectError + 513, , "executeSQL(): Recordset failed to open"
    End If

    ' Determine if enumeration values are present
    ctx.loop = GetLoopLimits(ctx, rs)
    
    ' Execute any iteration query passed in the SQL SELECT
    ' Performs iteration of parameterized query + mapping to data worksheet.
    IterativeSearch connectionObject, ctx, rs, row, recordCnt
    
    ' Execute any recursion query passed in the SQL SELECT
    ' Perfroms recursion + mapping to data worksheet.
    RecursiveSearch connectionObject, ctx, rs, rsRecursion
    
    Dim finalRs As Object

    If rsRecursion Is Nothing Then
        ' No recursion query was run, emit the results of the primary query
        Set finalRs = rs
    Else
        ' A set of recursive queries was executed. We have to merge the results
        ' of the primary query, and recursive queries into a single set of
        ' results so that cluster and subclusters are honored across all the
        ' queries.
        MergeRecordsets rs, rsRecursion, rsMerged
        Set finalRs = rsMerged
    End If

    ' =========================================================================
    ' Process the results
    ' =========================================================================
    
    If HasField(finalRs, ctx.fields.CreateEdges) Then
        ' Create a chain of edges from record to next record
        CreateEdges ctx, finalRs, row, recordCnt
        
    ElseIf HasField(finalRs, ctx.fields.CreateRank) Then
        ' Create subgroup of nodes on a common specified rank
        CreateRank ctx, finalRs, row, recordCnt
        
    ElseIf DetectMultiLevel(finalRs, ctx.fields.Cluster) Then
        ' Use the new CLUSTERn multi-level processor
        ProcessMultiLevelRecordset ctx, finalRs, row, recordCnt
        
    Else
        ' Use the legacy CLUSTER / SUBCLUSTER processor
        MapResultsToDataWorksheet ctx, finalRs, row, recordCnt
    End If
    
    ' Return success status in local language
    executeSQL = GetMessage("msgboxSqlStatusSuccess")
    
Cleanup:
    On Error Resume Next
    SafeCloseRecordset rs
    SafeCloseRecordset rsRecursion
    SafeCloseRecordset rsMerged
    Application.StatusBar = False
    On Error GoTo 0
    Exit Function

executeSQLError:
    ' GetMessage will reset the error state, save the message
    If err.number <> 0 Then
        errDescription = err.Description
        errNumber = err.number
    End If
    
    Dim logMessage As String
    logMessage = "executeSQL(): " & errDescription & vbNewLine & "  Excel Data File     : " & filePath

    Dim statusMessage As String
    statusMessage = errDescription _
                & vbNewLine & vbNewLine & ClassifyError(errDescription) _
                & vbNewLine & vbNewLine & "Err.Number=" & errNumber _
                & vbNewLine & vbNewLine & "datafile=" & filePath

    LogDiagnostic logMessage, errorNumber:=errNumber, attempt:=attempts, sql:=sqlStatement, errorCategory:=ClassifyError(errDescription)
    executeSQL = GetMessage("msgboxSqlStatusFailure") & " - " & statusMessage
    GoTo Cleanup
#Else
    executeSQL = GetMessage("msgboxSqlStatusFailure") & " - ADO is not supported on macOS."
#End If

End Function

' ==========================================================================
' PROCEDURE: SafeCloseRecordset
' PURPOSE:
'   Defensively closes and destroys an ADO Recordset object.
'
' TECHNICAL WORKFLOW:
'   1. NULL CHECK: Verifies the object exists before attempting operations.
'   2. ATOMIC STATE CAPTURE: Reads the .State property once to prevent
'      race conditions or COM instability during multiple reads.
'   3. BITWISE VALIDATION: Uses 'adStateOpen' comparison to ensure
'      the recordset is active before calling .Close.
'   4. RESOURCE RECLAMATION: Sets the object to 'Nothing' to signal
'      to Excel's garbage collector that the memory can be freed.
' ==========================================================================
Private Sub SafeCloseRecordset(ByRef rs As Object)
#If Win32 Or Win64 Then
    On Error Resume Next
    If Not rs Is Nothing Then
        Dim st As Long
        st = rs.State                  ' capture once (in case multiple reads cause issues)
        If (st And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    err.Clear
#End If
End Sub

' ==========================================================================
' PROCEDURE: SleepMilliseconds
' PURPOSE:
'   Implements a non-blocking pause in execution to stabilize COM operations.
'
' TECHNICAL WORKFLOW:
'   1. TIMER CAPTURE: Records the current system 'Timer' value.
'   2. YIELDING LOOP: Executes a 'Do While' loop until the specified
'      millisecond threshold is reached.
'   3. UI RESPONSIVENESS: Involves 'DoEvents' within the loop to ensure
'      Excel can process background tasks (repaints, clicks) during the wait.
'
' USAGE:
'   - Used by the 'getConnection' and 'executeSQL' retry loops to prevent
'     CPU-intensive collisions when reconnecting to locked database files.
' ==========================================================================
Private Sub SleepMilliseconds(ByVal ms As Long)
#If Win32 Or Win64 Then
    Dim t As Single
    t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
#End If
End Sub

' ==========================================================================
' FUNCTION: GetLoopLimits
' PURPOSE:
'   Extracts and validates iteration parameters for "Enumeration Mode."
'
' TECHNICAL WORKFLOW:
'   1. FEATURE DETECTION: Checks for the presence of the 'ENUMERATE' field
'      in the recordset to decide if looping is active.
'   2. PARAMETER HARVESTING: Retrieves the 'startAt', 'stopAt', and 'stepBy'
'      values required to define the mathematical range.
'   3. GOVERNOR ENFORCEMENT: Applies 'LOOP_MAX_STEPS' (or a user-defined max)
'      to prevent infinite loops if the SQL result is poorly configured.
'   4. DIRECTIONAL VALIDATION: Verifies that the 'stepBy' value logically
'      moves the current value toward the 'stopAt' value; if a mismatch
'      occurs, it safely reverts to a single-iteration mode.
'   5. DIAGNOSTIC LOGGING: Reports loop configuration errors (e.g., Step=0)
'      to the ADO logger while maintaining execution stability.
' ==========================================================================
Private Function GetLoopLimits(ByRef ctx As sqlContext, _
                               ByVal rs As Object) As EnumerateParameters

    Dim s As EnumerateParameters
    ' Default: no loop mode; EmitRows/consumers decide what to do with a single pass
    s.Enabled = False
    s.startAt = 1                   ' Default single-iteration loop
    s.stopAt = 1
    s.stepBy = 1
    s.count = 0
    s.max = LOOP_MAX_STEPS

    ' ------------------------------------------------------------
    ' Preconditions
    ' ------------------------------------------------------------
    If rs.EOF Then
        ' No rows -> callers can still choose to emit once with defaults
        GetLoopLimits = s
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' ENUMERATE field missing -> normal mode, single iteration, no diagnostics
    ' ------------------------------------------------------------
    If Not HasField(rs, ctx.fields.enumerateSwitch) Then
        GetLoopLimits = s
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' ENUMERATE field present -> read it
    ' ------------------------------------------------------------
    Dim enumerateSwitch As Boolean
    enumerateSwitch = GetFieldValueBoolean(rs, ctx.fields.enumerateSwitch)

    ' ENUMERATE = FALSE -> normal mode, single iteration, no diagnostics
    If Not enumerateSwitch Then
        GetLoopLimits = s
        Exit Function
    End If

    ' From here on, ENUMERATE is explicitly enabled
    s.Enabled = True

    ' ------------------------------------------------------------
    ' ENUMERATE = TRUE -> attempt to fetch loop parameters
    ' ------------------------------------------------------------
    Dim hasStart As Boolean: hasStart = HasField(rs, ctx.fields.enumerateStartAt)
    Dim hasStop  As Boolean: hasStop = HasField(rs, ctx.fields.enumerateStopAt)
    Dim hasStep  As Boolean: hasStep = HasField(rs, ctx.fields.enumerateStepBy)

    ' Missing parameters -> single iteration + diagnostic
    If Not hasStart Or Not hasStop Or Not hasStep Then
        LogDiagnostic "ENUMERATE=TRUE but loop parameters are missing; defaulting to single-iteration mode (1 -> 1 step 1)."
        GetLoopLimits = s
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' Extract supplied values
    ' ------------------------------------------------------------
    s.startAt = GetFieldValueLong(rs, ctx.fields.enumerateStartAt)
    s.stopAt = GetFieldValueLong(rs, ctx.fields.enumerateStopAt)
    s.stepBy = GetFieldValueLong(rs, ctx.fields.enumerateStepBy)

    ' ------------------------------------------------------------
    ' Caller can override the loop governor
    ' ------------------------------------------------------------
    Dim hasMax As Boolean: hasMax = HasField(rs, ctx.fields.enumerateMax)
    If hasMax Then
        s.max = GetFieldValueLong(rs, ctx.fields.enumerateMax)
        
        ' Max must be positive
        If s.max < 0 Then
            LogDiagnostic "Invalid loop: max < 0. Using internal limit."
            s.max = LOOP_MAX_STEPS
        End If
    End If

    ' ------------------------------------------------------------
    ' SAFETY VALIDATION
    ' Only log diagnostics when something is wrong.
    ' ------------------------------------------------------------

    ' Step cannot be zero
    If s.stepBy = 0 Then
        LogDiagnostic "Invalid loop: stepBy = 0. Defaulting to stepBy = 1."
        s.stepBy = 1
        GetLoopLimits = s
        Exit Function
    End If

    ' Direction mismatch
    If (s.stepBy > 0 And s.stopAt <= s.startAt) Then
        LogDiagnostic "Invalid loop: positive step but stopAt <= startAt. Defaulting to single-iteration mode."
        s.startAt = 1: s.stopAt = 1: s.stepBy = 1
        GetLoopLimits = s
        Exit Function
    End If

    If (s.stepBy < 0 And s.stopAt >= s.startAt) Then
        LogDiagnostic "Invalid loop: negative step but stopAt >= startAt. Defaulting to single-iteration mode."
        s.startAt = 1: s.stopAt = 1: s.stepBy = 1
        GetLoopLimits = s
        Exit Function
    End If

    GetLoopLimits = s
End Function

' ==========================================================================
' PROCEDURE: IterativeSearch
' PURPOSE:
'   Executes a two-stage "Identity-to-Data" search pattern.
'
' TECHNICAL WORKFLOW:
'   1. SIGNATURE DETECTION: Verifies the recordset contains the required
'      fields: 'ITERATE', 'ID_QUERY', and 'DATA_QUERY'.
'   2. TEMPLATE CAPTURE: Retrieves the primary ID query and the parameterized
'      data template from the initial result set.
'   3. IDENTITY RESOLUTION: Executes the 'ID_QUERY' (via GetHeaderRS) to
'      generate a list of keys (e.g., a list of Employee IDs).
'   4. BRANCHING LOGIC:
'      - Concat Mode: Collapses results for each ID into a single multi-line
'        entity using 'ProcessInConcatMode'.
'      - Classic Mode: Emits standard Node/Edge rows for each result found
'        via 'ProcessInClassicMode'.
'   5. RESOURCE HYGIENE: Systematic closure of the header recordset to
'      prevent memory leaks during high-frequency iteration.
' ==========================================================================
Private Sub IterativeSearch( _
    ByRef connectionObject As Object, _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    If rs Is Nothing Or rs.State <> ObjectStateEnum.adStateOpen Or _
       (rs.EOF And rs.BOF) Then Exit Sub

    If Not (HasField(rs, ctx.fields.iterate) And _
            HasField(rs, ctx.fields.idQuery) And _
            HasField(rs, ctx.fields.dataQuery)) Then Exit Sub

    Dim idQuery      As String
    idQuery = GetFieldValueString(rs, ctx.fields.idQuery)

    Dim dataTemplate As String
    dataTemplate = GetFieldValueString(rs, ctx.fields.dataQuery)

    If Len(idQuery) = 0 Or Len(dataTemplate) = 0 Then Exit Sub

    Dim cs As ConcatSettings
    cs = ReadConcatSettings(ctx, rs)

    Dim headerRS As Object
    Set headerRS = GetHeaderRS(connectionObject, idQuery)
    If headerRS Is Nothing Then Exit Sub

    If cs.Enabled Then
        ProcessInConcatMode connectionObject, ctx, dataTemplate, headerRS, cs, row, recordCnt
    Else
        ProcessInClassicMode connectionObject, ctx, dataTemplate, headerRS, row, recordCnt
    End If

    SafeCloseRecordset headerRS
End Sub

' ==========================================================================
' PROCEDURE: ProcessInClassicMode
' PURPOSE:
'   Executes parameterized data queries for a collection of unique IDs.
'
' TECHNICAL WORKFLOW:
'   1. ID AGGREGATION: Calls 'CollectUniqueIDs' to distill the header
'      recordset into a unique set of keys, preventing redundant queries.
'   2. PARAMETERIZED LOOP: Iterates through each unique ID:
'      - Injects the ID into the 'dataTemplate' using 'RunParameterizedQuery'.
'      - Validates the resulting ADO Recordset state.
'   3. DATA MAPPING: Streams the results of each sub-query directly to
'      the 'Data' worksheet via 'MapResultsToDataWorksheet'.
'   4. MEMORY HYGIENE: Closes and destroys each sub-recordset immediately
'      after mapping to maintain a low memory footprint.
' ==========================================================================
Private Sub ProcessInClassicMode( _
    conn As Object, ctx As sqlContext, _
    dataTemplate As String, headerRS As Object, _
    ByRef row As Long, ByRef recordCnt As Long)

    Dim idList As Object: Set idList = CollectUniqueIDs(headerRS)
    If idList Is Nothing Or idList.count = 0 Then Exit Sub

    Dim id As Variant, rsData As Object
    For Each id In idList.Keys
        Set rsData = RunParameterizedQuery(conn, ctx, dataTemplate, id)
        If Not rsData Is Nothing Then
            If rsData.State = ObjectStateEnum.adStateOpen Then
                MapResultsToDataWorksheet ctx, rsData, row, recordCnt
            End If
            SafeCloseRecordset rsData
        End If
    Next
End Sub

' ==========================================================================
' PROCEDURE: ProcessInConcatMode
' PURPOSE:
'   Collapses multi-row detail data into a single, unified recordset.
'
' TECHNICAL WORKFLOW:
'   1. RECORDSET AUGMENTATION: Calls 'CreateAugmentedRS' to execute sub-queries
'      for each header ID and append the aggregated results as a new field.
'   2. DATA AGGREGATION: Joins related detail records using user-defined
'      prefixes, suffixes, and separators (e.g., creating a bulleted list
'      inside a single Node label).
'   3. WORKSHEET STREAMING: Maps the final, "flattened" recordset to the
'      Data worksheet as a single row per primary ID.
'   4. RESOURCE CLEANUP: systematic closure of the augmented recordset to
'      reclaim memory after the transformation is complete.
'
' USAGE:
'   - Triggered when 'CONCAT_FIELD' is defined in the SQL worksheet.
'   - Ideal for displaying a list of attributes (e.g., 'Skills' or 'Tags')
'     inside a single Node rather than creating separate edges for each.
' ==========================================================================
Private Sub ProcessInConcatMode( _
    conn As Object, ctx As sqlContext, _
    dataTemplate As String, headerRS As Object, _
    ByRef cs As ConcatSettings, _
    ByRef row As Long, ByRef recordCnt As Long)

    Dim augRS As Object
    Set augRS = CreateAugmentedRS(conn, ctx, headerRS, dataTemplate, _
                                  cs.ConcatField, cs.TargetField, _
                                  cs.prefix, cs.suffix, cs.separator)
    
    If Not augRS Is Nothing Then
        If augRS.State = ObjectStateEnum.adStateOpen Then
            MapResultsToDataWorksheet ctx, augRS, row, recordCnt
        End If
        SafeCloseRecordset augRS
    End If
End Sub

' ==========================================================================
' FUNCTION: CollectUniqueIDs
' PURPOSE:
'   Extracts a unique set of identifiers from a source recordset to
'   optimize subsequent sub-query execution.
'
' TECHNICAL WORKFLOW:
'   1. STATE VALIDATION: Verifies the 'headerRS' is open and valid.
'   2. DICTIONARY INITIALIZATION: Creates a 'Scripting.Dictionary' to
'      manage the unique key-value pairs (O(1) lookup performance).
'   3. ITERATION: Scans the 'ID' field of the source recordset.
'   4. SANITIZATION: Uses 'SafeStr' to handle potential Null or
'      non-string data types.
'   5. DE-DUPLICATION: Adds only new IDs to the dictionary, effectively
'      flattening the list into a unique set.
'
' USAGE:
'   - Called by 'ProcessInClassicMode' to minimize the number of
'     database calls during parameterized searches.
' ==========================================================================
Private Function CollectUniqueIDs( _
    ByVal headerRS As Object) As Object

    If headerRS Is Nothing Then Exit Function
    If headerRS.State <> ObjectStateEnum.adStateOpen Then Exit Function

    Dim idList As Object
    Set idList = CreateObject("Scripting.Dictionary")

    Dim id As String
    headerRS.MoveFirst
    Do While Not headerRS.EOF
        id = SafeStr(headerRS.fields("ID").value)
        If Len(id) > 0 Then
            If Not idList.Exists(id) Then idList.Add id, True
        End If
        headerRS.MoveNext
    Loop

    Set CollectUniqueIDs = idList
End Function

' ==========================================================================
' FUNCTION: ReadConcatSettings
' PURPOSE:
'   Determines if "Concat Mode" is active and extracts formatting rules
'   from the ADO Recordset.
'
' TECHNICAL WORKFLOW:
'   1. FEATURE DETECTION: Checks for the 'CONCAT_ENABLE' field.
'   2. FUZZY BOOLEAN LOGIC: Evaluates the "Enabled" flag across multiple
'      types (Booleans, "Yes/No" strings, or Numeric 1/0) for user flexibility.
'   3. PROPERTY EXTRACTION: Captures the 'ConcatField' (source), 'TargetField'
'      (destination), and visual delimiters (Prefix, Suffix, Separator).
'   4. VALIDATION: Performs an early exit if the core mapping fields are
'      empty, ensuring the pipeline doesn't fail later.
'
' USAGE:
'   - Called by 'IterativeSearch' to configure the 'ProcessInConcatMode'
'     transformation lifecycle.
' ==========================================================================
Private Function ReadConcatSettings(ByRef ctx As sqlContext, ByRef rs As Object) As ConcatSettings
    Dim s As ConcatSettings
    s.Enabled = False
    
    If Not HasField(rs, ctx.fields.concatenateSwitch) Then
        ReadConcatSettings = s
        Exit Function
    End If

    Dim rawValue As Variant
    rawValue = rs.fields(ctx.fields.concatenateSwitch).value
    
    ' Treat as true if: real True, "1", "TRUE", "true", "yes", non-zero number
    If IsNull(rawValue) Then
        s.Enabled = False
    ElseIf VarType(rawValue) = vbBoolean Then
        s.Enabled = rawValue = True
    ElseIf IsNumeric(rawValue) Then
        s.Enabled = CLng(rawValue) <> 0
    Else
        Dim strVal As String
        strVal = LCase$(Trim$(CStr(rawValue)))
        s.Enabled = (strVal = "true") Or (strVal = "yes") Or (strVal = "1")
    End If
    
    If Not s.Enabled Then
        ReadConcatSettings = s
        Exit Function
    End If
    
    s.ConcatField = Trim$(Nz(GetFieldValueString(rs, ctx.fields.concatenateField), ""))
    s.TargetField = Trim$(Nz(GetFieldValueString(rs, ctx.fields.concatenateMapTo), ""))
    s.prefix = Nz(GetFieldValueString(rs, ctx.fields.concatenatePrefix), "")
    s.suffix = Nz(GetFieldValueString(rs, ctx.fields.concatenateSuffix), "")
    s.separator = Nz(GetFieldValueString(rs, ctx.fields.concatenateSeparator), "")
    
    ' Optional: early exit if required fields missing
    If Len(s.ConcatField) = 0 Or Len(s.TargetField) = 0 Then
        s.Enabled = False
    End If
    
    ReadConcatSettings = s
End Function

' ==========================================================================
' FUNCTION: Nz (Null-to-Zero/String)
' PURPOSE:
'   A robust implementation of the Access-style 'Nz' function for VBA.
'
' TECHNICAL WORKFLOW:
'   1. NULL CHECK: Evaluates the input variant for 'IsNull'.
'   2. LENGTH CHECK: Coerces the value to a string and checks for zero length.
'   3. FALLBACK: Returns the specified 'def' (default) value if the input
'      is empty; otherwise, returns the original value.
'
' USAGE:
'   - Essential for processing ADO Recordsets where database fields may
'     contain Null values that would otherwise crash VBA string operations.
' ==========================================================================
Private Function Nz(v, Optional def = "") As String: Nz = IIf(IsNull(v) Or Len(v & "") = 0, def, v): End Function

' ==========================================================================
' FUNCTION: GetHeaderRS
' PURPOSE:
'   Executes the initial "Identity Query" to retrieve a list of target IDs
'   for iterative processing.
'
' TECHNICAL WORKFLOW:
'   1. STABILITY PAUSE: Implements a brief 10ms 'SleepMilliseconds' to
'      stabilize the ADO connection before triggering the sub-query.
'   2. RECORDSET EXECUTION: Uses a late-bound 'ADODB.Recordset' to run the
'      'idQuery' in a static, read-only mode.
'   3. STRUCTURAL VALIDATION:
'      - Verifies the 'ID' field exists in the result set.
'      - Checks for empty results (EOF/BOF) and safely closes handles.
'   4. CURSOR MANAGEMENT: Forces 'MoveFirst' to ensure the subsequent
'      processing loop starts from the beginning of the list.
'   5. ERROR LOGGING: Captures SQL syntax errors or missing table issues
'      and routes them to the 'LogDiagnostic' forensic logger.
' ==========================================================================
Private Function GetHeaderRS( _
    ByRef connectionObject As Object, _
    ByVal idQuery As String) As Object
    
    Dim rs As Object

    On Error GoTo GetHeaderRSError
    
    DoEvents
    SleepMilliseconds 10
    
    ' Execute the ID query
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open idQuery, connectionObject, adOpenStatic, adLockReadOnly
    
    ' Guard against empty or invalid recordsets
    If rs Is Nothing Then Exit Function
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Function
    If rs.EOF And rs.BOF Then
        SafeCloseRecordset rs
        Exit Function
    End If
    
    ' Check for required ID field (case-insensitive)
    If Not HasField(rs, "ID") Then
        SafeCloseRecordset rs
        Exit Function
    End If
    
    ' Always start at the beginning
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        SafeCloseRecordset rs
        Exit Function
    End If
    On Error GoTo GetHeaderRSError
    
    Set GetHeaderRS = rs
    
    Exit Function
    
GetHeaderRSError:
        LogDiagnostic _
            "GetHeaderRS SQL failed: " & err.Description & vbNewLine, _
            errorNumber:=err.number, _
            sql:=idQuery, _
            errorCategory:="Iteration / SQL"
            
    On Error Resume Next
    SafeCloseRecordset rs
    Set GetHeaderRS = Nothing
End Function

' ==========================================================================
' FUNCTION: CreateAugmentedRS
' PURPOSE:
'   Constructs a new, in-memory Recordset that merges header data with
'   aggregated detail strings.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA CLONING: Iterates through the 'headerRS' to replicate its
'      field structure in a new 'ADODB.Recordset'.
'   2. VIRTUAL FIELD INJECTION: Appends a high-capacity 'adVarChar' field
'      (8KB) to store the result of the concatenation.
'   3. ITERATIVE SUB-QUERYING: For every row in the header:
'      - Executes a 'RunParameterizedQuery' using the row's unique ID.
'      - Aggregates the resulting detail rows via 'ConcatenateFieldValues'.
'   4. STRING WRAPPING: Applies user-defined 'prefix' and 'suffix' to the
'      final concatenated string.
'   5. DATA CONSOLIDATION: Populates the augmented recordset with the
'      merged data, providing a single source for the mapping engine.
' ==========================================================================
Private Function CreateAugmentedRS( _
    ByRef conn As Object, _
    ByRef ctx As sqlContext, _
    ByVal headerRS As Object, _
    ByVal dataTemplate As String, _
    ByVal concatFld As String, _
    ByVal targetFld As String, _
    ByVal prefix As String, _
    ByVal suffix As String, _
    ByVal separator As String) As Object

    If headerRS Is Nothing Or headerRS.State <> ObjectStateEnum.adStateOpen Then Exit Function

    On Error GoTo ErrHandler

    Dim aug As Object
    Set aug = CreateObject("ADODB.Recordset")

    ' Copy all fields from header
    Dim f As Object
    For Each f In headerRS.fields
        aug.fields.Append f.name, f.Type, f.DefinedSize
    Next

    ' Ensure target field exists
    If Not HasField(aug, targetFld) Then
        aug.fields.Append targetFld, ADODataTypeEnum.adVarChar, 8192  ' 8K to be able to handle long concatenations
    End If

    aug.Open

    headerRS.MoveFirst
    Do While Not headerRS.EOF
        aug.AddNew

        ' Copy header fields
        For Each f In headerRS.fields
            aug(f.name) = headerRS(f.name)
        Next

        ' Run detail query
        Dim rsDetail As Object
        Set rsDetail = RunParameterizedQuery(conn, ctx, dataTemplate, headerRS("ID"))

        ' Build concatenated string using extracted function
        Dim concat As String
        concat = ConcatenateFieldValues(rsDetail, concatFld, separator)

        ' Apply prefix/suffix
        concat = prefix & concat & suffix

        aug(targetFld) = concat
        aug.Update

        SafeCloseRecordset rsDetail

        headerRS.MoveNext
    Loop

    ' Prepare for reading/mapping
    If Not (aug.EOF And aug.BOF) Then
        aug.MoveFirst
    End If

    Set CreateAugmentedRS = aug
    Exit Function

ErrHandler:
    LogDiagnostic _
        "CreateAugmentedRS failed: " & err.Description, _
        errorNumber:=err.number, _
        errorCategory:="Iteration / Concatenation"

    On Error Resume Next
    SafeCloseRecordset aug
    Set CreateAugmentedRS = Nothing
End Function

' ==========================================================================
' FUNCTION: ConcatenateFieldValues
' PURPOSE:
'   Collapses multiple rows of a specific field into a single delimited string.
'
' TECHNICAL WORKFLOW:
'   1. VALIDATION: Verifies the recordset is open, non-empty, and contains
'      the requested field name.
'   2. SAFE ITERATION: Loops through all records using 'SafeStr' to handle
'      potential Null values without crashing.
'   3. DELIMITER LOGIC: Uses a boolean 'first' flag to ensure the 'separator'
'      is only placed between values (avoiding trailing delimiters).
'   4. VALUE FILTERING: Only appends non-empty strings, ensuring the final
'      label is compact and clean.
'
' USAGE:
'   - Called by 'CreateAugmentedRS' during "Concat Mode" processing.
'   - Enables the creation of multi-line labels (e.g., using '\n' as a
'     separator) from a list of related records.
' ==========================================================================
Private Function ConcatenateFieldValues( _
    ByVal rs As Object, _
    ByVal fieldName As String, _
    ByVal separator As String) As String

    Dim result As String
    Dim first As Boolean
    first = True

    If rs Is Nothing Then Exit Function
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Function
    If rs.EOF And rs.BOF Then Exit Function

    If Not HasField(rs, fieldName) Then Exit Function

    rs.MoveFirst
    Do While Not rs.EOF
        Dim v As String
        v = SafeStr(rs.fields(fieldName).value)

        ' Only include non-empty values (user can filter NULLs in SQL if desired)
        If Len(v) > 0 Then
            If Not first Then result = result & separator
            result = result & v
            first = False
        End If

        rs.MoveNext
    Loop

    ConcatenateFieldValues = result
End Function

' ==========================================================================
' FUNCTION: GetIDList
' PURPOSE:
'   Executes a primary "Identity Query" and returns a unique set of IDs.
'
' TECHNICAL WORKFLOW:
'   1. STABILITY DELAY: Executes a brief 10ms 'SleepMilliseconds' to
'      prevent COM collisions before opening the identity recordset.
'   2. FIELD DISCOVERY: Scans the result set for a field named 'id'
'      (case-insensitive) to use as the primary key for iteration.
'   3. DE-DUPLICATION: Uses a 'Scripting.Dictionary' to store unique
'      values, ensuring each ID is only processed once.
'   4. DATA SANITIZATION: Employs 'SafeStr' to handle Nulls or non-string
'      data types during the extraction loop.
'   5. ERROR LOGGING: Captures query failures or missing 'id' fields and
'      routes forensic data to the 'LogDiagnostic' subsystem.
'
' USAGE:
'   - Provides the 'Identity List' used to drive parameterized sub-queries
'     in the 'IterativeSearch' workflow.
' ==========================================================================
Private Function GetIDList( _
    ByRef connectionObject As Object, _
    ByVal idQuery As String) As Object

    Dim rs As Object
    On Error GoTo GetIDListError

    DoEvents
    SleepMilliseconds 10

    ' Execute the ID query
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open idQuery, connectionObject, adOpenStatic, adLockReadOnly

    ' Guard against empty or invalid recordsets
    If rs Is Nothing Then Exit Function
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Function
    If rs.EOF And rs.BOF Then
        SafeCloseRecordset rs
        Exit Function
    End If

    ' Create dictionary
    Dim idList As Object
    Set idList = CreateObject("Scripting.Dictionary")

    ' Find the ID field
    Dim idFieldIndex As Long
    idFieldIndex = -1

    Dim f As Long
    For f = 0 To rs.fields.count - 1
        If LCase$(SafeStr(rs.fields(f).name)) = "id" Then
            idFieldIndex = f
            Exit For
        End If
    Next f

    ' No ID field -> return Nothing
    If idFieldIndex = -1 Then
        SafeCloseRecordset rs
        Exit Function
    End If

    ' Always start at the beginning
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        SafeCloseRecordset rs
        Exit Function
    End If
    On Error GoTo GetIDListError

    ' Collect IDs
    Dim id As String
    Do While Not rs.EOF

        id = SafeStr(rs.fields(idFieldIndex).value)

        If Len(id) > 0 Then
            If Not idList.Exists(id) Then
                idList.Add id, True
            End If
        End If

        rs.MoveNext
    Loop

    ' Cleanup
    SafeCloseRecordset rs

    ' Return the dictionary
    Set GetIDList = idList
    Exit Function

GetIDListError:
    LogDiagnostic _
        "GetIDList SQL failed: " & err.Description & vbNewLine, _
        errorNumber:=err.number, _
        sql:=idQuery, _
        errorCategory:="Iteration / SQL"

    On Error Resume Next
    SafeCloseRecordset rs
    Set GetIDList = Nothing
End Function

' ==========================================================================
' FUNCTION: RunParameterizedQuery
' PURPOSE:
'   Resolves a SQL template by injecting a specific ID and executes the
'   resulting query.
'
' TECHNICAL WORKFLOW:
'   1. TOKEN SUBSTITUTION: Replaces the 'idPlaceholder' token (e.g., "{ID}")
'      within the dataTemplate with the active identity value.
'   2. LATE-BOUND EXECUTION: Instantiates a fresh 'ADODB.Recordset' and
'      opens it in a static, read-only state.
'   3. STATE VALIDATION: Verifies the resulting recordset is truly open,
'      providing a guard against providers that fail silently.
'   4. DIAGNOSTIC TRACING: If the query fails (e.g., due to syntax in the
'      template), the finalized SQL string is captured and logged for
'      forensic analysis.
'
' USAGE:
'   - The core "Worker" function for 'ProcessInClassicMode' and
'     'CreateAugmentedRS'.
' ==========================================================================
Private Function RunParameterizedQuery( _
    ByRef connectionObject As Object, _
    ByRef ctx As sqlContext, _
    ByVal dataQueryTemplate As String, _
    ByVal id As Variant) As Object

    Dim rsData As Object
    Dim sql As String

    On Error GoTo RunQueryError

    ' Substitute placeholder for the derived ID (Null-safe)
    sql = replace(dataQueryTemplate, ctx.fields.idPlaceholder, SafeStr(id), , , vbTextCompare)

    ' Create fresh recordset
    Set rsData = CreateObject("ADODB.Recordset")
    rsData.CursorLocation = adUseClient

    rsData.Open sql, connectionObject, adOpenStatic, adLockReadOnly

    ' Guard against providers that return a closed recordset
    If rsData Is Nothing Then
        Set RunParameterizedQuery = Nothing
        Exit Function
    End If

    If rsData.State <> ObjectStateEnum.adStateOpen Then
        SafeCloseRecordset rsData
        Set RunParameterizedQuery = Nothing
        Exit Function
    End If

    Set RunParameterizedQuery = rsData
    Exit Function

RunQueryError:
    LogDiagnostic _
        "RunParameterizedQuery failed: " & err.Description, _
        errorNumber:=err.number, _
        sql:=sql, _
        errorCategory:=ClassifyError(err.Description)

    On Error Resume Next
    SafeCloseRecordset rsData
    Set RunParameterizedQuery = Nothing
End Function

' ==========================================================================
' PROCEDURE: RecursiveSearch
' PURPOSE:
'   Orchestrates the discovery of hierarchical tree structures from flat
'   relational tables.
'
' TECHNICAL WORKFLOW:
'   1. SIGNATURE DETECTION: Verifies the recordset contains recursive
'      parameters: 'TREE_QUERY', 'WHERE_VALUE', and 'WHERE_COLUMN'.
'   2. PARAMETER EXTRACTION: Retrieves the SQL template and identifies the
'      starting "Anchor" point for the recursion.
'   3. INFINITE LOOP PROTECTION: Initializes a 'Scripting.Dictionary'
'      (searchedIDs) to track visited nodes, ensuring the engine doesn't
'      re-process the same branch in cyclical data.
'   4. DEPTH GOVERNOR: Establishes a 'maxDepth' limit (defaulting to 100)
'      to prevent runaway execution in massive or malformed datasets.
'   5. EXECUTION HANDOFF: Launches the recursive engine
'      (PerformRecursiveSearch) to crawl the data structure.
'
' USAGE:
'   - Essential for visualizing organizational charts, folder structures,
'     and bill-of-materials (BOM) from standard SQL databases.
' ==========================================================================
Private Sub RecursiveSearch(ByRef connectionObject As Object, _
                                  ByRef ctx As sqlContext, _
                                  ByVal rs As Object, _
                                  ByRef rsRecursion As Object)
    
    If rs.EOF Then Exit Sub
    If Not HasField(rs, ctx.fields.treeQuery) Then Exit Sub
    
    Dim recursionSql As String
    Dim whereValue As String
    Dim whereColumn As String
    
    ' Extract the query and parameters. Exit if not provided
    recursionSql = GetFieldValueString(rs, ctx.fields.treeQuery)
    If Len(recursionSql) = 0 Then Exit Sub
    
    whereValue = GetFieldValueString(rs, ctx.fields.whereValue)
    If Len(whereValue) = 0 Then Exit Sub
    
    whereColumn = GetFieldValueString(rs, ctx.fields.whereColumn)
    If Len(whereColumn) = 0 Then Exit Sub
    
    ' Create a collection to track what has been searched, so we
    ' don't fall into an infinite loop.
    Dim searchedIDs As Object
    Set searchedIDs = CreateObject("Scripting.Dictionary")
    
    ' Place limits on how many recursive calls can be made
    Dim maxDepth As Long
    maxDepth = GetFieldValueLong(rs, ctx.fields.maxDepth)
    
    If maxDepth = 0 Then
        maxDepth = DEFAULT_MAX_RECURSION_DEPTH
    End If
    
    Dim currentDepth As Long
    currentDepth = 0
    
    ' Execute SQL recursively until all branches of the tree are followed
    PerformRecursiveSearch connectionObject, ctx, recursionSql, whereValue, whereColumn, currentDepth, maxDepth, rsRecursion, searchedIDs
    
End Sub
  
' ==========================================================================
' SECTION: TYPE-SAFE FIELD ACCESSORS
' ==========================================================================

''
' STRING ACCESSOR: Safely retrieves and trims a field value.
' 1. Field Check: Verifies the field exists via 'HasField' to avoid runtime errors.
' 2. Sanitization: Coerces the value to a String and applies 'Trim$' to remove
'    unwanted whitespace from database padding.
'
Private Function GetFieldValueString(ByVal recordSet As Object, ByRef fieldName As String) As String
    If HasField(recordSet, fieldName) Then
        GetFieldValueString = Trim$(CStr(recordSet.fields(fieldName).value))
    Else
        GetFieldValueString = vbNullString
    End If
End Function

''
' NUMERIC ACCESSOR: Safely retrieves a field as a Long integer.
' 1. Logic: Attempts to cast the field value using 'CLng'.
' 2. Error Resilience: Employs a local ErrorHandler to return 0 if the field
'    contains non-numeric data or is missing.
'
Private Function GetFieldValueLong(ByVal recordSet As Object, ByRef fieldName As String) As Long
    On Error GoTo ErrorHandler
    If HasField(recordSet, fieldName) Then
        GetFieldValueLong = CLng(recordSet.fields(fieldName).value)
    Else
        GetFieldValueLong = 0
    End If
    Exit Function
    
ErrorHandler:
    GetFieldValueLong = 0
End Function

''
' BOOLEAN ACCESSOR: Safely retrieves a field as a Boolean.
' 1. Null-Safety: Specifically checks for 'IsNull' to ensure database Nulls
'    are treated as 'False' rather than triggering a VBA error.
' 2. Flexibility: Casts the result via 'CBool', allowing for bitwise,
'    numeric, or logical Boolean interpretation from the driver.
'
Private Function GetFieldValueBoolean(ByVal recordSet As Object, ByRef fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    If HasField(recordSet, fieldName) Then
        Dim v As Variant
        v = recordSet.fields(fieldName).value

        ' Null-safe: treat Null as False
        If IsNull(v) Then
            GetFieldValueBoolean = False
        Else
            GetFieldValueBoolean = CBool(v)
        End If
    Else
        GetFieldValueBoolean = False
    End If

    Exit Function

ErrorHandler:
    GetFieldValueBoolean = False
End Function

' ==========================================================================
' PROCEDURE: PerformRecursiveSearch
' PURPOSE:
'   The core recursive algorithm for following data relationships (e.g.,
'   Parent-Child) through a relational database.
'
' TECHNICAL WORKFLOW:
'   1. CYCLE DETECTION: Calls 'WasAlreadySearched' to check the 'searchedIDs'
'      registry, preventing infinite loops in circular data.
'   2. DEPTH GUARD: Increments the current depth and exits if it exceeds
'      the 'maxDepth' safety threshold.
'   3. TOKEN INJECTION: Dynamically replaces the '{WHERE_VALUE}' placeholder
'      in the SQL string with the current ID to find the next level of data.
'   4. SCHEMA UNIFICATION: If it's the first result found, it clones the
'      recordset structure to create a master 'recursionRecordSet' buffer.
'   5. ROW ACCUMULATION: Iterates through the children, copying their
'      data into the master buffer and then recursively calling itself
'      using the child's ID as the new 'nextValue'.
'   6. ERROR FORENSICS: Captures the exact SQL, depth, and specific ID
'      being processed during a failure for detailed log reporting.
' ==========================================================================
Private Sub PerformRecursiveSearch( _
    ByRef connectionObject As Object, _
    ByRef ctx As sqlContext, _
    ByVal sqlStatement As String, _
    ByRef whereValue As String, _
    ByVal whereColumn As String, _
    ByVal depth As Long, _
    ByVal maxDepth As Long, _
    ByRef recursionRecordSet As Object, _
    ByRef searchedIDs As Object)

    On Error GoTo RecursionError

    DoEvents
    SleepMilliseconds 10

    ' Base case: already searched?
    If WasAlreadySearched(whereValue, searchedIDs) Then Exit Sub

    ' Depth limit
    Dim currentDepth As Long
    currentDepth = depth + 1
    If currentDepth > maxDepth Then Exit Sub

    ' Expand placeholder
    Dim query As String
    query = replace(sqlStatement, "{" & ctx.fields.whereValue & "}", SafeStr(whereValue), , , vbTextCompare)

    ' Mark this ID as searched
    AddToSearchedList whereValue, searchedIDs

    ' Execute recursive query
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.Open query, connectionObject, adOpenStatic, adLockReadOnly

    ' Guard against invalid or empty recordsets
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then
        SafeCloseRecordset rs
        Exit Sub
    End If

    ' Initialize merged recordset structure
    If recursionRecordSet Is Nothing Then
        Set recursionRecordSet = CreateObject("ADODB.Recordset")

        Dim fieldNumber As Long
        For fieldNumber = 0 To rs.fields.count - 1
            recursionRecordSet.fields.Append _
                rs.fields(fieldNumber).name, _
                rs.fields(fieldNumber).Type, _
                rs.fields(fieldNumber).DefinedSize
        Next fieldNumber

        recursionRecordSet.Open
    End If

    ' Iterate through results
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        SafeCloseRecordset rs
        Exit Sub
    End If
    On Error GoTo RecursionError

    Do While Not rs.EOF

        ' Append row to merged recordset
        recursionRecordSet.AddNew
        For fieldNumber = 0 To rs.fields.count - 1
            recursionRecordSet.fields(fieldNumber).value = _
                SafeFieldValue(rs, rs.fields(fieldNumber).name)
        Next fieldNumber
        recursionRecordSet.Update

        ' Recurse using safe field value
        Dim nextValue As String
        nextValue = SafeFieldValue(rs, whereColumn)

        PerformRecursiveSearch _
            connectionObject, ctx, sqlStatement, nextValue, _
            whereColumn, currentDepth, maxDepth, recursionRecordSet, searchedIDs

        rs.MoveNext
    Loop

    SafeCloseRecordset rs
    Exit Sub

RecursionError:
    LogDiagnostic _
        "Recursive SQL failed: " & err.Description & vbNewLine & _
        "  Query: " & query & vbNewLine & _
        "  whereValue   = " & whereValue & vbNewLine & _
        "  whereColumn  = " & whereColumn & vbNewLine & _
        "  currentDepth = " & CStr(currentDepth) & vbNewLine & _
        "  maxDepth     = " & CStr(maxDepth) & vbNewLine, _
        errorNumber:=err.number, _
        sql:=query, _
        errorCategory:="Recursion / SQL"

    On Error Resume Next
    SafeCloseRecordset rs
    On Error GoTo 0
End Sub

' ==========================================================================
' PROCEDURE: AddToSearchedList
' PURPOSE:
'   Registers a unique identifier in the "visited" registry.
'
' TECHNICAL WORKFLOW:
'   1. TYPE COERCION: Converts the ID to a String (CStr) to ensure
'      dictionary keys remain consistent regardless of source data type.
'   2. REGISTRATION: Adds the key to the 'searchedIDs' Dictionary object
'      with a dummy boolean value.
'
' USAGE:
'   - Called by 'PerformRecursiveSearch' after a node is identified.
'   - Essential for cycle detection in hierarchical datasets.
' ==========================================================================
Private Sub AddToSearchedList(ByRef rowId As Variant, ByVal searchedIDs As Object)
    ' Add the ID to the dictionary
    searchedIDs.Add CStr(rowId), True
End Sub

' ==========================================================================
' FUNCTION: WasAlreadySearched
' PURPOSE:
'   Determines if a node ID has already been visited during a recursive crawl.
'
' TECHNICAL WORKFLOW:
'   1. KEY COERCION: Converts the 'rowId' to a String to match the
'      Dictionary's internal key format.
'   2. EXISTENCE CHECK: Uses the 'Exists' method of the 'searchedIDs'
'      Dictionary for an O(1) high-speed lookup.
'
' USAGE:
'   - The primary safety check in 'PerformRecursiveSearch'.
'   - If True, the search branch is terminated immediately to break
'     circular references.
' ==========================================================================
Private Function WasAlreadySearched(ByRef rowId As Variant, ByVal searchedIDs As Object) As Boolean
    ' Check if the ID is already in the dictionary
    WasAlreadySearched = searchedIDs.Exists(CStr(rowId))
End Function

' ==========================================================================
' PROCEDURE: MapResultsToDataWorksheet
' PURPOSE:
'   Determines the optimal mapping strategy to translate SQL records into
'   Excel 'Data' worksheet rows.
'
' TECHNICAL WORKFLOW:
'   1. STATE VALIDATION: Verifies the ADO Recordset is open and contains data.
'   2. SIGNATURE DETECTION: Checks for "Pseudo-SQL" fields (CreateEdges or
'      CreateRank) and diverts to those specialized generators if found.
'   3. HIERARCHY ANALYSIS: Identifies the presence of 'CLUSTER' and
'      'SUBCLUSTER' columns in the result set.
'   4. ROUTING LOGIC: Uses a nested decision tree to dispatch the data to
'      one of four specific processors based on grouping requirements:
'      - ProcessClusterYesSubclusterYes
'      - ProcessClusterYesSubclusterNo
'      - ProcessClusterNoSubclusterYes
'      - ProcessClusterNoSubclusterNo
'
' USAGE:
'   - The final stage of the 'executeSQL' pipeline.
'   - Ensures that the visual hierarchy of the graph matches the
'     relational grouping in the database.
' ==========================================================================
Private Sub MapResultsToDataWorksheet( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Exit early if invalid or empty
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Special-case overrides: edges and rank
    If HasField(rs, ctx.fields.CreateEdges) Then
        CreateEdges ctx, rs, row, recordCnt
        Exit Sub
    End If

    If HasField(rs, ctx.fields.CreateRank) Then
        CreateRank ctx, rs, row, recordCnt
        Exit Sub
    End If

    ' Determine cluster/subcluster presence
    Dim hasCluster As Boolean
    Dim hasSubcluster As Boolean

    hasCluster = HasField(rs, ctx.fields.Cluster)
    hasSubcluster = HasField(rs, ctx.fields.subcluster)

    ' Always start from BOF before dispatching
    rs.MoveFirst

    ' Dispatch to the correct processing routine
    If hasCluster Then
        If hasSubcluster Then
            ProcessClusterYesSubclusterYes ctx, rs, row, recordCnt
        Else
            ProcessClusterYesSubclusterNo ctx, rs, row, recordCnt
        End If
    Else
        If hasSubcluster Then
            ProcessClusterNoSubclusterYes ctx, rs, row, recordCnt
        Else
            ProcessClusterNoSubclusterNo ctx, rs, row, recordCnt
        End If
    End If

End Sub

' ==========================================================================
' PROCEDURE: ProcessClusterYesSubclusterYes
' PURPOSE:
'   Orchestrates the mapping of SQL results into a three-tier hierarchy:
'   Cluster -> Subcluster -> Rows.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Uses 'GetClusterInfo' and 'GetSubClusterInfoForCluster'
'      to build a dictionary-based map of the result set's hierarchy.
'   2. NESTED EMISSION:
'      - Opens a 'Cluster' block and iterates through its 'Subclusters'.
'      - Opens each 'Subcluster' block and emits the corresponding data rows.
'      - Handles "orphan" rows that belong to a cluster but no specific subcluster.
'   3. ORPHAN MANAGEMENT: Performs a final pass via 'GetOrphanSubClusterInfo'
'      to capture rows that belong to a subcluster but have no parent cluster.
'   4. ROOT-LEVEL EMISSION: Processes any remaining rows where both cluster
'      and subcluster fields are null.
'   5. COORDINATION: Manages 'EmitClusterOpen/Close' calls to ensure valid
'      Graphviz syntax (matching braces) in the final Data worksheet output.
' ==========================================================================
Private Sub ProcessClusterYesSubclusterYes( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    Dim clusterKey As Variant
    Dim subclusterKey As Variant

    Dim clusterCnt As Long
    Dim subclusterCnt As Long

    Dim clusterList As Dictionary
    Dim orphanClusterList As Dictionary

    Dim clusterInstance As Cluster
    Dim clusterRecord As Cluster
    Dim subclusterRecord As Cluster

    ' Guard against invalid recordsets
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Collect distinct clusters
    Set clusterList = GetClusterInfo(rs, ctx.fields)

    If clusterList.count > 0 Then
        ' Attach subcluster dictionaries to each cluster
        For Each clusterKey In clusterList.Keys()
            Set clusterInstance = clusterList.item(clusterKey)
            Set clusterInstance.subclusters = GetSubClusterInfoForCluster( _
                                                rs, _
                                                ctx.fields, _
                                                CStr(clusterKey))
        Next clusterKey
    End If

    ' Emit clusters (with or without subclusters)
    For Each clusterKey In clusterList.Keys()

        clusterCnt = clusterCnt + 1
        Set clusterRecord = clusterList.item(CStr(clusterKey))

        EmitClusterOpen clusterRecord, ctx.dataLayout, row, _
                        ctx.fields.clusterPlaceholder, clusterCnt

        If clusterRecord.subclusters.count = 0 Then
            ' No subclusters: emit all rows for this cluster
            On Error Resume Next
            rs.MoveFirst
            If err.number <> 0 Then
                err.Clear
                Exit For
            End If
            On Error GoTo 0

            Do While Not rs.EOF
                If SafeFieldValue(rs, ctx.fields.Cluster) = CStr(clusterKey) Then
                    EmitRows ctx, rs, row, recordCnt
                End If
                rs.MoveNext
            Loop

        Else
            ' Has subclusters: group rows by cluster + subcluster
            subclusterCnt = 0

            For Each subclusterKey In clusterRecord.subclusters.Keys()

                Set subclusterRecord = clusterRecord.subclusters.item(subclusterKey)

                On Error Resume Next
                rs.MoveFirst
                If err.number <> 0 Then
                    err.Clear
                    Exit For
                End If
                On Error GoTo 0

                subclusterCnt = subclusterCnt + 1

                EmitClusterOpen subclusterRecord, ctx.dataLayout, row, _
                                ctx.fields.subclusterPlaceholder, subclusterCnt

                Do While Not rs.EOF
                    If SafeFieldValue(rs, ctx.fields.Cluster) = CStr(clusterKey) _
                       And SafeFieldValue(rs, ctx.fields.subcluster) = CStr(subclusterKey) Then
                        EmitRows ctx, rs, row, recordCnt
                    End If
                    rs.MoveNext
                Loop

                EmitClusterClose subclusterRecord, ctx.dataLayout, row, _
                                 ctx.fields.subclusterPlaceholder, subclusterCnt

                ' Emit rows in this cluster with NULL subcluster
                On Error Resume Next
                rs.MoveFirst
                If err.number <> 0 Then
                    err.Clear
                    Exit For
                End If
                On Error GoTo 0

                Do While Not rs.EOF
                    If SafeFieldValue(rs, ctx.fields.Cluster) = CStr(clusterKey) _
                       And SafeFieldValue(rs, ctx.fields.subcluster) = "" Then
                        EmitRows ctx, rs, row, recordCnt
                    End If
                    rs.MoveNext
                Loop

            Next subclusterKey
        End If

        EmitClusterClose clusterRecord, ctx.dataLayout, row, _
                         ctx.fields.clusterPlaceholder, clusterCnt
    Next clusterKey

    ' Handle case where cluster has no data, but subcluster does
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Set orphanClusterList = GetOrphanSubClusterInfo(rs, ctx.fields)
    subclusterCnt = 0

    For Each subclusterKey In orphanClusterList.Keys()

        Set subclusterRecord = orphanClusterList.item(subclusterKey)

        On Error Resume Next
        rs.MoveFirst
        If err.number <> 0 Then
            err.Clear
            Exit For
        End If
        On Error GoTo 0

        subclusterCnt = subclusterCnt + 1

        EmitClusterOpen subclusterRecord, ctx.dataLayout, row, _
                        ctx.fields.subclusterPlaceholder, subclusterCnt

        Do While Not rs.EOF
            If SafeFieldValue(rs, ctx.fields.Cluster) = "" _
               And SafeFieldValue(rs, ctx.fields.subcluster) = CStr(subclusterKey) Then
                EmitRows ctx, rs, row, recordCnt
            End If
            rs.MoveNext
        Loop

        EmitClusterClose subclusterRecord, ctx.dataLayout, row, _
                         ctx.fields.subclusterPlaceholder, subclusterCnt
    Next subclusterKey

    ' Handle rows where both cluster and subcluster are NULL
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not rs.EOF
        If SafeFieldValue(rs, ctx.fields.Cluster) = "" _
           And SafeFieldValue(rs, ctx.fields.subcluster) = "" Then
            EmitRows ctx, rs, row, recordCnt
        End If
        rs.MoveNext
    Loop

End Sub

' ==========================================================================
' PROCEDURE: ProcessClusterYesSubclusterNo
' PURPOSE:
'   Groups SQL results into a single tier of clusters on the Data worksheet.
'
' TECHNICAL WORKFLOW:
'   1. CLUSTER DISCOVERY: Calls 'GetClusterInfo' to extract the unique set
'      of group identifiers from the recordset.
'   2. BLOCK EMISSION: Iterates through each unique cluster:
'      - Opens a 'Cluster' block via 'EmitClusterOpen' (creating the '{' row).
'      - Scans the recordset to emit only the rows belonging to that specific key.
'      - Closes the block via 'EmitClusterClose' (creating the '}' row).
'   3. ORPHAN PASS: Performs a final scan of the recordset to emit rows where
'      the cluster field is null, placing them at the root level of the graph.
'   4. RESILIENCE: Implements safe 'MoveFirst' checks to ensure the ADO
'      cursor state is maintained across multiple scans.
' ==========================================================================
Private Sub ProcessClusterYesSubclusterNo( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid recordsets
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    Dim clusterList As Dictionary
    Set clusterList = GetClusterInfo(rs, ctx.fields)

    Dim clusterCnt As Long
    clusterCnt = 0

    Dim clusterKey As Variant
    Dim clusterRecord As Cluster

    ' Emit each cluster block
    For Each clusterKey In clusterList.Keys()

        clusterCnt = clusterCnt + 1
        Set clusterRecord = clusterList.item(CStr(clusterKey))

        EmitClusterOpen clusterRecord, ctx.dataLayout, row, _
                        ctx.fields.clusterPlaceholder, clusterCnt

        ' Safe MoveFirst
        On Error Resume Next
        rs.MoveFirst
        If err.number <> 0 Then
            err.Clear
            Exit For
        End If
        On Error GoTo 0

        ' Emit rows belonging to this cluster
        Do While Not rs.EOF
            If SafeFieldValue(rs, ctx.fields.Cluster) = CStr(clusterKey) Then
                EmitRows ctx, rs, row, recordCnt
            End If
            rs.MoveNext
        Loop

        EmitClusterClose clusterRecord, ctx.dataLayout, row, _
                         ctx.fields.clusterPlaceholder, clusterCnt
    Next clusterKey

    ' Emit orphan rows (cluster column is Null)
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not rs.EOF
        If SafeFieldValue(rs, ctx.fields.Cluster) = "" Then
            EmitRows ctx, rs, row, recordCnt
        End If
        rs.MoveNext
    Loop

End Sub

' ==========================================================================
' PROCEDURE: ProcessClusterNoSubclusterYes
' PURPOSE:
'   Groups SQL results into subclusters at the root level of the graph.
'
' TECHNICAL WORKFLOW:
'   1. SUBCLUSTER DISCOVERY: Calls 'GetSubclusterInfo' to extract unique
'      identifiers for the secondary grouping tier.
'   2. NESTED EMISSION:
'      - Iterates through unique subcluster keys.
'      - Emits 'EmitClusterOpen' (using the Subcluster placeholder).
'      - Filters the recordset to emit matching rows via 'EmitRows'.
'      - Closes the block via 'EmitClusterClose'.
'   3. ORPHAN DATA PASS: Scans for rows where the subcluster field is Null,
'      ensuring they are placed outside any bounding boxes at the root.
'   4. CURSOR STABILITY: Employs 'On Error Resume Next' with 'rs.MoveFirst'
'      to handle potential forward-only recordset limitations gracefully.
' ==========================================================================
Private Sub ProcessClusterNoSubclusterYes( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid recordsets
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    Dim subclusterList As Dictionary
    Set subclusterList = GetSubclusterInfo(rs, ctx.fields)

    Dim subclusterCnt As Long
    subclusterCnt = 0

    Dim subclusterKey As Variant
    Dim subclusterRecord As Cluster

    ' Emit each subcluster block
    For Each subclusterKey In subclusterList.Keys()

        subclusterCnt = subclusterCnt + 1
        Set subclusterRecord = subclusterList.item(CStr(subclusterKey))

        EmitClusterOpen subclusterRecord, ctx.dataLayout, row, _
                        ctx.fields.subclusterPlaceholder, subclusterCnt

        ' Safe MoveFirst
        On Error Resume Next
        rs.MoveFirst
        If err.number <> 0 Then
            err.Clear
            Exit For
        End If
        On Error GoTo 0

        ' Emit rows belonging to this subcluster
        Do While Not rs.EOF
            If SafeFieldValue(rs, ctx.fields.subcluster) = CStr(subclusterKey) Then
                EmitRows ctx, rs, row, recordCnt
            End If
            rs.MoveNext
        Loop

        EmitClusterClose subclusterRecord, ctx.dataLayout, row, _
                         ctx.fields.subclusterPlaceholder, subclusterCnt
    Next subclusterKey

    ' Emit orphan rows (subcluster column is Null)
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not rs.EOF
        If SafeFieldValue(rs, ctx.fields.subcluster) = "" Then
            EmitRows ctx, rs, row, recordCnt
        End If
        rs.MoveNext
    Loop

End Sub

' ==========================================================================
' PROCEDURE: ProcessClusterNoSubclusterNo
' PURPOSE:
'   Maps SQL results directly to the Data worksheet without graph grouping.
'
' TECHNICAL WORKFLOW:
'   1. STATE VALIDATION: Performs a final check to ensure the Recordset is
'      open and contains active records.
'   2. CURSOR RESET: Forces 'rs.MoveFirst' to ensure the mapping starts from
'      the first returned record, with an early exit if the cursor is invalid.
'   3. LINEAR EMISSION: Iterates through the recordset in a single pass:
'      - Calls 'EmitRows' for every record to populate the worksheet.
'      - Increments the 'row' and 'recordCnt' trackers for the session.
'
' USAGE:
'   - Triggered when the SQL query lacks both 'CLUSTER' and 'SUBCLUSTER' fields.
'   - The highest-performance mapping mode for flat node/edge lists.
' ==========================================================================
Private Sub ProcessClusterNoSubclusterNo( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid or empty recordsets
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Always start at the beginning
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Emit each row
    Do While Not rs.EOF
        EmitRows ctx, rs, row, recordCnt
        rs.MoveNext
    Loop

End Sub

' ==========================================================================
' PROCEDURE: CreateEdges
' PURPOSE:
'   Generates a continuous chain of relationships (A->B, B->C) from a
'   sequence of nodes.
'
' TECHNICAL WORKFLOW:
'   1. SIGNATURE DETECTION: Triggered when the 'CREATE_EDGES' pseudo-field
'      is present in the SQL result set.
'   2. LOOP MODE (Enumeration):
'      - Mathematically generates IDs using the 'enumeratePlaceholder' (e.g., {i}).
'      - Automatically links ID(i) to ID(i+1) without requiring database rows.
'      - Respects the 'LOOP_MAX_STEPS' governor to prevent infinite cycles.
'   3. NORMAL MODE (Record-based):
'      - Peeks at the next record to establish the 'Related Item' (Head).
'      - Uses the current record's ID as the 'Item' (Tail).
'      - Automatically "slides" the tail to the previous head for each
'        subsequent row to create a connected path.
'
' USAGE:
'   - Ideal for visualizing process flows, timelines, or mathematical
'     sequences where every node links to its successor.
' ==========================================================================
Private Sub CreateEdges( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid or empty recordsets
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Safe MoveFirst
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim item As String
    item = GetFieldValueString(rs, ctx.headings.item)

    Dim relatedItem As String

    ' ------------------------------------------------------------
    ' LOOP MODE
    ' ------------------------------------------------------------
    If ctx.loop.Enabled Then

        Dim stopValue As Long
        stopValue = ctx.loop.stopAt - 1

        Dim i As Long
        For i = ctx.loop.startAt To stopValue Step ctx.loop.stepBy

            ctx.loop.count = ctx.loop.count + 1
            If ctx.loop.count > ctx.loop.max Then Exit For

            EmitOneRow ctx, rs, row, recordCnt, i

            DataSheet.Cells.item(row, ctx.dataLayout.itemColumn) = _
                replace(SafeStr(item), ctx.fields.enumeratePlaceholder, CStr(i), , , vbTextCompare)

            DataSheet.Cells.item(row, ctx.dataLayout.isRelatedToItemColumn) = _
                replace(SafeStr(item), ctx.fields.enumeratePlaceholder, CStr(i + 1), , , vbTextCompare)
            
            row = row + 1
        Next i

    ' ------------------------------------------------------------
    ' NORMAL MODE
    ' ------------------------------------------------------------
    Else

        ' Safe MoveNext (skip first row)
        On Error Resume Next
        rs.MoveNext
        If err.number <> 0 Then
            err.Clear
            Exit Sub
        End If
        On Error GoTo 0

        Dim emittedRow As Long
        
        Do While Not rs.EOF

            relatedItem = GetFieldValueString(rs, ctx.headings.item)

            emittedRow = row   ' capture before EmitRows increments it
            EmitRows ctx, rs, row, recordCnt

            DataSheet.Cells.item(emittedRow, ctx.dataLayout.itemColumn) = item
            DataSheet.Cells.item(emittedRow, ctx.dataLayout.isRelatedToItemColumn) = relatedItem

            item = relatedItem

            rs.MoveNext
        Loop

    End If

End Sub

' ==========================================================================
' PROCEDURE: CreateRank
' PURPOSE:
'   Collapses a recordset into a single 'Native' DOT command to enforce
'   node alignment (ranking).
'
' TECHNICAL WORKFLOW:
'   1. SIGNATURE DETECTION: Triggered when the 'CREATE_RANK' pseudo-field
'      is identified in the SQL result set.
'   2. RANK EXTRACTION: Retrieves the desired Graphviz rank type (e.g., 'same')
'      from the first record.
'   3. STRING AGGREGATION: Iterates through the entire recordset to collect
'      every 'Item' ID into a semicolon-delimited list.
'   4. NATIVE WRAPPING: Wraps the collection in the native Graphviz rank
'      syntax: '{ rank="type"; id1; id2; ... }'.
'   5. EMISSION: Writes a single row to the Data worksheet using the '>'
'      prefix to signal the 'modCreateGraph' engine to treat this as raw DOT code.
'
' USAGE:
'   - Essential for creating "swimlanes" or ensuring that specific nodes
'     (like Start/End) remain at the far edges of the graph.
' ==========================================================================
Private Sub CreateRank( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Exit early if invalid or empty
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Safe MoveFirst
    On Error Resume Next
    rs.MoveFirst
    If err.number <> 0 Then
        err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Establish the rank (Null-safe)
    Dim rank As String
    rank = LCase$(SafeFieldValue(rs, "RANK"))

    ' Collect node identifiers
    Dim item As String
    Dim subgraph As String
    subgraph = "{ rank=" & AddQuotes(rank) & ";"

    Do While Not rs.EOF
        item = SafeFieldValue(rs, ctx.headings.item)
        subgraph = subgraph & " " & AddQuotes(item) & ";"
        rs.MoveNext
    Loop

    subgraph = subgraph & " }"

    ' Emit the row
    recordCnt = recordCnt + 1
    DataSheet.Cells.item(row, ctx.dataLayout.itemColumn) = ">"
    DataSheet.Cells.item(row, ctx.dataLayout.labelColumn) = subgraph
    row = row + 1
End Sub

' ==========================================================================
' FUNCTION: GetClusterInfo
' PURPOSE:
'   Extracts unique parent clusters and their visual attributes from a
'   database recordset.
'
' TECHNICAL WORKFLOW:
'   1. PRE-COMPUTATION: Normalizes field names to lowercase to ensure
'      case-insensitive matching during recordset traversal.
'   2. DATA DISCOVERY: Iterates through the entire recordset to identify
'      unique entries in the 'CLUSTER' column.
'   3. METADATA CAPTURE: For every unique cluster ID found, it captures:
'      - clusterLabel: The display text for the group.
'      - clusterStyleName: The Style worksheet reference for the group.
'      - clusterAttributes: Any 'Extra Attributes' specific to the cluster.
'      - clusterTooltip: Hover-text for SVG/Web interactivity.
'   4. OBJECT CACHING: Instantiates a 'Cluster' class object for each unique
'      group and stores it in a high-speed Dictionary.
'   5. CURSOR HYGIENE: Resets the recordset to the beginning (MoveFirst)
'      after the scan to ensure the subsequent mapping loop is ready.
' ==========================================================================
Private Function GetClusterInfo(ByVal rs As Object, _
                                ByRef fields As sqlFieldName) As Dictionary
    Dim clusters As Dictionary
    Set clusters = New Dictionary

    ' Exit early if empty
    If rs.EOF Then
        Set GetClusterInfo = clusters
        Exit Function
    End If

    ' Precompute lowercase field names for comparison
    Dim clusterField As String:        clusterField = LCase$(fields.Cluster)
    Dim clusterLabelField As String:   clusterLabelField = LCase$(fields.clusterLabel)
    Dim clusterStyleField As String:   clusterStyleField = LCase$(fields.clusterStyleName)
    Dim clusterAttrField As String:    clusterAttrField = LCase$(fields.clusterAttributes)
    Dim clusterTooltipField As String: clusterTooltipField = LCase$(fields.clusterTooltip)

    ' Check once whether CLUSTER LABEL exists
    Dim hasClusterLabel As Boolean
    hasClusterLabel = HasField(rs, fields.clusterLabel)

    Dim fieldObject As Variant
    Dim clusterId As String
    Dim clusterLabel As String
    Dim clusterStyleName As String
    Dim clusterAttributes As String
    Dim clusterTooltip As String

    rs.MoveFirst

    Do While Not rs.EOF
        clusterId = vbNullString
        clusterLabel = vbNullString
        clusterStyleName = vbNullString
        clusterAttributes = vbNullString
        clusterTooltip = vbNullString

        ' Extract cluster metadata
        For Each fieldObject In rs.fields
            Select Case LCase$(fieldObject.name)

                Case clusterField
                    If Not IsNull(fieldObject.value) Then
                        clusterId = CStr(fieldObject.value)
                    End If
                    
                Case clusterLabelField
                    If Not IsNull(fieldObject.value) Then
                        clusterLabel = CStr(fieldObject.value)
                    End If

                Case clusterStyleField
                    If Not IsNull(fieldObject.value) Then
                        clusterStyleName = CStr(fieldObject.value)
                    End If

                Case clusterAttrField
                    If Not IsNull(fieldObject.value) Then
                        clusterAttributes = CStr(fieldObject.value)
                    End If

                Case clusterTooltipField
                    If Not IsNull(fieldObject.value) Then
                        clusterTooltip = CStr(fieldObject.value)
                    End If

            End Select
        Next fieldObject

        ' Add cluster if label exists
        If clusterId <> vbNullString Then
            If Not clusters.Exists(clusterId) Then
                Dim clusterObject As Cluster
                Set clusterObject = New Cluster

                clusterObject.id = clusterId
                clusterObject.label = clusterLabel
                clusterObject.styleName = clusterStyleName
                clusterObject.attributes = clusterAttributes
                clusterObject.Tooltip = clusterTooltip

                clusters.Add clusterId, clusterObject
            End If
        End If

        rs.MoveNext
    Loop

    rs.MoveFirst
    Set GetClusterInfo = clusters
End Function

' ==========================================================================
' FUNCTION: GetSubclusterInfo
' PURPOSE:
'   Extracts unique subclusters and their visual properties from a
'   database recordset.
'
' TECHNICAL WORKFLOW:
'   1. PRE-CHECK: Validates the ADO Recordset state (Open/Non-empty) and
'      initializes a Scripting.Dictionary for unique keys.
'   2. HEURISTIC DISCOVERY: Performs a case-insensitive scan of field names
'      to find 'SUBCLUSTER' and its related attribute columns.
'   3. METADATA CAPTURE: For every unique subcluster ID detected, it captures:
'      - id: The identifier used for grouping logic.
'      - label: Display text for the subcluster header.
'      - styleName: Style worksheet reference.
'      - attributes/tooltip: Visual properties and SVG interactivity.
'   4. OBJECT CACHING: Hydrates a 'Cluster' class instance for each unique
'      ID to provide O(1) attribute lookup during the mapping phase.
'   5. CURSOR STABILITY: Resets the recordset (MoveFirst) after the scan
'      to prevent EOF errors in subsequent processing loops.
' ==========================================================================
Private Function GetSubclusterInfo( _
    ByVal rs As Object, _
    ByRef fields As sqlFieldName) As Dictionary

    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    ' Exit early if empty or invalid
    If rs Is Nothing Then
        Set GetSubclusterInfo = subclusters
        Exit Function
    End If

    If rs.State <> ObjectStateEnum.adStateOpen Then
        Set GetSubclusterInfo = subclusters
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        Set GetSubclusterInfo = subclusters
        Exit Function
    End If

    ' Precompute lowercase field names for comparison
    Dim subField As String:        subField = LCase$(SafeStr(fields.subcluster))
    Dim subLabelField As String:   subLabelField = LCase$(SafeStr(fields.subclusterLabel))
    Dim subStyleField As String:   subStyleField = LCase$(SafeStr(fields.subclusterStyleName))
    Dim subAttrField As String:    subAttrField = LCase$(SafeStr(fields.subclusterAttributes))
    Dim subTooltipField As String: subTooltipField = LCase$(SafeStr(fields.subclusterTooltip))

    ' Check once whether SUBCLUSTER LABEL exists
    Dim hasSubLabel As Boolean
    hasSubLabel = HasField(rs, fields.subclusterLabel)

    Dim fieldObject As Variant
    Dim subId As String
    Dim subLabel As String
    Dim subStyle As String
    Dim subAttr As String
    Dim subTooltip As String

    rs.MoveFirst

    Do While Not rs.EOF

        subId = ""
        subLabel = ""
        subStyle = ""
        subAttr = ""
        subTooltip = ""

        ' Extract subcluster metadata
        For Each fieldObject In rs.fields
            Select Case LCase$(SafeStr(fieldObject.name))

                Case subField
                    subId = SafeStr(fieldObject.value)

                Case subLabelField
                    If hasSubLabel Then
                        subLabel = SafeStr(fieldObject.value)
                    End If

                Case subStyleField
                    subStyle = SafeStr(fieldObject.value)

                Case subAttrField
                    subAttr = SafeStr(fieldObject.value)

                Case subTooltipField
                    subTooltip = SafeStr(fieldObject.value)

            End Select
        Next fieldObject

        ' Add subcluster if id exists (use id as dictionary key)
        If Len(subId) > 0 Then
            If Not subclusters.Exists(subId) Then
                Dim clusterObject As Cluster
                Set clusterObject = New Cluster

                clusterObject.id = subId
                clusterObject.label = subLabel
                clusterObject.styleName = subStyle
                clusterObject.attributes = subAttr
                clusterObject.Tooltip = subTooltip

                subclusters.Add subId, clusterObject
            End If
        End If

        rs.MoveNext
    Loop

    rs.MoveFirst
    Set GetSubclusterInfo = subclusters
End Function

' ==========================================================================
' FUNCTION: GetSubClusterInfoForCluster
' PURPOSE:
'   Extracts the specific sub-groups belonging to a single parent cluster.
'
' TECHNICAL WORKFLOW:
'   1. SCOPED SCAN: Filters the recordset to only analyze rows where the
'      'CLUSTER' column matches the specified 'clusterName'.
'   2. NESTED DISCOVERY: Performs a case-insensitive search for secondary
'      grouping metadata (Subcluster ID, Label, Style, etc.) within that scope.
'   3. METADATA HYDRATION: Creates 'Cluster' class objects for each unique
'      sub-group found, capturing visual attributes like labels and tooltips.
'   4. DUPLICATION PROTECTION: Uses a Dictionary to ensure each subcluster
'      is only registered once per parent, maintaining a lean memory state.
'   5. CURSOR PRESERVATION: Resets the recordset (MoveFirst) after the scan
'      to keep the ADO cursor ready for the next parent or emission loop.
' ==========================================================================
Private Function GetSubClusterInfoForCluster( _
    ByVal rs As Object, _
    ByRef fields As sqlFieldName, _
    ByVal clusterName As String) As Dictionary

    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    ' Exit early if empty or invalid
    If rs Is Nothing Then
        Set GetSubClusterInfoForCluster = subclusters
        Exit Function
    End If

    If rs.State <> ObjectStateEnum.adStateOpen Then
        Set GetSubClusterInfoForCluster = subclusters
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        Set GetSubClusterInfoForCluster = subclusters
        Exit Function
    End If

    ' Precompute lowercase field names for comparison
    Dim clusterField As String:    clusterField = LCase$(SafeStr(fields.Cluster))
    Dim subField As String:        subField = LCase$(SafeStr(fields.subcluster))
    Dim subLabelField As String:   subLabelField = LCase$(SafeStr(fields.subclusterLabel))
    Dim subStyleField As String:   subStyleField = LCase$(SafeStr(fields.subclusterStyleName))
    Dim subAttrField As String:    subAttrField = LCase$(SafeStr(fields.subclusterAttributes))
    Dim subTooltipField As String: subTooltipField = LCase$(SafeStr(fields.subclusterTooltip))

    ' Check once whether SUBCLUSTER LABEL exists (optional)
    Dim hasSubLabel As Boolean
    hasSubLabel = HasField(rs, fields.subclusterLabel)

    Dim fieldObject As Variant
    Dim subId As String
    Dim subLabel As String
    Dim subStyle As String
    Dim subAttr As String
    Dim subTooltip As String

    rs.MoveFirst

    Do While Not rs.EOF

        ' Only process rows belonging to this cluster (Null-safe)
        If SafeFieldValue(rs, fields.Cluster) = clusterName Then

            subId = vbNullString
            subLabel = vbNullString
            subStyle = vbNullString
            subAttr = vbNullString
            subTooltip = vbNullString

            ' Extract subcluster metadata
            For Each fieldObject In rs.fields
                Select Case LCase$(SafeStr(fieldObject.name))

                    Case subField
                        subId = SafeStr(fieldObject.value)

                    Case subLabelField
                        ' Only attempt label if field actually exists
                        If hasSubLabel Then
                            subLabel = SafeStr(fieldObject.value)
                        End If

                    Case subStyleField
                        subStyle = SafeStr(fieldObject.value)

                    Case subAttrField
                        subAttr = SafeStr(fieldObject.value)

                    Case subTooltipField
                        subTooltip = SafeStr(fieldObject.value)
                End Select
            Next fieldObject

            ' Add subcluster if id exists (use id as the dictionary key)
            If Len(subId) > 0 Then
                If Not subclusters.Exists(subId) Then
                    Dim clusterObject As Cluster
                    Set clusterObject = New Cluster

                    clusterObject.id = subId
                    clusterObject.label = subLabel
                    clusterObject.styleName = subStyle
                    clusterObject.attributes = subAttr
                    clusterObject.Tooltip = subTooltip

                    subclusters.Add subId, clusterObject
                End If
            End If

        End If

        rs.MoveNext
    Loop

    rs.MoveFirst
    Set GetSubClusterInfoForCluster = subclusters
End Function

' ==========================================================================
' FUNCTION: GetOrphanSubClusterInfo
' PURPOSE:
'   Extracts metadata for subclusters that exist without a parent cluster.
'
' TECHNICAL WORKFLOW:
'   1. TARGETED SCAN: Filters the recordset for a specific logic gate:
'      [Cluster is Empty] AND [Subcluster is Not Empty].
'   2. METADATA CAPTURE: Identifies unique Subcluster IDs and hydrates 'Cluster'
'      objects with associated labels, styles, and attributes.
'   3. DE-DUPLICATION: Uses a Dictionary to ensure each orphan group is
'      cataloged only once, regardless of how many rows it contains.
'   4. CURSOR RESET: Re-initializes the recordset position to 'MoveFirst' to
'      allow the primary mapping engine to process the data.
'
' USAGE:
'   - Called during 'ProcessClusterYesSubclusterYes' to prevent data loss
'     when the hierarchy is incomplete or "top-heavy."
' ==========================================================================
Private Function GetOrphanSubClusterInfo( _
    ByVal rs As Object, _
    ByRef fields As sqlFieldName) As Dictionary

    ' Build a list of subclusters where the cluster column is null
    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    ' Exit early if empty or invalid
    If rs Is Nothing Then
        Set GetOrphanSubClusterInfo = subclusters
        Exit Function
    End If

    If rs.State <> ObjectStateEnum.adStateOpen Then
        Set GetOrphanSubClusterInfo = subclusters
        Exit Function
    End If

    If rs.EOF And rs.BOF Then
        Set GetOrphanSubClusterInfo = subclusters
        Exit Function
    End If

    ' Precompute lowercase field names for comparison
    Dim clusterField As String:    clusterField = LCase$(SafeStr(fields.Cluster))
    Dim subField As String:        subField = LCase$(SafeStr(fields.subcluster))
    Dim subLabelField As String:   subLabelField = LCase$(SafeStr(fields.subclusterLabel))
    Dim subStyleField As String:   subStyleField = LCase$(SafeStr(fields.subclusterStyleName))
    Dim subAttrField As String:    subAttrField = LCase$(SafeStr(fields.subclusterAttributes))
    Dim subTooltipField As String: subTooltipField = LCase$(SafeStr(fields.subclusterTooltip))

    ' Check once whether SUBCLUSTER LABEL exists
    Dim hasSubLabel As Boolean
    hasSubLabel = HasField(rs, fields.subclusterLabel)

    Dim fieldObject As Variant
    Dim subId As String
    Dim subLabel As String
    Dim subStyle As String
    Dim subAttr As String
    Dim subTooltip As String

    rs.MoveFirst

    Do While Not rs.EOF

        ' Only process rows where cluster is NULL/empty and subcluster is NOT NULL/empty
        If SafeFieldValue(rs, fields.Cluster) = "" _
           And Len(SafeFieldValue(rs, fields.subcluster)) > 0 Then

            subId = ""
            subLabel = ""
            subStyle = ""
            subAttr = ""
            subTooltip = ""

            ' Extract subcluster metadata
            For Each fieldObject In rs.fields
                Select Case LCase$(SafeStr(fieldObject.name))

                    Case subField
                        subId = SafeStr(fieldObject.value)

                    Case subLabelField
                        If hasSubLabel Then
                            subLabel = SafeStr(fieldObject.value)
                        End If

                    Case subStyleField
                        subStyle = SafeStr(fieldObject.value)

                    Case subAttrField
                        subAttr = SafeStr(fieldObject.value)

                    Case subTooltipField
                        subTooltip = SafeStr(fieldObject.value)

                End Select
            Next fieldObject

            ' Add subcluster if id exists (use id as key)
            If Len(subId) > 0 Then
                If Not subclusters.Exists(subId) Then
                    Dim clusterObject As Cluster
                    Set clusterObject = New Cluster

                    clusterObject.id = subId
                    clusterObject.label = subLabel
                    clusterObject.styleName = subStyle
                    clusterObject.attributes = subAttr
                    clusterObject.Tooltip = subTooltip

                    subclusters.Add subId, clusterObject
                End If
            End If

        End If

        rs.MoveNext
    Loop

    rs.MoveFirst
    Set GetOrphanSubClusterInfo = subclusters
End Function

' ==========================================================================
' PROCEDURE: EmitClusterOpen
' PURPOSE:
'   Writes the structural 'Open' row for a Graphviz Cluster/Subgraph.
'
' TECHNICAL WORKFLOW:
'   1. BRACE INSERTION: Places an 'OPEN_BRACE' ({) in the 'Item' column to
'      initiate a new Graphviz scope.
'   2. METADATA MAPPING: Populates the 'Label', 'Tooltip', and 'Extra Attributes'
'      columns with data from the 'clusterRecord' object.
'   3. DYNAMIC STYLE INJECTION:
'      - Retrieves the global 'Suffix Open' string from Settings.
'      - Performs token substitution (e.g., replacing '{i}' with the current
'        count) to support unique style naming for nested groups.
'   4. ROW MANAGEMENT: Automatically increments the global 'row' counter
'      after the write to prepare for the subsequent data records.
'
' USAGE:
'   - The primary tool for creating visual "bounding boxes" in the graph.
'   - Used by all Clustering routines (Legacy and Multi-Level).
' ==========================================================================
Private Sub EmitClusterOpen( _
    ByVal clusterRecord As Cluster, _
    ByRef dataLayout As dataWorksheet, _
    ByRef row As Long, _
    ByVal findStr As String, _
    ByRef replaceLong As Long)

    Dim newStyle As String
    Dim suffix As String

    ' Null-safe suffix retrieval
    suffix = SafeStr(SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_OPEN).value)

    With DataSheet
        .Cells(row, dataLayout.itemColumn).value = OPEN_BRACE
        .Cells(row, dataLayout.labelColumn).value = SafeStr(clusterRecord.label)

        .Cells(row, dataLayout.extraAttributesColumn).value = _
            replace(SafeStr(clusterRecord.attributes), _
                    findStr, SafeStr(replaceLong), , , vbTextCompare)

        .Cells(row, dataLayout.tooltipColumn).value = SafeStr(clusterRecord.Tooltip)

        If Len(SafeStr(clusterRecord.styleName)) > 0 Then
            newStyle = replace(SafeStr(clusterRecord.styleName), _
                               findStr, SafeStr(replaceLong), , , vbTextCompare) _
                       & suffix

            .Cells(row, dataLayout.styleNameColumn).value = newStyle
        End If
    End With

    row = row + 1
End Sub

' ==========================================================================
' PROCEDURE: EmitClusterClose
' PURPOSE:
'   Writes the structural 'Close' row for a Graphviz Cluster/Subgraph.
'
' TECHNICAL WORKFLOW:
'   1. BRACE INSERTION: Places a 'CLOSE_BRACE' (}) in the 'Item' column to
'      signal the end of the current Graphviz scope.
'   2. DYNAMIC STYLE INJECTION:
'      - Retrieves the global 'Suffix Close' string (e.g., '_CLOSE') from Settings.
'      - Substitutes tokens (e.g., replacing '{i}' with the cluster count) to
'        ensure the closing style matches the opening one.
'   3. STATE RESTORATION: Provides the necessary symmetry for Graphviz syntax;
'      without this row, the DOT source would remain unclosed and fail to render.
'   4. ROW MANAGEMENT: Increments the global 'row' counter after execution
'      to keep the worksheet cursor aligned for the next set of data.
'
' USAGE:
'   - Called immediately after a cluster's rows have been fully emitted.
'   - Essential for maintaining valid, human-readable DOT source code.
' ==========================================================================
Private Sub EmitClusterClose( _
    ByVal clusterRecord As Cluster, _
    ByRef dataLayout As dataWorksheet, _
    ByRef row As Long, _
    ByVal findStr As String, _
    ByRef replaceLong As Long)

    Dim suffix As String
    Dim newStyle As String

    ' Null-safe suffix retrieval
    suffix = SafeStr(SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_CLOSE).value)

    With DataSheet
        .Cells(row, dataLayout.itemColumn).value = CLOSE_BRACE

        ' Only emit style if non-empty
        If Len(SafeStr(clusterRecord.styleName)) > 0 Then
            newStyle = replace( _
                SafeStr(clusterRecord.styleName), _
                findStr, SafeStr(replaceLong), , , vbTextCompare _
            ) & suffix

            .Cells(row, dataLayout.styleNameColumn).value = newStyle
        End If
    End With

    row = row + 1
End Sub

' ==========================================================================
' PROCEDURE: EmitRows
' PURPOSE:
'   Translates ADO records into physical rows on the 'Data' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. INFINITE LOOP PROTECTION: Validates the 'stepBy' value and ensures
'      the loop direction (positive vs negative) is mathematically sound.
'   2. ENUMERATION CONTROL: Executes a 'For...Next' loop based on the
'      'GetLoopLimits' parameters. For standard data, this runs once.
'   3. GOVERNOR ENFORCEMENT: Monitors the 'ctx.loop.count' against the
'      maximum allowed steps to prevent system hangs on large datasets.
'   4. LINEAR MAPPING: Calls 'EmitOneRow' for each iteration to perform the
'      actual cell writes.
'   5. COORDINATE MANAGEMENT: Increments the global 'row' counter after
'      each emission to maintain the worksheet's vertical cursor.
' ==========================================================================
Private Sub EmitRows( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef position As Long)

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

        EmitOneRow ctx, rs, row, position, i
        row = row + 1
    Next i

End Sub

' ==========================================================================
' PROCEDURE: EmitOneRow
' PURPOSE:
'   Maps a single ADO record to a specific row on the 'Data' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. COUNTER MANAGEMENT: Increments the global 'position' tracker to
'      provide a unique index for the current session.
'   2. TOKEN REPLACEMENT: Performs two-tier string substitution:
'      - {record}: Replaced by the current record count.
'      - {i}: Replaced by the current 'enumStep' (if Enumeration is active).
'   3. COLUMN MAPPING: Uses a 'Select Case' block to match SQL field names
'      against the 'Data' worksheet schema.
'   4. AUTOMATIC TEXT SPLITTING: Specifically for 'Label' and 'xLabel' fields:
'      - Detects the 'SPLIT_LENGTH' parameter in the recordset.
'      - Calls 'SplitMultilineText' to wrap long strings into readable blocks.
'      - Supports custom 'Line Endings' (e.g., \n or \r\n).
'   5. NULL SAFETY: Uses 'SafeStr' for all values to ensure database Nulls
'      do not trigger VBA runtime errors.
' ==========================================================================
Private Sub EmitOneRow( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef position As Long, _
    ByVal enumStep As Long)

    ' Increment the result set position (i.e. recordCnt)
    position = position + 1

    With DataSheet
        Dim fld As Object
        Dim v As String
        Dim targetCol As Long

        For Each fld In rs.fields

            ' Common transformation: null -> "", placeholder replacement
            v = SafeStr(fld.value)

            If Len(v) > 0 Then
                v = replace(v, ctx.fields.recordsetPlaceholder, SafeStr(position), , , vbTextCompare)
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
        Next fld
    End With

End Sub

' ==========================================================================
' FUNCTION: GetSQLWorksheetHeadings
' PURPOSE:
'   Builds a dynamic map of worksheet headers to allow the SQL engine to
'   identify target columns by their current text labels.
'
' TECHNICAL WORKFLOW:
'   1. COM OPTIMIZATION: Reads the entire heading row into a 2D Variant
'      array in a single call, significantly reducing overhead for wide sheets.
'   2. DYNAMIC BINDING: Populates the 'DataWorksheetHeadings' UDT by
'      mapping internal logic keys (e.g., .tailLabel) to the physical column
'      indices defined in 'dataLayout'.
'   3. NORMALIZATION: Applies 'NormalizeHeading' to each value to strip
'      whitespace and standardize casing for reliable comparison.
'
' USAGE:
'   - Called during 'RunSQL' to establish the "column identity" contract
'     before any data is emitted.
' ==========================================================================
Private Function GetSQLWorksheetHeadings(ByRef dataLayout As dataWorksheet) As DataWorksheetHeadings
    Dim rowValues As Variant
    Dim r As Long: r = dataLayout.headingRow

    ' Read the entire header row in one COM call
    rowValues = DataSheet.rows(r).Value2

    With GetSQLWorksheetHeadings
        .flag = NormalizeHeading(rowValues(1, dataLayout.flagColumn))
        .item = NormalizeHeading(rowValues(1, dataLayout.itemColumn))
        .label = NormalizeHeading(rowValues(1, dataLayout.labelColumn))
        .xLabel = NormalizeHeading(rowValues(1, dataLayout.xLabelColumn))
        .tailLabel = NormalizeHeading(rowValues(1, dataLayout.tailLabelColumn))
        .headLabel = NormalizeHeading(rowValues(1, dataLayout.headLabelColumn))
        .Tooltip = NormalizeHeading(rowValues(1, dataLayout.tooltipColumn))
        .isRelatedToItem = NormalizeHeading(rowValues(1, dataLayout.isRelatedToItemColumn))
        .styleName = NormalizeHeading(rowValues(1, dataLayout.styleNameColumn))
        .extraAttributes = NormalizeHeading(rowValues(1, dataLayout.extraAttributesColumn))
        .errorMessage = NormalizeHeading(rowValues(1, dataLayout.errorMessageColumn))
    End With
End Function

' ==========================================================================
' FUNCTION: NormalizeHeading
' PURPOSE:
'   Standardizes worksheet header text for reliable string matching.
'
' TECHNICAL WORKFLOW:
'   1. NULL/ERROR GUARD: Returns an empty string if the cell contains an
'      Excel error (#REF!, #VALUE!), a database Null, or is blank.
'   2. TYPE COERCION: Explicitly casts the variant to a String to ensure
'      numeric headers don't cause type-mismatch errors.
'   3. STRING CLEANING: Applies 'LCase$' and 'Trim$' to eliminate case
'      sensitivity and leading/trailing whitespace.
'
' USAGE:
'   - Called by 'GetSQLWorksheetHeadings' to prepare the mapping dictionary.
'   - Crucial for matching user-edited column names to internal SQL logic.
' ==========================================================================
Private Function NormalizeHeading(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or v = vbNullString Then
        NormalizeHeading = vbNullString
    Else
        NormalizeHeading = Trim$(LCase$(CStr(v)))
    End If
End Function

' ==========================================================================
' FUNCTION: IsUserSQLError
' PURPOSE:
'   Determines if a SQL failure was caused by a malformed query or
'   schema mismatch rather than a transient system error.
'
' TECHNICAL WORKFLOW:
'   1. NORMALIZATION: Converts the incoming error message to lowercase to
'      ensure robust pattern matching across different ADO providers.
'   2. PATTERN MATCHING: Scans the message for specific diagnostic phrases:
'      - SCHEMA: "could not find the object", "not a valid name".
'      - SYNTAX: "syntax error", "missing operator", "reserved word".
'      - FORMAT: "external table is not in the expected format".
'      - LOGIC: "type mismatch", "undefined function".
'   3. LOGIC GATE: Returns True if any match is found, signaling the
'      'executeSQL' engine to skip retries and report the error to the user.
'
' USAGE:
'   - Used by the 'executeSQL' retry loop to prevent wasting time on
'     persistent syntax errors that retries cannot solve.
' ==========================================================================
Private Function IsUserSQLError(ByVal errMsg As String) As Boolean
    Dim m As String
    m = LCase$(errMsg)

    ' Patterns grouped by category
    If ContainsAny(m, Array( _
        "could not find the object", _
        "not a valid name", _
        "external table is not in the expected format", _
        "does not recognize", _
        "no value given", _
        "input must contain", _
        "syntax error", _
        "missing operator", _
        "type mismatch", _
        "input must contain at least one table or query", _
        "conflicts with distinct", _
        "statement includes a reserved word", _
        "undefined function", _
        "could not find installable isam", _
        "not enough information to identify" _
    )) Then
        IsUserSQLError = True
    End If
End Function

' ==========================================================================
' FUNCTION: ClassifyError
' PURPOSE:
'   Translates technical ADO error strings into helpful troubleshooting advice.
'
' TECHNICAL WORKFLOW:
'   1. CATEGORIZATION: Groups raw error substrings into four primary failure domains:
'      - TABLE ERRORS: Missing sheets, invalid names, or corrupt formats.
'      - COLUMN ERRORS: References to fields that do not exist in the source.
'      - SYNTAX ERRORS: Malformed SQL, reserved words, or type mismatches.
'      - SYSTEM ERRORS: ISAM driver issues or ambiguous range definitions.
'   2. ADVICE INJECTION: Returns a user-friendly string (e.g., "Verify the column name...")
'      that explains *how* to fix the problem.
'   3. FALLBACK: Returns an empty string if the error is unclassified,
'      allowing the system to fall back to the raw technical message.
'
' USAGE:
'   - Used by 'executeSQL' to populate the 'Status' column on the SQL worksheet.
'   - Essential for the "No-Code" user experience, removing the need for
'     database expertise to debug queries.
' ==========================================================================
Private Function ClassifyError(ByVal errMsg As String) As String
    Dim m As String
    m = LCase$(errMsg)

    ' --- Worksheet / table not found ---
    If ContainsAny(m, Array( _
        "could not find the object", _
        "not a valid name", _
        "external table is not in the expected format")) Then

        ClassifyError = "Cannot locate the worksheet/table. Check that the data file is correctly set and the worksheet/table name is valid."
        Exit Function
    End If

    ' --- Column not found ---
    If ContainsAny(m, Array( _
        "does not recognize", _
        "no value given")) Then

        ClassifyError = "Column not found. Verify the column name is correct and the worksheet/table is properly specified."
        Exit Function
    End If

    ' --- SQL syntax errors ---
    If ContainsAny(m, Array( _
        "syntax error", _
        "missing operator", _
        "type mismatch", _
        "input must contain at least one table or query", _
        "conflicts with distinct", _
        "statement includes a reserved word", _
        "undefined function")) Then

        ClassifyError = "SQL Syntax error"
        Exit Function
    End If

    ' --- ISAM / range issues ---
    If ContainsAny(m, Array( _
        "could not find installable isam", _
        "not enough information to identify")) Then

        ClassifyError = "ISAM / range issues"
        Exit Function
    End If

    ClassifyError = vbNullString
End Function

' ==========================================================================
' FUNCTION: ContainsAny
' PURPOSE:
'   Determines if a target string contains any element from a list of patterns.
'
' TECHNICAL WORKFLOW:
'   1. ITERATION: Loops through a Variant array of 'patterns'.
'   2. HEURISTIC SEARCH: Uses the 'InStr' function for fast substring detection.
'   3. EARLY EXIT: Returns True and terminates as soon as the first match
'      is found, optimizing performance for long error message lists.
'
' USAGE:
'   - Powering 'IsUserSQLError' and 'ClassifyError' to identify failure
'     categories within cryptic ADO driver messages.
' ==========================================================================
Private Function ContainsAny(ByVal Text As String, ByVal patterns As Variant) As Boolean
    Dim p As Variant
    For Each p In patterns
        If InStr(Text, p) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next p
End Function

' ==========================================================================
' FUNCTION: SafeStr
' PURPOSE:
'   Provides a "Bulletproof" conversion of any Variant to a String.
'
' TECHNICAL WORKFLOW:
'   1. PRE-EMPTIVE CHECKS: Systematically identifies problematic VBA states:
'      - IsMissing: Handles optional parameters.
'      - IsNull: Essential for ADO Recordsets where fields are empty.
'      - vbError: Prevents crashes if a cell contains #REF!, #NAME->, etc.
'      - Empty: Handles uninitialized variables.
'   2. CASTING: If the value is valid, it performs a standard 'CStr' conversion.
'   3. FAIL-SAFE: Employs a 'Last-chance' error handler that returns an
'      empty string even if the coercion itself fails.
'
' USAGE:
'   - Used globally throughout 'modWorksheetSQL' to ensure that database
'     content can be safely concatenated or written to cells.
' ==========================================================================
Public Function SafeStr(ByVal v As Variant) As String
    ' Convert Null, Empty, Missing, or Error to ""
    ' Convert anything else to a string without throwing
    
    On Error GoTo CleanFail

    If IsMissing(v) Then
        SafeStr = ""
        Exit Function
    End If

    If IsNull(v) Then
        SafeStr = ""
        Exit Function
    End If

    If VarType(v) = vbError Then
        SafeStr = ""
        Exit Function
    End If

    If v = Empty Then
        SafeStr = ""
        Exit Function
    End If

    ' Normal case
    SafeStr = CStr(v)
    Exit Function

CleanFail:
    ' Last-chance safety net -> never propagate errors
    SafeStr = ""
End Function

' ==========================================================================
' FUNCTION: SafeFieldValue
' PURPOSE:
'   The ultimate "Safe-Read" wrapper for ADO Recordset fields.
'
' TECHNICAL WORKFLOW:
'   1. STATE VALIDATION: Verifies the Recordset object exists and is
'      currently in an 'adStateOpen' state.
'   2. FIELD EXISTENCE PROBE: Uses a localized 'On Error Resume Next' to
'      test if the requested field name exists in the schema.
'   3. DATA COERCION: If the field is valid, it passes the value through
'      'SafeStr' to handle Nulls, Errors, or Empty states.
'   4. SILENT FAILURE: Designed to never throw an error; it returns
'      an empty string ("") as the safest possible fallback for the graph.
'
' USAGE:
'   - Used by 'PerformRecursiveSearch' and 'ProcessMultiLevelRecordset'
'     to identify parent/child IDs and visual attributes.
' ==========================================================================
Public Function SafeFieldValue(ByVal rs As Object, ByVal fieldName As String) As String
    ' Null-safe accessor for recordset fields.
    ' Returns "" for:
    '   - Null
    '   - Missing field
    '   - Closed/invalid recordset
    '   - Error variants
    '   - Empty
    ' Never throws. Never mutates caller intent.

    On Error GoTo CleanFail

    ' Validate recordset
    If rs Is Nothing Then
        SafeFieldValue = ""
        Exit Function
    End If

    If rs.State <> ObjectStateEnum.adStateOpen Then
        SafeFieldValue = ""
        Exit Function
    End If

    ' Validate field name
    If Len(SafeStr(fieldName)) = 0 Then
        SafeFieldValue = ""
        Exit Function
    End If

    ' Check if field exists
    Dim fld As Object
    On Error Resume Next
    Set fld = rs.fields(fieldName)
    If err.number <> 0 Then
        err.Clear
        SafeFieldValue = ""
        Exit Function
    End If
    On Error GoTo CleanFail

    ' Extract value safely
    SafeFieldValue = SafeStr(fld.value)
    Exit Function

CleanFail:
    ' Last-chance safety net -> never propagate errors
    SafeFieldValue = ""
End Function

' ==========================================================================
' FUNCTION: IsSqlRowActive
' PURPOSE:
'   Determines if a worksheet row contains an executable SQL SELECT statement.
'
' TECHNICAL WORKFLOW:
'   1. COMMENT FILTERING: Checks for the '#' indicator in the 'Comment'
'      column to identify rows manually disabled by the user.
'   2. CONTENT VALIDATION: Verifies the 'SQL Statement' column is not
'      empty, using 'SafeStr' to prevent errors on null/error cells.
'   3. SYNTAX VERIFICATION: Ensures the statement begins with the
'      'SELECT' keyword, isolating data-fetching queries from
'      environmental commands (like SET_DATA_FILE).
'   4. STATE REPORT: Returns True only if the row is uncommented,
'      populated, and contains valid query syntax.
'
' USAGE:
'   - Used by the Ribbon UI to determine if the 'Run SQL' button should
'     be enabled or disabled.
'   - Used to determine when to display floating buttons.
'   - Serves as the primary filter for batch execution loops.
' ==========================================================================
Public Function IsSqlRowActive(ByVal row As Long) As Boolean
    IsSqlRowActive = False     ' Establish default
    
    ' Check to see if the row is commented out
    Dim commentIndicator As String
    commentIndicator = SqlSheet.Cells(row, GetSettingColNum(SETTINGS_SQL_COL_COMMENT)).value
    If commentIndicator = "#" Then Exit Function
    
    ' Get SQL statement
    Dim sqlStatement As String
    sqlStatement = Trim$(SafeStr(SqlSheet.Cells.item(row, GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)).value))
    If sqlStatement = "" Then Exit Function
    
    ' See if it starts with SELECT
    Dim sqlUCase As String
    sqlUCase = UCase$(sqlStatement)
    If Not StartsWith(sqlUCase, SQL_SELECT) Then Exit Function

    ' All tests passed
    IsSqlRowActive = True
End Function

' ==========================================================================
' PROCEDURE: RunOneSqlStatement
' PURPOSE:
'   Triggers the SQL execution engine for the currently selected worksheet row.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT CAPTURE: Identifies the 'ActiveCell.row' to determine which
'      SQL statement the user is currently focused on.
'   2. ENGINE INVOCATION: Routes the row index to 'RunSQLAsExtension',
'      which handles the setup, execution, and cleanup for that specific query.
'
' USAGE:
'   - Linked to the 'Run Active SQL' button in the Ribbon.
'   - Essential for iterative query development and debugging.
' ==========================================================================
Public Sub RunOneSqlStatement()
    Dim activeRow As Long
    activeRow = ActiveCell.row
    RunSQLAsExtension rowNumber:=activeRow
End Sub

' ==========================================================================
' PROCEDURE: RunSQLAsExtension
' PURPOSE:
'   Wraps the SQL execution engine in a high-stability application state.
'
' TECHNICAL WORKFLOW:
'   1. STATE PRESERVATION: Backs up critical Excel settings (Cursor,
'      Calculation, ScreenUpdating, Events, and AutoRecover).
'   2. LOCKDOWN:
'      - Disables ScreenUpdating and Events to maximize execution speed.
'      - Switches to 'Manual Calculation' to prevent formula lag.
'      - Disables 'AutoRecover' to prevent ADO file-lock contention.
'   3. EXECUTION: Calls the core 'RunSQL' logic for the specified row.
'   4. UI SYNCHRONIZATION: Notifies the Ribbon to update controls (e.g.,
'      enabling the 'Reset Pool' button) after execution.
'   5. RESTORATION: Systematically returns all Excel settings to their
'      original states once the process completes.
' ==========================================================================
Public Sub RunSQLAsExtension(Optional ByVal rowNumber As Long = 0)
    ' Change the cursor to the wait cursor (hourglass)
    Dim originalCursorType As Long
    originalCursorType = Application.Cursor
    Application.Cursor = xlWait
    
    ' Add a guard against Excel recalculation
    Dim originalCalculation As Long
    originalCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' AutoSave/AutoRecover can lock the workbook while ADO is reading it.
    ' Disable AutoRecover during SQL
    Dim originalAutoRecover As Boolean
    originalAutoRecover = Application.AutoRecover.Enabled
    Application.AutoRecover.Enabled = False

    ' Disable screen updating
    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Disable events
    Dim originalEnableEvents As Boolean
    originalEnableEvents = Application.enableEvents
    Application.enableEvents = False

    ' Execute ALL the SQL commands — pass the row number
    RunSQL rowNumber
    
    ' Refresh the ribbon controls based on SQL execution activity
    InvalidateRibbonControl RIBBON_CTL_SQL_CONN_POOL_RESET
    
    ' Restore prior states
    Application.enableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    Application.AutoRecover.Enabled = originalAutoRecover
    Application.Calculation = originalCalculation
    Application.Cursor = originalCursorType
End Sub

' ==========================================================================
' FUNCTION: GetLastRowInColumn
' PURPOSE:
'   Determines the index of the final populated row in a specific column.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY CHECK: Validates that the 'colNum' exists within the
'      worksheet's physical limits to prevent runtime errors.
'   2. UPWARD SCAN: Uses the 'End(xlUp)' method from the bottom of the
'      sheet to find the first non-empty cell.
'   3. EMPTY STATE RESOLUTION: Specifically checks if row 1 is truly
'      populated; if 'End(xlUp)' hits row 1 but the cell is empty, the
'      function correctly returns 0.
'
' USAGE:
'   - Called by 'RunSQL' to determine the 'lastRow' of the SQL statement list.
'   - Prevents the batch processor from wasting cycles on empty worksheet rows.
' ==========================================================================
Private Function GetLastRowInColumn( _
    ByVal ws As Worksheet, _
    ByVal colNum As Long) As Long

    ' Returns the last row that contains any value in the specified column
    ' Returns 0 if the column is completely empty

    If colNum < 1 Or colNum > ws.columns.count Then
        GetLastRowInColumn = 0
        Exit Function
    End If

    Dim last As Long
    last = ws.Cells(ws.rows.count, colNum).End(xlUp).row

    ' If .End(xlUp) lands on row 1 and that cell is empty -> column is empty
    If last = 1 And IsEmpty(ws.Cells(1, colNum).value) Then
        GetLastRowInColumn = 0
    Else
        GetLastRowInColumn = last
    End If

End Function

' ==========================================================================
' PROCEDURE: ParsePlaceholderLine
' PURPOSE:
'   Extracts custom variable definitions from the SQL worksheet.
'
' TECHNICAL WORKFLOW:
'   1. COMMAND STRIPPING: Removes the 'SET PLACEHOLDER' keyword to isolate
'      the assignment logic.
'   2. DELIMITER IDENTIFICATION: Locates the '=' character to distinguish
'      between the placeholder name (key) and its replacement text (value).
'   3. STRING SANITIZATION: Trims both parts to ensure leading or trailing
'      spaces in the Excel cell don't affect the final SQL injection.
'   4. DICTIONARY UPSERT:
'      - If the name already exists, it updates the value (Overwrite).
'      - If new, it adds the pair to the 'placeholders' registry.
'
' USAGE:
'   - Triggered by 'RunSQL' when it encounters the 'SET PLACEHOLDER' command.
'   - Enables syntax like: "SET PLACEHOLDER DeptID = 101"
' ==========================================================================
Private Sub ParsePlaceholderLine(ByRef placeholders As Dictionary, ByVal line As String)
    Dim work As String
    Dim eqPos As Long
    Dim namePart As String
    Dim valuePart As String

    ' Strip the prefix "SET PLACEHOLDER"
    work = Trim$(Mid$(line, Len(SQL_SET_PLACEHOLDER) + 1))

    ' Find the equals sign
    eqPos = InStr(1, work, "=", vbTextCompare)
    If eqPos = 0 Then Exit Sub   ' malformed, ignore

    ' Split into name and value
    namePart = Trim$(Left$(work, eqPos - 1))
    valuePart = Trim$(Mid$(work, eqPos + 1))

    ' Add or replace
    If placeholders.Exists(namePart) Then
        placeholders(namePart) = valuePart
    Else
        placeholders.Add namePart, valuePart
    End If
End Sub

' ==========================================================================
' PROCEDURE: ParseClusterLevelLimitLine
' PURPOSE:
'   Extracts and applies a depth limit for the Multi-Level Clustering algorithm.
'
' TECHNICAL WORKFLOW:
'   1. KEYWORD STRIPPING: Removes the 'SET CLUSTER LEVEL LIMIT' prefix to
'      isolate the numeric assignment.
'   2. DELIMITER SEARCH: Identifies the '=' character to locate the
'      specified limit value.
'   3. TYPE CONVERSION: Extracts the text following the equals sign and
'      casts it to a 'Long' integer via 'CLng'.
'   4. STATE UPDATE: Overwrites the 'clusterLevelLimit' property in the
'      global SQL context.
'
' USAGE:
'   - Triggered when 'RunSQL' encounters the 'SQL_SET_CLUSTER_LEVEL_LIMIT' command.
'   - Essential for preventing performance degradation in extremely deep
'     hierarchical datasets (e.g., stopping the search at Level 5).
' ==========================================================================
Private Sub ParseClusterLevelLimitLine(ByVal line As String, ByRef clusterLevelLimit As Long)
    Dim work As String
    Dim eqPos As Long
    Dim valuePart As String
    Dim limit As Long

    ' Strip the prefix "SET CLUSTER LEVEL LIMIT"
    work = Trim$(Mid$(line, Len(SQL_SET_CLUSTER_LEVEL_LIMIT) + 1))

    ' Find the equals sign
    eqPos = InStr(1, work, "=", vbTextCompare)
    If eqPos = 0 Then Exit Sub   ' malformed, ignore

    ' Get the value after the equals sign, and convert to a number
    valuePart = Trim$(Mid$(work, eqPos + 1))
    clusterLevelLimit = CLng(valuePart)
End Sub

' ==========================================================================
' PROCEDURE: CleanupPlaceholders
' PURPOSE:
'   Systematically destroys the dictionary used for dynamic SQL tokens.
'
' TECHNICAL WORKFLOW:
'   1. NULL VALIDATION: Verifies the dictionary exists before attempting
'      to access its methods.
'   2. DICTIONARY PURGE: Calls 'RemoveAll' to clear every key-value pair
'      (e.g., Department IDs, Date ranges) stored during the session.
'   3. OBJECT DESTRUCTION: Sets the 'placeholders' reference to 'Nothing',
'      signaling Excel's garbage collector to reclaim the memory.
'
' USAGE:
'   - Called at the conclusion of 'RunSQL'.
'   - Ensures that placeholder values from one data sheet do not "leak"
'     into the execution context of another.
' ==========================================================================
Private Sub CleanupPlaceholders(ByRef placeholders As Dictionary)
    If Not placeholders Is Nothing Then
        placeholders.RemoveAll
        Set placeholders = Nothing
    End If
End Sub

' ==========================================================================
' PROCEDURE: ApplyPlaceholders
' PURPOSE:
'   Resolves all custom variables within a SQL statement before execution.
'
' TECHNICAL WORKFLOW:
'   1. PRECONDITION CHECK: Exits immediately if no placeholders are defined,
'      ensuring zero overhead for standard queries.
'   2. TOKEN WRAPPING: Automatically wraps each dictionary key in curly
'      braces (e.g., "DeptID" becomes "{DeptID}") to identify the target token.
'   3. CASE-INSENSITIVE INJECTION: Uses 'vbTextCompare' to perform the
'      replacement, allowing users to be flexible with casing in their SQL.
'   4. GLOBAL REPLACEMENT: Scans the entire 'sqlText' string to ensure every
'      instance of a token is updated with its assigned value.
'
' USAGE:
'   - Called by 'RunSQL' immediately before 'executeSQL'.
'   - Enables the creation of "Master Templates" where a single variable change
'     updates dozens of dependent queries.
' ==========================================================================
Private Sub ApplyPlaceholders(ByRef sqlText As String, ByRef placeholders As Dictionary)
    Dim key As Variant
    Dim token As String

    If placeholders Is Nothing Then Exit Sub
    If placeholders.count = 0 Then Exit Sub

    For Each key In placeholders.Keys
        token = "{" & CStr(key) & "}"
        sqlText = replace(sqlText, token, placeholders(key), , , vbTextCompare)
    Next key
End Sub

' ==========================================================================
' FUNCTION: HasField
' PURPOSE:
'   Determines if a specific field name exists within an ADO Recordset.
'
' TECHNICAL WORKFLOW:
'   1. OBJECT GUARD: Performs a null-check on the Recordset to prevent
'      runtime errors if the query failed to open.
'   2. CASE-INSENSITIVE MATCHING: Normalizes both the 'fieldName' and
'      the Recordset's field names to lowercase using 'LCase$'.
'   3. COLLECTION SCAN: Iterates through the 'rs.fields' collection using
'      a 'For Each' loop to perform an exhaustive search.
'   4. EARLY EXIT: Returns 'True' and terminates immediately upon finding
'      the first match, optimizing performance for wide tables.
'
' USAGE:
'   - Used by 'executeSQL' to detect "Pseudo-SQL" triggers (e.g., CREATE_EDGES).
'   - Powering 'SafeFieldValue' to prevent errors when accessing optional data.
' ==========================================================================
Public Function HasField(ByVal rs As Object, ByVal fieldName As String) As Boolean
    If rs Is Nothing Then Exit Function
    
    Dim fld As Object
    Dim searchName As String
    searchName = LCase$(Trim$(fieldName))
    
    For Each fld In rs.fields
        If LCase$(CStr(fld.name)) = searchName Then
            HasField = True
            Exit Function
        End If
    Next fld
End Function

' ==========================================================================
' FUNCTION: GetLineEnding
' PURPOSE:
'   Determines the character sequence used to terminate lines in wrapped text.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA PROBE: Uses 'HasField' to check if a specific "line-ending"
'      instruction was included in the SQL result set.
'   2. DYNAMIC OVERRIDE: If the field exists and is populated, it captures
'      that value (e.g., "\n" for Graphviz or "vbCrLf" for standard text).
'   3. FALLBACK LOGIC: If no specific instruction is found, it reverts
'      to a provided default (typically the global NEWLINE constant).
'
' USAGE:
'   - Used by 'EmitOneRow' to configure the 'SplitMultilineText' process.
'   - Enables SQL queries to dictate visual formatting for complex nodes.
' ==========================================================================
Public Function GetLineEnding( _
    ByVal rs As Object, _
    ByVal fieldName As String, _
    ByVal defaultLineEnding As String) As String

    Dim temp As String

    If HasField(rs, fieldName) Then
        temp = SafeFieldValue(rs, fieldName)
        If Len(temp) > 0 Then
            GetLineEnding = temp
            Exit Function
        End If
    End If

    GetLineEnding = defaultLineEnding
End Function

' ==========================================================================
' FUNCTION: GetSplitLength
' PURPOSE:
'   Extracts the maximum line length for text-wrapping from a SQL recordset.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DETECTION: Verifies if the 'SPLIT_LENGTH' field exists in the
'      returned database results.
'   2. TYPE VALIDATION: Confirms the value is numeric and greater than zero
'      to avoid logical errors or infinite loops in the wrapping engine.
'   3. SANITIZATION: Uses 'SafeFieldValue' to ensure Nulls or non-numeric
'      strings from the database are handled without crashing.
'   4. DEFAULT FALLBACK: Returns 0 if the field is missing or invalid,
'      signaling the mapping engine to keep the text on a single line.
'
' USAGE:
'   - Used by 'EmitOneRow' to decide whether to trigger 'SplitMultilineText'.
'   - Allows SQL-driven automation to control the visual "density" of a node.
' ==========================================================================
Public Function GetSplitLength( _
    ByVal rs As Object, _
    ByVal fieldName As String) As Long

    Dim temp As String
    Dim value As Long

    If HasField(rs, fieldName) Then
        temp = SafeFieldValue(rs, fieldName)

        If Len(temp) > 0 Then
            If IsNumeric(temp) Then
                value = CLng(temp)
                If value > 0 Then
                    GetSplitLength = value
                    Exit Function
                End If
            End If
        End If
    End If

    GetSplitLength = 0
End Function


