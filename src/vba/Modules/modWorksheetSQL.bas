Attribute VB_Name = "modWorksheetSQL"
' Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.SQL")

Option Explicit

Private Const RETRY_DELAY_MS As Long = 100
Private Const SQL_DELAY_MS As Long = 20
Private Const DEFAULT_MAX_RECURSION_DEPTH As Long = 100
Private Const LOOP_MAX_STEPS As Long = 10000 ' Put an upper limit on DO loop to prevent infinite loops

Private Type EnumerateParameters
    enabled As Boolean
    startAt As Long
    stopAt As Long
    stepBy As Long
    max As Long
    count As Long
End Type

Private Type sqlContext
    dataLayout As dataWorksheet
    fields As sqlFieldName
    headings As DataWorksheetHeadings
    sqlLayout As sqlWorksheet
    loop As EnumerateParameters
End Type

''' Button Actions - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'''  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Sub RunSQL()
    ' Disable logging from prior run
    SetLoggingEnabled False

    ' Get the column layout of the 'data' worksheet
    Dim context As sqlContext
    context.dataLayout = GetSettingsForDataWorksheet(DataSheet.name)

    ' Get the heading values of the 'data' worksheet columns.
    context.headings = GetSQLWorksheetHeadings(context.dataLayout)

    ' Get the column layout of the 'sql' worksheet
    context.sqlLayout = GetSettingsForSqlWorksheet()

    ' Get the list of special field names used for determining clusters and subclusters.
    context.fields = GetSettingsForSqlFields(True)

    ' Determine the last row with data
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    ' Disable automatic graph rendering as cells change.
    Dim runMode As String
    runMode = SafeStr(SettingsSheet.Range(SETTINGS_RUN_MODE).value)
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_MANUAL

    ' Clear out the info from previous run
    ClearSQLStatus
    ClearDataWorksheet DataSheet.name

    Dim dataRow As Long
    dataRow = context.dataLayout.firstRow

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

    For sqlRow = context.sqlLayout.firstRow To lastRow

        ' Skip initializations if the SQL row is commented out
        If SafeStr(SqlSheet.Cells.item(sqlRow, context.sqlLayout.flagColumn).value) <> FLAG_COMMENT Then

            ' Establish the full path to the Excel file containing the data
            filePath = GetExcelFilePath(sqlRow, context.sqlLayout, dataFile)

            ' Get SQL statement, and convert to upper case
            sqlStatement = Trim$(SafeStr(SqlSheet.Cells.item(sqlRow, context.sqlLayout.sqlStatementColumn).value))
            sqlUCase = UCase$(sqlStatement)

            ' Get default SUCCESS message
            message = GetMessage("msgboxSqlStatusSuccess")
        End If

        If SafeStr(SqlSheet.Cells.item(sqlRow, context.sqlLayout.flagColumn).value) = FLAG_COMMENT Then
            sqlStatement = vbNullString
            sqlUCase = vbNullString
            message = GetMessage("msgboxSqlStatusSkipped")

        ElseIf Len(sqlStatement) = 0 Then
            message = vbNullString

        ElseIf Not PassesFilter(sqlRow, filterColumn) Then
            message = GetMessage("msgboxSqlStatusFiltered")

        ElseIf sqlUCase = SQL_SET_DATA_FILE Then
            dataFile = SafeStr(SqlSheet.Cells.item(sqlRow, context.sqlLayout.excelFileColumn).value)

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

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_ALL_VIEWS_AS_DIRECTED_GRAPH) Then
            PublishAllViewsAsDirectedGraph (sqlStatement)

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_ALL_VIEWS_AS_UNDIRECTED_GRAPH) Then
            PublishAllViewsAsUndirectedGraph sqlStatement

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_ALL_VIEWS) Then
            PublishAllViews sqlStatement, SQL_PUBLISH_ALL_VIEWS

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_AS_DIRECTED_GRAPH) Then
            PublishAsDirectedGraph sqlStatement

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH_AS_UNDIRECTED_GRAPH) Then
            PublishAsUndirectedGraph sqlStatement

        ElseIf StartsWith(sqlUCase, SQL_PUBLISH) Then
            Publish sqlStatement, SQL_PUBLISH

        ElseIf StartsWith(sqlUCase, SQL_PREVIEW_AS_DIRECTED_GRAPH) Then
            PreviewAs TOGGLE_DIRECTED

        ElseIf StartsWith(sqlUCase, SQL_PREVIEW_AS_UNDIRECTED_GRAPH) Then
            PreviewAs TOGGLE_UNDIRECTED

        ElseIf StartsWith(sqlUCase, SQL_PREVIEW) Then
            CreateGraphWorksheet

        ElseIf Not StartsWith(sqlUCase, SQL_SELECT) Then
            message = GetMessage("msgboxSqlStatusSkipped") & " - " & GetMessage("msgboxSqlMustBeginWithSelect")

        Else
            ' SELECT branch
            Set connectionObject = Nothing

            ' Get connection to data source
            On Error Resume Next
            Set connectionObject = getConnection(filePath, context.fields.maxConnectionMinutes)
            Dim connErrDescription As String: connErrDescription = Err.Description
            Dim connErrNumber As Long: connErrNumber = Err.number
            On Error GoTo 0

            ' Execute the SQL query
            If connectionObject Is Nothing Then
                LogDiagnostic connErrDescription, errorNumber:=connErrNumber, errorCategory:="Data / Connection"
                message = GetMessage("msgboxSqlStatusFailure") & " - " & connErrDescription
            Else
                Err.Clear
                message = executeSQL(context, filePath, connectionObject, sqlStatement, dataRow)
            End If
        End If

        ' Display the status of the SQL query
        SqlSheet.Cells.item(sqlRow, context.sqlLayout.statusColumn).value = message

        ' Breathe. a small delay before each SQL execution can reduce COM collisions on slower machines
        DoEvents
        SleepMilliseconds SQL_DELAY_MS
    Next sqlRow

    ' Clean up connection pool if using narrow-scoped pooling
    If GetSettingBoolean(SETTINGS_SQL_CLOSE_CONNECTIONS) Then CleanupConnectionPool

    ' Restore the run mode setting
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = runMode
End Sub

Private Sub PreviewAs(ByVal graphType As String)
    Dim originalGraphType As String
    originalGraphType = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = graphType
    CreateGraphWorksheet
    SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = originalGraphType
End Sub

Private Sub PublishAsDirectedGraph(ByRef commandStatement As String)
    PublishAs commandStatement, SQL_PUBLISH_AS_DIRECTED_GRAPH, TOGGLE_DIRECTED
End Sub

Private Sub PublishAsUndirectedGraph(ByRef commandStatement As String)
    PublishAs commandStatement, SQL_PUBLISH_AS_UNDIRECTED_GRAPH, TOGGLE_UNDIRECTED
End Sub

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

Private Sub Publish(ByRef commandStatement As String, ByRef phrase As String)
    Dim firstColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    Dim lastColumn As Long
    lastColumn = firstColumn

    PublishViews commandStatement, phrase, firstColumn, lastColumn
End Sub

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

Private Sub PublishAllViewsAsDirectedGraph(ByRef commandStatement As String)
    PublishAllViewsAs commandStatement, SQL_PUBLISH_ALL_VIEWS_AS_DIRECTED_GRAPH, TOGGLE_DIRECTED
End Sub

Private Sub PublishAllViewsAsUndirectedGraph(ByRef commandStatement As String)
    PublishAllViewsAs commandStatement, SQL_PUBLISH_ALL_VIEWS_AS_UNDIRECTED_GRAPH, TOGGLE_UNDIRECTED
End Sub

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

Private Sub PublishAllViews(ByRef commandStatement As String, ByRef phrase As String)
    Dim firstColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    
    Dim lastColumn As Long
    lastColumn = GetLastViewColumn(firstColumn)

    PublishViews commandStatement, phrase, firstColumn, lastColumn
End Sub

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

Private Sub SetPrefix(ByRef commandStatement As String, ByRef phrase As String)
    ' Override the filename prefix with the value provided after the PUBLISH phrase
    If Len(commandStatement) > Len(phrase) Then
        Dim prefix As String
        prefix = Mid$(commandStatement, Len(phrase) + 1)
        SettingsSheet.Range(SETTINGS_FILE_NAME).value = Trim$(prefix)
    End If
End Sub

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

''' SQL PROCESSING - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'''  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' https://technet.microsoft.com/en-us/library/ee692882.aspx

Private Function executeSQL( _
    ByRef context As sqlContext, _
    ByVal filePath As String, _
    ByRef connectionObject As Object, _
    ByVal sqlStatement As String, _
    ByRef row As Long) As String

#If Win32 Or Win64 Then

    On Error GoTo executeSQLError
    
    Dim rsQueryResults As Object
    Dim rsRecursionResults As Object
    Dim rsMergedResults As Object
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
    Set rsQueryResults = CreateObject("ADODB.Recordset")
    
    Err.Clear
    For attempts = 1 To context.fields.retryLimit
        On Error Resume Next
        
        ' Execute the SQL SELECT query
        rsQueryResults.Open source:=sqlStatement, ActiveConnection:=connectionObject, CursorType:=CursorTypeEnum.adOpenStatic, LockType:=LockTypeEnum.adLockReadOnly, options:=CommandTypeEnum.adCmdText
        
        ' Immediately save error state, as processing an error could trigger it being cleared
        errNumber = Err.number
        errDescription = Err.Description
        
        ' Break from retry loop if query succeeded
        If errNumber = 0 Then Exit For
        
        ' Check for non-retryable SQL errors (bad sheet, bad column, syntax, etc.)
        userError = IsUserSQLError(errDescription)
        If userError Then
            Exit For
        End If

        LogDiagnostic "executeSQL(): rsQueryResults.Open - " & errDescription, errorNumber:=errNumber, attempt:=attempts, sql:=sqlStatement, errorCategory:=ClassifyError(Err.Description)
        Err.Clear
        SleepMilliseconds RETRY_DELAY_MS
    Next attempts
    
    ' Reset status bar
    Application.StatusBar = False
    Err.Clear
    
    ' If userError, then the SQL is bad. Stop processing.
    If userError Then GoTo executeSQLError
    If errNumber <> 0 Then GoTo executeSQLError
    
    ' If the recordset failed to open but didn?t trigger userError, rsQueryResults.State might be 0.
    
    If rsQueryResults Is Nothing Or rsQueryResults.State <> ObjectStateEnum.adStateOpen Then
        Err.Raise vbObjectError + 513, , "executeSQL(): Recordset failed to open"
    End If

    ' Determine if enumeration values are present
    context.loop = GetLoopLimits(context, rsQueryResults)
    
    ' Execute any iteration query passed in the SQL SELECT
    ' Performs iteration of parameterized query + mapping to data worksheet.
    IterativeSearch connectionObject, context, rsQueryResults, row, recordCnt
    
    ' Execute any recursion query passed in the SQL SELECT
    ' Perfroms recursion + mapping to data worksheet.
    RecursiveSearch connectionObject, context, rsQueryResults, rsRecursionResults
    
    If rsRecursionResults Is Nothing Then
        ' No recursion query was run, emit the results of the primary query
        MapResultsToDataWorksheet context, rsQueryResults, row, recordCnt
    Else
        ' A set of recursive queries was executed. We have to merge the results
        ' of the primary query, and recursive queries into a single set of
        ' results so that cluster and subclusters are honored across all the
        ' queries.
        MergeRecordsets rsQueryResults, rsRecursionResults, rsMergedResults
        MapResultsToDataWorksheet context, rsMergedResults, row, recordCnt
    End If
    
    ' Return success status in local language
    executeSQL = GetMessage("msgboxSqlStatusSuccess")
    
Cleanup:
    On Error Resume Next
    SafeCloseRecordset rsQueryResults
    SafeCloseRecordset rsRecursionResults
    SafeCloseRecordset rsMergedResults
    Application.StatusBar = False
    On Error GoTo 0
    Exit Function

executeSQLError:
    ' GetMessage will reset the error state, save the message
    If Err.number <> 0 Then
        errDescription = Err.Description
        errNumber = Err.number
    End If
    
    Dim logMessage As String
    logMessage = "executeSQL(): " & errDescription & vbNewLine & "  Excel Data File     : " & filePath

    Dim statusMessage As String
    statusMessage = errDescription _
                & vbNewLine & vbNewLine & "Err.Number=" & errNumber _
                & vbNewLine & vbNewLine & "datafile=" & filePath
    
    LogDiagnostic logMessage, errorNumber:=errNumber, attempt:=attempts, sql:=sqlStatement, errorCategory:=ClassifyError(errDescription)
    executeSQL = GetMessage("msgboxSqlStatusFailure") & " - " & statusMessage
    GoTo Cleanup
#Else
    executeSQL = GetMessage("msgboxSqlStatusFailure") & " - ADO is not supported on macOS."
#End If

End Function

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
    Err.Clear
#End If
End Sub

Private Sub SleepMilliseconds(ByVal ms As Long)
#If Win32 Or Win64 Then
    Dim t As Single
    t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
#End If
End Sub

Private Function GetLoopLimits(ByRef context As sqlContext, _
                               ByVal rsQueryResults As Object) As EnumerateParameters

    Dim s As EnumerateParameters
    ' Default: no loop mode; EmitRows/consumers decide what to do with a single pass
    s.enabled = False
    s.startAt = 1                   ' Default single-iteration loop
    s.stopAt = 1
    s.stepBy = 1
    s.count = 0
    s.max = LOOP_MAX_STEPS

    ' ------------------------------------------------------------
    ' Preconditions
    ' ------------------------------------------------------------
    If rsQueryResults.EOF Then
        ' No rows -> callers can still choose to emit once with defaults
        GetLoopLimits = s
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' ENUMERATE field missing -> normal mode, single iteration, no diagnostics
    ' ------------------------------------------------------------
    If Not HasField(rsQueryResults, context.fields.enumerateSwitch) Then
        GetLoopLimits = s
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' ENUMERATE field present -> read it
    ' ------------------------------------------------------------
    Dim enumerateSwitch As Boolean
    enumerateSwitch = GetFieldValueBoolean(rsQueryResults, context.fields.enumerateSwitch)

    ' ENUMERATE = FALSE -> normal mode, single iteration, no diagnostics
    If Not enumerateSwitch Then
        GetLoopLimits = s
        Exit Function
    End If

    ' From here on, ENUMERATE is explicitly enabled
    s.enabled = True

    ' ------------------------------------------------------------
    ' ENUMERATE = TRUE -> attempt to fetch loop parameters
    ' ------------------------------------------------------------
    Dim hasStart As Boolean: hasStart = HasField(rsQueryResults, context.fields.enumerateStartAt)
    Dim hasStop  As Boolean: hasStop = HasField(rsQueryResults, context.fields.enumerateStopAt)
    Dim hasStep  As Boolean: hasStep = HasField(rsQueryResults, context.fields.enumerateStepBy)

    ' Missing parameters -> single iteration + diagnostic
    If Not hasStart Or Not hasStop Or Not hasStep Then
        LogDiagnostic "ENUMERATE=TRUE but loop parameters are missing; defaulting to single-iteration mode (1 -> 1 step 1)."
        GetLoopLimits = s
        Exit Function
    End If

    ' ------------------------------------------------------------
    ' Extract supplied values
    ' ------------------------------------------------------------
    s.startAt = GetFieldValueLong(rsQueryResults, context.fields.enumerateStartAt)
    s.stopAt = GetFieldValueLong(rsQueryResults, context.fields.enumerateStopAt)
    s.stepBy = GetFieldValueLong(rsQueryResults, context.fields.enumerateStepBy)

    ' ------------------------------------------------------------
    ' Caller can override the loop governor
    ' ------------------------------------------------------------
    Dim hasMax As Boolean: hasMax = HasField(rsQueryResults, context.fields.enumerateMax)
    If hasMax Then
        s.max = GetFieldValueLong(rsQueryResults, context.fields.enumerateMax)
        
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

Private Sub IterativeSearch( _
    ByRef connectionObject As Object, _
    ByRef context As sqlContext, _
    ByVal rsQueryResults As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Preconditions
    If rsQueryResults Is Nothing Then Exit Sub
    If rsQueryResults.State <> adStateOpen Then Exit Sub
    If rsQueryResults.EOF And rsQueryResults.BOF Then Exit Sub

    If Not HasField(rsQueryResults, context.fields.iterate) Then Exit Sub
    If Not HasField(rsQueryResults, context.fields.idQuery) Then Exit Sub
    If Not HasField(rsQueryResults, context.fields.dataQuery) Then Exit Sub

    ' Extract the get-ID-list query
    Dim idQuery As String
    idQuery = GetFieldValueString(rsQueryResults, context.fields.idQuery)
    If Len(idQuery) = 0 Then Exit Sub

    ' Extract the parameterized data query
    Dim dataQueryTemplate As String
    dataQueryTemplate = GetFieldValueString(rsQueryResults, context.fields.dataQuery)
    If Len(dataQueryTemplate) = 0 Then Exit Sub

    ' Run the query which returns the list of IDs
    Dim idList As Object
    Set idList = GetIDList(connectionObject, idQuery)
    If idList Is Nothing Then Exit Sub
    If idList.count = 0 Then Exit Sub

    ' Run the data query for each ID
    Dim id As Variant
    Dim rsData As Object

    For Each id In idList.Keys
        Set rsData = RunParameterizedQuery(connectionObject, context, dataQueryTemplate, id)

        If Not rsData Is Nothing Then
            If rsData.State = adStateOpen Then
                MapResultsToDataWorksheet context, rsData, row, recordCnt
            End If
        End If

        SafeCloseRecordset rsData
    Next id

End Sub

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
    If rs.State <> adStateOpen Then Exit Function
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

    ' No ID field ? return Nothing
    If idFieldIndex = -1 Then
        SafeCloseRecordset rs
        Exit Function
    End If

    ' Always start at the beginning
    On Error Resume Next
    rs.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
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
        "GetIDList SQL failed: " & Err.Description & vbNewLine, _
        errorNumber:=Err.number, _
        sql:=idQuery, _
        errorCategory:="Iteration / SQL"

    On Error Resume Next
    SafeCloseRecordset rs
    Set GetIDList = Nothing
End Function

Private Function RunParameterizedQuery( _
    ByRef connectionObject As Object, _
    ByRef context As sqlContext, _
    ByVal dataQueryTemplate As String, _
    ByVal id As Variant) As Object

    Dim rsData As Object
    Dim sql As String

    On Error GoTo RunQueryError

    ' Substitute placeholder for the derived ID (Null-safe)
    sql = replace(dataQueryTemplate, context.fields.idPlaceholder, SafeStr(id), , , vbTextCompare)

    ' Create fresh recordset
    Set rsData = CreateObject("ADODB.Recordset")
    rsData.CursorLocation = adUseClient

    rsData.Open sql, connectionObject, adOpenStatic, adLockReadOnly

    ' Guard against providers that return a closed recordset
    If rsData Is Nothing Then
        Set RunParameterizedQuery = Nothing
        Exit Function
    End If

    If rsData.State <> adStateOpen Then
        SafeCloseRecordset rsData
        Set RunParameterizedQuery = Nothing
        Exit Function
    End If

    Set RunParameterizedQuery = rsData
    Exit Function

RunQueryError:
    LogDiagnostic _
        "RunParameterizedQuery failed: " & Err.Description, _
        errorNumber:=Err.number, _
        sql:=sql, _
        errorCategory:=ClassifyError(Err.Description)

    On Error Resume Next
    SafeCloseRecordset rsData
    Set RunParameterizedQuery = Nothing
End Function

Private Sub RecursiveSearch(ByRef connectionObject As Object, _
                                  ByRef context As sqlContext, _
                                  ByVal rsQueryResults As Object, _
                                  ByRef rsRecursionResults As Object)
    
    If rsQueryResults.EOF Then Exit Sub
    If Not HasField(rsQueryResults, context.fields.treeQuery) Then Exit Sub
    
    Dim recursionSql As String
    Dim whereValue As String
    Dim whereColumn As String
    
    ' Extract the query and parameters. Exit if not provided
    recursionSql = GetFieldValueString(rsQueryResults, context.fields.treeQuery)
    If Len(recursionSql) = 0 Then Exit Sub
    
    whereValue = GetFieldValueString(rsQueryResults, context.fields.whereValue)
    If Len(whereValue) = 0 Then Exit Sub
    
    whereColumn = GetFieldValueString(rsQueryResults, context.fields.whereColumn)
    If Len(whereColumn) = 0 Then Exit Sub
    
    ' Create a collection to track what has been searched, so we
    ' don't fall into an infinite loop.
    Dim searchedIDs As Object
    Set searchedIDs = CreateObject("Scripting.Dictionary")
    
    ' Place limits on how many recursive calls can be made
    Dim maxDepth As Long
    maxDepth = GetFieldValueLong(rsQueryResults, context.fields.maxDepth)
    
    If maxDepth = 0 Then
        maxDepth = DEFAULT_MAX_RECURSION_DEPTH
    End If
    
    Dim currentDepth As Long
    currentDepth = 0
    
    ' Execute SQL recursively until all branches of the tree are followed
    PerformRecursiveSearch connectionObject, context, recursionSql, whereValue, whereColumn, currentDepth, maxDepth, rsRecursionResults, searchedIDs
    
End Sub
  
Private Function GetFieldValueString(ByVal recordSet As Object, ByRef fieldName As String) As String
    If HasField(recordSet, fieldName) Then
        GetFieldValueString = Trim$(CStr(recordSet.fields(fieldName).value))
    Else
        GetFieldValueString = vbNullString
    End If
End Function

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

Private Sub PerformRecursiveSearch( _
    ByRef connectionObject As Object, _
    ByRef context As sqlContext, _
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
    query = replace(sqlStatement, "{" & context.fields.whereValue & "}", SafeStr(whereValue), , , vbTextCompare)

    ' Mark this ID as searched
    AddToSearchedList whereValue, searchedIDs

    ' Execute recursive query
    Dim rsQueryResults As Object
    Set rsQueryResults = CreateObject("ADODB.Recordset")
    rsQueryResults.CursorLocation = adUseClient
    rsQueryResults.Open query, connectionObject, adOpenStatic, adLockReadOnly

    ' Guard against invalid or empty recordsets
    If rsQueryResults Is Nothing Then Exit Sub
    If rsQueryResults.State <> adStateOpen Then Exit Sub
    If rsQueryResults.EOF And rsQueryResults.BOF Then
        SafeCloseRecordset rsQueryResults
        Exit Sub
    End If

    ' Initialize merged recordset structure
    If recursionRecordSet Is Nothing Then
        Set recursionRecordSet = CreateObject("ADODB.Recordset")

        Dim fieldNumber As Long
        For fieldNumber = 0 To rsQueryResults.fields.count - 1
            recursionRecordSet.fields.Append _
                rsQueryResults.fields(fieldNumber).name, _
                rsQueryResults.fields(fieldNumber).Type, _
                rsQueryResults.fields(fieldNumber).DefinedSize
        Next fieldNumber

        recursionRecordSet.Open
    End If

    ' Iterate through results
    On Error Resume Next
    rsQueryResults.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        SafeCloseRecordset rsQueryResults
        Exit Sub
    End If
    On Error GoTo RecursionError

    Do While Not rsQueryResults.EOF

        ' Append row to merged recordset
        recursionRecordSet.AddNew
        For fieldNumber = 0 To rsQueryResults.fields.count - 1
            recursionRecordSet.fields(fieldNumber).value = _
                SafeFieldValue(rsQueryResults, rsQueryResults.fields(fieldNumber).name)
        Next fieldNumber
        recursionRecordSet.Update

        ' Recurse using safe field value
        Dim nextValue As String
        nextValue = SafeFieldValue(rsQueryResults, whereColumn)

        PerformRecursiveSearch _
            connectionObject, context, sqlStatement, nextValue, _
            whereColumn, currentDepth, maxDepth, recursionRecordSet, searchedIDs

        rsQueryResults.MoveNext
    Loop

    SafeCloseRecordset rsQueryResults
    Exit Sub

RecursionError:
    LogDiagnostic _
        "Recursive SQL failed: " & Err.Description & vbNewLine & _
        "  Query: " & query & vbNewLine & _
        "  whereValue   = " & whereValue & vbNewLine & _
        "  whereColumn  = " & whereColumn & vbNewLine & _
        "  currentDepth = " & CStr(currentDepth) & vbNewLine & _
        "  maxDepth     = " & CStr(maxDepth) & vbNewLine, _
        errorNumber:=Err.number, _
        sql:=query, _
        errorCategory:="Recursion / SQL"

    On Error Resume Next
    SafeCloseRecordset rsQueryResults
    On Error GoTo 0
End Sub

Private Sub AddToSearchedList(ByRef rowId As Variant, ByVal searchedIDs As Object)
    ' Add the ID to the dictionary
    searchedIDs.Add CStr(rowId), True
End Sub

Private Function WasAlreadySearched(ByRef rowId As Variant, ByVal searchedIDs As Object) As Boolean
    ' Check if the ID is already in the dictionary
    WasAlreadySearched = searchedIDs.Exists(CStr(rowId))
End Function

Private Sub MapResultsToDataWorksheet( _
    ByRef context As sqlContext, _
    ByVal rsQueryResults As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Exit early if invalid or empty
    If rsQueryResults Is Nothing Then Exit Sub
    If rsQueryResults.State <> adStateOpen Then Exit Sub
    If rsQueryResults.EOF And rsQueryResults.BOF Then Exit Sub

    ' Special-case overrides: edges and rank
    If HasField(rsQueryResults, context.fields.CreateEdges) Then
        CreateEdges context, rsQueryResults, row, recordCnt
        Exit Sub
    End If

    If HasField(rsQueryResults, context.fields.CreateRank) Then
        CreateRank context, rsQueryResults, row, recordCnt
        Exit Sub
    End If

    ' Determine cluster/subcluster presence
    Dim hasCluster As Boolean
    Dim hasSubcluster As Boolean

    hasCluster = HasField(rsQueryResults, context.fields.Cluster)
    hasSubcluster = HasField(rsQueryResults, context.fields.subcluster)

    ' Always start from BOF before dispatching
    rsQueryResults.MoveFirst

    ' Dispatch to the correct processing routine
    If hasCluster Then
        If hasSubcluster Then
            ProcessClusterYesSubclusterYes context, rsQueryResults, row, recordCnt
        Else
            ProcessClusterYesSubclusterNo context, rsQueryResults, row, recordCnt
        End If
    Else
        If hasSubcluster Then
            ProcessClusterNoSubclusterYes context, rsQueryResults, row, recordCnt
        Else
            ProcessClusterNoSubclusterNo context, rsQueryResults, row, recordCnt
        End If
    End If

End Sub

Private Sub ProcessClusterYesSubclusterYes( _
    ByRef context As sqlContext, _
    ByVal recordSetObject As Object, _
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
    If recordSetObject Is Nothing Then Exit Sub
    If recordSetObject.State <> adStateOpen Then Exit Sub
    If recordSetObject.EOF And recordSetObject.BOF Then Exit Sub

    ' Collect distinct clusters
    Set clusterList = GetClusterInfo(recordSetObject, context.fields)

    If clusterList.count > 0 Then
        ' Attach subcluster dictionaries to each cluster
        For Each clusterKey In clusterList.Keys()
            Set clusterInstance = clusterList.item(clusterKey)
            Set clusterInstance.subclusters = GetSubClusterInfoForCluster( _
                                                recordSetObject, _
                                                context.fields, _
                                                CStr(clusterKey))
        Next clusterKey
    End If

    ' Emit clusters (with or without subclusters)
    For Each clusterKey In clusterList.Keys()

        clusterCnt = clusterCnt + 1
        Set clusterRecord = clusterList.item(CStr(clusterKey))

        EmitClusterOpen clusterRecord, context.dataLayout, row, _
                        context.fields.clusterPlaceholder, clusterCnt

        If clusterRecord.subclusters.count = 0 Then
            ' No subclusters: emit all rows for this cluster
            On Error Resume Next
            recordSetObject.MoveFirst
            If Err.number <> 0 Then
                Err.Clear
                Exit For
            End If
            On Error GoTo 0

            Do While Not recordSetObject.EOF
                If SafeFieldValue(recordSetObject, context.fields.Cluster) = CStr(clusterKey) Then
                    EmitRows context, recordSetObject, row, recordCnt
                End If
                recordSetObject.MoveNext
            Loop

        Else
            ' Has subclusters: group rows by cluster + subcluster
            subclusterCnt = 0

            For Each subclusterKey In clusterRecord.subclusters.Keys()

                Set subclusterRecord = clusterRecord.subclusters.item(subclusterKey)

                On Error Resume Next
                recordSetObject.MoveFirst
                If Err.number <> 0 Then
                    Err.Clear
                    Exit For
                End If
                On Error GoTo 0

                subclusterCnt = subclusterCnt + 1

                EmitClusterOpen subclusterRecord, context.dataLayout, row, _
                                context.fields.subclusterPlaceholder, subclusterCnt

                Do While Not recordSetObject.EOF
                    If SafeFieldValue(recordSetObject, context.fields.Cluster) = CStr(clusterKey) _
                       And SafeFieldValue(recordSetObject, context.fields.subcluster) = CStr(subclusterKey) Then
                        EmitRows context, recordSetObject, row, recordCnt
                    End If
                    recordSetObject.MoveNext
                Loop

                EmitClusterClose subclusterRecord, context.dataLayout, row, _
                                 context.fields.subclusterPlaceholder, subclusterCnt

                ' Emit rows in this cluster with NULL subcluster
                On Error Resume Next
                recordSetObject.MoveFirst
                If Err.number <> 0 Then
                    Err.Clear
                    Exit For
                End If
                On Error GoTo 0

                Do While Not recordSetObject.EOF
                    If SafeFieldValue(recordSetObject, context.fields.Cluster) = CStr(clusterKey) _
                       And SafeFieldValue(recordSetObject, context.fields.subcluster) = "" Then
                        EmitRows context, recordSetObject, row, recordCnt
                    End If
                    recordSetObject.MoveNext
                Loop

            Next subclusterKey
        End If

        EmitClusterClose clusterRecord, context.dataLayout, row, _
                         context.fields.clusterPlaceholder, clusterCnt
    Next clusterKey

    ' Handle case where cluster has no data, but subcluster does
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Set orphanClusterList = GetOrphanSubClusterInfo(recordSetObject, context.fields)
    subclusterCnt = 0

    For Each subclusterKey In orphanClusterList.Keys()

        Set subclusterRecord = orphanClusterList.item(subclusterKey)

        On Error Resume Next
        recordSetObject.MoveFirst
        If Err.number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0

        subclusterCnt = subclusterCnt + 1

        EmitClusterOpen subclusterRecord, context.dataLayout, row, _
                        context.fields.subclusterPlaceholder, subclusterCnt

        Do While Not recordSetObject.EOF
            If SafeFieldValue(recordSetObject, context.fields.Cluster) = "" _
               And SafeFieldValue(recordSetObject, context.fields.subcluster) = CStr(subclusterKey) Then
                EmitRows context, recordSetObject, row, recordCnt
            End If
            recordSetObject.MoveNext
        Loop

        EmitClusterClose subclusterRecord, context.dataLayout, row, _
                         context.fields.subclusterPlaceholder, subclusterCnt
    Next subclusterKey

    ' Handle rows where both cluster and subcluster are NULL
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not recordSetObject.EOF
        If SafeFieldValue(recordSetObject, context.fields.Cluster) = "" _
           And SafeFieldValue(recordSetObject, context.fields.subcluster) = "" Then
            EmitRows context, recordSetObject, row, recordCnt
        End If
        recordSetObject.MoveNext
    Loop

End Sub

Private Sub ProcessClusterYesSubclusterNo( _
    ByRef context As sqlContext, _
    ByVal recordSetObject As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid recordsets
    If recordSetObject Is Nothing Then Exit Sub
    If recordSetObject.State <> adStateOpen Then Exit Sub
    If recordSetObject.EOF And recordSetObject.BOF Then Exit Sub

    Dim clusterList As Dictionary
    Set clusterList = GetClusterInfo(recordSetObject, context.fields)

    Dim clusterCnt As Long
    clusterCnt = 0

    Dim clusterKey As Variant
    Dim clusterRecord As Cluster

    ' Emit each cluster block
    For Each clusterKey In clusterList.Keys()

        clusterCnt = clusterCnt + 1
        Set clusterRecord = clusterList.item(CStr(clusterKey))

        EmitClusterOpen clusterRecord, context.dataLayout, row, _
                        context.fields.clusterPlaceholder, clusterCnt

        ' Safe MoveFirst
        On Error Resume Next
        recordSetObject.MoveFirst
        If Err.number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0

        ' Emit rows belonging to this cluster
        Do While Not recordSetObject.EOF
            If SafeFieldValue(recordSetObject, context.fields.Cluster) = CStr(clusterKey) Then
                EmitRows context, recordSetObject, row, recordCnt
            End If
            recordSetObject.MoveNext
        Loop

        EmitClusterClose clusterRecord, context.dataLayout, row, _
                         context.fields.clusterPlaceholder, clusterCnt
    Next clusterKey

    ' Emit orphan rows (cluster column is Null)
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not recordSetObject.EOF
        If SafeFieldValue(recordSetObject, context.fields.Cluster) = "" Then
            EmitRows context, recordSetObject, row, recordCnt
        End If
        recordSetObject.MoveNext
    Loop

End Sub

Private Sub ProcessClusterNoSubclusterYes( _
    ByRef context As sqlContext, _
    ByVal recordSetObject As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid recordsets
    If recordSetObject Is Nothing Then Exit Sub
    If recordSetObject.State <> adStateOpen Then Exit Sub
    If recordSetObject.EOF And recordSetObject.BOF Then Exit Sub

    Dim subclusterList As Dictionary
    Set subclusterList = GetSubclusterInfo(recordSetObject, context.fields)

    Dim subclusterCnt As Long
    subclusterCnt = 0

    Dim subclusterKey As Variant
    Dim subclusterRecord As Cluster

    ' Emit each subcluster block
    For Each subclusterKey In subclusterList.Keys()

        subclusterCnt = subclusterCnt + 1
        Set subclusterRecord = subclusterList.item(CStr(subclusterKey))

        EmitClusterOpen subclusterRecord, context.dataLayout, row, _
                        context.fields.subclusterPlaceholder, subclusterCnt

        ' Safe MoveFirst
        On Error Resume Next
        recordSetObject.MoveFirst
        If Err.number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo 0

        ' Emit rows belonging to this subcluster
        Do While Not recordSetObject.EOF
            If SafeFieldValue(recordSetObject, context.fields.subcluster) = CStr(subclusterKey) Then
                EmitRows context, recordSetObject, row, recordCnt
            End If
            recordSetObject.MoveNext
        Loop

        EmitClusterClose subclusterRecord, context.dataLayout, row, _
                         context.fields.subclusterPlaceholder, subclusterCnt
    Next subclusterKey

    ' Emit orphan rows (subcluster column is Null)
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not recordSetObject.EOF
        If SafeFieldValue(recordSetObject, context.fields.subcluster) = "" Then
            EmitRows context, recordSetObject, row, recordCnt
        End If
        recordSetObject.MoveNext
    Loop

End Sub

Private Sub ProcessClusterNoSubclusterNo( _
    ByRef context As sqlContext, _
    ByVal recordSetObject As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid or empty recordsets
    If recordSetObject Is Nothing Then Exit Sub
    If recordSetObject.State <> adStateOpen Then Exit Sub
    If recordSetObject.EOF And recordSetObject.BOF Then Exit Sub

    ' Always start at the beginning
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Emit each row
    Do While Not recordSetObject.EOF
        EmitRows context, recordSetObject, row, recordCnt
        recordSetObject.MoveNext
    Loop

End Sub

Private Sub CreateEdges( _
    ByRef context As sqlContext, _
    ByVal recordSetObject As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Guard against invalid or empty recordsets
    If recordSetObject Is Nothing Then Exit Sub
    If recordSetObject.State <> adStateOpen Then Exit Sub
    If recordSetObject.EOF And recordSetObject.BOF Then Exit Sub

    ' Safe MoveFirst
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Dim item As String
    item = GetFieldValueString(recordSetObject, context.headings.item)

    Dim relatedItem As String
    Dim emittedRow As Long

    ' ------------------------------------------------------------
    ' LOOP MODE
    ' ------------------------------------------------------------
    If context.loop.enabled Then

        Dim stopValue As Long
        stopValue = context.loop.stopAt - 1

        Dim i As Long
        For i = context.loop.startAt To stopValue Step context.loop.stepBy

            context.loop.count = context.loop.count + 1
            If context.loop.count > context.loop.max Then Exit For

            emittedRow = row   ' capture before EmitOneRow increments it
            EmitOneRow context, recordSetObject, row, recordCnt, i

            DataSheet.Cells.item(emittedRow, context.dataLayout.itemColumn) = _
                replace(SafeStr(item), context.fields.enumeratePlaceholder, CStr(i), , , vbTextCompare)

            DataSheet.Cells.item(emittedRow, context.dataLayout.isRelatedToItemColumn) = _
                replace(SafeStr(item), context.fields.enumeratePlaceholder, CStr(i + 1), , , vbTextCompare)

        Next i

    ' ------------------------------------------------------------
    ' NORMAL MODE
    ' ------------------------------------------------------------
    Else

        ' Safe MoveNext (skip first row)
        On Error Resume Next
        recordSetObject.MoveNext
        If Err.number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0

        Do While Not recordSetObject.EOF

            relatedItem = GetFieldValueString(recordSetObject, context.headings.item)

            emittedRow = row   ' capture before EmitRows increments it
            EmitRows context, recordSetObject, row, recordCnt

            DataSheet.Cells.item(emittedRow, context.dataLayout.itemColumn) = item
            DataSheet.Cells.item(emittedRow, context.dataLayout.isRelatedToItemColumn) = relatedItem

            item = relatedItem

            recordSetObject.MoveNext
        Loop

    End If

End Sub

Private Sub CreateRank( _
    ByRef context As sqlContext, _
    ByVal recordSetObject As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Exit early if invalid or empty
    If recordSetObject Is Nothing Then Exit Sub
    If recordSetObject.State <> adStateOpen Then Exit Sub
    If recordSetObject.EOF And recordSetObject.BOF Then Exit Sub

    ' Safe MoveFirst
    On Error Resume Next
    recordSetObject.MoveFirst
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    ' Establish the rank (Null-safe)
    Dim rank As String
    rank = LCase$(SafeFieldValue(recordSetObject, "RANK"))

    ' Collect node identifiers
    Dim item As String
    Dim subgraph As String
    subgraph = "{ rank=" & AddQuotes(rank) & ";"

    Do While Not recordSetObject.EOF
        item = SafeFieldValue(recordSetObject, context.headings.item)
        subgraph = subgraph & " " & AddQuotes(item) & ";"
        recordSetObject.MoveNext
    Loop

    subgraph = subgraph & " }"

    ' Emit the row
    recordCnt = recordCnt + 1
    DataSheet.Cells.item(row, context.dataLayout.itemColumn) = ">"
    DataSheet.Cells.item(row, context.dataLayout.labelColumn) = subgraph
    row = row + 1
End Sub

Private Function GetClusterInfo(ByVal recordSetObject As Object, _
                                ByRef fields As sqlFieldName) As Dictionary
    Dim clusters As Dictionary
    Set clusters = New Dictionary

    ' Exit early if empty
    If recordSetObject.EOF Then
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
    hasClusterLabel = HasField(recordSetObject, fields.clusterLabel)

    Dim fieldObject As Variant
    Dim clusterId As String
    Dim clusterLabel As String
    Dim clusterStyleName As String
    Dim clusterAttributes As String
    Dim clusterTooltip As String

    recordSetObject.MoveFirst

    Do While Not recordSetObject.EOF
        clusterId = vbNullString
        clusterLabel = vbNullString
        clusterStyleName = vbNullString
        clusterAttributes = vbNullString
        clusterTooltip = vbNullString

        ' Extract cluster metadata
        For Each fieldObject In recordSetObject.fields
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
                clusterObject.tooltip = clusterTooltip

                clusters.Add clusterId, clusterObject
            End If
        End If

        recordSetObject.MoveNext
    Loop

    recordSetObject.MoveFirst
    Set GetClusterInfo = clusters
End Function

Private Function GetSubclusterInfo( _
    ByVal recordSetObject As Object, _
    ByRef fields As sqlFieldName) As Dictionary

    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    ' Exit early if empty or invalid
    If recordSetObject Is Nothing Then
        Set GetSubclusterInfo = subclusters
        Exit Function
    End If

    If recordSetObject.State <> adStateOpen Then
        Set GetSubclusterInfo = subclusters
        Exit Function
    End If

    If recordSetObject.EOF And recordSetObject.BOF Then
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
    hasSubLabel = HasField(recordSetObject, fields.subclusterLabel)

    Dim fieldObject As Variant
    Dim subId As String
    Dim subLabel As String
    Dim subStyle As String
    Dim subAttr As String
    Dim subTooltip As String

    recordSetObject.MoveFirst

    Do While Not recordSetObject.EOF

        subId = ""
        subLabel = ""
        subStyle = ""
        subAttr = ""
        subTooltip = ""

        ' Extract subcluster metadata
        For Each fieldObject In recordSetObject.fields
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
                clusterObject.tooltip = subTooltip

                subclusters.Add subId, clusterObject
            End If
        End If

        recordSetObject.MoveNext
    Loop

    recordSetObject.MoveFirst
    Set GetSubclusterInfo = subclusters
End Function

Private Function GetSubClusterInfoForCluster( _
    ByVal recordSetObject As Object, _
    ByRef fields As sqlFieldName, _
    ByVal clusterName As String) As Dictionary

    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    ' Exit early if empty or invalid
    If recordSetObject Is Nothing Then
        Set GetSubClusterInfoForCluster = subclusters
        Exit Function
    End If

    If recordSetObject.State <> adStateOpen Then
        Set GetSubClusterInfoForCluster = subclusters
        Exit Function
    End If

    If recordSetObject.EOF And recordSetObject.BOF Then
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
    hasSubLabel = HasField(recordSetObject, fields.subclusterLabel)

    Dim fieldObject As Variant
    Dim subId As String
    Dim subLabel As String
    Dim subStyle As String
    Dim subAttr As String
    Dim subTooltip As String

    recordSetObject.MoveFirst

    Do While Not recordSetObject.EOF

        ' Only process rows belonging to this cluster (Null-safe)
        If SafeFieldValue(recordSetObject, fields.Cluster) = clusterName Then

            subId = vbNullString
            subLabel = vbNullString
            subStyle = vbNullString
            subAttr = vbNullString
            subTooltip = vbNullString

            ' Extract subcluster metadata
            For Each fieldObject In recordSetObject.fields
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
                    clusterObject.tooltip = subTooltip

                    subclusters.Add subId, clusterObject
                End If
            End If

        End If

        recordSetObject.MoveNext
    Loop

    recordSetObject.MoveFirst
    Set GetSubClusterInfoForCluster = subclusters
End Function

Private Function GetOrphanSubClusterInfo( _
    ByVal recordSetObject As Object, _
    ByRef fields As sqlFieldName) As Dictionary

    ' Build a list of subclusters where the cluster column is null
    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    ' Exit early if empty or invalid
    If recordSetObject Is Nothing Then
        Set GetOrphanSubClusterInfo = subclusters
        Exit Function
    End If

    If recordSetObject.State <> adStateOpen Then
        Set GetOrphanSubClusterInfo = subclusters
        Exit Function
    End If

    If recordSetObject.EOF And recordSetObject.BOF Then
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
    hasSubLabel = HasField(recordSetObject, fields.subclusterLabel)

    Dim fieldObject As Variant
    Dim subId As String
    Dim subLabel As String
    Dim subStyle As String
    Dim subAttr As String
    Dim subTooltip As String

    recordSetObject.MoveFirst

    Do While Not recordSetObject.EOF

        ' Only process rows where cluster is NULL/empty and subcluster is NOT NULL/empty
        If SafeFieldValue(recordSetObject, fields.Cluster) = "" _
           And Len(SafeFieldValue(recordSetObject, fields.subcluster)) > 0 Then

            subId = ""
            subLabel = ""
            subStyle = ""
            subAttr = ""
            subTooltip = ""

            ' Extract subcluster metadata
            For Each fieldObject In recordSetObject.fields
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
                    clusterObject.tooltip = subTooltip

                    subclusters.Add subId, clusterObject
                End If
            End If

        End If

        recordSetObject.MoveNext
    Loop

    recordSetObject.MoveFirst
    Set GetOrphanSubClusterInfo = subclusters
End Function

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

        .Cells(row, dataLayout.tooltipColumn).value = SafeStr(clusterRecord.tooltip)

        If Len(SafeStr(clusterRecord.styleName)) > 0 Then
            newStyle = replace(SafeStr(clusterRecord.styleName), _
                               findStr, SafeStr(replaceLong), , , vbTextCompare) _
                       & suffix

            .Cells(row, dataLayout.styleNameColumn).value = newStyle
        End If
    End With

    row = row + 1
End Sub

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

Private Function HasField(ByVal recordSetObject As Object, ByVal fieldName As String) As Boolean
    Dim fieldObject As Variant

    For Each fieldObject In recordSetObject.fields
        If Trim$(LCase$(CStr(fieldObject.name))) = LCase$(fieldName) Then
            HasField = True
            Exit For
        End If
    Next
End Function

Private Sub EmitRows( _
    ByRef context As sqlContext, _
    ByVal rs As Object, _
    ByRef targetRow As Long, _
    ByRef position As Long)

    Dim i As Long

    ' Safety: prevent infinite loop
    If context.loop.stepBy = 0 Then Exit Sub

    ' Safety: prevent direction mismatch infinite loop
    If context.loop.stepBy > 0 Then
        If context.loop.startAt > context.loop.stopAt Then Exit Sub
    Else
        If context.loop.startAt < context.loop.stopAt Then Exit Sub
    End If

    For i = context.loop.startAt To context.loop.stopAt Step context.loop.stepBy
        context.loop.count = context.loop.count + 1
        If context.loop.count > context.loop.max Then Exit For

        EmitOneRow context, rs, targetRow, position, i
    Next i

End Sub

Private Sub EmitOneRow( _
    ByRef context As sqlContext, _
    ByVal rs As Object, _
    ByRef targetRow As Long, _
    ByRef position As Long, _
    ByVal enumStep As Long)

    ' Increment the result set position (i.e. recordCnt)
    position = position + 1

    ' ---------------------------------------------------------------
    ' Determine multiline splitting behavior for THIS row / THIS query
    ' ---------------------------------------------------------------
    Dim doSplit     As Boolean: doSplit = False
    Dim splitLength As Long:    splitLength = 0
    Dim lineEnding  As String:  lineEnding = NEWLINE    ' default

    Dim temp As String

    ' Check & read split length if the column exists in this resultset
    If HasField(rs, context.fields.splitLength) Then
        temp = SafeFieldValue(rs, context.fields.splitLength)
        If Len(temp) > 0 Then
            If IsNumeric(temp) Then
                splitLength = CLng(temp)
                If splitLength > 0 Then doSplit = True
            End If
        End If
    End If

    ' Check & read custom line ending if present
    If HasField(rs, context.fields.lineEnding) Then
        temp = SafeFieldValue(rs, context.fields.lineEnding)
        If Len(temp) > 0 Then
            lineEnding = temp
        End If
    End If

    ' ---------------------------------------------------------------
    ' Process all fields and write to sheet
    ' ---------------------------------------------------------------
    With DataSheet
        Dim fld As Object
        Dim v As String
        Dim targetCol As Long

        For Each fld In rs.fields

            ' Common transformation: null ? "", placeholder replacement
            v = SafeStr(fld.value)

            If Len(v) > 0 Then
                v = replace(v, context.fields.recordsetPlaceholder, SafeStr(position), , , vbTextCompare)
                If context.loop.enabled Then
                    v = replace(v, context.fields.enumeratePlaceholder, SafeStr(enumStep), , , vbTextCompare)
                End If
            End If

            Select Case LCase$(fld.name)

                Case context.headings.flag
                    .Cells(targetRow, context.dataLayout.flagColumn).value = v

                Case context.headings.item
                    .Cells(targetRow, context.dataLayout.itemColumn).value = v

                Case context.headings.label, context.headings.xLabel
                    targetCol = IIf(LCase$(fld.name) = context.headings.label, _
                                    context.dataLayout.labelColumn, _
                                    context.dataLayout.xLabelColumn)

                    ' Apply multiline splitting only when requested & meaningful
                    If doSplit Then
                        v = SplitMultilineText(v, splitLength, lineEnding)
                    End If

                    .Cells(targetRow, targetCol).value = v

                Case context.headings.tailLabel
                    .Cells(targetRow, context.dataLayout.tailLabelColumn).value = v

                Case context.headings.headLabel
                    .Cells(targetRow, context.dataLayout.headLabelColumn).value = v

                Case context.headings.tooltip
                    .Cells(targetRow, context.dataLayout.tooltipColumn).value = v

                Case context.headings.isRelatedToItem
                    .Cells(targetRow, context.dataLayout.isRelatedToItemColumn).value = v

                Case context.headings.styleName
                    .Cells(targetRow, context.dataLayout.styleNameColumn).value = v

                Case context.headings.extraAttributes
                    .Cells(targetRow, context.dataLayout.extraAttributesColumn).value = v

                Case context.headings.errorMessage
                    .Cells(targetRow, context.dataLayout.errorMessageColumn).value = v

                ' Case Else: ignore unknown columns (intentional, general-purpose)
            End Select
        Next fld

        ' Increment the row counter once the row has been fully emitted
        targetRow = targetRow + 1
    End With

End Sub

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
        .tooltip = NormalizeHeading(rowValues(1, dataLayout.tooltipColumn))
        .isRelatedToItem = NormalizeHeading(rowValues(1, dataLayout.isRelatedToItemColumn))
        .styleName = NormalizeHeading(rowValues(1, dataLayout.styleNameColumn))
        .extraAttributes = NormalizeHeading(rowValues(1, dataLayout.extraAttributesColumn))
        .errorMessage = NormalizeHeading(rowValues(1, dataLayout.errorMessageColumn))
    End With
End Function

Private Function NormalizeHeading(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or v = vbNullString Then
        NormalizeHeading = vbNullString
    Else
        NormalizeHeading = Trim$(LCase$(CStr(v)))
    End If
End Function

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

Private Function ClassifyError(ByVal errMsg As String) As String
    Dim m As String
    m = LCase$(errMsg)

    ' --- Worksheet / table not found ---
    If ContainsAny(m, Array( _
        "could not find the object", _
        "not a valid name", _
        "external table is not in the expected format")) Then

        ClassifyError = "Worksheet (table) not found"
        Exit Function
    End If

    ' --- Column not found ---
    If ContainsAny(m, Array( _
        "does not recognize", _
        "no value given")) Then

        ClassifyError = "Column not found"
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

Private Function ContainsAny(ByVal Text As String, ByVal patterns As Variant) As Boolean
    Dim p As Variant
    For Each p In patterns
        If InStr(Text, p) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next p
End Function

Private Function SafeStr(ByVal v As Variant) As String
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


Private Function SafeFieldValue(ByVal recordSetObject As Object, ByVal fieldName As String) As String
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
    If recordSetObject Is Nothing Then
        SafeFieldValue = ""
        Exit Function
    End If

    If recordSetObject.State <> adStateOpen Then
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
    Set fld = recordSetObject.fields(fieldName)
    If Err.number <> 0 Then
        Err.Clear
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




