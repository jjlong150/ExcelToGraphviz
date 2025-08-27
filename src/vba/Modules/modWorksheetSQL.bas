Attribute VB_Name = "modWorksheetSQL"
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.SQL")

Option Explicit

Private Type sqlContext
    dataLayout As dataWorksheet
    fields As sqlFieldName
    headings As DataWorksheetHeadings
    sqlLayout As sqlWorksheet
End Type


''' Button Actions - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'''  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Sub RunSQL()
    Dim message As String
    Dim filePath As String
    
    Dim context As sqlContext
    
    Dim connectionObject As Object  ' Connection
    
    ' Get the column layout of the 'data' worksheet
    context.dataLayout = GetSettingsForDataWorksheet(DataSheet.name)
    
    ' Get the heading values of the 'data' worksheet columns.
    ' SQL queries map to the localized column name so that non-english speaking
    ' people are not forced to use english column names in SQL queries.
    context.headings = GetSQLWorksheetHeadings(context.dataLayout)
    
    ' Get the column layout of the 'sql' worksheet
    context.sqlLayout = GetSettingsForSqlWorksheet()
    
    ' Get the list of special field names used for determining
    ' clusters and subclusters.
    context.fields = GetSettingsForSqlFields()
    
    ' Determine the last row with data
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    ' Disable automatic graph rendering as cells change.
    Dim runMode As String
    runMode = SettingsSheet.Range(SETTINGS_RUN_MODE).value
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_MANUAL
    
    ' Clear out the info from previous run
    ClearSQLStatus
    ClearDataWorksheet DataSheet.name
    
    Dim dataRow As Long
    dataRow = context.dataLayout.firstRow
    
    ' The column used to filter which SQL statements should be run
    Dim filterColumn As Long
    If Len(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value) = 0 Then
        filterColumn = 0
    Else
        filterColumn = GetSettingColNum(SETTINGS_SQL_COL_FILTER)
    End If

    ' Loop through the data rows of SQL statements
    Dim sqlStatement As String
    Dim sqlUCase As String
    Dim sqlRow As Long
    Dim dataFile As String
    
    For sqlRow = context.sqlLayout.firstRow To lastRow
        ' Skip initializations if the SQL row is commented out
        If SqlSheet.Cells.item(sqlRow, context.sqlLayout.flagColumn).value <> FLAG_COMMENT Then
            ' Establish the full path to the Excel file containing the data
            filePath = GetExcelFilePath(sqlRow, context.sqlLayout, dataFile)
            
            ' Establish connection to the file containing the relational data using
            ' connection pooling.
            Set connectionObject = getConnection(filePath)
            
            ' Get SQL statement, and convert to upper case
            sqlStatement = Trim$(SqlSheet.Cells.item(sqlRow, context.sqlLayout.sqlStatementColumn).value)
            sqlUCase = UCase$(sqlStatement)
            
            ' Get default SUCCESS message
            message = GetMessage("msgboxSqlStatusSuccess")
        End If
    
        If SqlSheet.Cells.item(sqlRow, context.sqlLayout.flagColumn).value = FLAG_COMMENT Then
            message = GetMessage("msgboxSqlStatusSkipped")
        
        ElseIf Len(sqlStatement) = 0 Then
            message = vbNullString
        
        ElseIf Not PassesFilter(sqlRow, filterColumn) Then
            message = GetMessage("msgboxSqlStatusFiltered")
        
        ElseIf sqlUCase = SQL_SET_DATA_FILE Then
            dataFile = SqlSheet.Cells.item(sqlRow, context.sqlLayout.excelFileColumn).value
        
        ElseIf sqlUCase = SQL_RESET Then
            ClearDataWorksheet DataSheet.name
       
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
        
        ElseIf Not FileExists(filePath) Then
            message = GetMessage("msgboxSqlFileNotFound")
            message = replace(message, "{filePath}", filePath)
            message = GetMessage("msgboxSqlStatusFailure") & " - " & message
        
        Else
            message = executeSQL(context, filePath, connectionObject, sqlStatement, dataRow)
        End If
        
        ' Display the status of the SQL query
        SqlSheet.Cells.item(sqlRow, context.sqlLayout.statusColumn).value = message
        
        ' Breathe
        DoEvents
    Next sqlRow
   
    ' Clean up connection pool if using narrow-scoped pooling, otherwise
    ' connections will be cleaned up when the workbook is closed.
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
    PassesFilter = False
    If filterColumn <= 0 Then
        PassesFilter = True
    Else
        If Trim$(SqlSheet.Cells.item(sqlRow, filterColumn).value) = SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value Then
            PassesFilter = True
        End If
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

Private Function executeSQL(ByRef context As sqlContext, _
                            ByVal filePath As String, _
                            ByRef connectionObject As Object, _
                            ByVal sqlStatement As String, _
                            ByRef row As Long) As String
    
    On Error GoTo executeSQLError
    
    Dim rsQueryResults As Object   ' Record Set
    Dim rsRecursionResults As Object
    Dim rsMergedResults As Object
        
    Dim recordCnt As Long
    recordCnt = 0
   
    ' A Microsoft bug is causing it to take 12 seconds to get a connection, so provide feedback
    Application.StatusBar = replace(GetMessage("statusbarSqlEstablishingConnection"), "{filePath}", filePath)
        
    ' Reset status bar now that the connection has been made
    Application.StatusBar = False
                
    ' Define a recordset for a SQL SELECT statement using late binding
    ' as we do not know which version of Excel this spreadsheet
    ' will be running on
    Set rsQueryResults = CreateObject("ADODB.Recordset")
    
    ' Execute the SQL SELECT query
'   recordSetObject.Open source:=sqlStatement, ActiveConnection:=connectionObject, CursorType:=CursorTypeEnum.adOpenForwardOnly, LockType:=LockTypeEnum.adLockOptimistic, options:=CommandTypeEnum.adCmdText

    rsQueryResults.Open source:=sqlStatement, ActiveConnection:=connectionObject, CursorType:=CursorTypeEnum.adOpenForwardOnly, LockType:=LockTypeEnum.adLockOptimistic, options:=CommandTypeEnum.adCmdText
        
    ' Execute any recursion query passed in the SQL SELECT
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
    
executeSQLError:

    If Err.number <> 0 Then
        ' GetMessage will reset the error state, save the message
        Dim errMsg As String
        errMsg = Err.Description
        
        Dim errNumber As Long
        errNumber = Err.number
        
        executeSQL = GetMessage("msgboxSqlStatusFailure") & " - " & errMsg & vbNewLine & vbNewLine & "Err.Number=" & errNumber & vbNewLine & vbNewLine & "datafile=" & filePath
    End If
    
    On Error Resume Next
    
    ' Close the rsQueryResults record set
    If Not rsQueryResults Is Nothing Then
        If rsQueryResults.State = ObjectStateEnum.adStateOpen Then
            rsQueryResults.Close
            Set rsQueryResults = Nothing
        End If
    End If
    
    ' Close the rsRecursionResults record set
    If Not rsRecursionResults Is Nothing Then
        If rsRecursionResults.State = ObjectStateEnum.adStateOpen Then
            rsRecursionResults.Close
            Set rsRecursionResults = Nothing
        End If
    End If
    
    ' Close the rsMergedResults record set
    If Not rsMergedResults Is Nothing Then
        If rsMergedResults.State = ObjectStateEnum.adStateOpen Then
            rsMergedResults.Close
            Set rsMergedResults = Nothing
        End If
    End If
    
    On Error GoTo 0

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
        maxDepth = 100
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

Private Sub PerformRecursiveSearch(ByRef connectionObject As Object, ByRef context As sqlContext, ByVal sqlStatement As String, ByRef whereValue As String, ByVal whereColumn As String, ByVal depth As Long, ByVal maxDepth As Long, ByRef recursionRecordSet As Object, ByRef searchedIDs As Object)

    ' Base case: if 'whereValue' value was already searched, exit the function to avoid loop
    If WasAlreadySearched(whereValue, searchedIDs) Then Exit Sub
    
    ' Impose a limit on how many levels of the tree can be recursed
    Dim currentDepth As Long
    currentDepth = depth + 1
    If currentDepth > maxDepth Then Exit Sub
    
    ' Expand placeholders in the query with actual values
    Dim query As String
    query = replace(sqlStatement, "{" & context.fields.whereValue & "}", whereValue)
    
    ' Add current ID to the list of searched IDs
    AddToSearchedList whereValue, searchedIDs
    
    ' Create a record set with the results of this query.
    ' These results will get merged into the full recursion recordset.
    Dim rsQueryResults As Object
    Set rsQueryResults = CreateObject("ADODB.Recordset")
    rsQueryResults.Open query, connectionObject, 1, 1

    ' If combinedRS is Nothing, initialize it with the structure of the first recordset
    If recursionRecordSet Is Nothing Then
        Set recursionRecordSet = CreateObject("ADODB.Recordset")
        Dim fieldNumber As Long
        
        ' Create fields in the combined recordset
        For fieldNumber = 0 To rsQueryResults.fields.count - 1
            recursionRecordSet.fields.Append rsQueryResults.fields(fieldNumber).name, rsQueryResults.fields(fieldNumber).Type, rsQueryResults.fields(fieldNumber).DefinedSize
        Next fieldNumber
        recursionRecordSet.Open
    End If

    ' Loop through each record, adding the results to the recursion recordset
    Do While Not rsQueryResults.EOF
        recursionRecordSet.AddNew
        For fieldNumber = 0 To rsQueryResults.fields.count - 1
            recursionRecordSet.fields(fieldNumber).value = rsQueryResults.fields(fieldNumber).value
        Next fieldNumber
        recursionRecordSet.Update
        
        ' Recursively call the function for the current value pair
        PerformRecursiveSearch connectionObject, context, sqlStatement, rsQueryResults.fields(whereColumn).value, _
                               whereColumn, currentDepth, maxDepth, recursionRecordSet, searchedIDs
        rsQueryResults.MoveNext
    Loop

    ' Close the recordset
    rsQueryResults.Close
    Set rsQueryResults = Nothing
End Sub

Private Sub AddToSearchedList(ByRef rowId As Variant, ByVal searchedIDs As Object)
    ' Add the ID to the dictionary
    searchedIDs.Add CStr(rowId), True
End Sub

Private Function WasAlreadySearched(ByRef rowId As Variant, ByVal searchedIDs As Object) As Boolean
    ' Check if the ID is already in the dictionary
    WasAlreadySearched = searchedIDs.Exists(CStr(rowId))
End Function

Private Sub MapResultsToDataWorksheet(ByRef context As sqlContext, _
                                      ByVal rsQueryResults As Object, _
                                      ByRef row As Long, _
                                      ByRef recordCnt As Long)
                                      
    If rsQueryResults.EOF Then Exit Sub
    
    ' Determine if the query specified chaining the items
    If HasField(rsQueryResults, context.fields.CreateEdges) Then
        CreateEdges context, rsQueryResults, row, recordCnt
        Exit Sub
    End If
    
    ' Determine if the query is asking to create a subgraph which puts
    ' all the items on the same rank. Currently this feature does not create separate
    ' subgraphs for clusters or subclusters (the juice is not worth the squeeze).
    If HasField(rsQueryResults, context.fields.CreateRank) Then
        CreateRank context, rsQueryResults, row, recordCnt
        Exit Sub
    End If

    ' Determine if the query specified clusters and/or subclusters
    Dim hasCluster As Boolean
    hasCluster = HasField(rsQueryResults, context.fields.Cluster)
    
    Dim hasSubcluster As Boolean
    hasSubcluster = HasField(rsQueryResults, context.fields.subcluster)

    ' Ensure the recordset is at the beginning
    rsQueryResults.MoveFirst
    
    ' Work the four possible combinations to emit the clustered or unclustered results
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
  
Private Sub ProcessClusterYesSubclusterYes(ByRef context As sqlContext, _
                                           ByVal recordSetObject As Object, _
                                           ByRef row As Long, _
                                           ByRef recordCnt As Long)
    Dim clusterKey As Variant
    Dim subclusterKey As Variant
    
    Dim clusterCnt As Long
    clusterCnt = 0

    ' Scan the result set, and collect a distinct set of values for fields defined as "cluster"
    Dim clusterList As Dictionary
    Set clusterList = GetClusterInfo(recordSetObject, context.fields)

    If clusterList.count > 0 Then
        ' Determine if the cluster also has subclusters
        Dim clusterInstance As Cluster
        For Each clusterKey In clusterList.Keys()
            ' Retrieve the "cluster" fields for this key
            Set clusterInstance = clusterList.item(clusterKey)
            ' Add the dictionary of subcluster info to this cluster
            Set clusterInstance.subclusters = GetSubClusterInfoForCluster(recordSetObject, context.fields, CStr(clusterKey))
        Next
    End If

    Dim clusterRecord As Cluster
    For Each clusterKey In clusterList.Keys()
        clusterCnt = clusterCnt + 1
        Set clusterRecord = clusterList.item(CStr(clusterKey))

        EmitClusterOpen clusterRecord, context.dataLayout, row, context.fields.clusterPlaceholder, clusterCnt
        If clusterRecord.subclusters.count = 0 Then ' Results do not need to be grouped in subclusters
            recordSetObject.MoveFirst
            Do While recordSetObject.EOF = False
                If recordSetObject.fields(context.fields.Cluster).value = CStr(clusterKey) Then
                    recordCnt = recordCnt + 1
                    EmitRow context, recordSetObject, row, recordCnt
                    row = row + 1
                End If
                recordSetObject.MoveNext
            Loop
        Else ' Results should be grouped in subclusters
            Dim subclusterCnt As Long
            subclusterCnt = 0

            Dim subclusterRecord As Cluster
            For Each subclusterKey In clusterRecord.subclusters.Keys()
                ' Create a row to start the subcluster
                Set subclusterRecord = clusterRecord.subclusters.item(subclusterKey)
                recordSetObject.MoveFirst
    
                subclusterCnt = subclusterCnt + 1
                EmitClusterOpen subclusterRecord, context.dataLayout, row, context.fields.subclusterPlaceholder, subclusterCnt
                Do While recordSetObject.EOF = False
                    If recordSetObject.fields(context.fields.Cluster).value = CStr(clusterKey) And recordSetObject.fields(context.fields.subcluster).value = CStr(subclusterKey) Then
                        recordCnt = recordCnt + 1
                        EmitRow context, recordSetObject, row, recordCnt
                        row = row + 1
                    End If
                    recordSetObject.MoveNext
                Loop
                EmitClusterClose subclusterRecord, context.dataLayout, row, context.fields.subclusterPlaceholder, subclusterCnt
    
                ' Iterate through the query results again for the nodes which are not part of the subcluster
                recordSetObject.MoveFirst
                Do While recordSetObject.EOF = False
                    If recordSetObject.fields(context.fields.Cluster).value = CStr(clusterKey) And IsNull(recordSetObject.fields(context.fields.subcluster).value) Then
                        recordCnt = recordCnt + 1
                        EmitRow context, recordSetObject, row, recordCnt
                        row = row + 1
                    End If
                    recordSetObject.MoveNext
                Loop
            Next
        End If
        EmitClusterClose clusterRecord, context.dataLayout, row, context.fields.clusterPlaceholder, clusterCnt
    Next
    
    ' Handle case where cluster has no data, but subcluster does specify data
    recordSetObject.MoveFirst
    Dim orphanClusterList As Dictionary
    Set orphanClusterList = GetOrphanSubClusterInfo(recordSetObject, context.fields)
    subclusterCnt = 0

    For Each subclusterKey In orphanClusterList.Keys()
        Set subclusterRecord = orphanClusterList.item(subclusterKey)
        recordSetObject.MoveFirst
        subclusterCnt = subclusterCnt + 1
        EmitClusterOpen subclusterRecord, context.dataLayout, row, context.fields.subclusterPlaceholder, subclusterCnt
        Do While recordSetObject.EOF = False
            If IsNull(recordSetObject.fields(context.fields.Cluster)) And recordSetObject.fields(context.fields.subcluster).value = CStr(subclusterKey) Then
                recordCnt = recordCnt + 1
                EmitRow context, recordSetObject, row, recordCnt
                row = row + 1
            End If
            recordSetObject.MoveNext
        Loop
        EmitClusterClose subclusterRecord, context.dataLayout, row, context.fields.subclusterPlaceholder, subclusterCnt
    Next

    ' Handle case where query specified cluster and subcluster, but the data row is null for these columns
    recordSetObject.MoveFirst
    Do While recordSetObject.EOF = False
        If IsNull(recordSetObject.fields(context.fields.Cluster)) And IsNull(recordSetObject.fields(context.fields.subcluster)) Then
            recordCnt = recordCnt + 1
            EmitRow context, recordSetObject, row, recordCnt
            row = row + 1
        End If
        recordSetObject.MoveNext
    Loop

End Sub

Private Sub ProcessClusterYesSubclusterNo(ByRef context As sqlContext, _
                                          ByVal recordSetObject As Object, _
                                          ByRef row As Long, _
                                          ByRef recordCnt As Long)
    Dim clusterList As Dictionary
    Set clusterList = GetClusterInfo(recordSetObject, context.fields)

    Dim clusterCnt As Long
    clusterCnt = 0

    ' Emit the clusters
    Dim clusterKey As Variant
    Dim clusterRecord As Cluster
    For Each clusterKey In clusterList.Keys()
        clusterCnt = clusterCnt + 1
        Set clusterRecord = clusterList.item(CStr(clusterKey))

        EmitClusterOpen clusterRecord, context.dataLayout, row, context.fields.clusterPlaceholder, clusterCnt
        recordSetObject.MoveFirst
        Do While recordSetObject.EOF = False
            If recordSetObject.fields(context.fields.Cluster).value = CStr(clusterKey) Then
                recordCnt = recordCnt + 1
                EmitRow context, recordSetObject, row, recordCnt
                row = row + 1
            End If
            recordSetObject.MoveNext
        Loop
        EmitClusterClose clusterRecord, context.dataLayout, row, context.fields.clusterPlaceholder, clusterCnt
    Next
    
    ' Emit the records which are not in a cluster
    recordSetObject.MoveFirst
    Do While recordSetObject.EOF = False
        If IsNull(recordSetObject.fields(context.fields.Cluster).value) Then
            recordCnt = recordCnt + 1
            EmitRow context, recordSetObject, row, recordCnt
            row = row + 1
        End If
        recordSetObject.MoveNext
    Loop

End Sub

Private Sub ProcessClusterNoSubclusterYes(ByRef context As sqlContext, _
                                          ByVal recordSetObject As Object, _
                                          ByRef row As Long, _
                                          ByRef recordCnt As Long)
    Dim subclusterList As Dictionary
    Set subclusterList = GetSubclusterInfo(recordSetObject, context.fields)

    Dim subclusterCnt As Long
    subclusterCnt = 0

    ' Emit the subclusters
    Dim subclusterKey As Variant
    Dim subclusterRecord As Cluster
    For Each subclusterKey In subclusterList.Keys()
        subclusterCnt = subclusterCnt + 1
        Set subclusterRecord = subclusterList.item(CStr(subclusterKey))

        EmitClusterOpen subclusterRecord, context.dataLayout, row, context.fields.subclusterPlaceholder, subclusterCnt
        recordSetObject.MoveFirst
        Do While recordSetObject.EOF = False
            If recordSetObject.fields(context.fields.subcluster).value = CStr(subclusterKey) Then
                recordCnt = recordCnt + 1
                EmitRow context, recordSetObject, row, recordCnt
                row = row + 1
            End If
            recordSetObject.MoveNext
        Loop
        EmitClusterClose subclusterRecord, context.dataLayout, row, context.fields.subclusterPlaceholder, subclusterCnt
    Next

    ' Handle case where query specified subcluster, but the subcluster column data is null
    recordSetObject.MoveFirst
    Do While recordSetObject.EOF = False
        If IsNull(recordSetObject.fields(context.fields.subcluster)) Then
            recordCnt = recordCnt + 1
            EmitRow context, recordSetObject, row, recordCnt
            row = row + 1
        End If
        recordSetObject.MoveNext
    Loop
End Sub

Private Sub ProcessClusterNoSubclusterNo(ByRef context As sqlContext, _
                                         ByVal recordSetObject As Object, _
                                         ByRef row As Long, _
                                         ByRef recordCnt As Long)
    recordSetObject.MoveFirst
    Do While recordSetObject.EOF = False
        recordCnt = recordCnt + 1
        EmitRow context, recordSetObject, row, recordCnt
        row = row + 1
        recordSetObject.MoveNext
    Loop
End Sub

Private Sub CreateEdges(ByRef context As sqlContext, _
                         ByVal recordSetObject As Object, _
                         ByRef row As Long, _
                         ByRef recordCnt As Long)
                         
    If recordSetObject.EOF Then Exit Sub

    Dim item As String
    Dim relatedItem As String
    
    recordSetObject.MoveFirst
    item = CStr(recordSetObject.fields(context.headings.item).value)
    
    recordSetObject.MoveNext
    Do While recordSetObject.EOF = False
        relatedItem = CStr(recordSetObject.fields(context.headings.item).value)
        
        ' Emit the row
        recordCnt = recordCnt + 1
        EmitRow context, recordSetObject, row, recordCnt
        
        ' Override the Item and Related Item cells with the previous and current item IDs
        DataSheet.Cells.item(row, context.dataLayout.itemColumn) = item
        DataSheet.Cells.item(row, context.dataLayout.isRelatedToItemColumn) = relatedItem
        
        ' Advance to the next result
        item = relatedItem
        row = row + 1
        recordSetObject.MoveNext
    Loop
End Sub

Private Sub CreateRank(ByRef context As sqlContext, _
                         ByVal recordSetObject As Object, _
                         ByRef row As Long, _
                         ByRef recordCnt As Long)
                         
    If recordSetObject.EOF Then Exit Sub

    ' Establish the rank
    recordSetObject.MoveFirst
    Dim rank As String
    rank = LCase$(CStr(recordSetObject.fields("RANK")))
    
    ' Collect the node identifiers
    Dim item As String
    Dim subgraph As String
    subgraph = "{ rank=" & AddQuotes(rank) & ";"
    Do While recordSetObject.EOF = False
        item = CStr(recordSetObject.fields(context.headings.item).value)
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

Private Function GetClusterInfo(ByVal recordSetObject As Object, ByRef fields As sqlFieldName) As Dictionary
    Dim clusters As Dictionary
    Set clusters = New Dictionary

    Dim fieldObject As Variant

    Dim clusterLabel As String
    Dim clusterTooltip As String
    Dim clusterStyleName As String
    Dim clusterAttributes As String
    
    If Not recordSetObject.EOF Then
        recordSetObject.MoveFirst
        Do While recordSetObject.EOF = False
            clusterLabel = vbNullString
            clusterStyleName = vbNullString
            clusterAttributes = vbNullString
            clusterTooltip = vbNullString
            
            For Each fieldObject In recordSetObject.fields
                If LCase$(fieldObject.name) = fields.Cluster Then
                    If Not HasField(recordSetObject, fields.clusterLabel) Then
                        If Not IsNull(fieldObject.value) Then
                            clusterLabel = CStr(fieldObject.value)
                        End If
                    End If
                ElseIf LCase$(fieldObject.name) = fields.clusterLabel Then
                    If Not IsNull(fieldObject.value) Then
                        clusterLabel = CStr(fieldObject.value)
                    End If
                ElseIf LCase$(fieldObject.name) = fields.clusterStyleName Then
                    If Not IsNull(fieldObject.value) Then
                         clusterStyleName = CStr(fieldObject.value)
                   End If
                ElseIf LCase$(fieldObject.name) = fields.clusterAttributes Then
                    If Not IsNull(fieldObject.value) Then
                        clusterAttributes = CStr(fieldObject.value)
                    End If
                ElseIf LCase$(fieldObject.name) = fields.clusterTooltip Then
                    If Not IsNull(fieldObject.value) Then
                        clusterTooltip = CStr(fieldObject.value)
                    End If
                End If
            Next
            
            If clusterLabel <> vbNullString Then
                Dim clusterObject As Cluster
                Set clusterObject = New Cluster
                
                clusterObject.label = clusterLabel
                clusterObject.styleName = clusterStyleName
                clusterObject.attributes = clusterAttributes
                clusterObject.tooltip = clusterTooltip
                
                If Not clusters.Exists(clusterLabel) Then
                    clusters.Add clusterLabel, clusterObject
                End If
            End If
            recordSetObject.MoveNext
        Loop
        recordSetObject.MoveFirst
    End If
    
    Set GetClusterInfo = clusters
End Function

Private Function GetSubclusterInfo(ByVal recordSetObject As Object, ByRef fields As sqlFieldName) As Dictionary
    Dim subclusters As Dictionary
    Set subclusters = New Dictionary

    Dim fieldObject As Variant

    Dim subclusterLabel As String
    Dim subclusterTooltip As String
    Dim subclusterStyleName As String
    Dim subclusterAttributes As String
    
    If Not recordSetObject.EOF Then
        recordSetObject.MoveFirst
        Do While recordSetObject.EOF = False
            subclusterLabel = vbNullString
            subclusterStyleName = vbNullString
            subclusterAttributes = vbNullString
            subclusterTooltip = vbNullString
            
            For Each fieldObject In recordSetObject.fields
                If LCase$(fieldObject.name) = fields.subcluster Then
                    If Not HasField(recordSetObject, fields.subclusterLabel) Then
                        If Not IsNull(fieldObject.value) Then
                            subclusterLabel = CStr(fieldObject.value)
                        End If
                    End If
                ElseIf LCase$(fieldObject.name) = fields.subclusterLabel Then
                    If Not IsNull(fieldObject.value) Then
                        subclusterLabel = CStr(fieldObject.value)
                    End If
                ElseIf LCase$(fieldObject.name) = fields.subclusterStyleName Then
                    If Not IsNull(fieldObject.value) Then
                         subclusterStyleName = CStr(fieldObject.value)
                   End If
                ElseIf LCase$(fieldObject.name) = fields.subclusterAttributes Then
                    If Not IsNull(fieldObject.value) Then
                        subclusterAttributes = CStr(fieldObject.value)
                    End If
                ElseIf LCase$(fieldObject.name) = fields.subclusterTooltip Then
                    If Not IsNull(fieldObject.value) Then
                        subclusterTooltip = CStr(fieldObject.value)
                    End If
                End If
            Next
            
            If subclusterLabel <> vbNullString Then
                Dim clusterObject As Cluster
                Set clusterObject = New Cluster
                
                clusterObject.label = subclusterLabel
                clusterObject.styleName = subclusterStyleName
                clusterObject.attributes = subclusterAttributes
                clusterObject.tooltip = subclusterTooltip
                
                If Not subclusters.Exists(subclusterLabel) Then
                    subclusters.Add subclusterLabel, clusterObject
                End If
            End If
            recordSetObject.MoveNext
        Loop
        recordSetObject.MoveFirst
    End If
    
    Set GetSubclusterInfo = subclusters
End Function

Private Function GetSubClusterInfoForCluster(ByVal recordSetObject As Object, ByRef fields As sqlFieldName, ByVal clusterName As String) As Dictionary
    Dim clusters As Dictionary
    Set clusters = New Dictionary

    Dim fieldObject As Variant

    Dim clusterLabel As String
    Dim clusterTooltip As String
    Dim clusterStyleName As String
    Dim clusterAttributes As String
    
    Dim position As Long
    position = 0
    
    If Not recordSetObject.EOF Then
        recordSetObject.MoveFirst
        Do While recordSetObject.EOF = False
            If recordSetObject.fields(fields.Cluster).value = clusterName Then
                clusterLabel = vbNullString
                clusterStyleName = vbNullString
                clusterAttributes = vbNullString
                clusterTooltip = vbNullString
                position = position + 1
                
                For Each fieldObject In recordSetObject.fields
                    If LCase$(fieldObject.name) = fields.subcluster Then
                        If Not IsNull(fieldObject.value) Then
                            clusterLabel = CStr(fieldObject.value)
                        End If
                    ElseIf LCase$(fieldObject.name) = fields.subclusterStyleName Then
                        If Not IsNull(fieldObject.value) Then
                            clusterStyleName = CStr(fieldObject.value)
                        End If
                    ElseIf LCase$(fieldObject.name) = fields.subclusterAttributes Then
                        If Not IsNull(fieldObject.value) Then
                            clusterAttributes = CStr(fieldObject.value)
                        End If
                    ElseIf LCase$(fieldObject.name) = fields.subclusterTooltip Then
                        If Not IsNull(fieldObject.value) Then
                            clusterTooltip = CStr(fieldObject.value)
                        End If
                    End If
                Next
                
                ' A cluster field was found, add it to the dictionary of cluster info
                If clusterLabel <> vbNullString Then
                    Dim clusterObject As Cluster
                    Set clusterObject = New Cluster
                    
                    clusterObject.label = clusterLabel
                    clusterObject.styleName = clusterStyleName
                    clusterObject.attributes = clusterAttributes
                    clusterObject.tooltip = clusterTooltip
                    
                    If Not clusters.Exists(clusterLabel) Then
                        clusters.Add clusterLabel, clusterObject
                    End If
                End If
            End If
            recordSetObject.MoveNext
        Loop
        recordSetObject.MoveFirst
    End If
    
    Set GetSubClusterInfoForCluster = clusters
End Function

Private Function GetOrphanSubClusterInfo(ByVal recordSetObject As Object, ByRef fields As sqlFieldName) As Dictionary
    ' Build a list of subclusters where the cluster column was null
    Dim clusters As Dictionary
    Set clusters = New Dictionary

    Dim fieldObject As Variant

    Dim subclusterLabel As String
    Dim subclusterTooltip As String
    Dim subclusterStyleName As String
    Dim subclusterAttributes As String
    
    If Not recordSetObject.EOF Then
        recordSetObject.MoveFirst
        Do While recordSetObject.EOF = False
            If IsNull(recordSetObject.fields(fields.Cluster)) And Not IsNull(recordSetObject.fields(fields.subcluster)) Then
                subclusterLabel = vbNullString
                subclusterStyleName = vbNullString
                subclusterAttributes = vbNullString
                subclusterTooltip = vbNullString
                
                For Each fieldObject In recordSetObject.fields
                    If LCase$(fieldObject.name) = fields.subcluster Then
                        If Not IsNull(fieldObject.value) Then
                            subclusterLabel = CStr(fieldObject.value)
                        End If
                    ElseIf LCase$(fieldObject.name) = fields.subclusterStyleName Then
                        If Not IsNull(fieldObject.value) Then
                            subclusterStyleName = CStr(fieldObject.value)
                        End If
                    ElseIf LCase$(fieldObject.name) = fields.subclusterAttributes Then
                        If Not IsNull(fieldObject.value) Then
                            subclusterAttributes = CStr(fieldObject.value)
                        End If
                    ElseIf LCase$(fieldObject.name) = fields.subclusterTooltip Then
                        If Not IsNull(fieldObject.value) Then
                            subclusterTooltip = CStr(fieldObject.value)
                        End If
                    End If
                Next
                
                Dim clusterObject As Cluster
                Set clusterObject = New Cluster
                
                clusterObject.label = subclusterLabel
                clusterObject.styleName = subclusterStyleName
                clusterObject.attributes = subclusterAttributes
                clusterObject.tooltip = subclusterTooltip
                
                If Not clusters.Exists(subclusterLabel) Then
                    clusters.Add subclusterLabel, clusterObject
                End If
            End If
            recordSetObject.MoveNext
        Loop
    End If
    
    Set GetOrphanSubClusterInfo = clusters
End Function

Private Sub EmitClusterOpen(ByVal clusterRecord As Cluster, ByRef dataLayout As dataWorksheet, ByRef row As Long, ByVal findStr As String, ByRef replaceLong As Long)
    DataSheet.Cells.item(row, dataLayout.itemColumn) = OPEN_BRACE
    DataSheet.Cells.item(row, dataLayout.labelColumn) = clusterRecord.label
    DataSheet.Cells.item(row, dataLayout.extraAttributesColumn) = replace(clusterRecord.attributes, findStr, CStr(replaceLong), 1, -1, vbTextCompare)
    DataSheet.Cells.item(row, dataLayout.tooltipColumn) = clusterRecord.tooltip
    
    If clusterRecord.styleName <> vbNullString Then
        ' Append the suffix to the style name
        DataSheet.Cells.item(row, dataLayout.styleNameColumn) = replace(clusterRecord.styleName, findStr, CStr(replaceLong), 1, -1, vbTextCompare) & SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_OPEN).value
    End If
    row = row + 1
End Sub

Private Sub EmitClusterClose(ByVal clusterRecord As Cluster, ByRef dataLayout As dataWorksheet, ByRef row As Long, ByVal findStr As String, ByRef replaceLong As Long)
    DataSheet.Cells.item(row, dataLayout.itemColumn) = CLOSE_BRACE
    
    If clusterRecord.styleName <> vbNullString Then
        ' Append the suffix to the style name
        DataSheet.Cells.item(row, dataLayout.styleNameColumn) = replace(clusterRecord.styleName, findStr, CStr(replaceLong), 1, -1, vbTextCompare) & SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_CLOSE).value
    End If
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

Private Sub EmitRow(ByRef context As sqlContext, ByVal recordSetObject As Object, ByVal row As Long, ByVal position As Long)
    
    Dim fieldObject As Variant      ' Field object within a Record Set record Fields collection
    Dim fieldObjectValue As String
    Dim splitLength As Long
    Dim lineEnding As String
        
    ' Transfer the results to the specified location. Destination worksheet,
    ' row, and column are passed in as parameters
    For Each fieldObject In recordSetObject.fields
        If IsNull(fieldObject.value) Then
            fieldObjectValue = vbNullString
        Else
            fieldObjectValue = replace(CStr(fieldObject.value), context.fields.recordsetPlaceholder, CStr(position), 1, -1, vbTextCompare)
        End If
        
        Select Case LCase$(fieldObject.name)
        Case context.headings.flag
            DataSheet.Cells.item(row, context.dataLayout.flagColumn) = CStr(fieldObjectValue)
            
        Case context.headings.item
            DataSheet.Cells.item(row, context.dataLayout.itemColumn) = CStr(fieldObjectValue)
            
        Case context.headings.label
            If HasField(recordSetObject, context.fields.splitLength) Then
                splitLength = CLng(recordSetObject.fields(context.fields.splitLength).value)
                If splitLength < 0 Then
                    splitLength = 0
                End If
                
                If HasField(recordSetObject, context.fields.lineEnding) Then
                    lineEnding = CStr(recordSetObject.fields(context.fields.lineEnding).value)
                Else
                    lineEnding = NEWLINE
                End If
                
                DataSheet.Cells.item(row, context.dataLayout.labelColumn) = SplitMultilineText(CStr(fieldObjectValue), splitLength, lineEnding)
            Else
                DataSheet.Cells.item(row, context.dataLayout.labelColumn) = CStr(fieldObjectValue)
            End If
            
        Case context.headings.xLabel
            If HasField(recordSetObject, context.fields.splitLength) Then
                splitLength = CLng(recordSetObject.fields(context.fields.splitLength).value)
                If splitLength < 0 Then
                    splitLength = 0
                End If
                
                If HasField(recordSetObject, context.fields.lineEnding) Then
                    lineEnding = CStr(recordSetObject.fields(context.fields.lineEnding).value)
                Else
                    lineEnding = NEWLINE
                End If
                
                DataSheet.Cells.item(row, context.dataLayout.xLabelColumn) = SplitMultilineText(CStr(fieldObjectValue), splitLength, lineEnding)
            Else
                DataSheet.Cells.item(row, context.dataLayout.xLabelColumn) = CStr(fieldObjectValue)
            End If
            
        Case context.headings.tailLabel
            DataSheet.Cells.item(row, context.dataLayout.tailLabelColumn) = CStr(fieldObjectValue)
            
        Case context.headings.headLabel
            DataSheet.Cells.item(row, context.dataLayout.headLabelColumn) = CStr(fieldObjectValue)
            
        Case context.headings.tooltip
            DataSheet.Cells.item(row, context.dataLayout.tooltipColumn) = CStr(fieldObjectValue)
            
        Case context.headings.isRelatedToItem
            DataSheet.Cells.item(row, context.dataLayout.isRelatedToItemColumn) = CStr(fieldObjectValue)
            
        Case context.headings.styleName
            DataSheet.Cells.item(row, context.dataLayout.styleNameColumn) = CStr(fieldObjectValue)
            
        Case context.headings.extraAttributes
            DataSheet.Cells.item(row, context.dataLayout.extraAttributesColumn) = CStr(fieldObjectValue)
        
        Case context.headings.errorMessage
            DataSheet.Cells.item(row, context.dataLayout.errorMessageColumn) = CStr(fieldObjectValue)
        End Select
    Next
End Sub

Private Function GetSQLWorksheetHeadings(ByRef dataLayout As dataWorksheet) As DataWorksheetHeadings
    GetSQLWorksheetHeadings.flag = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.flagColumn).value))
    GetSQLWorksheetHeadings.item = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.itemColumn).value))
    GetSQLWorksheetHeadings.label = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.labelColumn).value))
    GetSQLWorksheetHeadings.xLabel = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.xLabelColumn).value))
    GetSQLWorksheetHeadings.tailLabel = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.tailLabelColumn).value))
    GetSQLWorksheetHeadings.headLabel = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.headLabelColumn).value))
    GetSQLWorksheetHeadings.tooltip = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.tooltipColumn).value))
    GetSQLWorksheetHeadings.isRelatedToItem = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.isRelatedToItemColumn).value))
    GetSQLWorksheetHeadings.styleName = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.styleNameColumn).value))
    GetSQLWorksheetHeadings.extraAttributes = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.extraAttributesColumn).value))
    GetSQLWorksheetHeadings.errorMessage = Trim$(LCase$(DataSheet.Cells.item(dataLayout.headingRow, dataLayout.errorMessageColumn).value))
End Function


