Attribute VB_Name = "modExchangeExport"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Exchange")

Option Explicit

Public Sub ExportData()
    ' Disable screen updates
    OptimizeCode_Begin
    
    ' Prompt user to enter a filename
    Dim exportFile As String
    
#If Mac Then
    exportFile = RunAppleScriptTask("getSaveAsFileName", ".json")
#Else
    exportFile = GetSaveAsFilename("Excel Files (*.json), *json")
#End If
    
    ' Did they specify a filename?
    If exportFile <> vbNullString Then
        Dim exportIt As Boolean
        exportIt = False
        
        If FileExists(exportFile) Then  ' File exists. Confirm that it should be overwritten
            Dim answer As Long
            answer = MsgBox(replace(GetMessage("msgboxFileAlreadyExists"), "{exportFile}", exportFile), vbYesNo + vbQuestion, GetLabel("msgboxFileAlreadyExists"))
            
            If answer = vbYes Then      ' User has confirmed that the file can be overwritten
                exportIt = True
            End If
        Else                            ' No file exists having this name, proceed
            exportIt = True
        End If
        
        ' Export the contents
        If exportIt Then
#If Mac Then
            WriteTextToFile GetAllDataAsJson, exportFile
#Else
            WriteTextToUTF8FileFileWithoutBOM GetAllDataAsJson, exportFile
#End If
            MsgBox GetMessage("msgboxExportComplete") & vbNewLine & exportFile, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        End If
    End If
    
    ' Enable screen updates
    OptimizeCode_End
End Sub

Private Function GetAllDataAsJson() As String
    Dim ini As settings
    ini = GetSettings(DataSheet.name)
    
    Dim exchange As ExchangeOptions
    exchange = GetExchangeOptions()
    
    Dim body As Dictionary
    Set body = New Dictionary
    
    If exchange.includeMetadata Then
        body.Add JSON_SECTION_METADATA, GetExportMetaData
    End If
    
    Dim worksheetDictionary As Dictionary
    Set worksheetDictionary = New Dictionary
    
    Dim includeWorksheets As Boolean
    '@Ignore AssignmentNotUsed
    includeWorksheets = False
    
    ' Export 'data' worksheet
    If GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET) Then
        worksheetDictionary.Add WORKSHEET_DATA, GetDataRows(ini, exchange)
        includeWorksheets = True
    End If
    
    ' Export 'styles' worksheet
    If GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_STYLES_WORKSHEET) Then
        worksheetDictionary.Add WORKSHEET_STYLES, GetStylesRows(ini, exchange)
        includeWorksheets = True
    End If
    
    ' Export 'sql' worksheet
    If GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SQL_WORKSHEET) Then
        worksheetDictionary.Add WORKSHEET_SQL, GetSqlRows(ini, exchange)
        includeWorksheets = True
    End If
    
    ' Export 'svg' worksheet
    If GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SVG_WORKSHEET) Then
        worksheetDictionary.Add WORKSHEET_SVG, GetSvgRows(ini, exchange)
        includeWorksheets = True
    End If
    
    ' Add 'content' structure if any worksheet data was exported
    If includeWorksheets Then
        body.Add JSON_SECTION_CONTENT, worksheetDictionary
    End If
    
    ' Worksheet options
    If GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_GRAPH_OPTIONS) Then
        body.Add WORKSHEET_SETTINGS, GetWorksheetSettings(ini)
    End If
    
    ' Worksheet layouts
    If GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS) Then
        body.Add JSON_SECTION_LAYOUTS, GetWorksheetLayouts(ini)
    End If
    
    GetAllDataAsJson = ConvertToJson(body, Whitespace:=2)
End Function

Private Function GetExportMetaData() As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_METADATA_NAME, JSON_METADATA_NAME_VALUE
    dictionaryObj.Add JSON_METADATA_TYPE, JSON_METADATA_TYPE_VALUE
    dictionaryObj.Add JSON_METADATA_VERSION, JSON_METADATA_VERSION_VERSION
    dictionaryObj.Add JSON_METADATA_USER, Application.username
    dictionaryObj.Add JSON_METADATA_DATE, format(date, "yyyy-mm-dd")
    dictionaryObj.Add JSON_METADATA_TIME, format(time, "hh:mm:ss")
    dictionaryObj.Add JSON_METADATA_OS, Application.OperatingSystem
    dictionaryObj.Add JSON_METADATA_EXCEL, Application.version
    dictionaryObj.Add JSON_METADATA_FILENAME, ThisWorkbook.name
    
    Set GetExportMetaData = dictionaryObj
End Function

Private Function GetDataRows(ByRef ini As settings, ByRef exchange As ExchangeOptions) As Collection
    Dim data As dataRow
    Dim Items As Collection
    Set Items = New Collection
    Dim dictionaryObj As Dictionary

    ' Iterate through the rows of data
    Dim row As Long
    For row = ini.data.firstRow To ini.data.lastRow
        data = GetDataRow(ini, ini.data.worksheetName, row)
        Set dictionaryObj = ConvertDataRowToDictionary(exchange, data, row)
        If dictionaryObj.count > 0 Then
            Items.Add dictionaryObj
        End If
    Next row

    Set GetDataRows = Items
End Function

Private Function GetSqlRows(ByRef ini As settings, ByRef exchange As ExchangeOptions) As Collection
    Dim sql As sqlRow
    Dim Items As Collection
    Set Items = New Collection
    Dim dictionaryObj As Dictionary

    ' Iterate through the rows of data
    Dim row As Long
    For row = ini.sql.firstRow To ini.sql.lastRow
        sql = GetSqlRow(ini, row)
        Set dictionaryObj = ConvertSqlRowToDictionary(exchange, sql, row)
        If dictionaryObj.count > 0 Then
            Items.Add dictionaryObj
        End If
    Next row

    Set GetSqlRows = Items
End Function

Private Function GetSvgRows(ByRef ini As settings, ByRef exchange As ExchangeOptions) As Collection
    Dim svg As svgRow
    Dim Items As Collection
    Set Items = New Collection
    Dim dictionaryObj As Dictionary

    ' Iterate through the rows of data
    Dim row As Long
    For row = ini.svg.firstRow To ini.svg.lastRow
        svg = GetSvgRow(ini, row)
        Set dictionaryObj = ConvertSvgRowToDictionary(exchange, svg, row)
        If dictionaryObj.count > 0 Then
            Items.Add dictionaryObj
        End If
    Next row

    Set GetSvgRows = Items
End Function

Private Function GetStylesRows(ByRef ini As settings, ByRef exchange As ExchangeOptions) As Collection
    Dim switches() As String
    Dim style As StylesRow
    Dim Items As Collection
    Set Items = New Collection
    Dim dictionaryObj As Dictionary

    ' Iterate through the rows of styles
    Dim row As Long
    For row = ini.styles.firstRow To ini.styles.lastRow
        style = GetStylesRow(ini, row)
        switches = GetStylesRowViews(ini, row)
        Set dictionaryObj = ConvertStylesRowToDictionary(exchange, style, switches, row)
        If dictionaryObj.count > 0 Then
            Items.Add dictionaryObj
        End If
    Next row

    Set GetStylesRows = Items
End Function

Private Function GetWorksheetLayouts(ByRef ini As settings) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add WORKSHEET_DATA, GetLayoutData(ini)
    dictionaryObj.Add WORKSHEET_STYLES, GetLayoutStyles(ini)
    dictionaryObj.Add WORKSHEET_SQL, GetLayoutSql(ini)
    dictionaryObj.Add WORKSHEET_SVG, GetLayoutSvg(ini)
    dictionaryObj.Add WORKSHEET_SOURCE, GetLayoutSource(ini)

    Set GetWorksheetLayouts = dictionaryObj
End Function

Private Function GetWorksheetSettings(ByRef ini As settings) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add WORKSHEET_DATA, GetSettingsDictionaryData(ini)
    dictionaryObj.Add WORKSHEET_GRAPH, GetSettingsDictionaryGraph()
    dictionaryObj.Add WORKSHEET_SETTINGS, GetSettingsDictionarySettings(ini)
    dictionaryObj.Add WORKSHEET_SOURCE, GetSettingsDictionarySource(ini)
    dictionaryObj.Add WORKSHEET_SQL, GetSettingsDictionarySql()
    dictionaryObj.Add "extensions", GetSettingsDictionaryExtensions()
    
    Set GetWorksheetSettings = dictionaryObj
End Function

Private Function GetLayoutData(ByRef ini As settings) As Dictionary
    Dim rowItems As Collection
    Set rowItems = New Collection
    
    rowItems.Add GetLayoutRowData(DataSheet.name, JSON_HEADING, ini.data.headingRow)
    rowItems.Add GetLayoutRowData(DataSheet.name, JSON_FIRST, ini.data.firstRow)
    
    Dim columnItems As Collection
    Set columnItems = New Collection
    
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_FLAG, ini.data.headingRow, ini.data.flagColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_ITEM, ini.data.headingRow, ini.data.itemColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_LABEL, ini.data.headingRow, ini.data.labelColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_OUTSIDE_LABEL, ini.data.headingRow, ini.data.xLabelColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_TAIL_LABEL, ini.data.headingRow, ini.data.tailLabelColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_HEAD_LABEL, ini.data.headingRow, ini.data.headLabelColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_TOOLTIP, ini.data.headingRow, ini.data.tooltipColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_RELATED_ITEM, ini.data.headingRow, ini.data.isRelatedToItemColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_STYLE_NAME, ini.data.headingRow, ini.data.styleNameColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_EXTRA_ATTRIBUTES, ini.data.headingRow, ini.data.extraAttributesColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_MESSAGE, ini.data.headingRow, ini.data.errorMessageColumn)
    columnItems.Add GetLayoutColumnData(DataSheet.name, JSON_DATA_GRAPH_DISPLAY_COLUMN, ini.data.headingRow, ini.data.graphDisplayColumn)
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ROWS, rowItems
    dictionaryObj.Add JSON_COLUMNS, columnItems
    
    Set GetLayoutData = dictionaryObj
End Function

Private Function GetLayoutStyles(ByRef ini As settings) As Dictionary
    Dim rowItems As Collection
    Set rowItems = New Collection
    
    rowItems.Add GetLayoutRowData(StylesSheet.name, JSON_HEADING, ini.styles.headingRow)
    rowItems.Add GetLayoutRowData(StylesSheet.name, JSON_FIRST, ini.styles.firstRow)

    Dim columnItems As Collection
    Set columnItems = New Collection
    
    columnItems.Add GetLayoutColumnData(StylesSheet.name, JSON_STYLES_FLAG, ini.styles.headingRow, ini.styles.flagColumn)
    columnItems.Add GetLayoutColumnData(StylesSheet.name, JSON_STYLES_NAME, ini.styles.headingRow, ini.styles.nameColumn)
    columnItems.Add GetLayoutColumnData(StylesSheet.name, JSON_STYLES_FORMAT, ini.styles.headingRow, ini.styles.formatColumn)
    columnItems.Add GetLayoutColumnData(StylesSheet.name, JSON_STYLES_TYPE, ini.styles.headingRow, ini.styles.typeColumn)
    
    Dim i As Long
    Dim viewCount As Long
    Dim lastColumn As Long
    lastColumn = GetLastColumn(StylesSheet.name, ini.styles.headingRow)
    
    viewCount = 0
    For i = ini.styles.firstYesNoColumn To lastColumn
        columnItems.Add GetLayoutColumnData(StylesSheet.name, JSON_STYLES_VIEW & viewCount, ini.styles.headingRow, i)
        viewCount = viewCount + 1
    Next i
    
    Dim namedRange As Dictionary
    Set namedRange = New Dictionary
    
    Dim ranges As Collection
    Set ranges = New Collection
    
    ranges.Add namedRange
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ROWS, rowItems
    dictionaryObj.Add JSON_COLUMNS, columnItems
    dictionaryObj.Add JSON_RANGES, ranges
    
    Set GetLayoutStyles = dictionaryObj
End Function

Private Function GetLayoutSource(ByRef ini As settings) As Dictionary
    Dim rowItems As Collection
    Set rowItems = New Collection
    
    rowItems.Add GetLayoutRowData(SourceSheet.name, JSON_HEADING, ini.source.headingRow)
    rowItems.Add GetLayoutRowData(SourceSheet.name, JSON_FIRST, ini.source.firstRow)
    
    Dim columnItems As Collection
    Set columnItems = New Collection
    
    columnItems.Add GetLayoutColumnData(SourceSheet.name, JSON_SOURCE_LINE_NUMBER, ini.source.headingRow, ini.source.lineNumberColumn)
    columnItems.Add GetLayoutColumnData(SourceSheet.name, JSON_SOURCE_SOURCE, ini.source.headingRow, ini.source.sourceColumn)
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ROWS, rowItems
    dictionaryObj.Add JSON_COLUMNS, columnItems
    
    Set GetLayoutSource = dictionaryObj
End Function

Private Function GetLayoutSql(ByRef ini As settings) As Dictionary
    Dim rowItems As Collection
    Set rowItems = New Collection
    
    rowItems.Add GetLayoutRowData(SqlSheet.name, JSON_HEADING, ini.sql.headingRow)
    rowItems.Add GetLayoutRowData(SqlSheet.name, JSON_FIRST, ini.sql.firstRow)
    
    Dim columnItems As Collection
    Set columnItems = New Collection
    
    columnItems.Add GetLayoutColumnData(SqlSheet.name, JSON_LAYOUT_SQL_FLAG, ini.sql.headingRow, ini.sql.flagColumn)
    columnItems.Add GetLayoutColumnData(SqlSheet.name, JSON_LAYOUT_SQL_SQL_STATEMENT, ini.sql.headingRow, ini.sql.sqlStatementColumn)
    columnItems.Add GetLayoutColumnData(SqlSheet.name, JSON_LAYOUT_SQL_EXCEL_FILE, ini.sql.headingRow, ini.sql.excelFileColumn)
    columnItems.Add GetLayoutColumnData(SqlSheet.name, JSON_LAYOUT_SQL_STATUS, ini.sql.headingRow, ini.sql.statusColumn)
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ROWS, rowItems
    dictionaryObj.Add JSON_COLUMNS, columnItems
    
    Set GetLayoutSql = dictionaryObj
End Function

Private Function GetLayoutSvg(ByRef ini As settings) As Dictionary
    Dim rowItems As Collection
    Set rowItems = New Collection
    
    rowItems.Add GetLayoutRowData(SvgSheet.name, JSON_HEADING, ini.svg.headingRow)
    rowItems.Add GetLayoutRowData(SvgSheet.name, JSON_FIRST, ini.svg.firstRow)
    
    Dim columnItems As Collection
    Set columnItems = New Collection
    
    columnItems.Add GetLayoutColumnData(SvgSheet.name, JSON_LAYOUT_SVG_FLAG, ini.svg.headingRow, ini.svg.flagColumn)
    columnItems.Add GetLayoutColumnData(SvgSheet.name, JSON_LAYOUT_SVG_FIND, ini.svg.headingRow, ini.svg.findColumn)
    columnItems.Add GetLayoutColumnData(SvgSheet.name, JSON_LAYOUT_SVG_REPLACE, ini.svg.headingRow, ini.svg.replaceColumn)
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ROWS, rowItems
    dictionaryObj.Add JSON_COLUMNS, columnItems
    
    Set GetLayoutSvg = dictionaryObj
End Function

Private Function GetSettingsDictionaryData(ByRef ini As settings) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    ' Graph to Worksheet
    Dim graphToWorksheet As Dictionary
    Set graphToWorksheet = New Dictionary
    
    graphToWorksheet.Add JSON_SETTINGS_RUN_MODE, SettingsSheet.Range(SETTINGS_RUN_MODE).value
    graphToWorksheet.Add JSON_SETTINGS_IMAGE_TYPE, ini.graph.imageTypeWorksheet
    graphToWorksheet.Add JSON_SETTINGS_IMAGE_WORKSHEET, ini.graph.imageWorksheet
    graphToWorksheet.Add JSON_SETTINGS_SCALE_IMAGE, ini.graph.scaleImage
    
    ' Graph to File
    Dim graphToFile As Dictionary
    Set graphToFile = New Dictionary
    
    graphToFile.Add JSON_SETTINGS_DIRECTORY, ini.output.directory
    graphToFile.Add JSON_SETTINGS_FILE_NAME_PREFIX, ini.output.fileNamePrefix
    graphToFile.Add JSON_SETTINGS_IMAGE_TYPE, ini.graph.imageTypeFile
    graphToFile.Add JSON_SETTINGS_APPEND_OPTIONS, ini.output.appendOptions
    graphToFile.Add JSON_SETTINGS_APPEND_TIME_STAMP, ini.output.appendTimeStamp

    ' Layout
    Dim layout As Dictionary
    Set layout = New Dictionary
    
    layout.Add JSON_SETTINGS_ENGINE, SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE)
    
    ' Maintain backward compatibility. Direction and rankdir were consolidated into rankdir to improve performance
    ' when the port to Apple Mac was performed. Old versions of the spreadsheet will still expect direction to be
    ' present, so the value will be derived from rankdir instead of coming from a cell as was done previously.
        Select Case UCase$(SettingsSheet.Range(SETTINGS_RANKDIR).value)
        Case "TB"
            layout.Add JSON_SETTINGS_DIRECTION, "top to bottom"
        Case "BT"
            layout.Add JSON_SETTINGS_DIRECTION, "bottom to top"
        Case "LR"
            layout.Add JSON_SETTINGS_DIRECTION, "left to right"
        Case "RL"
            layout.Add JSON_SETTINGS_DIRECTION, "right to left"
        Case Else
             layout.Add JSON_SETTINGS_DIRECTION, vbNullString
    End Select

    layout.Add JSON_SETTINGS_RANKDIR, SettingsSheet.Range(SETTINGS_RANKDIR).value
    layout.Add JSON_SETTINGS_SPLINES, ini.graph.splines
    
    ' Options
    Dim options As Dictionary
    Set options = New Dictionary

    ' Options -> Graph
    Dim optionsGraph As Dictionary
    Set optionsGraph = New Dictionary
    
    optionsGraph.Add JSON_SETTINGS_CENTER, ini.graph.center
    optionsGraph.Add JSON_SETTINGS_CLUSTER_RANK, ini.graph.clusterrank
    optionsGraph.Add JSON_SETTINGS_COMPOUND, ini.graph.compound
    optionsGraph.Add JSON_SETTINGS_DIM, ini.graph.layoutDim
    optionsGraph.Add JSON_SETTINGS_DIMEN, ini.graph.layoutDimen
    optionsGraph.Add JSON_SETTINGS_FORCE_LABELS, ini.graph.forceLabels
    optionsGraph.Add JSON_SETTINGS_GRAPH_TYPE, SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value
    optionsGraph.Add JSON_SETTINGS_MODE, ini.graph.mode
    optionsGraph.Add JSON_SETTINGS_MODEL, ini.graph.model
    optionsGraph.Add JSON_SETTINGS_NEWRANK, ini.graph.newrank
    optionsGraph.Add JSON_SETTINGS_ORDERING, ini.graph.ordering
    optionsGraph.Add JSON_SETTINGS_ORIENTATION, ini.graph.orientation
    optionsGraph.Add JSON_SETTINGS_OUTPUT_ORDER, ini.graph.outputOrder
    optionsGraph.Add JSON_SETTINGS_OVERLAP, ini.graph.overlap
    optionsGraph.Add JSON_SETTINGS_SMOOTHING, ini.graph.smoothing
    optionsGraph.Add JSON_SETTINGS_TRANSPARENT_BACKGROUND, ini.graph.transparentBackground
    optionsGraph.Add JSON_SETTINGS_INCLUDE_IMAGE_PATH, ini.graph.includeGraphImagePath

    ' Options -> Nodes
    Dim optionsNodes As Dictionary
    Set optionsNodes = New Dictionary
    
    optionsNodes.Add JSON_SETTINGS_INCLUDE_ORPHAN_NODES, ini.graph.includeOrphanNodes
    optionsNodes.Add JSON_SETTINGS_INCLUDE_NODE_LABELS, ini.graph.includeNodeLabels
    optionsNodes.Add JSON_SETTINGS_INCLUDE_NODE_XLABELS, ini.graph.includeNodeXLabels
    optionsNodes.Add JSON_SETTINGS_BLANK_NODE_LABELS, SettingsSheet.Range(SETTINGS_BLANK_NODE_LABELS).value

    ' Options -> Edges
    Dim optionsEdges As Dictionary
    Set optionsEdges = New Dictionary
    
    optionsEdges.Add JSON_SETTINGS_ADD_STRICT, ini.graph.addStrict
    optionsEdges.Add JSON_SETTINGS_CONCENTRATE, ini.graph.concentrate
    optionsEdges.Add JSON_SETTINGS_INCLUDE_ORPHAN_EDGES, ini.graph.includeOrphanEdges
    optionsEdges.Add JSON_SETTINGS_INCLUDE_EDGE_HEAD_LABELS, ini.graph.includeEdgeHeadLabels
    optionsEdges.Add JSON_SETTINGS_INCLUDE_EDGE_LABELS, ini.graph.includeEdgeLabels
    optionsEdges.Add JSON_SETTINGS_INCLUDE_EDGE_XLABELS, ini.graph.includeEdgeXLabels
    optionsEdges.Add JSON_SETTINGS_INCLUDE_EDGE_TAIL_LABELS, ini.graph.includeEdgeTailLabels
    optionsEdges.Add JSON_SETTINGS_INCLUDE_EDGE_PORTS, ini.graph.includeEdgePorts
    optionsEdges.Add JSON_SETTINGS_BLANK_EDGE_LABELS, SettingsSheet.Range(SETTINGS_BLANK_EDGE_LABELS).value

    ' Collect graph, nodes, and edge under the options parent
    options.Add JSON_SETTINGS_SECTION_GRAPH, optionsGraph
    options.Add JSON_SETTINGS_SECTION_NODES, optionsNodes
    options.Add JSON_SETTINGS_SECTION_EDGES, optionsEdges
    
    ' Style
    Dim styles As Dictionary
    Set styles = New Dictionary
    styles.Add JSON_SETTINGS_SELECTED_VIEW_COLUMN, SettingsSheet.Range(SETTINGS_STYLES_COL_SHOW_STYLE).value
    styles.Add JSON_SETTINGS_INCLUDE_STYLE_FORMAT, ini.graph.includeStyleFormat
    styles.Add JSON_SETTINGS_INCLUDE_EXTRA_ATTRIBUTES, ini.graph.includeExtraAttributes
    styles.Add JSON_SETTINGS_STYLES_SUFFIX_OPEN, ini.styles.suffixOpen
    styles.Add JSON_SETTINGS_STYLES_SUFFIX_CLOSE, ini.styles.suffixClose

    ' Debug
    Dim debugOptions As Dictionary
    Set debugOptions = New Dictionary
    
    debugOptions.Add JSON_SETTINGS_DEBUG_SWITCH, SettingsSheet.Range(SETTINGS_DEBUG).value
    debugOptions.Add JSON_SETTINGS_FILE_DISPOSITION, SettingsSheet.Range(SETTINGS_FILE_DISPOSITION).value
    
    ' Console
    Dim consoleOptions As Dictionary
    Set consoleOptions = New Dictionary
    
    consoleOptions.Add JSON_SETTINGS_LOG_TO_CONSOLE, ini.console.logToConsole
    consoleOptions.Add JSON_SETTINGS_APPEND_CONSOLE, ini.console.appendConsole
    consoleOptions.Add JSON_SETTINGS_GRAPHVIZ_VERBOSE, ini.console.graphvizVerbose
    
    ' Show/Hide Columns
    Dim columns As Dictionary
    Set columns = New Dictionary
    
    columns.Add JSON_STYLES_FLAG, SettingsSheet.Range(SETTINGS_DATA_SHOW_COMMENT).value
    columns.Add JSON_DATA_ITEM, SettingsSheet.Range(SETTINGS_DATA_SHOW_ITEM).value
    columns.Add JSON_DATA_LABEL, SettingsSheet.Range(SETTINGS_DATA_SHOW_LABEL).value
    columns.Add JSON_DATA_OUTSIDE_LABEL, SettingsSheet.Range(SETTINGS_DATA_SHOW_OUTSIDE_LABEL).value
    columns.Add JSON_DATA_TAIL_LABEL, SettingsSheet.Range(SETTINGS_DATA_SHOW_TAIL_LABEL).value
    columns.Add JSON_DATA_HEAD_LABEL, SettingsSheet.Range(SETTINGS_DATA_SHOW_HEAD_LABEL).value
    columns.Add JSON_DATA_RELATED_ITEM, SettingsSheet.Range(SETTINGS_DATA_SHOW_IS_RELATED_TO_ITEM).value
    columns.Add JSON_DATA_STYLE_NAME, SettingsSheet.Range(SETTINGS_DATA_SHOW_STYLE).value
    columns.Add JSON_DATA_EXTRA_ATTRIBUTES, SettingsSheet.Range(SETTINGS_DATA_SHOW_EXTRA_STYLE_ATTRIBUTES).value
    columns.Add JSON_DATA_MESSAGE, SettingsSheet.Range(SETTINGS_DATA_SHOW_MESSAGES).value

    Dim worksheets As Dictionary
    Set worksheets = New Dictionary
    
    worksheets.Add JSON_WORKSHEETS_ABOUT, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_ABOUT).value
    worksheets.Add JSON_WORKSHEETS_ATTRIBUTES, SettingsSheet.Range(SETTINGS_HELP_ATTRIBUTES).value
    worksheets.Add JSON_WORKSHEETS_COLORS, SettingsSheet.Range(SETTINGS_HELP_COLORS).value
    worksheets.Add JSON_WORKSHEETS_CONSOLE, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_ABOUT).value
    worksheets.Add JSON_WORKSHEETS_DIAGNOSTICS, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS).value
    worksheets.Add JSON_WORKSHEETS_EXCHANGE, SettingsSheet.Range(SETTINGS_TABS_TOGGLE_EXCHANGE).value
    worksheets.Add JSON_WORKSHEETS_LISTS, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LISTS).value
    worksheets.Add JSON_WORKSHEETS_LOCALE_DE_DE, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_DE_DE).value
    worksheets.Add JSON_WORKSHEETS_LOCALE_EN_GB, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_EN_GB).value
    worksheets.Add JSON_WORKSHEETS_LOCALE_EN_US, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_EN_US).value
    worksheets.Add JSON_WORKSHEETS_LOCALE_FR_FR, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_FR_FR).value
    worksheets.Add JSON_WORKSHEETS_LOCALE_IT_IT, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_IT_IT).value
    worksheets.Add JSON_WORKSHEETS_LOCALE_PL_PL, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LOCALE_PL_PL).value
    worksheets.Add JSON_WORKSHEETS_SETTINGS, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).value
    worksheets.Add JSON_WORKSHEETS_SHAPES, SettingsSheet.Range(SETTINGS_HELP_SHAPES).value
    worksheets.Add JSON_WORKSHEETS_SOURCE, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SOURCE).value
    worksheets.Add JSON_WORKSHEETS_SQL, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SQL).value
    worksheets.Add JSON_WORKSHEETS_STYLE_DESIGNER, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).value
    worksheets.Add JSON_WORKSHEETS_STYLES, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLES).value
    worksheets.Add JSON_WORKSHEETS_SVG, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SVG).value
    
    ' Language
    Dim language As Dictionary
    Set language = New Dictionary
    
    language.Add JSON_SETTINGS_LANGUAGE, SettingsSheet.Range(SETTINGS_LANGUAGE).value

    ' Collect the dictionaries
    dictionaryObj.Add JSON_SETTINGS_SECTION_GRAPH_TO_WORKSHEET, graphToWorksheet
    dictionaryObj.Add JSON_SETTINGS_SECTION_GRAPH_TO_FILE, graphToFile
    dictionaryObj.Add JSON_SETTINGS_SECTION_LAYOUT, layout
    dictionaryObj.Add JSON_SETTINGS_SECTION_OPTIONS, options
    dictionaryObj.Add JSON_SETTINGS_SECTION_STYLES, styles
    dictionaryObj.Add JSON_SETTINGS_SECTION_DEBUG, debugOptions
    dictionaryObj.Add JSON_SETTINGS_SECTION_COLUMNS, columns
    dictionaryObj.Add JSON_SETTINGS_SECTION_LANGUAGE, language
    dictionaryObj.Add JSON_SETTINGS_SECTION_CONSOLE, consoleOptions
    dictionaryObj.Add JSON_SETTINGS_SECTION_WORKSHEETS, worksheets

    Set GetSettingsDictionaryData = dictionaryObj
End Function

Private Function GetSettingsDictionaryGraph() As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ZOOM, GetZoom(GraphSheet.name)
   
    Set GetSettingsDictionaryGraph = dictionaryObj
End Function

Private Function GetZoom(ByVal worksheetName As String) As Long
    ' Save the name of the current worksheet
    Dim previousSheet As String
    previousSheet = ActiveSheet.name
    
    ' Switch to the sheet we need to get the zoom value from
    ActiveWorkbook.Sheets.[_Default](worksheetName).Activate
    GetZoom = ActiveWindow.zoom
    
    ' Switch back to the original worksheet
    ActiveWorkbook.Sheets.[_Default](previousSheet).Activate
End Function

Private Function GetSettingsDictionarySettings(ByRef ini As settings) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_SETTINGS_GV_PATH, ini.CommandLine.GraphvizPath
    dictionaryObj.Add JSON_SETTINGS_IMAGE_PATH, ini.graph.imagePath
    dictionaryObj.Add JSON_SETTINGS_GRAPH_OPTIONS, ini.graph.options
    dictionaryObj.Add JSON_SETTINGS_PICTURE_NAME, ini.graph.pictureName
    dictionaryObj.Add JSON_SETTINGS_COMMAND_LINE_PARAMETERS, ini.CommandLine.parameters
    
    Set GetSettingsDictionarySettings = dictionaryObj
End Function

Private Function GetSettingsDictionarySource(ByRef ini As settings) As Dictionary
    Dim buttons As Collection
    Dim button As Dictionary
    Dim i As Long
    
    Set buttons = New Collection
    
    For i = 1 To 6
        Set button = New Dictionary
        button.Add JSON_ID, BUTTON_PREFIX_SOURCE_WEB & i
        button.Add JSON_SETTINGS_BUTTON_TEXT, SettingsSheet.Range(BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_TEXT).value
        button.Add JSON_SETTINGS_URL, SettingsSheet.Range(BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_URL).value
        button.Add JSON_SETTINGS_SCREEN_TIP, SettingsSheet.Range(BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_SCREENTIP).value
        button.Add JSON_SETTINGS_SUPER_TIP, SettingsSheet.Range(BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_SUPERTIP).value
        button.Add JSON_SETTINGS_VISIBLE, SettingsSheet.Range(BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_VISIBLE).value
        buttons.Add button
    Next i
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_SETTINGS_INDENT, ini.source.indent
    dictionaryObj.Add JSON_SETTINGS_BUTTONS, buttons
    
    Set GetSettingsDictionarySource = dictionaryObj
End Function

Private Function GetSettingsDictionarySql() As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    Dim fields As sqlFieldName
    fields = GetSettingsForSqlFields()

    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER, fields.Cluster
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_LABEL, fields.clusterLabel
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME, fields.clusterStyleName
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES, fields.clusterAttributes
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP, fields.clusterTooltip
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_PLACEHOLDER, fields.clusterPlaceholder
    
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER, fields.subcluster
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_LABEL, fields.subclusterLabel
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME, fields.subclusterStyleName
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES, fields.subclusterAttributes
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP, fields.subclusterTooltip
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_PLACEHOLDER, fields.subclusterPlaceholder
    
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_RECORDSET_PLACEHOLDER, fields.recordsetPlaceholder
    
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH, fields.splitLength
    dictionaryObj.Add JSON_SETTINGS_SQL_FIELD_NAME_LINE_ENDING, fields.lineEnding
    
    dictionaryObj.Add JSON_SETTINGS_SQL_FILTER_COLUMN, fields.filterColumn
    dictionaryObj.Add JSON_SETTINGS_SQL_FILTER_VALUE, fields.filterValue
   
    Set GetSettingsDictionarySql = dictionaryObj
End Function

Private Function GetSettingsDictionaryExtensions() As Dictionary
    Dim button As Dictionary
    Dim i As Long
    
    Dim buttonsWeb As Collection
    Set buttonsWeb = New Collection
    
    For i = 1 To 6
        Set button = New Dictionary
        button.Add JSON_ID, BUTTON_PREFIX_EXT_WEB & i
        button.Add JSON_SETTINGS_BUTTON_TEXT, SettingsSheet.Range(BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_TEXT).value
        button.Add JSON_SETTINGS_URL, SettingsSheet.Range(BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_URL).value
        button.Add JSON_SETTINGS_SCREEN_TIP, SettingsSheet.Range(BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_SCREENTIP).value
        button.Add JSON_SETTINGS_SUPER_TIP, SettingsSheet.Range(BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_SUPERTIP).value
        button.Add JSON_SETTINGS_VISIBLE, SettingsSheet.Range(BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_VISIBLE).value
        buttonsWeb.Add button
    Next i
    
    Dim extWeb As Dictionary
    Set extWeb = New Dictionary
    
    extWeb.Add JSON_SETTINGS_GROUP_NAME, SettingsSheet.Range(SETTINGS_EXT_TAB_GROUP_NAME_WEB)
    extWeb.Add JSON_SETTINGS_BUTTONS, buttonsWeb
    
    ' -----------------------------------------------------------
    Dim buttonsCode As Collection
    Set buttonsCode = New Collection
    
    For i = 1 To 6
        Set button = New Dictionary
        button.Add JSON_ID, BUTTON_PREFIX_EXT_CODE & i
        button.Add JSON_SETTINGS_BUTTON_TEXT, SettingsSheet.Range(BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_TEXT).value
        button.Add JSON_SETTINGS_SUB, SettingsSheet.Range(BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SUB).value
        button.Add JSON_SETTINGS_SCREEN_TIP, SettingsSheet.Range(BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SCREENTIP).value
        button.Add JSON_SETTINGS_SUPER_TIP, SettingsSheet.Range(BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SUPERTIP).value
        button.Add JSON_SETTINGS_VISIBLE, SettingsSheet.Range(BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_VISIBLE).value
        buttonsCode.Add button
    Next i
    
    Dim extCode As Dictionary
    Set extCode = New Dictionary
    
    extCode.Add JSON_SETTINGS_GROUP_NAME, SettingsSheet.Range(SETTINGS_EXT_TAB_GROUP_NAME_CODE)
    extCode.Add JSON_SETTINGS_BUTTONS, buttonsCode
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_SETTINGS_EXT_TAB_NAME, SettingsSheet.Range(SETTINGS_EXT_TAB_NAME)
    dictionaryObj.Add JSON_SETTINGS_EXT_TAB_GROUP_NAME_CODE, extCode
    dictionaryObj.Add JSON_SETTINGS_EXT_TAB_GROUP_NAME_WEB, extWeb
    
    Set GetSettingsDictionaryExtensions = dictionaryObj
End Function

Private Function IsRowEnabled(ByVal flag As String) As Boolean
    
    IsRowEnabled = Not (flag = FLAG_COMMENT)
    
End Function

Private Function GetLayoutRowData(ByRef worksheetName As String, ByRef rowId As String, ByRef row As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ID, rowId
    dictionaryObj.Add JSON_ROW, row
    
    If (row > 0) Then
        dictionaryObj.Add JSON_HEIGHT, ActiveWorkbook.Sheets.[_Default](worksheetName).rows(row).height
        dictionaryObj.Add JSON_HIDDEN, ActiveWorkbook.Sheets.[_Default](worksheetName).rows(row).Hidden
    End If
    
    Set GetLayoutRowData = dictionaryObj
End Function

Private Function GetLayoutColumnData(ByRef worksheetName As String, ByRef columnId As String, ByRef row As Long, ByRef col As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    dictionaryObj.Add JSON_ID, columnId
    dictionaryObj.Add JSON_COLUMN, col
    dictionaryObj.Add JSON_HEADING, Trim$(ActiveWorkbook.Sheets.[_Default](worksheetName).Cells(row, col).value)
    dictionaryObj.Add JSON_WIDTH, ActiveWorkbook.Sheets.[_Default](worksheetName).columns(col).ColumnWidth
    dictionaryObj.Add JSON_HIDDEN, ActiveWorkbook.Sheets.[_Default](worksheetName).columns(col).Hidden
    dictionaryObj.Add JSON_WRAP_TEXT, ActiveWorkbook.Sheets.[_Default](worksheetName).columns(col).WrapText

    Set GetLayoutColumnData = dictionaryObj
End Function

Private Function ConvertDataRowToDictionary(ByRef exchange As ExchangeOptions, ByRef data As dataRow, ByRef row As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    If exchange.data.row.number Then
        dictionaryObj.Add JSON_ROW, row
    End If
    
    If exchange.data.row.visible Then
        dictionaryObj.Add JSON_HIDDEN, DataSheet.rows.Item(row).Hidden
    End If
    
    If exchange.data.row.height Then
        dictionaryObj.Add JSON_HEIGHT, DataSheet.rows.Item(row).height
    End If
    
    Dim enabled As Boolean
    enabled = IsRowEnabled(data.comment)
    If Not enabled Then
        dictionaryObj.Add JSON_ENABLED, enabled
    End If
    
    If data.Item <> vbNullString Then
        dictionaryObj.Add JSON_DATA_ITEM, data.Item
    End If
    
    If data.label <> vbNullString Then
        dictionaryObj.Add JSON_DATA_LABEL, data.label
    End If
    
    If data.xLabel <> vbNullString Then
        dictionaryObj.Add JSON_DATA_OUTSIDE_LABEL, data.xLabel
    End If
    
    If data.tailLabel <> vbNullString Then
        dictionaryObj.Add JSON_DATA_TAIL_LABEL, data.tailLabel
    End If
    
    If data.headLabel <> vbNullString Then
        dictionaryObj.Add JSON_DATA_HEAD_LABEL, data.headLabel
    End If
    
    If data.tooltip <> vbNullString Then
        dictionaryObj.Add JSON_DATA_TOOLTIP, data.tooltip
    End If
    
    If data.relatedItem <> vbNullString Then
        dictionaryObj.Add JSON_DATA_RELATED_ITEM, data.relatedItem
    End If
    
    If data.styleName <> vbNullString Then
        dictionaryObj.Add JSON_DATA_STYLE_NAME, data.styleName
    End If
    
    If data.extraAttrs <> vbNullString Then
        dictionaryObj.Add JSON_DATA_EXTRA_ATTRIBUTES, ParseAttributeString(data.extraAttrs)
    End If
    
    Set ConvertDataRowToDictionary = dictionaryObj
End Function

Private Function ConvertStylesRowToDictionary(ByRef exchange As ExchangeOptions, ByRef style As StylesRow, ByRef switches() As String, ByRef row As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    If exchange.styles.row.number Then
        dictionaryObj.Add JSON_ROW, row
    End If
    
    If exchange.styles.row.visible Then
        dictionaryObj.Add JSON_HIDDEN, StylesSheet.rows.Item(row).Hidden
    End If
    
    If exchange.styles.row.height Then
        dictionaryObj.Add JSON_HEIGHT, StylesSheet.rows.Item(row).height
    End If
    
    Dim enabled As Boolean
    enabled = IsRowEnabled(style.comment)
    If Not enabled Then
        dictionaryObj.Add JSON_ENABLED, enabled
    End If
    
    If style.styleName <> vbNullString Then
        dictionaryObj.Add JSON_STYLES_NAME, style.styleName
    End If
    
    If style.styleType <> vbNullString Then
        dictionaryObj.Add JSON_STYLES_TYPE, style.styleType
    End If
    
    If style.format <> vbNullString Then
        dictionaryObj.Add JSON_STYLES_FORMAT, ParseAttributeString(style.format)
    End If
    
    Dim switchCollection As Collection
    Set switchCollection = New Collection
    
    If IsArrayAllocated(switches) Then
        Dim i As Long
        For i = LBound(switches) To UBound(switches) - 1
            switchCollection.Add switches(i)
        Next i

        dictionaryObj.Add JSON_STYLES_VIEW_SWITCHES, switchCollection
    End If
    
    Set ConvertStylesRowToDictionary = dictionaryObj
End Function

Private Function ConvertSqlRowToDictionary(ByRef exchange As ExchangeOptions, ByRef sql As sqlRow, ByRef row As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    If exchange.sql.row.number Then
        dictionaryObj.Add JSON_ROW, row
    End If
    
    If exchange.sql.row.visible Then
        'dictionaryObj.Add JSON_HIDDEN, ActiveWorkbook.Sheets.[_Default](SqlSheet.name).rows(row).Hidden
        dictionaryObj.Add JSON_HIDDEN, SqlSheet.rows.Item(row).Hidden
    End If
    
    If exchange.sql.row.height Then
        dictionaryObj.Add JSON_HEIGHT, ActiveWorkbook.Sheets.[_Default](SqlSheet.name).rows(row).height
    End If
    
    Dim enabled As Boolean
    enabled = IsRowEnabled(sql.comment)
    If Not enabled Then
        dictionaryObj.Add JSON_ENABLED, enabled
    End If
    
    If sql.sqlStatement <> vbNullString Then
        dictionaryObj.Add JSON_SQL_SQL_STATEMENT, sql.sqlStatement
    End If
    
    If sql.excelFile <> vbNullString Then
        dictionaryObj.Add JSON_SQL_EXCEL_FILE, sql.excelFile
    End If
    
    If sql.status <> vbNullString Then
        dictionaryObj.Add JSON_SQL_STATUS, sql.status
    End If

    Dim filterItems As Collection
    Set filterItems = New Collection
    
    Dim col As Long
    For col = 5 To 26
        filterItems.Add SqlSheet.Cells.Item(row, col).value
    Next col

    dictionaryObj.Add JSON_SQL_FILTERS, filterItems

    Set ConvertSqlRowToDictionary = dictionaryObj
End Function

Private Function ConvertSvgRowToDictionary(ByRef exchange As ExchangeOptions, ByRef svg As svgRow, ByRef row As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    If exchange.svg.row.number Then
        dictionaryObj.Add JSON_ROW, row
    End If
    
    If exchange.svg.row.visible Then
        dictionaryObj.Add JSON_HIDDEN, SvgSheet.rows.Item(row).Hidden
    End If
    
    If exchange.svg.row.height Then
        dictionaryObj.Add JSON_HEIGHT, SvgSheet.rows.Item(row).height
    End If
    
    Dim enabled As Boolean
    enabled = IsRowEnabled(svg.comment)
    If Not enabled Then
        dictionaryObj.Add JSON_ENABLED, enabled
    End If
    
    If svg.find <> vbNullString Then
        dictionaryObj.Add JSON_SVG_FIND, svg.find
    End If
    
    If svg.replace <> vbNullString Then
        dictionaryObj.Add JSON_SVG_REPLACE, svg.replace
    End If
    
    Set ConvertSvgRowToDictionary = dictionaryObj
End Function

Private Function GetStylesRow(ByRef ini As settings, ByVal row As Long) As StylesRow

    GetStylesRow.comment = StylesSheet.Cells.Item(row, ini.styles.flagColumn).value
    GetStylesRow.styleName = StylesSheet.Cells.Item(row, ini.styles.nameColumn).value
    GetStylesRow.format = StylesSheet.Cells.Item(row, ini.styles.formatColumn).value
    GetStylesRow.styleType = StylesSheet.Cells.Item(row, ini.styles.typeColumn).value

End Function

Private Function GetSqlRow(ByRef ini As settings, ByVal row As Long) As sqlRow

    GetSqlRow.comment = SqlSheet.Cells.Item(row, ini.sql.flagColumn).value
    GetSqlRow.excelFile = SqlSheet.Cells.Item(row, ini.sql.excelFileColumn).value
    GetSqlRow.sqlStatement = SqlSheet.Cells.Item(row, ini.sql.sqlStatementColumn).value
    GetSqlRow.status = SqlSheet.Cells.Item(row, ini.sql.statusColumn).value

End Function

Private Function GetSvgRow(ByRef ini As settings, ByVal row As Long) As svgRow

    GetSvgRow.comment = SvgSheet.Cells.Item(row, ini.svg.flagColumn).value
    GetSvgRow.find = SvgSheet.Cells.Item(row, ini.svg.findColumn).value
    GetSvgRow.replace = SvgSheet.Cells.Item(row, ini.svg.replaceColumn).value

End Function

Private Function GetStylesRowViews(ByRef ini As settings, ByVal row As Long) As String()

    Dim switches() As String
    Dim arraySize As Long
    
    Dim lastColumn As Long
    lastColumn = GetLastColumn(StylesSheet.name, row)
    arraySize = lastColumn - ini.styles.firstYesNoColumn + 1
    
    If arraySize > 0 Then
        Dim col As Long
        Dim i As Long
        i = 0
        ReDim switches(arraySize)
        For col = ini.styles.firstYesNoColumn To lastColumn
            switches(i) = StylesSheet.Cells.Item(row, col).value
            i = i + 1
        Next col
    End If
    
    GetStylesRowViews = switches
End Function

Private Function IsArrayAllocated(ByRef arr As Variant) As Boolean
    On Error GoTo ErrorHandler
    IsArrayAllocated = IsArray(arr) And _
                        Not IsError(LBound(arr, 1)) And _
                        LBound(arr, 1) <= UBound(arr, 1)

    Exit Function
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Function


