Attribute VB_Name = "modWorksheetSettings"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Settings")
'@IgnoreModule ProcedureNotUsed

Option Explicit

Public Sub SelectImageDirectory()
    Dim directoryName As String
    
    ' Let the user select a directory
    directoryName = ChooseDirectory(vbNullString)
    
    If directoryName <> vbNullString Then
        ' Update the cell with the directory name chosen
        SetCellString SettingsSheet.name, SETTINGS_IMAGE_PATH, directoryName
    End If
    
End Sub

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

Public Function GetSettingsForStylesWorksheet() As stylesWorksheet
    GetSettingsForStylesWorksheet.headingRow = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    GetSettingsForStylesWorksheet.firstRow = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_FIRST))
    
    GetSettingsForStylesWorksheet.lastRow = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_LAST))
    If GetSettingsForStylesWorksheet.lastRow = 0 Then
        With StylesSheet.UsedRange
            GetSettingsForStylesWorksheet.lastRow = .Cells.Item(.Cells.count).row
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

Public Function GetSettingColNum(ByVal namedRange As String) As Long
    GetSettingColNum = ActiveSheet.Range(SettingsSheet.Range(namedRange).value & 1).Column
End Function

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

Public Function GetSettingsForSqlWorksheet() As sqlWorksheet
    GetSettingsForSqlWorksheet.headingRow = CLng(SettingsSheet.Range(SETTINGS_SQL_ROW_HEADING))
    GetSettingsForSqlWorksheet.firstRow = CLng(SettingsSheet.Range(SETTINGS_SQL_ROW_FIRST))
    With SqlSheet.UsedRange
        GetSettingsForSqlWorksheet.lastRow = .Cells.Item(.Cells.count).row
    End With
    GetSettingsForSqlWorksheet.flagColumn = GetSettingColNum(SETTINGS_SQL_COL_COMMENT)
    GetSettingsForSqlWorksheet.sqlStatementColumn = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
    GetSettingsForSqlWorksheet.excelFileColumn = GetSettingColNum(SETTINGS_SQL_COL_EXCEL_FILE)
    GetSettingsForSqlWorksheet.statusColumn = GetSettingColNum(SETTINGS_SQL_COL_STATUS)
End Function

Public Function GetSettingsForSqlFields() As sqlFieldName
    GetSettingsForSqlFields.Cluster = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_CLUSTER).value))
    GetSettingsForSqlFields.clusterLabel = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_CLUSTER_LABEL).value))
    GetSettingsForSqlFields.clusterStyleName = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME).value))
    GetSettingsForSqlFields.clusterAttributes = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES).value))
    GetSettingsForSqlFields.clusterTooltip = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP).value))

    GetSettingsForSqlFields.subcluster = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER).value))
    GetSettingsForSqlFields.subclusterLabel = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_LABEL).value))
    GetSettingsForSqlFields.subclusterStyleName = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME).value))
    GetSettingsForSqlFields.subclusterAttributes = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES).value))
    GetSettingsForSqlFields.subclusterTooltip = Trim$(LCase$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP).value))
    
    GetSettingsForSqlFields.clusterPlaceholder = Trim$(SettingsSheet.Range(SETTINGS_SQL_COUNT_PLACEHOLDER_CLUSTER).value)
    GetSettingsForSqlFields.subclusterPlaceholder = Trim$(SettingsSheet.Range(SETTINGS_SQL_COUNT_PLACEHOLDER_SUBCLUSTER).value)
    GetSettingsForSqlFields.recordsetPlaceholder = Trim$(SettingsSheet.Range(SETTINGS_SQL_COUNT_PLACEHOLDER_RECORDSET).value)
    
    GetSettingsForSqlFields.splitLength = Trim$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH).value)
    GetSettingsForSqlFields.lineEnding = Trim$(SettingsSheet.Range(SETTINGS_SQL_FIELD_NAME_LINE_ENDING).value)
    
    GetSettingsForSqlFields.filterColumn = Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value)
    GetSettingsForSqlFields.filterValue = Trim$(SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value)
End Function

Public Function GetSettingsForSvgWorksheet() As svgWorksheet
    GetSettingsForSvgWorksheet.headingRow = svgLayoutRow.headingRow
    GetSettingsForSvgWorksheet.firstRow = svgLayoutRow.firstDataRow
    With SvgSheet.UsedRange
        GetSettingsForSvgWorksheet.lastRow = .Cells.Item(.Cells.count).row
    End With
    GetSettingsForSvgWorksheet.flagColumn = svgLayoutColumn.flagColumn
    GetSettingsForSvgWorksheet.findColumn = svgLayoutColumn.findColumn
    GetSettingsForSvgWorksheet.replaceColumn = svgLayoutColumn.replaceColumn
End Function

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

Public Function GetSettingsForConsole() As consoleOptions
    GetSettingsForConsole.logToConsole = GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
    GetSettingsForConsole.appendConsole = GetSettingBoolean(SETTINGS_APPEND_CONSOLE)
    GetSettingsForConsole.graphvizVerbose = GetSettingBoolean(SETTINGS_GRAPHVIZ_VERBOSE)
End Function

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


Public Function GetGraphvizEngine() As String
    GetGraphvizEngine = SETTINGS_DEFAULT_GRAPHVIZ_ENGINE
End Function

Public Function GetSettingsForCommandLine() As CommandLine
    GetSettingsForCommandLine.parameters = SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).value
    GetSettingsForCommandLine.GraphvizPath = SettingsSheet.Range(SETTINGS_GV_PATH).value
End Function

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

Public Function GetSettingBoolean(ByVal cellName As String) As Boolean
    
    GetSettingBoolean = False
    
    Select Case UCase$(Trim$(SettingsSheet.Range(cellName).value))
        Case "ON", "YES", "TRUE", "AUTO", "SHOW", "INCLUDE", "DEFAULT"
            GetSettingBoolean = True
        Case Else
            GetSettingBoolean = False
    End Select
    
End Function

Public Sub DisplayTabRows(ByVal isVisible As Boolean, ByVal rowFrom As Long, ByVal rowTo As Long)
    Dim row As Long
    For row = rowFrom To rowTo
        SettingsSheet.rows.Item(row).Hidden = Not isVisible
    Next row
End Sub

Public Sub DisplayGraphOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_IMAGE_PATH).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_PICTURE_NAME).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabGraphOptions").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabGraphOptions").visible = Not isVisible
End Sub

Public Sub DisplayCmdLineOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_GV_PATH).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabCmdLineOptions").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabCmdLineOptions").visible = Not isVisible
End Sub

Public Sub DisplayStylesOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_STYLES_COL_COMMENT).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabStylesWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabStylesWorksheet").visible = Not isVisible
End Sub

Public Sub DisplayDataOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_DATA_COL_COMMENT).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_DATA_COL_GRAPH).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabDataWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabDataWorksheet").visible = Not isVisible
End Sub

Public Sub DisplaySourceOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_SOURCE_ROW_HEADING).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_SOURCE_INDENT).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabSourceWorksheet").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabSourceWorksheet").visible = Not isVisible
End Sub

Public Sub DisplaySqlOptions(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_SQL_COL_COMMENT).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_SQL_COUNT_PLACEHOLDER_RECORDSET).row + 1
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

Public Sub DisplayGraphvizTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_TAB_GRAPHVIZ).row
    rowTo = SettingsSheet.Range(SETTINGS_TAB_SOURCE).row - 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabGraphvizTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabGraphvizTab").visible = Not isVisible
End Sub

Public Sub DisplaySourceTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range(SETTINGS_TAB_SOURCE).row
    rowTo = SettingsSheet.Range(SETTINGS_EXT_TAB_NAME).row - 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabSourceTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabSourceTab").visible = Not isVisible
End Sub

Public Sub DisplayExtensionsTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
  
    rowFrom = SettingsSheet.Range(SETTINGS_EXT_TAB_NAME).row - 1
    rowTo = SettingsSheet.Range(SETTINGS_TAB_EXCHANGE).row - 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabExtensionsTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabExtensionsTab").visible = Not isVisible
End Sub

Public Sub DisplayExchangeTab(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    
    rowFrom = SettingsSheet.Range("SettingsExchangeTab").row - 1
    rowTo = SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS).row + 1
    DisplayTabRows isVisible, rowFrom, rowTo
    
    SettingsSheet.Shapes.Range("enabledTabExchangeTab").visible = isVisible
    SettingsSheet.Shapes.Range("disabledTabExchangeTab").visible = Not isVisible
End Sub

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

