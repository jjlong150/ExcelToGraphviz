Attribute VB_Name = "modExchangeImport"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Exchange")

Option Explicit

Public Sub ImportData()
    Dim returnMessage As String
    Dim currentLanguage As String

    ' Disable screen updates
    OptimizeCode_Begin
    
    ' Have the user choose a file to import
    Dim importFile As String
    importFile = GetImportFilename(Application.ActiveWorkbook.path)
    
    ' Record what language is currently in use
    currentLanguage = SettingsSheet.Range(SETTINGS_LANGUAGE).value
    
    ' Import the JSON contents
    If importFile <> vbNullString Then
        returnMessage = ImportDataProcessFile(importFile)
    End If
    
    ' If the import specified a change in language, update localizations
    If currentLanguage <> SettingsSheet.Range(SETTINGS_LANGUAGE).value Then
        Localize
    End If
    
    ' Update the ribbon
    RefreshRibbon

    ' Enable screen updates
    OptimizeCode_End
    
    ' Provide any information messages
    If returnMessage <> vbNullString Then
        MsgBox returnMessage, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If
End Sub

Private Function ImportDataProcessFile(ByVal importFile As String) As String
    
    ' Default the return message to null string
    ImportDataProcessFile = vbNullString
    
    Dim jsonString As String
    jsonString = ReadUTF8File(importFile)
    
    ' Kludge - JSON export routine has bug related to in-cell carriage returns
    jsonString = replace(jsonString, "\r\r\r\r\r\u000A", "\u000A")
    
    Dim jsonObject As Object
    Set jsonObject = TryParseJson(jsonString)
    
    If (jsonObject Is Nothing) Then ' Error was already reported
        Exit Function
    End If
    
    ' Test the metadata to see if we should bother proceeding
    If jsonObject.Exists(JSON_SECTION_METADATA) Then
        Dim metadata As Dictionary
        Set metadata = jsonObject.item(JSON_SECTION_METADATA)
    
        Dim name As String
        name = metadata.item("name")
               
        If name <> "E2GXF" Then
            ImportDataProcessFile = "The JSON in this exchange file is not recognized. " & _
                    vbNewLine & vbNewLine & _
                    "Found:    name=""" & name & """" & vbNewLine & _
                    "Expected: name=""E2GXF"""
        End If
    End If
    
    PerformImports jsonObject
    RefreshRibbon

End Function

Private Function TryParseJson(ByVal jsonString As String) As Object
    On Error GoTo parseExit
    
    Set TryParseJson = Nothing
    
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(jsonString)
    
    Set TryParseJson = jsonObject
    Exit Function

parseExit:
    Dim errorMsg As String
    errorMsg = GetMessage("msgboxCannotImportJSON") & vbNewLine & vbNewLine & Err.Description
    MsgBox errorMsg, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    
End Function

Private Sub PerformImports(ByVal dictionaryObj As Dictionary)

    Dim exchange As ExchangeOptions
    exchange = GetExchangeOptions()

    ' Order of the imports matters. We want to import the settings and layouts
    ' and then apply them to the worksheet content when it is imported
    
    ' Import metadata
    If dictionaryObj.Exists(JSON_SECTION_METADATA) Then
        ImportMetadata dictionaryObj.item(JSON_SECTION_METADATA)
        If Application.Calculation = xlManual Then
            SettingsSheet.Calculate
        End If
    End If
    
    ' Import worksheet layouts
    If dictionaryObj.Exists(JSON_SECTION_LAYOUTS) Then
        ImportLayouts dictionaryObj.item(JSON_SECTION_LAYOUTS), exchange
        If Application.Calculation = xlManual Then
            SettingsSheet.Calculate
        End If
    End If
    
    ' Import settings
    If dictionaryObj.Exists(JSON_SECTION_SETTINGS) Then
        ImportSettings dictionaryObj.item(JSON_SECTION_SETTINGS), exchange
        If Application.Calculation = xlManual Then
            SettingsSheet.Calculate
        End If
    End If
    
    ' Import worksheet contents
    If dictionaryObj.Exists(JSON_SECTION_CONTENT) Then
        ' Refresh the ini settings based upon what settings were imported
        ' as column order may have changed.
        Dim ini As settings
        ini = GetSettings(DataSheet.name)

        ImportContent dictionaryObj.item(JSON_SECTION_CONTENT), ini, exchange
        If Application.Calculation = xlManual Then
            SettingsSheet.Calculate
        End If
    End If
    
    ' Refresh the graph
    CreateGraphWorksheet
End Sub

Private Sub ImportMetadata(ByVal dictionaryObj As Dictionary)
    
    Dim key As Variant
    For Each key In dictionaryObj.Keys()
        Select Case key
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_NAME
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_TYPE
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_VERSION
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_USER
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_DATE
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_TIME
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_EXCEL
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_OS
            '@Ignore EmptyCaseBlock
            Case JSON_METADATA_FILENAME
            Case Else
                MsgBox GetMessage("msgboxUnexpectedMetaData") & vbNewLine & vbNewLine & key & "=" & dictionaryObj.item(key), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        End Select
    Next
    
End Sub

Private Sub ImportContent(ByVal dictionaryObj As Dictionary, ByRef ini As settings, ByRef exchange As ExchangeOptions)

    ' Each worksheet has its own section in the json object
    Dim worksheetName As Variant
    For Each worksheetName In dictionaryObj.Keys()
        Select Case worksheetName
            Case WORKSHEET_DATA
                If exchange.data.include Then
                    ImportContentData ini, exchange, dictionaryObj.item(worksheetName)
                End If
            
            Case WORKSHEET_SQL
                If exchange.sql.include Then
                    ImportContentSql ini, exchange, dictionaryObj.item(worksheetName)
                End If
            
            Case WORKSHEET_SVG
                If exchange.svg.include Then
                    ImportContentSvg ini, exchange, dictionaryObj.item(worksheetName)
                End If
            
            Case WORKSHEET_STYLES
                If exchange.styles.include Then
                    ImportContentStyles ini, exchange, dictionaryObj.item(worksheetName)
                    StylesSheet.Activate
                    ClearStylesPreview
                    GenerateStylesPreviewAll
                    ClearStatusBar
                End If
        End Select
    Next
    
End Sub

Private Sub ImportSettings(ByVal dictionaryObj As Dictionary, ByRef exchange As ExchangeOptions)

    ' Quick abort if user does not want the settings imported
    If Not exchange.includeSettings Then
        Exit Sub
    End If
    
    ' Settings are organized based upon what worksheet they live on
    Dim worksheetName As Variant
    For Each worksheetName In dictionaryObj.Keys()
        Select Case worksheetName
            Case WORKSHEET_DATA
                 ImportSettingsData dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_GRAPH
                ImportSettingsGraph dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_SETTINGS
                ImportSettingsSettings dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_SOURCE
                ImportSettingsSource dictionaryObj.item(worksheetName)
                    
            Case WORKSHEET_SQL
                ImportSettingsSql dictionaryObj.item(worksheetName)
                    
            Case "extensions"
                ImportSettingsExtensions dictionaryObj.item(worksheetName)
        End Select
    Next
    
End Sub

Private Sub ImportSettingsGraph(ByVal dictionaryObj As Dictionary)
    If dictionaryObj.Exists(JSON_ZOOM) Then
        SetZoom GraphSheet.name, dictionaryObj.item(JSON_ZOOM)
    End If
End Sub

Private Sub ImportSettingsSettings(ByVal dictionaryObj As Dictionary)
    RestoreSetting SETTINGS_GV_PATH, dictionaryObj.item(JSON_SETTINGS_GV_PATH)
    RestoreSetting SETTINGS_IMAGE_PATH, dictionaryObj.item(JSON_SETTINGS_IMAGE_PATH)
    RestoreSetting SETTINGS_GRAPH_OPTIONS, dictionaryObj.item(JSON_SETTINGS_GRAPH_OPTIONS)
    RestoreSetting SETTINGS_PICTURE_NAME, dictionaryObj.item(JSON_SETTINGS_PICTURE_NAME)
    RestoreSetting SETTINGS_COMMAND_LINE_PARAMETERS, dictionaryObj.item(JSON_SETTINGS_COMMAND_LINE_PARAMETERS)
End Sub

Private Sub ImportSettingsSource(ByVal dictionaryObj As Dictionary)
    Dim buttons As Collection
    Dim button As Dictionary
    Dim i As Long

    Set buttons = dictionaryObj.item(JSON_SETTINGS_BUTTONS)

    RestoreSetting SETTINGS_SOURCE_INDENT, dictionaryObj.item(JSON_SETTINGS_INDENT)
    
    For i = 1 To 6
        Set button = buttons.item(i)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_TEXT, button.item(JSON_SETTINGS_BUTTON_TEXT)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_URL, button.item(JSON_SETTINGS_URL)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_SCREENTIP, button.item(JSON_SETTINGS_SCREEN_TIP)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_SUPERTIP, button.item(JSON_SETTINGS_SUPER_TIP)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_VISIBLE, button.item(JSON_SETTINGS_VISIBLE)
    Next i
End Sub

Private Sub ImportSettingsSql(ByVal dictionaryObj As Dictionary)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_LABEL, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_LABEL)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP)
    
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_LABEL, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_LABEL)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP)
    
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_LINE_ENDING, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_LINE_ENDING)

    RestoreSetting SETTINGS_SQL_COUNT_PLACEHOLDER_CLUSTER, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_PLACEHOLDER)
    RestoreSetting SETTINGS_SQL_COUNT_PLACEHOLDER_SUBCLUSTER, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_PLACEHOLDER)
    RestoreSetting SETTINGS_SQL_COUNT_PLACEHOLDER_RECORDSET, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_RECORDSET_PLACEHOLDER)
    
    If dictionaryObj.Exists(JSON_SETTINGS_SQL_FIELD_NAME_TREE_QUERY) Then
        RestoreSetting SETTINGS_SQL_FIELD_NAME_TREE_QUERY, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_TREE_QUERY)
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SQL_FIELD_NAME_WHERE_COLUMN) Then
        RestoreSetting SETTINGS_SQL_FIELD_NAME_WHERE_COLUMN, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_WHERE_COLUMN)
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SQL_FIELD_NAME_WHERE_VALUE) Then
        RestoreSetting SETTINGS_SQL_FIELD_NAME_WHERE_VALUE, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_WHERE_VALUE)
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SQL_FIELD_NAME_MAX_DEPTH) Then
        RestoreSetting SETTINGS_SQL_FIELD_NAME_MAX_DEPTH, dictionaryObj.item(JSON_SETTINGS_SQL_FIELD_NAME_MAX_DEPTH)
    End If
        
    RestoreSetting SETTINGS_SQL_COL_FILTER, dictionaryObj.item(JSON_SETTINGS_SQL_FILTER_COLUMN)
    RestoreSetting SETTINGS_SQL_FILTER_VALUE, dictionaryObj.item(JSON_SETTINGS_SQL_FILTER_VALUE)

    If dictionaryObj.Exists(JSON_SETTINGS_SQL_CLOSE_CONNECTIONS) Then
        RestoreSetting SETTINGS_SQL_CLOSE_CONNECTIONS, BooleanToYesNo(dictionaryObj.item(JSON_SETTINGS_SQL_CLOSE_CONNECTIONS))
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SQL_DATASOURCE_DIRECTORY) Then
        RestoreSetting SETTINGS_DATASOURCE_DIRECTORY, dictionaryObj.item(JSON_SETTINGS_SQL_DATASOURCE_DIRECTORY)
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SQL_DATASOURCE_FILE) Then
        RestoreSetting SETTINGS_DATASOURCE_FILE, dictionaryObj.item(JSON_SETTINGS_SQL_DATASOURCE_FILE)
    End If
End Sub

Private Sub ImportSettingsExtensions(ByVal dictionaryObj As Dictionary)
    Dim buttons As Collection
    Dim button As Dictionary
    Dim i As Long

    RestoreSetting SETTINGS_EXT_TAB_NAME, dictionaryObj.item(JSON_SETTINGS_EXT_TAB_NAME)
    
    Dim group As Dictionary
    
    If dictionaryObj.Exists(JSON_SETTINGS_EXT_TAB_GROUP_NAME_WEB) Then
        Set group = dictionaryObj.item(JSON_SETTINGS_EXT_TAB_GROUP_NAME_WEB)
        Set buttons = group.item(JSON_SETTINGS_BUTTONS)
        RestoreSetting SETTINGS_EXT_TAB_GROUP_NAME_WEB, group.item(JSON_SETTINGS_GROUP_NAME)
    
        For i = 1 To 6
            Set button = buttons.item(i)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_TEXT, button.item(JSON_SETTINGS_BUTTON_TEXT)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_URL, button.item(JSON_SETTINGS_URL)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_SCREENTIP, button.item(JSON_SETTINGS_SCREEN_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_SUPERTIP, button.item(JSON_SETTINGS_SUPER_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_VISIBLE, button.item(JSON_SETTINGS_VISIBLE)
        Next i
    End If
    
     If dictionaryObj.Exists(JSON_SETTINGS_EXT_TAB_GROUP_NAME_CODE) Then
        Set group = dictionaryObj.item(JSON_SETTINGS_EXT_TAB_GROUP_NAME_CODE)
        Set buttons = group.item(JSON_SETTINGS_BUTTONS)
        RestoreSetting SETTINGS_EXT_TAB_GROUP_NAME_CODE, group.item(JSON_SETTINGS_GROUP_NAME)
    
        For i = 1 To 6
            Set button = buttons.item(i)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_TEXT, button.item(JSON_SETTINGS_BUTTON_TEXT)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SUB, button.item(JSON_SETTINGS_SUB)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SCREENTIP, button.item(JSON_SETTINGS_SCREEN_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SUPERTIP, button.item(JSON_SETTINGS_SUPER_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_VISIBLE, button.item(JSON_SETTINGS_VISIBLE)
        Next i
    End If
End Sub

Private Sub ImportSettingsData(ByVal dictionaryObj As Dictionary)
    Dim section As Dictionary
    Dim subSection As Dictionary
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_GRAPH_TO_WORKSHEET) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_GRAPH_TO_WORKSHEET)
        RestoreSetting SETTINGS_RUN_MODE, section.item(JSON_SETTINGS_RUN_MODE)
        RestoreSetting SETTINGS_IMAGE_TYPE, section.item(JSON_SETTINGS_IMAGE_TYPE)
        RestoreSetting SETTINGS_IMAGE_WORKSHEET, section.item(JSON_SETTINGS_IMAGE_WORKSHEET)
        
        If section.item(JSON_SETTINGS_SCALE_IMAGE) = vbNullString Then
            RestoreSetting SETTINGS_SCALE_IMAGE, "100"  ' For backward compatability with older export files
        Else
            RestoreSetting SETTINGS_SCALE_IMAGE, section.item(JSON_SETTINGS_SCALE_IMAGE)
        End If
    End If
                    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_GRAPH_TO_FILE) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_GRAPH_TO_FILE)
        RestoreSetting SETTINGS_OUTPUT_DIRECTORY, section.item(JSON_SETTINGS_DIRECTORY)
        RestoreSetting SETTINGS_FILE_NAME, section.item(JSON_SETTINGS_FILE_NAME_PREFIX)
        RestoreSetting SETTINGS_FILE_FORMAT, section.item(JSON_SETTINGS_IMAGE_TYPE)
        RestoreSetting SETTINGS_APPEND_OPTIONS, BooleanToYesNo(section.item(JSON_SETTINGS_APPEND_OPTIONS))
        RestoreSetting SETTINGS_APPEND_TIMESTAMP, BooleanToYesNo(section.item(JSON_SETTINGS_APPEND_TIME_STAMP))
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_LAYOUT) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_LAYOUT)
        RestoreSetting SETTINGS_GRAPHVIZ_ENGINE, section.item(JSON_SETTINGS_ENGINE)
        
        ' Maintain backward compatibility. Direction and rankdir were consolidated into rankdir to improve performance
        ' when the port to Apple Mac was performed. Old versions of the spreadsheet will still export direction and not rankdir
        ' so the value will be derived from direction to set the rankdir cell.
        Select Case LCase$(section.item(JSON_SETTINGS_DIRECTION))
            Case "top to bottom"
                RestoreSetting SETTINGS_RANKDIR, "TB"
            Case "bottom to top"
                RestoreSetting SETTINGS_RANKDIR, "BT"
            Case "left to right"
                RestoreSetting SETTINGS_RANKDIR, "LR"
            Case "right to left"
                RestoreSetting SETTINGS_RANKDIR, "RL"
            Case Else
                RestoreSetting SETTINGS_RANKDIR, vbNullString
        End Select
        RestoreSetting SETTINGS_SPLINES, section.item(JSON_SETTINGS_SPLINES)
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_OPTIONS) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_OPTIONS)
        If section.Exists(JSON_SETTINGS_SECTION_GRAPH) Then
            Set subSection = section.item(JSON_SETTINGS_SECTION_GRAPH)
            RestoreSetting SETTINGS_GRAPH_CENTER, BooleanToYesNo(subSection.item(JSON_SETTINGS_CENTER))
            RestoreSetting SETTINGS_GRAPH_CLUSTER_RANK, subSection.item(JSON_SETTINGS_CLUSTER_RANK)
            RestoreSetting SETTINGS_GRAPH_COMPOUND, BooleanToYesNo(subSection.item(JSON_SETTINGS_COMPOUND))
            RestoreSetting SETTINGS_GRAPH_DIM, subSection.item(JSON_SETTINGS_DIM)
            RestoreSetting SETTINGS_GRAPH_DIMEN, subSection.item(JSON_SETTINGS_DIMEN)
            RestoreSetting SETTINGS_GRAPH_FORCE_LABELS, BooleanToYesNo(subSection.item(JSON_SETTINGS_FORCE_LABELS))
            RestoreSetting SETTINGS_GRAPH_MODE, subSection.item(JSON_SETTINGS_MODE)
            RestoreSetting SETTINGS_GRAPH_MODEL, subSection.item(JSON_SETTINGS_MODEL)
            RestoreSetting SETTINGS_GRAPH_NEWRANK, BooleanToYesNo(subSection.item(JSON_SETTINGS_NEWRANK))
            RestoreSetting SETTINGS_GRAPH_ORDERING, subSection.item(JSON_SETTINGS_ORDERING)
            RestoreSetting SETTINGS_GRAPH_ORIENTATION, BooleanToYesNo(subSection.item(JSON_SETTINGS_ORIENTATION))
            RestoreSetting SETTINGS_GRAPH_OUTPUT_ORDER, subSection.item(JSON_SETTINGS_OUTPUT_ORDER)
            RestoreSetting SETTINGS_GRAPH_OVERLAP, subSection.item(JSON_SETTINGS_OVERLAP)
            RestoreSetting SETTINGS_GRAPH_SMOOTHING, subSection.item(JSON_SETTINGS_SMOOTHING)
            RestoreSetting SETTINGS_GRAPH_TRANSPARENT, BooleanToYesNo(subSection.item(JSON_SETTINGS_TRANSPARENT_BACKGROUND))
            RestoreSetting SETTINGS_GRAPH_INCLUDE_IMAGE_PATH, BooleanToYesNo(subSection.item(JSON_SETTINGS_INCLUDE_IMAGE_PATH))
            RestoreSetting SETTINGS_GRAPH_TYPE, subSection.item(JSON_SETTINGS_GRAPH_TYPE)
        End If
    
        If section.Exists(JSON_SETTINGS_SECTION_NODES) Then
            Set subSection = section.item(JSON_SETTINGS_SECTION_NODES)
            RestoreSetting SETTINGS_NODES_WITHOUT_RELATIONSHIPS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_ORPHAN_NODES))
            RestoreSetting SETTINGS_NODE_LABELS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_NODE_LABELS))
            RestoreSetting SETTINGS_NODE_XLABELS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_NODE_XLABELS))
            RestoreSetting SETTINGS_BLANK_NODE_LABELS, subSection.item(JSON_SETTINGS_BLANK_NODE_LABELS)
        End If
    
        If section.Exists(JSON_SETTINGS_SECTION_EDGES) Then
            Set subSection = section.item(JSON_SETTINGS_SECTION_EDGES)
            RestoreSetting SETTINGS_GRAPH_STRICT, BooleanToYesNo(subSection.item(JSON_SETTINGS_ADD_STRICT))
            RestoreSetting SETTINGS_GRAPH_CONCENTRATE, BooleanToYesNo(subSection.item(JSON_SETTINGS_CONCENTRATE))
            RestoreSetting SETTINGS_RELATIONSHIPS_WITHOUT_NODES, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_ORPHAN_EDGES))
            RestoreSetting SETTINGS_EDGE_HEAD_LABELS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_EDGE_HEAD_LABELS))
            RestoreSetting SETTINGS_EDGE_LABELS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_EDGE_LABELS))
            RestoreSetting SETTINGS_EDGE_TAIL_LABELS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_EDGE_TAIL_LABELS))
            RestoreSetting SETTINGS_EDGE_XLABELS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_EDGE_XLABELS))
            RestoreSetting SETTINGS_EDGE_PORTS, BooleanToIncludeExclude(subSection.item(JSON_SETTINGS_INCLUDE_EDGE_PORTS))
            RestoreSetting SETTINGS_BLANK_EDGE_LABELS, subSection.item(JSON_SETTINGS_BLANK_EDGE_LABELS)
        End If
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_STYLES) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_STYLES)
        RestoreSetting SETTINGS_STYLES_COL_SHOW_STYLE, section.item(JSON_SETTINGS_SELECTED_VIEW_COLUMN)
        RestoreSetting SETTINGS_INCLUDE_STYLE_FORMAT, BooleanToIncludeExclude(section.item(JSON_SETTINGS_INCLUDE_STYLE_FORMAT))
        RestoreSetting SETTINGS_INCLUDE_EXTRA_ATTRIBUTES, BooleanToIncludeExclude(section.item(JSON_SETTINGS_INCLUDE_EXTRA_ATTRIBUTES))
        
        ' Old exports do not have this value
        If section.Exists(JSON_SETTINGS_STYLES_SUFFIX_OPEN) Then
            RestoreSetting SETTINGS_STYLES_SUFFIX_OPEN, section.item(JSON_SETTINGS_STYLES_SUFFIX_OPEN)
        End If
        
        ' Old exports do not have this value
        If section.Exists(SETTINGS_STYLES_SUFFIX_CLOSE) Then
            RestoreSetting SETTINGS_STYLES_SUFFIX_CLOSE, section.item(JSON_SETTINGS_STYLES_SUFFIX_CLOSE)
        End If
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_DEBUG) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_DEBUG)
        RestoreSetting SETTINGS_DEBUG, section.item(JSON_SETTINGS_DEBUG_SWITCH)
        RestoreSetting SETTINGS_FILE_DISPOSITION, section.item(JSON_SETTINGS_FILE_DISPOSITION)
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_CONSOLE) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_CONSOLE)
        RestoreSetting SETTINGS_LOG_TO_CONSOLE, BooleanToYesNo(section.item(JSON_SETTINGS_LOG_TO_CONSOLE))
        RestoreSetting SETTINGS_APPEND_CONSOLE, BooleanToYesNo(section.item(JSON_SETTINGS_APPEND_CONSOLE))
        RestoreSetting SETTINGS_GRAPHVIZ_VERBOSE, BooleanToYesNo(section.item(JSON_SETTINGS_GRAPHVIZ_VERBOSE))
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_COLUMNS) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_COLUMNS)
        RestoreSetting SETTINGS_DATA_SHOW_COMMENT, LCase$(section.item(JSON_STYLES_FLAG))
        RestoreSetting SETTINGS_DATA_SHOW_ITEM, LCase$(section.item(JSON_DATA_ITEM))
        RestoreSetting SETTINGS_DATA_SHOW_LABEL, LCase$(section.item(JSON_DATA_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_OUTSIDE_LABEL, LCase$(section.item(JSON_DATA_OUTSIDE_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_TAIL_LABEL, LCase$(section.item(JSON_DATA_TAIL_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_HEAD_LABEL, LCase$(section.item(JSON_DATA_HEAD_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_IS_RELATED_TO_ITEM, LCase$(section.item(JSON_DATA_RELATED_ITEM))
        RestoreSetting SETTINGS_DATA_SHOW_STYLE, LCase$(section.item(JSON_DATA_STYLE_NAME))
        RestoreSetting SETTINGS_DATA_SHOW_EXTRA_STYLE_ATTRIBUTES, LCase$(section.item(JSON_DATA_EXTRA_ATTRIBUTES))
        RestoreSetting SETTINGS_DATA_SHOW_MESSAGES, LCase$(section.item(JSON_DATA_MESSAGE))
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_LANGUAGE) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_LANGUAGE)
        RestoreSetting SETTINGS_LANGUAGE, section.item(JSON_SETTINGS_LANGUAGE)
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_WORKSHEETS) Then
        Set section = dictionaryObj.item(JSON_SETTINGS_SECTION_WORKSHEETS)
        RestoreRequiredSetting SETTINGS_HELP_ATTRIBUTES, section.item(JSON_WORKSHEETS_ATTRIBUTES)
        RestoreRequiredSetting SETTINGS_HELP_COLORS, section.item(JSON_WORKSHEETS_COLORS)
        RestoreRequiredSetting SETTINGS_HELP_SHAPES, section.item(JSON_WORKSHEETS_SHAPES)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_ABOUT, section.item(JSON_WORKSHEETS_ABOUT)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_CONSOLE, section.item(JSON_WORKSHEETS_CONSOLE)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS, section.item(JSON_WORKSHEETS_DIAGNOSTICS)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LISTS, section.item(JSON_WORKSHEETS_LISTS)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LOCALE_DE_DE, section.item(JSON_WORKSHEETS_LOCALE_DE_DE)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LOCALE_EN_GB, section.item(JSON_WORKSHEETS_LOCALE_EN_GB)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LOCALE_EN_US, section.item(JSON_WORKSHEETS_LOCALE_EN_US)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LOCALE_FR_FR, section.item(JSON_WORKSHEETS_LOCALE_FR_FR)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LOCALE_IT_IT, section.item(JSON_WORKSHEETS_LOCALE_IT_IT)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_LOCALE_PL_PL, section.item(JSON_WORKSHEETS_LOCALE_PL_PL)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_SETTINGS, section.item(JSON_WORKSHEETS_SETTINGS)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_SOURCE, section.item(JSON_WORKSHEETS_SOURCE)
#If Mac Then
        ' SQL is only available on Windows. Prevent a file exported on a Windows PC from restoring
        ' a setting on a Mac which would be invalid if imported.
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_SQL, TOGGLE_HIDE
#Else
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_SQL, section.item(JSON_WORKSHEETS_SQL)
#End If
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER, section.item(JSON_WORKSHEETS_STYLE_DESIGNER)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_STYLES, section.item(JSON_WORKSHEETS_STYLES)
        RestoreRequiredSetting SETTINGS_TOOLS_TOGGLE_SVG, section.item(JSON_WORKSHEETS_SVG)
        
        ' Show or hide the worksheets based on the newly restored settings
        ShowOrHideWorksheets
    End If
End Sub

Private Sub SetZoom(ByVal worksheetName As String, ByVal zoom As Long)
    ' Save the name of the current worksheet
    Dim previousSheet As String
    previousSheet = ActiveSheet.name
    
    ' Switch to the sheet we need to get the zoom value from
    ActiveWorkbook.Sheets.[_Default](worksheetName).Activate
    ActiveWindow.zoom = zoom
    
    ' Switch back to the original worksheet
    ActiveWorkbook.Sheets.[_Default](previousSheet).Activate
End Sub

Private Function BooleanToYesNo(ByVal setting As Boolean) As String
    BooleanToYesNo = TOGGLE_NO
    If setting Then BooleanToYesNo = TOGGLE_YES
End Function

Private Function BooleanToIncludeExclude(ByVal setting As Boolean) As String
    BooleanToIncludeExclude = TOGGLE_EXCLUDE
    If setting Then BooleanToIncludeExclude = TOGGLE_INCLUDE
End Function

Private Sub RestoreSetting(ByVal cellName As String, ByVal cellValue As String)
    ' This sub can nullify existing settings
    SettingsSheet.Range(cellName).value = cellValue
End Sub

Private Sub RestoreRequiredSetting(ByVal cellName As String, ByVal cellValue As String)
    ' This sub prevents existing values from being nulled out
    If Trim$(cellValue) <> vbNullString Then
        SettingsSheet.Range(cellName).value = cellValue
    End If
End Sub

Private Sub ImportLayouts(ByVal dictionaryObj As Dictionary, ByRef exchange As ExchangeOptions)

    ' Quick abort if user does not want the layouts imported
    If Not exchange.includeLayouts Then
        Exit Sub
    End If
    
    Dim key As Variant
    Dim worksheetName As String
    
    For Each key In dictionaryObj.Keys()
        worksheetName = key
            
        ' Set the row heights
        ImportLayoutsRowHeights worksheetName, dictionaryObj.item(worksheetName)

        ' Column layouts
        Select Case worksheetName
            Case WORKSHEET_DATA
                ImportLayoutsData dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_STYLES
                ImportLayoutsStyles dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_SQL
                ImportLayoutsSql dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_SVG
                ImportLayoutsSvg dictionaryObj.item(worksheetName)
            
            Case WORKSHEET_SOURCE
                ImportLayoutsSource dictionaryObj.item(worksheetName)
        End Select
    Next
End Sub

Private Sub ImportLayoutsData(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetData
    
    Dim columns As Collection
    Set columns = dictionaryObj.item(JSON_COLUMNS)
    
    Dim data As dataWorksheet
    data = GetSettingsForDataWorksheet(GetDataWorksheetName())

    Dim i As Long
    
    For i = 1 To columns.count
        Select Case columns.item(i)(JSON_ID)
            Case JSON_DATA_FLAG
                DataSheet.Cells.item(data.headingRow, data.flagColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.flagColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.flagColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.flagColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
           
            Case JSON_DATA_ITEM
                DataSheet.Cells.item(data.headingRow, data.itemColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.itemColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.itemColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.itemColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
           
            Case JSON_DATA_LABEL
                DataSheet.Cells.item(data.headingRow, data.labelColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.labelColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.labelColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.labelColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_OUTSIDE_LABEL
                DataSheet.Cells.item(data.headingRow, data.xLabelColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.xLabelColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.xLabelColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.xLabelColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_TAIL_LABEL
                DataSheet.Cells.item(data.headingRow, data.tailLabelColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.tailLabelColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.tailLabelColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.tailLabelColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_HEAD_LABEL
                DataSheet.Cells.item(data.headingRow, data.headLabelColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.headLabelColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.headLabelColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.headLabelColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_TOOLTIP
                DataSheet.Cells.item(data.headingRow, data.tooltipColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.tooltipColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.tooltipColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.tooltipColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_RELATED_ITEM
                DataSheet.Cells.item(data.headingRow, data.isRelatedToItemColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.isRelatedToItemColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.isRelatedToItemColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.isRelatedToItemColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_STYLE_NAME
                DataSheet.Cells.item(data.headingRow, data.styleNameColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.styleNameColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.styleNameColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.styleNameColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_EXTRA_ATTRIBUTES
                DataSheet.Cells.item(data.headingRow, data.extraAttributesColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.extraAttributesColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.extraAttributesColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.extraAttributesColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_MESSAGE
                DataSheet.Cells.item(data.headingRow, data.errorMessageColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.errorMessageColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.errorMessageColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.errorMessageColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_GRAPH_DISPLAY_COLUMN
                DataSheet.Cells.item(data.headingRow, data.graphDisplayColumn).value = columns.item(i)(JSON_HEADING)
                DataSheet.columns.item(data.graphDisplayColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                DataSheet.columns.item(data.graphDisplayColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                DataSheet.columns.item(data.graphDisplayColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsStyles(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetStyles
    
    Dim columns As Collection
    Set columns = dictionaryObj.item(JSON_COLUMNS)
    
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    Dim i As Long
    
    Dim offset As Long
    offset = -1
    
    For i = 1 To columns.count
        Select Case columns.item(i)(JSON_ID)
            Case JSON_STYLES_FLAG
                StylesSheet.Cells.item(styles.headingRow, styles.flagColumn).value = columns.item(i)(JSON_HEADING)
                StylesSheet.columns.item(styles.flagColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                StylesSheet.columns.item(styles.flagColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                StylesSheet.columns.item(styles.flagColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
           
            Case JSON_STYLES_NAME
                StylesSheet.Cells.item(styles.headingRow, styles.nameColumn).value = columns.item(i)(JSON_HEADING)
                StylesSheet.columns.item(styles.nameColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                StylesSheet.columns.item(styles.nameColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                StylesSheet.columns.item(styles.nameColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
           
            Case JSON_STYLES_FORMAT
                StylesSheet.Cells.item(styles.headingRow, styles.formatColumn).value = columns.item(i)(JSON_HEADING)
                StylesSheet.columns.item(styles.formatColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                StylesSheet.columns.item(styles.formatColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                StylesSheet.columns.item(styles.formatColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_STYLES_TYPE
                StylesSheet.Cells.item(styles.headingRow, styles.typeColumn).value = columns.item(i)(JSON_HEADING)
                StylesSheet.columns.item(styles.typeColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                StylesSheet.columns.item(styles.typeColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                StylesSheet.columns.item(styles.typeColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case Else
                offset = offset + 1
                StylesSheet.Cells.item(styles.headingRow, styles.firstYesNoColumn + offset).value = columns.item(i)(JSON_HEADING)
                StylesSheet.columns.item(styles.firstYesNoColumn + offset).ColumnWidth = columns.item(i)(JSON_WIDTH)
                StylesSheet.columns.item(styles.firstYesNoColumn + offset).Hidden = columns.item(i)(JSON_HIDDEN)
                StylesSheet.columns.item(styles.firstYesNoColumn + offset).WrapText = columns.item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsSql(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetSql
    
    Dim columns As Collection
    Set columns = dictionaryObj.item(JSON_COLUMNS)
    
    Dim sql As sqlWorksheet
    sql = GetSettingsForSqlWorksheet()
    
    Dim i As Long
    
    For i = 1 To columns.count
        Select Case columns.item(i)(JSON_ID)
            Case JSON_LAYOUT_SQL_FLAG
                SqlSheet.Cells.item(sql.headingRow, sql.flagColumn).value = columns.item(i)(JSON_HEADING)
                SqlSheet.columns.item(sql.flagColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SqlSheet.columns.item(sql.flagColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SqlSheet.columns.item(sql.flagColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SQL_SQL_STATEMENT
                SqlSheet.Cells.item(sql.headingRow, sql.sqlStatementColumn).value = columns.item(i)(JSON_HEADING)
                SqlSheet.columns.item(sql.sqlStatementColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SqlSheet.columns.item(sql.sqlStatementColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SqlSheet.columns.item(sql.sqlStatementColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SQL_EXCEL_FILE
                SqlSheet.Cells.item(sql.headingRow, sql.excelFileColumn).value = columns.item(i)(JSON_HEADING)
                SqlSheet.columns.item(sql.excelFileColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SqlSheet.columns.item(sql.excelFileColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SqlSheet.columns.item(sql.excelFileColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SQL_STATUS
                SqlSheet.Cells.item(sql.headingRow, sql.statusColumn).value = columns.item(i)(JSON_HEADING)
                SqlSheet.columns.item(sql.statusColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SqlSheet.columns.item(sql.statusColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SqlSheet.columns.item(sql.statusColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsSvg(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetSvg
    
    Dim columns As Collection
    Set columns = dictionaryObj.item(JSON_COLUMNS)
    
    Dim svg As svgWorksheet
    svg = GetSettingsForSvgWorksheet()
    
    Dim i As Long
    
    For i = 1 To columns.count
        Select Case columns.item(i)(JSON_ID)
            Case JSON_LAYOUT_SVG_FLAG
                SvgSheet.Cells.item(svg.headingRow, svg.flagColumn).value = columns.item(i)(JSON_HEADING)
                SvgSheet.columns.item(svg.flagColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SvgSheet.columns.item(svg.flagColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SvgSheet.columns.item(svg.flagColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SVG_FIND
                SvgSheet.Cells.item(svg.headingRow, svg.findColumn).value = columns.item(i)(JSON_HEADING)
                SvgSheet.columns.item(svg.findColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SvgSheet.columns.item(svg.findColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SvgSheet.columns.item(svg.findColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SVG_REPLACE
                SvgSheet.Cells.item(svg.headingRow, svg.replaceColumn).value = columns.item(i)(JSON_HEADING)
                SvgSheet.columns.item(svg.replaceColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SvgSheet.columns.item(svg.replaceColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SvgSheet.columns.item(svg.replaceColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsSource(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetSource
    
    Dim columns As Collection
    Set columns = dictionaryObj.item(JSON_COLUMNS)
    
    Dim source As sourceWorksheet
    source = GetSettingsForSourceWorksheet()
    
    Dim i As Long
    For i = 1 To columns.count
        Select Case columns.item(i)(JSON_ID)
            Case JSON_SOURCE_LINE_NUMBER
                SourceSheet.Cells.item(source.headingRow, source.lineNumberColumn).value = columns.item(i)(JSON_HEADING)
                SourceSheet.columns.item(source.lineNumberColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SourceSheet.columns.item(source.lineNumberColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SourceSheet.columns.item(source.lineNumberColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
           
            Case JSON_SOURCE_SOURCE
                SourceSheet.Cells.item(source.headingRow, source.sourceColumn).value = columns.item(i)(JSON_HEADING)
                SourceSheet.columns.item(source.sourceColumn).ColumnWidth = columns.item(i)(JSON_WIDTH)
                SourceSheet.columns.item(source.sourceColumn).Hidden = columns.item(i)(JSON_HIDDEN)
                SourceSheet.columns.item(source.sourceColumn).WrapText = columns.item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsRowHeights(ByRef worksheetName As String, ByVal dictionaryObj As Dictionary)
    Dim rows As Collection
    Dim row As Dictionary
    
    Set rows = dictionaryObj.item(JSON_ROWS)
    Dim i As Long

    ' Set the row heights
    For i = 1 To rows.count
        Set row = rows.item(i)
        ActiveWorkbook.Sheets.[_Default](worksheetName).rows(row.item(JSON_ROW)).rowHeight = row.item(JSON_HEIGHT)
        ActiveWorkbook.Sheets.[_Default](worksheetName).rows(row.item(JSON_ROW)).Hidden = row.item(JSON_HIDDEN)
    Next i
End Sub

Private Function GetImportFilename(ByVal initialFolder As String) As String
    GetImportFilename = vbNullString
#If Mac Then
    GetImportFilename = RunAppleScriptTask("chooseOneFile", "json")
#Else
    Dim fileDialogHandle As FileDialog
    Set fileDialogHandle = Application.FileDialog(msoFileDialogFilePicker)
    fileDialogHandle.title = "Choose an Excel to Graphviz data exchange file"
    fileDialogHandle.AllowMultiSelect = False
    fileDialogHandle.Filters.Clear
    fileDialogHandle.Filters.Add "Excel to Graphviz Files", "*.json", 1
    fileDialogHandle.InitialFileName = initialFolder
    
    'Get the number of the button chosen
    If fileDialogHandle.show Then
        GetImportFilename = Trim$(fileDialogHandle.SelectedItems.item(1))
    End If

    Set fileDialogHandle = Nothing
#End If
End Function

Private Sub ImportContentData(ByRef ini As settings, ByRef exchange As ExchangeOptions, ByVal rows As Collection)
    Dim i As Long
    Dim key As Variant
    Dim row As Long
    Dim extraAttributes As String
    Dim dictionaryObj As Dictionary
    Dim firstRow As Long
    
    ' First possible row = 1
    '@Ignore AssignmentNotUsed
    firstRow = 1
    
    Dim lastRow As Long
    With DataSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    Select Case exchange.data.action
        Case IMPORT_REPLACE
            firstRow = ini.data.firstRow
            ClearWorksheetData ini

        Case IMPORT_APPEND
            firstRow = lastRow + 1
    End Select
    
    ' Loop through all the objects in collection
    For i = 1 To rows.count
        Select Case exchange.data.action
            Case IMPORT_REPLACE
                If rows.item(i).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.item(i)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + i - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + i - 1
        End Select
        
        DataSheet.Cells.item(row, 1).EntireRow.ClearContents

        For Each key In rows.item(i).Keys()
            Select Case key
                Case JSON_HIDDEN
                    DataSheet.rows.item(row).Hidden = rows.item(i)(key)
                    
                Case JSON_HEIGHT
                    DataSheet.rows.item(row).rowHeight = rows.item(i)(key)
                    
                Case JSON_ENABLED
                    If Not rows.item(i)(JSON_ENABLED) Then
                        DataSheet.Cells.item(row, ini.data.flagColumn).value = FLAG_COMMENT
                    End If
                    
                Case JSON_DATA_ITEM
                    DataSheet.Cells.item(row, ini.data.itemColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_LABEL
                    DataSheet.Cells.item(row, ini.data.labelColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_OUTSIDE_LABEL
                    DataSheet.Cells.item(row, ini.data.xLabelColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_TAIL_LABEL
                    DataSheet.Cells.item(row, ini.data.tailLabelColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_HEAD_LABEL
                    DataSheet.Cells.item(row, ini.data.headLabelColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_TOOLTIP
                    DataSheet.Cells.item(row, ini.data.tooltipColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_RELATED_ITEM
                    DataSheet.Cells.item(row, ini.data.isRelatedToItemColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_STYLE_NAME
                    DataSheet.Cells.item(row, ini.data.styleNameColumn).value = rows.item(i)(key)
                    
                Case JSON_DATA_EXTRA_ATTRIBUTES
                    Set dictionaryObj = rows.item(i)(key)
                    extraAttributes = DictionaryToAttributes(dictionaryObj)
                    DataSheet.Cells.item(row, ini.data.extraAttributesColumn).value = extraAttributes
            End Select
        Next
    Next i
End Sub

Private Function DictionaryToAttributes(ByVal dictionaryObj As Dictionary) As String
    DictionaryToAttributes = vbNullString
    
    Dim key As Variant
    For Each key In dictionaryObj.Keys()
        DictionaryToAttributes = DictionaryToAttributes & " " & key & "=" & AddQuotesConditionally(dictionaryObj.item(key))
    Next
    
    DictionaryToAttributes = Trim$(DictionaryToAttributes)
End Function

Public Sub ClearWorksheetSql(ByRef ini As settings)
    Dim lastColumn As Long
    Dim cellRange As String
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < ini.sql.firstRow Then
        lastRow = ini.sql.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(SqlSheet.name, ini.sql.headingRow)

    ' Remove any existing content
    cellRange = "A" & ini.sql.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    SqlSheet.Range(cellRange).ClearContents
    SqlSheet.rows.UseStandardHeight = True
End Sub

Public Sub ClearWorksheetSvg(ByRef ini As settings)
    Dim lastColumn As Long
    Dim cellRange As String
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With SvgSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < ini.svg.firstRow Then
        lastRow = ini.svg.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(SvgSheet.name, ini.svg.headingRow)

    ' Remove any existing content
    cellRange = "A" & ini.svg.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    SvgSheet.Range(cellRange).ClearContents
    SvgSheet.rows.UseStandardHeight = True
End Sub

Public Sub ClearWorksheetStyles(ByRef ini As settings)
    Dim lastColumn As Long
    Dim cellRange As String
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With StylesSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < ini.styles.firstRow Then
        lastRow = ini.styles.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(StylesSheet.name, ini.sql.headingRow)

    ' Remove any existing content
    cellRange = "A" & ini.sql.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    StylesSheet.Range(cellRange).ClearContents
    StylesSheet.rows.UseStandardHeight = True
End Sub

Public Sub ClearWorksheetData(ByRef ini As settings)
    Dim lastColumn As Long
    Dim cellRange As String
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With DataSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < ini.data.firstRow Then
        lastRow = ini.data.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(DataSheet.name, ini.sql.headingRow)

    ' Remove any existing content
    cellRange = "A" & ini.sql.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    DataSheet.Range(cellRange).ClearContents
    DataSheet.rows.UseStandardHeight = True
End Sub

Private Sub ImportContentSql(ByRef ini As settings, ByRef exchange As ExchangeOptions, ByVal rows As Collection)
    Dim i As Long
    Dim firstRow As Long
    Dim key As Variant
    Dim row As Long
    
    ' First possible row after headings = 2
    '@Ignore AssignmentNotUsed
    firstRow = 2
    
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    Select Case exchange.sql.action
        Case IMPORT_REPLACE
            firstRow = ini.sql.firstRow
            ClearWorksheetSql ini

        Case IMPORT_APPEND
            firstRow = lastRow + 1
    End Select
    
    For i = 1 To rows.count
        Select Case exchange.sql.action
            Case IMPORT_REPLACE
                If rows.item(i).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.item(i)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + i - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + i - 1
        End Select

        SqlSheet.Cells.item(row, 1).EntireRow.ClearContents
        For Each key In rows.item(i).Keys()
            Select Case key
                Case JSON_HIDDEN
                    SqlSheet.rows.item(row).Hidden = rows.item(i)(key)
                        
                Case JSON_HEIGHT
                    SqlSheet.rows.item(row).rowHeight = rows.item(i)(key)
                        
                Case JSON_ENABLED
                    If Not rows.item(i)(JSON_ENABLED) Then
                        SqlSheet.Cells.item(row, ini.sql.flagColumn).value = FLAG_COMMENT
                    End If
                    
                Case JSON_SQL_SQL_STATEMENT
                    SqlSheet.Cells.item(row, ini.sql.sqlStatementColumn).value = rows.item(i)(key)
                    
                Case JSON_SQL_EXCEL_FILE
                    SqlSheet.Cells.item(row, ini.sql.excelFileColumn).value = rows.item(i)(key)
                    
                Case JSON_SQL_STATUS
                    SqlSheet.Cells.item(row, ini.sql.statusColumn).value = rows.item(i)(key)
                    
                Case JSON_SQL_FILTERS
                    Dim filterValues As Collection
                    Set filterValues = rows.item(i)(key)
                    
                    Dim col As Long
                    col = 5     ' Start at column E
                    
                    Dim filter As Variant
                    For Each filter In filterValues
                        SqlSheet.Cells.item(row, col).value = filter
                        col = col + 1
                    Next filter
            End Select
        Next
    Next i
End Sub

Private Sub ImportContentSvg(ByRef ini As settings, ByRef exchange As ExchangeOptions, ByVal rows As Collection)
    Dim i As Long
    Dim firstRow As Long
    Dim key As Variant
    Dim row As Long
    
    ' First possible row after headings = 2
    '@Ignore AssignmentNotUsed
    firstRow = 2
    
    Dim lastRow As Long
    With SvgSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    Select Case exchange.svg.action
        Case IMPORT_REPLACE
            firstRow = ini.svg.firstRow
            ClearWorksheetSvg ini

        Case IMPORT_APPEND
            firstRow = lastRow + 1
    End Select
    
    For i = 1 To rows.count
        Select Case exchange.svg.action
            Case IMPORT_REPLACE
                If rows.item(i).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.item(i)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + i - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + i - 1
        End Select

        SvgSheet.Cells.item(row, 1).EntireRow.ClearContents
        For Each key In rows.item(i).Keys()
            Select Case key
                Case JSON_HIDDEN
                    SvgSheet.rows.item(row).Hidden = rows.item(i)(key)
                        
                Case JSON_HEIGHT
                    SvgSheet.rows.item(row).rowHeight = rows.item(i)(key)
                        
                Case JSON_ENABLED
                    If Not rows.item(i)(JSON_ENABLED) Then
                        SvgSheet.Cells.item(row, ini.svg.flagColumn).value = FLAG_COMMENT
                    End If
                    
                Case JSON_SVG_FIND
                    SvgSheet.Cells.item(row, ini.svg.findColumn).value = rows.item(i)(key)
                    
                Case JSON_SVG_REPLACE
                    SvgSheet.Cells.item(row, ini.svg.replaceColumn).value = rows.item(i)(key)
            End Select
        Next
    Next i
End Sub

Private Function GetStylesAppendRow(ByRef ini As settings) As Long

    Dim row As Long
    With StylesSheet.UsedRange
        row = .Cells.item(.Cells.count).row
    End With

    Do While row > ini.styles.firstRow
        If GetCell(StylesSheet.name, row, ini.styles.nameColumn) <> vbNullString Then
            GetStylesAppendRow = row + 1
            Exit Do
        End If
        row = row - 1
    Loop

End Function

Private Sub ImportContentStyles(ByRef ini As settings, ByRef exchange As ExchangeOptions, ByVal rows As Collection)
    Dim rowIndex As Long
    Dim firstRow As Long
    Dim switchIndex As Long
    Dim key As Variant
    Dim row As Long
    Dim switches As Collection
    Dim dictionaryObj As Dictionary
    Dim format As String

    ' First possible row after headings = 1
    '@Ignore AssignmentNotUsed
    firstRow = 2
    
    Select Case exchange.styles.action
        Case IMPORT_REPLACE
            firstRow = ini.styles.firstRow
            ClearWorksheetStyles ini
        Case IMPORT_APPEND
            firstRow = GetStylesAppendRow(ini)
    End Select
    
    For rowIndex = 1 To rows.count
        Select Case exchange.styles.action
            Case IMPORT_REPLACE
                If rows.item(rowIndex).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.item(rowIndex)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + rowIndex - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + rowIndex - 1
        End Select
            
        StylesSheet.Cells.item(row, 1).EntireRow.ClearContents
        For Each key In rows.item(rowIndex).Keys()
            Select Case key
                Case JSON_HIDDEN
                    StylesSheet.rows.item(row).Hidden = rows.item(rowIndex)(key)
                        
                Case JSON_HEIGHT
                    StylesSheet.rows.item(row).rowHeight = rows.item(rowIndex)(key)
                        
                Case JSON_ENABLED
                    If Not rows.item(rowIndex)(JSON_ENABLED) Then
                        StylesSheet.Cells.item(row, ini.styles.flagColumn).value = FLAG_COMMENT
                    End If
                    
                Case JSON_STYLES_NAME
                    StylesSheet.Cells.item(row, ini.styles.nameColumn).value = rows.item(rowIndex)(key)
                    
                Case JSON_STYLES_FORMAT
                    Set dictionaryObj = rows.item(rowIndex)(key)
                    format = DictionaryToAttributes(dictionaryObj)
                    StylesSheet.Cells.item(row, ini.styles.formatColumn).value = format
                    
                Case JSON_STYLES_TYPE
                    StylesSheet.Cells.item(row, ini.styles.typeColumn).value = rows.item(rowIndex)(key)
                
                Case JSON_STYLES_VIEW_SWITCHES
                    Set switches = rows.item(rowIndex)(JSON_STYLES_VIEW_SWITCHES)
                    For switchIndex = 1 To switches.count
                        StylesSheet.Cells.item(row, (ini.styles.firstYesNoColumn + switchIndex - 1)).value = switches.item(switchIndex)
                    Next switchIndex
            End Select
        Next
    Next rowIndex
End Sub


