Attribute VB_Name = "modExchangeImport"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

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
    currentLanguage = SettingsSheet.Range(SETTINGS_LANGUAGE).Value
    
    ' Import the JSON contents
    If importFile <> vbNullString Then
        returnMessage = ImportDataProcessFile(importFile)
    End If
    
    ' If the import specified a change in language, update localizations
    If currentLanguage <> SettingsSheet.Range(SETTINGS_LANGUAGE).Value Then
        Localize
    End If
    
    ' Update the ribbon
    RefreshRibbon tag:=RIBBON_TAB_GRAPHVIZ

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
    
    Dim jsonObject As Object
    Set jsonObject = TryParseJson(jsonString)
    
    If (jsonObject Is Nothing) Then ' Error was already reported
        Exit Function
    End If
    
    ' Test the metadata to see if we should bother proceeding
    If jsonObject.Exists(JSON_SECTION_METADATA) Then
        Dim metadata As Dictionary
        Set metadata = jsonObject.Item(JSON_SECTION_METADATA)
    
        Dim name As String
        name = metadata.Item("name")
               
        If name <> "E2GXF" Then
            ImportDataProcessFile = "The JSON in this exchange file is not recognized. " & _
                    vbNewLine & vbNewLine & _
                    "Found:    name=""" & name & """" & vbNewLine & _
                    "Expected: name=""E2GXF"""
        End If
    End If
    
    PerformImports jsonObject
    RefreshRibbon tag:="*Tab"

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
        ImportMetadata dictionaryObj.Item(JSON_SECTION_METADATA)
        If Application.Calculation = xlManual Then
            SettingsSheet.Calculate
        End If
    End If
    
    ' Import worksheet layouts
    If dictionaryObj.Exists(JSON_SECTION_LAYOUTS) Then
        ImportLayouts dictionaryObj.Item(JSON_SECTION_LAYOUTS), exchange
        If Application.Calculation = xlManual Then
            SettingsSheet.Calculate
        End If
    End If
    
    ' Import settings
    If dictionaryObj.Exists(JSON_SECTION_SETTINGS) Then
        ImportSettings dictionaryObj.Item(JSON_SECTION_SETTINGS), exchange
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

        ImportContent dictionaryObj.Item(JSON_SECTION_CONTENT), ini, exchange
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
                MsgBox GetMessage("msgboxUnexpectedMetaData") & vbNewLine & vbNewLine & key & "=" & dictionaryObj.Item(key), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
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
                    ImportContentData ini, exchange, dictionaryObj.Item(worksheetName)
                End If
            
            Case WORKSHEET_SQL
                If exchange.sql.include Then
                    ImportContentSql ini, exchange, dictionaryObj.Item(worksheetName)
                End If
            
            Case WORKSHEET_SVG
                If exchange.svg.include Then
                    ImportContentSvg ini, exchange, dictionaryObj.Item(worksheetName)
                End If
            
            Case WORKSHEET_STYLES
                If exchange.styles.include Then
                    ImportContentStyles ini, exchange, dictionaryObj.Item(worksheetName)
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
                 ImportSettingsData dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_GRAPH
                ImportSettingsGraph dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_SETTINGS
                ImportSettingsSettings dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_SOURCE
                ImportSettingsSource dictionaryObj.Item(worksheetName)
                    
            Case WORKSHEET_SQL
                ImportSettingsSql dictionaryObj.Item(worksheetName)
                    
            Case "extensions"
                ImportSettingsExtensions dictionaryObj.Item(worksheetName)
        End Select
    Next
    
End Sub

Private Sub ImportSettingsGraph(ByVal dictionaryObj As Dictionary)
    If dictionaryObj.Exists(JSON_ZOOM) Then
        SetZoom GraphSheet.name, dictionaryObj.Item(JSON_ZOOM)
    End If
End Sub

Private Sub ImportSettingsSettings(ByVal dictionaryObj As Dictionary)
    RestoreSetting SETTINGS_GV_PATH, dictionaryObj.Item(JSON_SETTINGS_GV_PATH)
    RestoreSetting SETTINGS_IMAGE_PATH, dictionaryObj.Item(JSON_SETTINGS_IMAGE_PATH)
    RestoreSetting SETTINGS_GRAPH_OPTIONS, dictionaryObj.Item(JSON_SETTINGS_GRAPH_OPTIONS)
    RestoreSetting SETTINGS_MAX_SECONDS, dictionaryObj.Item(JSON_SETTINGS_MAX_SECONDS)
    RestoreSetting SETTINGS_PICTURE_NAME, dictionaryObj.Item(JSON_SETTINGS_PICTURE_NAME)
    RestoreSetting SETTINGS_COMMAND_LINE_PARAMETERS, dictionaryObj.Item(JSON_SETTINGS_COMMAND_LINE_PARAMETERS)
End Sub

Private Sub ImportSettingsSource(ByVal dictionaryObj As Dictionary)
    Dim buttons As Collection
    Dim button As Dictionary
    Dim i As Long

    Set buttons = dictionaryObj.Item(JSON_SETTINGS_BUTTONS)

    RestoreSetting SETTINGS_SOURCE_INDENT, dictionaryObj.Item(JSON_SETTINGS_INDENT)
    
    For i = 1 To 6
        Set button = buttons.Item(i)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_TEXT, button.Item(JSON_SETTINGS_BUTTON_TEXT)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_URL, button.Item(JSON_SETTINGS_URL)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_SCREENTIP, button.Item(JSON_SETTINGS_SCREEN_TIP)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_SUPERTIP, button.Item(JSON_SETTINGS_SUPER_TIP)
        RestoreSetting BUTTON_PREFIX_SOURCE_WEB & i & BUTTON_SUFFIX_VISIBLE, button.Item(JSON_SETTINGS_VISIBLE)
    Next i
End Sub

Private Sub ImportSettingsSql(ByVal dictionaryObj As Dictionary)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_STYLE_NAME)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_ATTRIBUTES)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_TOOLTIP)
    
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_STYLE_NAME)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_ATTRIBUTES)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_TOOLTIP)
    
    RestoreSetting SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_SPLIT_LENGTH)
    RestoreSetting SETTINGS_SQL_FIELD_NAME_LINE_ENDING, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_LINE_ENDING)

    RestoreSetting SETTINGS_SQL_COUNT_PLACEHOLDER_CLUSTER, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_CLUSTER_PLACEHOLDER)
    RestoreSetting SETTINGS_SQL_COUNT_PLACEHOLDER_SUBCLUSTER, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_SUBCLUSTER_PLACEHOLDER)
    RestoreSetting SETTINGS_SQL_COUNT_PLACEHOLDER_RECORDSET, dictionaryObj.Item(JSON_SETTINGS_SQL_FIELD_NAME_RECORDSET_PLACEHOLDER)
    
    RestoreSetting SETTINGS_SQL_COL_FILTER, dictionaryObj.Item(JSON_SETTINGS_SQL_FILTER_COLUMN)
    RestoreSetting SETTINGS_SQL_FILTER_VALUE, dictionaryObj.Item(JSON_SETTINGS_SQL_FILTER_VALUE)
End Sub

Private Sub ImportSettingsExtensions(ByVal dictionaryObj As Dictionary)
    Dim buttons As Collection
    Dim button As Dictionary
    Dim i As Long

    RestoreSetting SETTINGS_EXT_TAB_NAME, dictionaryObj.Item(JSON_SETTINGS_EXT_TAB_NAME)
    
    Dim group As Dictionary
    
    If dictionaryObj.Exists(JSON_SETTINGS_EXT_TAB_GROUP_NAME_WEB) Then
        Set group = dictionaryObj.Item(JSON_SETTINGS_EXT_TAB_GROUP_NAME_WEB)
        Set buttons = group.Item(JSON_SETTINGS_BUTTONS)
        RestoreSetting SETTINGS_EXT_TAB_GROUP_NAME_WEB, group.Item(JSON_SETTINGS_GROUP_NAME)
    
        For i = 1 To 6
            Set button = buttons.Item(i)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_TEXT, button.Item(JSON_SETTINGS_BUTTON_TEXT)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_URL, button.Item(JSON_SETTINGS_URL)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_SCREENTIP, button.Item(JSON_SETTINGS_SCREEN_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_SUPERTIP, button.Item(JSON_SETTINGS_SUPER_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_WEB & i & BUTTON_SUFFIX_VISIBLE, button.Item(JSON_SETTINGS_VISIBLE)
        Next i
    End If
    
     If dictionaryObj.Exists(JSON_SETTINGS_EXT_TAB_GROUP_NAME_CODE) Then
        Set group = dictionaryObj.Item(JSON_SETTINGS_EXT_TAB_GROUP_NAME_CODE)
        Set buttons = group.Item(JSON_SETTINGS_BUTTONS)
        RestoreSetting SETTINGS_EXT_TAB_GROUP_NAME_CODE, group.Item(JSON_SETTINGS_GROUP_NAME)
    
        For i = 1 To 6
            Set button = buttons.Item(i)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_TEXT, button.Item(JSON_SETTINGS_BUTTON_TEXT)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SUB, button.Item(JSON_SETTINGS_SUB)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SCREENTIP, button.Item(JSON_SETTINGS_SCREEN_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_SUPERTIP, button.Item(JSON_SETTINGS_SUPER_TIP)
            RestoreSetting BUTTON_PREFIX_EXT_CODE & i & BUTTON_SUFFIX_VISIBLE, button.Item(JSON_SETTINGS_VISIBLE)
        Next i
    End If
End Sub

Private Sub ImportSettingsData(ByVal dictionaryObj As Dictionary)
    Dim section As Dictionary
    Dim subSection As Dictionary
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_GRAPH_TO_WORKSHEET) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_GRAPH_TO_WORKSHEET)
        RestoreSetting SETTINGS_RUN_MODE, section.Item(JSON_SETTINGS_RUN_MODE)
        RestoreSetting SETTINGS_IMAGE_TYPE, section.Item(JSON_SETTINGS_IMAGE_TYPE)
        RestoreSetting SETTINGS_IMAGE_WORKSHEET, section.Item(JSON_SETTINGS_IMAGE_WORKSHEET)
        
        If section.Item(JSON_SETTINGS_SCALE_IMAGE) = vbNullString Then
            RestoreSetting SETTINGS_SCALE_IMAGE, "100"  ' For backward compatability with older export files
        Else
            RestoreSetting SETTINGS_SCALE_IMAGE, section.Item(JSON_SETTINGS_SCALE_IMAGE)
        End If
    End If
                    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_GRAPH_TO_FILE) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_GRAPH_TO_FILE)
        RestoreSetting SETTINGS_OUTPUT_DIRECTORY, section.Item(JSON_SETTINGS_DIRECTORY)
        RestoreSetting SETTINGS_FILE_NAME, section.Item(JSON_SETTINGS_FILE_NAME_PREFIX)
        RestoreSetting SETTINGS_FILE_FORMAT, section.Item(JSON_SETTINGS_IMAGE_TYPE)
        RestoreSetting SETTINGS_APPEND_OPTIONS, BooleanToYesNo(section.Item(JSON_SETTINGS_APPEND_OPTIONS))
        RestoreSetting SETTINGS_APPEND_TIMESTAMP, BooleanToYesNo(section.Item(JSON_SETTINGS_APPEND_TIME_STAMP))
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_LAYOUT) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_LAYOUT)
        RestoreSetting SETTINGS_GRAPHVIZ_ENGINE, section.Item(JSON_SETTINGS_ENGINE)
        
        ' Maintain backward compatibility. Direction and rankdir were consolidated into rankdir to improve performance
        ' when the port to Apple Mac was performed. Old versions of the spreadsheet will still export direction and not rankdir
        ' so the value will be derived from direction to set the rankdir cell.
        Select Case LCase$(section.Item(JSON_SETTINGS_DIRECTION))
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
        RestoreSetting SETTINGS_SPLINES, section.Item(JSON_SETTINGS_SPLINES)
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_OPTIONS) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_OPTIONS)
        If section.Exists(JSON_SETTINGS_SECTION_GRAPH) Then
            Set subSection = section.Item(JSON_SETTINGS_SECTION_GRAPH)
            RestoreSetting SETTINGS_GRAPH_CENTER, BooleanToYesNo(subSection.Item(JSON_SETTINGS_CENTER))
            RestoreSetting SETTINGS_GRAPH_CLUSTER_RANK, subSection.Item(JSON_SETTINGS_CLUSTER_RANK)
            RestoreSetting SETTINGS_GRAPH_COMPOUND, BooleanToYesNo(subSection.Item(JSON_SETTINGS_COMPOUND))
            RestoreSetting SETTINGS_GRAPH_DIM, subSection.Item(JSON_SETTINGS_DIM)
            RestoreSetting SETTINGS_GRAPH_DIMEN, subSection.Item(JSON_SETTINGS_DIMEN)
            RestoreSetting SETTINGS_GRAPH_FORCE_LABELS, BooleanToYesNo(subSection.Item(JSON_SETTINGS_FORCE_LABELS))
            RestoreSetting SETTINGS_GRAPH_MODE, subSection.Item(JSON_SETTINGS_MODE)
            RestoreSetting SETTINGS_GRAPH_MODEL, subSection.Item(JSON_SETTINGS_MODEL)
            RestoreSetting SETTINGS_GRAPH_NEWRANK, BooleanToYesNo(subSection.Item(JSON_SETTINGS_NEWRANK))
            RestoreSetting SETTINGS_GRAPH_ORDERING, subSection.Item(JSON_SETTINGS_ORDERING)
            RestoreSetting SETTINGS_GRAPH_ORIENTATION, BooleanToYesNo(subSection.Item(JSON_SETTINGS_ORIENTATION))
            RestoreSetting SETTINGS_GRAPH_OUTPUT_ORDER, subSection.Item(JSON_SETTINGS_OUTPUT_ORDER)
            RestoreSetting SETTINGS_GRAPH_OVERLAP, subSection.Item(JSON_SETTINGS_OVERLAP)
            RestoreSetting SETTINGS_GRAPH_SMOOTHING, subSection.Item(JSON_SETTINGS_SMOOTHING)
            RestoreSetting SETTINGS_GRAPH_TRANSPARENT, BooleanToYesNo(subSection.Item(JSON_SETTINGS_TRANSPARENT_BACKGROUND))
            RestoreSetting SETTINGS_GRAPH_TYPE, subSection.Item(JSON_SETTINGS_GRAPH_TYPE)
        End If
    
        If section.Exists(JSON_SETTINGS_SECTION_NODES) Then
            Set subSection = section.Item(JSON_SETTINGS_SECTION_NODES)
            RestoreSetting SETTINGS_NODES_WITHOUT_RELATIONSHIPS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_ORPHAN_NODES))
            RestoreSetting SETTINGS_NODE_LABELS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_NODE_LABELS))
            RestoreSetting SETTINGS_NODE_XLABELS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_NODE_XLABELS))
            RestoreSetting SETTINGS_BLANK_NODE_LABELS, subSection.Item(JSON_SETTINGS_BLANK_NODE_LABELS)
        End If
    
        If section.Exists(JSON_SETTINGS_SECTION_EDGES) Then
            Set subSection = section.Item(JSON_SETTINGS_SECTION_EDGES)
            RestoreSetting SETTINGS_GRAPH_STRICT, BooleanToYesNo(subSection.Item(JSON_SETTINGS_ADD_STRICT))
            RestoreSetting SETTINGS_GRAPH_CONCENTRATE, BooleanToYesNo(subSection.Item(JSON_SETTINGS_CONCENTRATE))
            RestoreSetting SETTINGS_RELATIONSHIPS_WITHOUT_NODES, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_ORPHAN_EDGES))
            RestoreSetting SETTINGS_EDGE_HEAD_LABELS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_EDGE_HEAD_LABELS))
            RestoreSetting SETTINGS_EDGE_LABELS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_EDGE_LABELS))
            RestoreSetting SETTINGS_EDGE_TAIL_LABELS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_EDGE_TAIL_LABELS))
            RestoreSetting SETTINGS_EDGE_XLABELS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_EDGE_XLABELS))
            RestoreSetting SETTINGS_EDGE_PORTS, BooleanToIncludeExclude(subSection.Item(JSON_SETTINGS_INCLUDE_EDGE_PORTS))
            RestoreSetting SETTINGS_BLANK_EDGE_LABELS, subSection.Item(JSON_SETTINGS_BLANK_EDGE_LABELS)
        End If
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_STYLES) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_STYLES)
        RestoreSetting SETTINGS_STYLES_COL_SHOW_STYLE, section.Item(JSON_SETTINGS_SELECTED_VIEW_COLUMN)
        RestoreSetting SETTINGS_INCLUDE_STYLE_FORMAT, BooleanToIncludeExclude(section.Item(JSON_SETTINGS_INCLUDE_STYLE_FORMAT))
        RestoreSetting SETTINGS_INCLUDE_EXTRA_ATTRIBUTES, BooleanToIncludeExclude(section.Item(JSON_SETTINGS_INCLUDE_EXTRA_ATTRIBUTES))
        
        ' Old exports do not have this value
        If section.Exists(JSON_SETTINGS_STYLES_SUFFIX_OPEN) Then
            RestoreSetting SETTINGS_STYLES_SUFFIX_OPEN, section.Item(JSON_SETTINGS_STYLES_SUFFIX_OPEN)
        End If
        
        ' Old exports do not have this value
        If section.Exists(SETTINGS_STYLES_SUFFIX_CLOSE) Then
            RestoreSetting SETTINGS_STYLES_SUFFIX_CLOSE, section.Item(JSON_SETTINGS_STYLES_SUFFIX_CLOSE)
        End If
    End If
    
    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_DEBUG) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_DEBUG)
        RestoreSetting SETTINGS_DEBUG, section.Item(JSON_SETTINGS_DEBUG_SWITCH)
        RestoreSetting SETTINGS_FILE_DISPOSITION, section.Item(JSON_SETTINGS_FILE_DISPOSITION)
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_COLUMNS) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_COLUMNS)
        RestoreSetting SETTINGS_DATA_SHOW_COMMENT, LCase$(section.Item(JSON_STYLES_FLAG))
        RestoreSetting SETTINGS_DATA_SHOW_ITEM, LCase$(section.Item(JSON_DATA_ITEM))
        RestoreSetting SETTINGS_DATA_SHOW_LABEL, LCase$(section.Item(JSON_DATA_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_OUTSIDE_LABEL, LCase$(section.Item(JSON_DATA_OUTSIDE_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_TAIL_LABEL, LCase$(section.Item(JSON_DATA_TAIL_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_HEAD_LABEL, LCase$(section.Item(JSON_DATA_HEAD_LABEL))
        RestoreSetting SETTINGS_DATA_SHOW_IS_RELATED_TO_ITEM, LCase$(section.Item(JSON_DATA_RELATED_ITEM))
        RestoreSetting SETTINGS_DATA_SHOW_STYLE, LCase$(section.Item(JSON_DATA_STYLE_NAME))
        RestoreSetting SETTINGS_DATA_SHOW_EXTRA_STYLE_ATTRIBUTES, LCase$(section.Item(JSON_DATA_EXTRA_ATTRIBUTES))
        RestoreSetting SETTINGS_DATA_SHOW_MESSAGES, LCase$(section.Item(JSON_DATA_MESSAGE))
    End If

    If dictionaryObj.Exists(JSON_SETTINGS_SECTION_LANGUAGE) Then
        Set section = dictionaryObj.Item(JSON_SETTINGS_SECTION_LANGUAGE)
        RestoreSetting SETTINGS_LANGUAGE, section.Item(JSON_SETTINGS_LANGUAGE)
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
    SettingsSheet.Range(cellName).Value = cellValue
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
        ImportLayoutsRowHeights worksheetName, dictionaryObj.Item(worksheetName)

        ' Column layouts
        Select Case worksheetName
            Case WORKSHEET_DATA
                ImportLayoutsData dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_STYLES
                ImportLayoutsStyles dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_SQL
                ImportLayoutsSql dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_SVG
                ImportLayoutsSvg dictionaryObj.Item(worksheetName)
            
            Case WORKSHEET_SOURCE
                ImportLayoutsSource dictionaryObj.Item(worksheetName)
        End Select
    Next
End Sub

Private Sub ImportLayoutsData(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetData
    
    Dim columns As Collection
    Set columns = dictionaryObj.Item(JSON_COLUMNS)
    
    Dim data As dataWorksheet
    data = GetSettingsForDataWorksheet(GetDataWorksheetName())

    Dim i As Long
    
    For i = 1 To columns.Count
        Select Case columns.Item(i)(JSON_ID)
            Case JSON_DATA_FLAG
                DataSheet.Cells.Item(data.headingRow, data.flagColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.flagColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.flagColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.flagColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
           
            Case JSON_DATA_ITEM
                DataSheet.Cells.Item(data.headingRow, data.itemColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.itemColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.itemColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.itemColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
           
            Case JSON_DATA_LABEL
                DataSheet.Cells.Item(data.headingRow, data.labelColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.labelColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.labelColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.labelColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_OUTSIDE_LABEL
                DataSheet.Cells.Item(data.headingRow, data.xLabelColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.xLabelColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.xLabelColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.xLabelColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_TAIL_LABEL
                DataSheet.Cells.Item(data.headingRow, data.tailLabelColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.tailLabelColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.tailLabelColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.tailLabelColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_HEAD_LABEL
                DataSheet.Cells.Item(data.headingRow, data.headLabelColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.headLabelColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.headLabelColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.headLabelColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_TOOLTIP
                DataSheet.Cells.Item(data.headingRow, data.tooltipColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.tooltipColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.tooltipColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.tooltipColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_RELATED_ITEM
                DataSheet.Cells.Item(data.headingRow, data.isRelatedToItemColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.isRelatedToItemColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.isRelatedToItemColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.isRelatedToItemColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_STYLE_NAME
                DataSheet.Cells.Item(data.headingRow, data.styleNameColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.styleNameColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.styleNameColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.styleNameColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_EXTRA_ATTRIBUTES
                DataSheet.Cells.Item(data.headingRow, data.extraAttributesColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.extraAttributesColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.extraAttributesColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.extraAttributesColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_MESSAGE
                DataSheet.Cells.Item(data.headingRow, data.errorMessageColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.errorMessageColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.errorMessageColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.errorMessageColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_DATA_GRAPH_DISPLAY_COLUMN
                DataSheet.Cells.Item(data.headingRow, data.graphDisplayColumn).Value = columns.Item(i)(JSON_HEADING)
                DataSheet.columns.Item(data.graphDisplayColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                DataSheet.columns.Item(data.graphDisplayColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                DataSheet.columns.Item(data.graphDisplayColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsStyles(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetStyles
    
    Dim columns As Collection
    Set columns = dictionaryObj.Item(JSON_COLUMNS)
    
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    Dim i As Long
    
    Dim offset As Long
    offset = -1
    
    For i = 1 To columns.Count
        Select Case columns.Item(i)(JSON_ID)
            Case JSON_STYLES_FLAG
                StylesSheet.Cells.Item(styles.headingRow, styles.flagColumn).Value = columns.Item(i)(JSON_HEADING)
                StylesSheet.columns.Item(styles.flagColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                StylesSheet.columns.Item(styles.flagColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                StylesSheet.columns.Item(styles.flagColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
           
            Case JSON_STYLES_NAME
                StylesSheet.Cells.Item(styles.headingRow, styles.nameColumn).Value = columns.Item(i)(JSON_HEADING)
                StylesSheet.columns.Item(styles.nameColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                StylesSheet.columns.Item(styles.nameColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                StylesSheet.columns.Item(styles.nameColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
           
            Case JSON_STYLES_FORMAT
                StylesSheet.Cells.Item(styles.headingRow, styles.formatColumn).Value = columns.Item(i)(JSON_HEADING)
                StylesSheet.columns.Item(styles.formatColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                StylesSheet.columns.Item(styles.formatColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                StylesSheet.columns.Item(styles.formatColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_STYLES_TYPE
                StylesSheet.Cells.Item(styles.headingRow, styles.typeColumn).Value = columns.Item(i)(JSON_HEADING)
                StylesSheet.columns.Item(styles.typeColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                StylesSheet.columns.Item(styles.typeColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                StylesSheet.columns.Item(styles.typeColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case Else
                offset = offset + 1
                StylesSheet.Cells.Item(styles.headingRow, styles.firstYesNoColumn + offset).Value = columns.Item(i)(JSON_HEADING)
                StylesSheet.columns.Item(styles.firstYesNoColumn + offset).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                StylesSheet.columns.Item(styles.firstYesNoColumn + offset).Hidden = columns.Item(i)(JSON_HIDDEN)
                StylesSheet.columns.Item(styles.firstYesNoColumn + offset).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsSql(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetSql
    
    Dim columns As Collection
    Set columns = dictionaryObj.Item(JSON_COLUMNS)
    
    Dim sql As sqlWorksheet
    sql = GetSettingsForSqlWorksheet()
    
    Dim i As Long
    
    For i = 1 To columns.Count
        Select Case columns.Item(i)(JSON_ID)
            Case JSON_LAYOUT_SQL_FLAG
                SqlSheet.Cells.Item(sql.headingRow, sql.flagColumn).Value = columns.Item(i)(JSON_HEADING)
                SqlSheet.columns.Item(sql.flagColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SqlSheet.columns.Item(sql.flagColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SqlSheet.columns.Item(sql.flagColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SQL_SQL_STATEMENT
                SqlSheet.Cells.Item(sql.headingRow, sql.sqlStatementColumn).Value = columns.Item(i)(JSON_HEADING)
                SqlSheet.columns.Item(sql.sqlStatementColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SqlSheet.columns.Item(sql.sqlStatementColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SqlSheet.columns.Item(sql.sqlStatementColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SQL_EXCEL_FILE
                SqlSheet.Cells.Item(sql.headingRow, sql.excelFileColumn).Value = columns.Item(i)(JSON_HEADING)
                SqlSheet.columns.Item(sql.excelFileColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SqlSheet.columns.Item(sql.excelFileColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SqlSheet.columns.Item(sql.excelFileColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SQL_STATUS
                SqlSheet.Cells.Item(sql.headingRow, sql.statusColumn).Value = columns.Item(i)(JSON_HEADING)
                SqlSheet.columns.Item(sql.statusColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SqlSheet.columns.Item(sql.statusColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SqlSheet.columns.Item(sql.statusColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsSvg(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetSvg
    
    Dim columns As Collection
    Set columns = dictionaryObj.Item(JSON_COLUMNS)
    
    Dim svg As svgWorksheet
    svg = GetSettingsForSvgWorksheet()
    
    Dim i As Long
    
    For i = 1 To columns.Count
        Select Case columns.Item(i)(JSON_ID)
            Case JSON_LAYOUT_SVG_FLAG
                SvgSheet.Cells.Item(svg.headingRow, svg.flagColumn).Value = columns.Item(i)(JSON_HEADING)
                SvgSheet.columns.Item(svg.flagColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SvgSheet.columns.Item(svg.flagColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SvgSheet.columns.Item(svg.flagColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SVG_FIND
                SvgSheet.Cells.Item(svg.headingRow, svg.findColumn).Value = columns.Item(i)(JSON_HEADING)
                SvgSheet.columns.Item(svg.findColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SvgSheet.columns.Item(svg.findColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SvgSheet.columns.Item(svg.findColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
            
            Case JSON_LAYOUT_SVG_REPLACE
                SvgSheet.Cells.Item(svg.headingRow, svg.replaceColumn).Value = columns.Item(i)(JSON_HEADING)
                SvgSheet.columns.Item(svg.replaceColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SvgSheet.columns.Item(svg.replaceColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SvgSheet.columns.Item(svg.replaceColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsSource(ByVal dictionaryObj As Dictionary)
    
    LocalizeWorksheetSource
    
    Dim columns As Collection
    Set columns = dictionaryObj.Item(JSON_COLUMNS)
    
    Dim source As sourceWorksheet
    source = GetSettingsForSourceWorksheet()
    
    Dim i As Long
    For i = 1 To columns.Count
        Select Case columns.Item(i)(JSON_ID)
            Case JSON_SOURCE_LINE_NUMBER
                SourceSheet.Cells.Item(source.headingRow, source.lineNumberColumn).Value = columns.Item(i)(JSON_HEADING)
                SourceSheet.columns.Item(source.lineNumberColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SourceSheet.columns.Item(source.lineNumberColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SourceSheet.columns.Item(source.lineNumberColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
           
            Case JSON_SOURCE_SOURCE
                SourceSheet.Cells.Item(source.headingRow, source.sourceColumn).Value = columns.Item(i)(JSON_HEADING)
                SourceSheet.columns.Item(source.sourceColumn).ColumnWidth = columns.Item(i)(JSON_WIDTH)
                SourceSheet.columns.Item(source.sourceColumn).Hidden = columns.Item(i)(JSON_HIDDEN)
                SourceSheet.columns.Item(source.sourceColumn).WrapText = columns.Item(i)(JSON_WRAP_TEXT)
        End Select
    Next
End Sub

Private Sub ImportLayoutsRowHeights(ByRef worksheetName As String, ByVal dictionaryObj As Dictionary)
    Dim rows As Collection
    Dim row As Dictionary
    
    Set rows = dictionaryObj.Item(JSON_ROWS)
    Dim i As Long

    ' Set the row heights
    For i = 1 To rows.Count
        Set row = rows.Item(i)
        ActiveWorkbook.Sheets.[_Default](worksheetName).rows(row.Item(JSON_ROW)).RowHeight = row.Item(JSON_HEIGHT)
        ActiveWorkbook.Sheets.[_Default](worksheetName).rows(row.Item(JSON_ROW)).Hidden = row.Item(JSON_HIDDEN)
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
        GetImportFilename = Trim$(fileDialogHandle.SelectedItems.Item(1))
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
        lastRow = .Cells.Item(.Cells.Count).row
    End With

    Select Case exchange.data.action
        Case IMPORT_REPLACE
            firstRow = ini.data.firstRow
            ClearWorksheetData ini

        Case IMPORT_APPEND
            firstRow = lastRow + 1
    End Select
    
    ' Loop through all the objects in collection
    For i = 1 To rows.Count
        Select Case exchange.data.action
            Case IMPORT_REPLACE
                If rows.Item(i).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.Item(i)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + i - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + i - 1
        End Select
        
        DataSheet.Cells.Item(row, 1).EntireRow.ClearContents

        For Each key In rows.Item(i).Keys()
            Select Case key
                Case JSON_HIDDEN
                    DataSheet.rows.Item(row).Hidden = rows.Item(i)(key)
                    
                Case JSON_HEIGHT
                    DataSheet.rows.Item(row).RowHeight = rows.Item(i)(key)
                    
                Case JSON_ENABLED
                    If Not rows.Item(i)(JSON_ENABLED) Then
                        DataSheet.Cells.Item(row, ini.data.flagColumn).Value = FLAG_COMMENT
                    End If
                    
                Case JSON_DATA_ITEM
                    DataSheet.Cells.Item(row, ini.data.itemColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_LABEL
                    DataSheet.Cells.Item(row, ini.data.labelColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_OUTSIDE_LABEL
                    DataSheet.Cells.Item(row, ini.data.xLabelColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_TAIL_LABEL
                    DataSheet.Cells.Item(row, ini.data.tailLabelColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_HEAD_LABEL
                    DataSheet.Cells.Item(row, ini.data.headLabelColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_TOOLTIP
                    DataSheet.Cells.Item(row, ini.data.tooltipColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_RELATED_ITEM
                    DataSheet.Cells.Item(row, ini.data.isRelatedToItemColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_STYLE_NAME
                    DataSheet.Cells.Item(row, ini.data.styleNameColumn).Value = rows.Item(i)(key)
                    
                Case JSON_DATA_EXTRA_ATTRIBUTES
                    Set dictionaryObj = rows.Item(i)(key)
                    extraAttributes = DictionaryToAttributes(dictionaryObj)
                    DataSheet.Cells.Item(row, ini.data.extraAttributesColumn).Value = extraAttributes
            End Select
        Next
    Next i
End Sub

Private Function DictionaryToAttributes(ByVal dictionaryObj As Dictionary) As String
    DictionaryToAttributes = vbNullString
    
    Dim key As Variant
    For Each key In dictionaryObj.Keys()
        DictionaryToAttributes = DictionaryToAttributes & " " & key & "=" & AddQuotes(dictionaryObj.Item(key))
    Next
    
    DictionaryToAttributes = Trim$(DictionaryToAttributes)
End Function

Public Sub ClearWorksheetSql(ByRef ini As settings)
    Dim lastColumn As Long
    Dim cellRange As String
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.Item(.Cells.Count).row
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
        lastRow = .Cells.Item(.Cells.Count).row
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
        lastRow = .Cells.Item(.Cells.Count).row
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
        lastRow = .Cells.Item(.Cells.Count).row
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
        lastRow = .Cells.Item(.Cells.Count).row
    End With
    
    Select Case exchange.sql.action
        Case IMPORT_REPLACE
            firstRow = ini.sql.firstRow
            ClearWorksheetSql ini

        Case IMPORT_APPEND
            firstRow = lastRow + 1
    End Select
    
    For i = 1 To rows.Count
        Select Case exchange.sql.action
            Case IMPORT_REPLACE
                If rows.Item(i).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.Item(i)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + i - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + i - 1
        End Select

        SqlSheet.Cells.Item(row, 1).EntireRow.ClearContents
        For Each key In rows.Item(i).Keys()
            Select Case key
                Case JSON_HIDDEN
                    SqlSheet.rows.Item(row).Hidden = rows.Item(i)(key)
                        
                Case JSON_HEIGHT
                    SqlSheet.rows.Item(row).RowHeight = rows.Item(i)(key)
                        
                Case JSON_ENABLED
                    If Not rows.Item(i)(JSON_ENABLED) Then
                        SqlSheet.Cells.Item(row, ini.sql.flagColumn).Value = FLAG_COMMENT
                    End If
                    
                Case JSON_SQL_SQL_STATEMENT
                    SqlSheet.Cells.Item(row, ini.sql.sqlStatementColumn).Value = rows.Item(i)(key)
                    
                Case JSON_SQL_EXCEL_FILE
                    SqlSheet.Cells.Item(row, ini.sql.excelFileColumn).Value = rows.Item(i)(key)
                    
                Case JSON_SQL_STATUS
                    SqlSheet.Cells.Item(row, ini.sql.statusColumn).Value = rows.Item(i)(key)
                    
                Case JSON_SQL_FILTERS
                    Dim filterValues As Collection
                    Set filterValues = rows.Item(i)(key)
                    
                    Dim col As Long
                    col = 5     ' Start at column E
                    
                    Dim filter As Variant
                    For Each filter In filterValues
                        SqlSheet.Cells.Item(row, col).Value = filter
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
        lastRow = .Cells.Item(.Cells.Count).row
    End With
    
    Select Case exchange.svg.action
        Case IMPORT_REPLACE
            firstRow = ini.svg.firstRow
            ClearWorksheetSvg ini

        Case IMPORT_APPEND
            firstRow = lastRow + 1
    End Select
    
    For i = 1 To rows.Count
        Select Case exchange.svg.action
            Case IMPORT_REPLACE
                If rows.Item(i).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.Item(i)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + i - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + i - 1
        End Select

        SvgSheet.Cells.Item(row, 1).EntireRow.ClearContents
        For Each key In rows.Item(i).Keys()
            Select Case key
                Case JSON_HIDDEN
                    SvgSheet.rows.Item(row).Hidden = rows.Item(i)(key)
                        
                Case JSON_HEIGHT
                    SvgSheet.rows.Item(row).RowHeight = rows.Item(i)(key)
                        
                Case JSON_ENABLED
                    If Not rows.Item(i)(JSON_ENABLED) Then
                        SvgSheet.Cells.Item(row, ini.svg.flagColumn).Value = FLAG_COMMENT
                    End If
                    
                Case JSON_SVG_FIND
                    SvgSheet.Cells.Item(row, ini.svg.findColumn).Value = rows.Item(i)(key)
                    
                Case JSON_SVG_REPLACE
                    SvgSheet.Cells.Item(row, ini.svg.replaceColumn).Value = rows.Item(i)(key)
            End Select
        Next
    Next i
End Sub

Private Function GetStylesAppendRow(ByRef ini As settings) As Long

    Dim row As Long
    With StylesSheet.UsedRange
        row = .Cells.Item(.Cells.Count).row
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
    
    For rowIndex = 1 To rows.Count
        Select Case exchange.styles.action
            Case IMPORT_REPLACE
                If rows.Item(rowIndex).Exists(JSON_ROW) Then   ' If the row number is provided, use it
                    row = rows.Item(rowIndex)(JSON_ROW)
                Else                                    ' calculate a row number by array index and first row setting
                    row = firstRow + rowIndex - 1
                End If
            Case IMPORT_APPEND
                row = firstRow + rowIndex - 1
        End Select
            
        StylesSheet.Cells.Item(row, 1).EntireRow.ClearContents
        For Each key In rows.Item(rowIndex).Keys()
            Select Case key
                Case JSON_HIDDEN
                    StylesSheet.rows.Item(row).Hidden = rows.Item(rowIndex)(key)
                        
                Case JSON_HEIGHT
                    StylesSheet.rows.Item(row).RowHeight = rows.Item(rowIndex)(key)
                        
                Case JSON_ENABLED
                    If Not rows.Item(rowIndex)(JSON_ENABLED) Then
                        StylesSheet.Cells.Item(row, ini.styles.flagColumn).Value = FLAG_COMMENT
                    End If
                    
                Case JSON_STYLES_NAME
                    StylesSheet.Cells.Item(row, ini.styles.nameColumn).Value = rows.Item(rowIndex)(key)
                    
                Case JSON_STYLES_FORMAT
                    Set dictionaryObj = rows.Item(rowIndex)(key)
                    format = DictionaryToAttributes(dictionaryObj)
                    StylesSheet.Cells.Item(row, ini.styles.formatColumn).Value = format
                    
                Case JSON_STYLES_TYPE
                    StylesSheet.Cells.Item(row, ini.styles.typeColumn).Value = rows.Item(rowIndex)(key)
                
                Case JSON_STYLES_VIEW_SWITCHES
                    Set switches = rows.Item(rowIndex)(JSON_STYLES_VIEW_SWITCHES)
                    For switchIndex = 1 To switches.Count
                        StylesSheet.Cells.Item(row, (ini.styles.firstYesNoColumn + switchIndex - 1)).Value = switches.Item(switchIndex)
                    Next switchIndex
            End Select
        Next
    Next rowIndex
End Sub


