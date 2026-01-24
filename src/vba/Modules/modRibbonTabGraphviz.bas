Attribute VB_Name = "modRibbonTabGraphviz"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ParameterNotUsed, UseMeaningfulName, UnassignedVariableUsage, ProcedureNotUsed

Option Explicit

Private Const MAX_ZOOM As Long = 150
Private Const MIN_ZOOM As Long = 5
Private Const ZOOM_STEP As Long = 5
    

' ===========================================================================
' Callbacks for Show/Hide Labels

'@Ignore ParameterNotUsed
Public Sub showColumn_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    ClearWorksheetGraphs
    SettingsSheet.Range(control.id).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    ShowHideDataColumn (control.id)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showColumn_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ShowHideDataColumn (control.id)
    returnedVal = GetSettingBoolean(control.id)
End Sub

Public Sub ShowHideDataColumn(ByVal columnId As String)
    Dim ShowColumn As Boolean
    Dim columnRange As String
    Dim col As String
    
    OptimizeCode_Begin
    
    ' Map the menu item to the column name
    Select Case columnId
        Case RIBBON_CTL_SHOW_COMMENT
            col = SettingsSheet.Range(SETTINGS_DATA_COL_COMMENT).value
        Case RIBBON_CTL_SHOW_ITEM
            col = SettingsSheet.Range(SETTINGS_DATA_COL_ITEM).value
        Case RIBBON_CTL_SHOW_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL).value
        Case RIBBON_CTL_SHOW_OUTSIDE_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL_X).value
        Case RIBBON_CTL_SHOW_TAIL_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL_TAIL).value
        Case RIBBON_CTL_SHOW_HEAD_LABEL
            col = SettingsSheet.Range(SETTINGS_DATA_COL_LABEL_HEAD).value
        Case RIBBON_CTL_SHOW_TOOLTIP
            col = SettingsSheet.Range(SETTINGS_DATA_COL_TOOLTIP).value
        Case RIBBON_CTL_SHOW_IS_RELATED_TO_ITEM
            col = SettingsSheet.Range(SETTINGS_DATA_COL_IS_RELATED_TO).value
        Case RIBBON_CTL_SHOW_STYLE
            col = SettingsSheet.Range(SETTINGS_DATA_COL_STYLE).value
        Case RIBBON_CTL_SHOW_EXTRA_STYLE_ATTRIBUTES
            col = SettingsSheet.Range(SETTINGS_DATA_COL_EXTRA_ATTRIBUTES).value
        Case RIBBON_CTL_SHOW_MESSAGES
            col = SettingsSheet.Range(SETTINGS_DATA_COL_ERROR_MESSAGES).value
    End Select
    
    ' Activate the "data" worksheet
    ActiveWorkbook.Sheets.[_Default](GetDataWorksheetName()).Activate
    
    ' Compose the column range to show/hide
    columnRange = col & ":" & col
    
    ' Show/Hide column according the saved value that corresponds to the check mark in the dropdown list
    ActiveSheet.columns(columnRange).Select
    ShowColumn = GetSettingBoolean(columnId)
    Selection.EntireColumn.Hidden = Not ShowColumn
    
    ' Put the focus on the heading
    ActiveSheet.Range(col & CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))).Select

    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for data worksheet

'@Ignore ParameterNotUsed
Public Sub clearData_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    
    Dim worksheetName As String
    worksheetName = GetDataWorksheetName()
    
    ActiveWorkbook.Sheets.[_Default](worksheetName).Activate
    ActiveWorkbook.Sheets.[_Default](worksheetName).rows.UseStandardHeight = True

    ClearDataWorksheet (worksheetName)
    ClearWorksheetGraphs
    ClearSourceWorksheet
    ClearSourceForm
    Application.StatusBar = False
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for includeImagePath

'@Ignore ParameterNotUsed
Public Sub includeImagePath_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_INCLUDE_IMAGE_PATH).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub includeImagePath_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_INCLUDE_IMAGE_PATH)
End Sub

' ===========================================================================
' Callbacks for addOptions

'@Ignore ParameterNotUsed
Public Sub addOptions_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_APPEND_OPTIONS).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub addOptions_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_APPEND_OPTIONS)
End Sub

' ===========================================================================
' Callbacks for addTimestamp

'@Ignore ParameterNotUsed
Public Sub addTimestamp_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_APPEND_TIMESTAMP).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub addTimestamp_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_APPEND_TIMESTAMP)
End Sub

' ===========================================================================
' Callbacks for blankNodeLabels

'@Ignore ParameterNotUsed
Public Sub blankNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_NODE_LABELS).value = TOGGLE_BLANK_USE_BLANK

    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub blankNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_NODE_LABELS, TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for defaultNodeLabels

'@Ignore ParameterNotUsed
Public Sub defaultNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_NODE_LABELS).value = TOGGLE_BLANK_USE_DEFAULT
    
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_NODE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub defaultNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_NODE_LABELS, TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for blankEdgeLabels

'@Ignore ParameterNotUsed
Public Sub blankEdgeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_EDGE_LABELS).value = TOGGLE_BLANK_USE_BLANK
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub blankEdgeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_EDGE_LABELS, TOGGLE_BLANK_USE_BLANK)
End Sub

' ===========================================================================
' Callbacks for defaultEdgeLabels

'@Ignore ParameterNotUsed
Public Sub defaultEdgeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_BLANK_EDGE_LABELS).value = TOGGLE_BLANK_USE_DEFAULT
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_BLANK
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABELS_DEFAULT
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub defaultEdgeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_BLANK_EDGE_LABELS, TOGGLE_BLANK_USE_DEFAULT)
End Sub

' ===========================================================================
' Callbacks for clearMessages

'@Ignore ParameterNotUsed
Public Sub clearMessages_onAction(ByVal control As IRibbonControl)
    ClearErrors
End Sub

' ===========================================================================
' Callbacks for clearWorksheetGraphs

'@Ignore ParameterNotUsed
Public Sub clearWorksheetGraphs_onAction(ByVal control As IRibbonControl)
    ClearWorksheetGraphs
End Sub

' ===========================================================================
' Callbacks for directed

'@Ignore ParameterNotUsed
Public Sub directed_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        If SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED Then
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
        Else
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
        End If
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
    End If
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_UNDIRECTED
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub directed_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
End Sub

' ===========================================================================
' Callbacks for layout

Private Sub RefreshLayoutGroup()
    InvalidateRibbonControl RIBBON_CTL_GROUP_LAYOUT
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_CIRCO
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_DOT
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_FDP
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_NEATO
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_OSAGE
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_PATCHWORK
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_SFDP
    InvalidateRibbonControl RIBBON_CTL_LAYOUT_TWOPI
End Sub

Private Sub RefreshGraphTypeGroup()
    InvalidateRibbonControl RIBBON_CTL_GROUP_GRAPH_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_DIRECTED
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_UNDIRECTED
End Sub

Private Sub RefreshLayoutParametersGroup()
    InvalidateRibbonControl RIBBON_CTL_ALGORITHM_GROUP
    InvalidateRibbonControl RIBBON_CTL_GRAPH_MODE
    InvalidateRibbonControl RIBBON_CTL_GRAPH_MODEL
    InvalidateRibbonControl RIBBON_CTL_GRAPH_SMOOTHING
    InvalidateRibbonControl RIBBON_CTL_ALGORITHM_GROUP_SEPARATOR1
    InvalidateRibbonControl RIBBON_CTL_ALGORITHM_GROUP_SEPARATOR2
    InvalidateRibbonControl RIBBON_CTL_GRAPH_DIM
    InvalidateRibbonControl RIBBON_CTL_GRAPH_DIMEN
    InvalidateRibbonControl RIBBON_CTL_GRAPH_CLUSTER_RANK
    InvalidateRibbonControl RIBBON_CTL_NEWRANK
    InvalidateRibbonControl RIBBON_CTL_COMPOUND
    InvalidateRibbonControl RIBBON_CTL_OVERLAP
End Sub

'@Ignore ParameterNotUsed
Public Sub layout_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LCase$(Mid$(control.id, Len("layout") + 1))
    Else
        SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LAYOUT_DOT
    End If
    
    OptimizeCode_Begin
    RefreshLayoutGroup
    RefreshSplinesGroup
    RefreshGraphTypeGroup
    RefreshOutputorderGroup
    RefreshRankdirGroup
    RefreshOrderingGroup
    RefreshLayoutParametersGroup
    OptimizeCode_End
    
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub layout_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    Dim layout As String
    layout = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
    
    ' Backward compatibility. Map layout aliases to buttons provided
    Select Case layout
        Case "compound": layout = "polyline"
        Case "splines": layout = "true"
        Case "line": layout = "false"
    End Select
    
    pressed = layout = LCase$(Mid$(control.id, Len("layout") + 1))
End Sub

'@Ignore ParameterNotUsed
Public Sub layoutOptions_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    label = "layout=" & SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
End Sub

' ===========================================================================
' Callbacks for undirected

'@Ignore ParameterNotUsed
Public Sub undirected_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        If SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED Then
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
        Else
            SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
        End If
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_DIRECTED
    End If
    InvalidateRibbonControl RIBBON_CTL_GRAPH_TYPE_DIRECTED
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub undirected_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_TYPE).value = TOGGLE_UNDIRECTED
End Sub

'@Ignore ParameterNotUsed
Public Sub directed_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_PATCHWORK
            visible = False
        Case Else
            visible = True
    End Select
End Sub

' ===========================================================================
' Callbacks for splines

Private Sub RefreshSplinesGroup()
    InvalidateRibbonControl RIBBON_CTL_SPLINES_COMPOUND
    InvalidateRibbonControl RIBBON_CTL_SPLINES_CURVED
    InvalidateRibbonControl RIBBON_CTL_SPLINES_LINE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_NONE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_ORTHO
    InvalidateRibbonControl RIBBON_CTL_SPLINES_POLYLINE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_SPLINE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_TRUE
    InvalidateRibbonControl RIBBON_CTL_SPLINES_FALSE
End Sub

'@Ignore ParameterNotUsed
Public Sub splines_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_SPLINES).value = LCase$(Mid$(control.id, Len("spline") + 1))
    Else
        SettingsSheet.Range(SETTINGS_SPLINES).value = vbNullString
    End If
    RefreshSplinesGroup
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub splines_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_PATCHWORK
            visible = False
        Case Else
            visible = True
    End Select
End Sub

'@Ignore ParameterNotUsed
Public Sub splines_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If SettingsSheet.Range(SETTINGS_SPLINES).value = vbNullString And control.id = "splineFalse" Then
        pressed = True
    Else
        pressed = SettingsSheet.Range(SETTINGS_SPLINES).value = LCase$(Mid$(control.id, Len("spline") + 1))
    End If
End Sub


' ===========================================================================
' Callbacks for dirName

'@Ignore ParameterNotUsed
Public Sub getDir_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = True
    
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    If dirName = vbNullString Then Exit Sub

    ' Validate that the output directory exists
    If Not DirectoryExists(dirName) Then
        SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY).value = vbNullString
    End If

End Sub

'@Ignore ParameterNotUsed
Public Sub getDir_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    If dirName = vbNullString Then
        label = GetLabel("getDir")
    Else
        label = vbNullString
    End If
End Sub

' ===========================================================================
' Callbacks for getDirLabel

'@Ignore ParameterNotUsed
Public Sub getDirLabel_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim folder As String
    folder = Trim$(SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY))
    label = ShortenToLastTwoFolders(folder)
End Sub

Private Function ShortenToLastTwoFolders(ByVal fullPath As String) As String
    Dim parts() As String
    Dim n As Long
    Dim sep As String

    ' Ensure a path was passed in
    If Trim$(fullPath) = vbNullString Then
        ShortenToLastTwoFolders = vbNullString
        Exit Function
    End If
    
    ' We have a non-null string, parse it.
    sep = Application.pathSeparator
    parts = split(fullPath, sep)
    n = UBound(parts)

    Select Case n
        Case 0
            ShortenToLastTwoFolders = parts(0)

        Case 1
            ShortenToLastTwoFolders = parts(0) & sep & parts(1)

        Case Else
            ShortenToLastTwoFolders = "..." & sep & parts(n - 1) & sep & parts(n)
    End Select
End Function

' ===========================================================================
' Callbacks for fileFormat

'@Ignore ParameterNotUsed
Public Sub fileFormat_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_FILE_FORMAT).value = Mid$(controlId, Len("ff_") + 1)
End Sub

'@Ignore ParameterNotUsed
Public Sub fileFormat_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "ff_" & SettingsSheet.Range(SETTINGS_FILE_FORMAT).value
End Sub

' ===========================================================================
' Callbacks for filePrefix

'@Ignore ParameterNotUsed
Public Sub filePrefix_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range(SETTINGS_FILE_NAME).value = Text
End Sub

'@Ignore ParameterNotUsed
Public Sub filePrefix_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range(SETTINGS_FILE_NAME))
End Sub

' ===========================================================================
' Callbacks for getDir

'@Ignore ParameterNotUsed
Public Sub getDir_onAction(ByVal control As IRibbonControl)
    SelectDirectoryToCell SettingsSheet.name, SETTINGS_OUTPUT_DIRECTORY
    RefreshRibbon
End Sub

' ===========================================================================
' Callbacks for graphToFile

'@Ignore ParameterNotUsed
Public Sub graphToFile_onAction(ByVal control As IRibbonControl)
    Dim firstColumn As Long
    Dim lastColumn As Long
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    lastColumn = firstColumn
    
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents
    
    OptimizeCode_Begin
    CreateGraphFile firstColumn, lastColumn
    OptimizeCode_End
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
End Sub

'@Ignore ParameterNotUsed
Public Sub graphToFile_getEnabled(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = Not (IsAViewSpecified() = False)
End Sub

' ===========================================================================
' Callbacks for graphAllViewsToFile

'@Ignore ParameterNotUsed
Public Sub graphAllViewsToFile_onAction(ByVal control As IRibbonControl)

    Dim nonEmptyCellCount As Long
    Dim row As Long
    Dim col As Long
    Dim columnName As String
    Dim firstColumn As Long
    Dim lastColumn As Long
    
    row = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    nonEmptyCellCount = 0
    
    ' Get the configured location of the first view name column
    firstColumn = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    
    ' Count the non-empty cells beginning at the first view column
    For col = firstColumn To GetLastColumn(StylesSheet.name, row)
        columnName = StylesSheet.Cells.item(row, col)
        If columnName <> vbNullString Then
            nonEmptyCellCount = nonEmptyCellCount + 1
        End If
    Next col

    ' Calaculate the absolute column number of the last view column
    lastColumn = firstColumn + nonEmptyCellCount - 1
    
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents
    
    ' Graph all the views
    OptimizeCode_Begin
    CreateGraphFile firstColumn, lastColumn
    OptimizeCode_End
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
End Sub

' ===========================================================================
' Callbacks for graphToWorksheet

'@Ignore ParameterNotUsed
Public Sub graphToWorksheet_onAction(ByVal control As IRibbonControl)
    CreateGraphWorksheetQuickly
End Sub

'@Ignore ParameterNotUsed
Public Sub graphToWorksheet_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    enabled = IsAViewSpecified()
End Sub

' ===========================================================================
' Callbacks for graphAuto

'@Ignore ParameterNotUsed
Public Sub graphAuto_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RUN_MODE).value = Toggle(pressed, TOGGLE_AUTO, TOGGLE_MANUAL)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub graphAuto_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_RUN_MODE).value = TOGGLE_AUTO
End Sub

' ===========================================================================
' Callbacks for graphWorksheet

'@Ignore ParameterNotUsed
Public Sub graphWorksheet_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value = "data"
    Else
        SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value = "graph"
    End If
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub graphWorksheet_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef itemLabel As Variant)
    If index = 0 Then
        itemLabel = GetLabel("worksheetDataName")
    Else
        itemLabel = GetLabel("worksheetGraphName")
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub graphWorksheet_getItemCount(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = 2
End Sub

'@Ignore ParameterNotUsed
Public Sub graphWorksheet_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef itemIndex As Variant)
    If SettingsSheet.Range(SETTINGS_IMAGE_WORKSHEET).value = "data" Then
        itemIndex = 0
    Else
        itemIndex = 1
    End If
End Sub

' ===========================================================================
' Callbacks for imageFormat

'@Ignore ParameterNotUsed
Public Sub imageFormat_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value = Mid$(controlId, Len("img_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub imageFormat_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "img_" & SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value
End Sub

' ===========================================================================
' Callbacks for includeOrphanEdges

'@Ignore ParameterNotUsed
Public Sub includeOrphanEdges_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_RELATIONSHIPS_WITHOUT_NODES).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub includeOrphanEdges_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_RELATIONSHIPS_WITHOUT_NODES)
End Sub

' ===========================================================================
' Callbacks for includeOrphanNodes

'@Ignore ParameterNotUsed
Public Sub includeOrphanNodes_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_NODES_WITHOUT_RELATIONSHIPS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub includeOrphanNodes_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_NODES_WITHOUT_RELATIONSHIPS)
End Sub

' ===========================================================================
' Callbacks for keepGvFile

'@Ignore ParameterNotUsed
Public Sub keepGvFile_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_FILE_DISPOSITION).value = Toggle(pressed, TOGGLE_KEEP, TOGGLE_DELETE)
End Sub

'@Ignore ParameterNotUsed
Public Sub keepGvFile_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_FILE_DISPOSITION).value = TOGGLE_KEEP
End Sub

' ===========================================================================
' Callbacks for rankdir

'@Ignore ParameterNotUsed
Public Sub rankdir_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LAYOUT_DOT
End Sub

Public Sub rankdir_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_RANKDIR).value = Mid$(control.id, Len("rankdir") + 1)
    Else
        SettingsSheet.Range(SETTINGS_RANKDIR).value = vbNullString
    End If
    RefreshRankdirGroup
    AutoDraw
End Sub

Public Sub rankdir_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If SettingsSheet.Range(SETTINGS_RANKDIR).value = vbNullString And control.id = "rankdirTB" Then
        pressed = True
    Else
        pressed = SettingsSheet.Range(SETTINGS_RANKDIR).value = Mid$(control.id, Len("rankdir") + 1)
    End If
End Sub

Private Sub RefreshRankdirGroup()
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_GROUP
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_TB
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_BT
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_LR
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_RL
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_DUMMY1
    InvalidateRibbonControl RIBBON_CTL_RANKDIR_DUMMY2
End Sub

' ===========================================================================
' Callbacks for showNodeLabels

'@Ignore ParameterNotUsed
Public Sub showNodeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_NODE_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showNodeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_NODE_LABELS)
End Sub

' ===========================================================================
' Callbacks for showNodeXLabels

'@Ignore ParameterNotUsed
Public Sub showNodeXLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_NODE_XLABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showNodeXLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_NODE_XLABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_LABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeXLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeXLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_XLABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeXLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_XLABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeHeadLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeHeadLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_HEAD_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeHeadLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_HEAD_LABELS)
End Sub

' ===========================================================================
' Callbacks for showEdgeTailLabels

'@Ignore ParameterNotUsed
Public Sub showEdgeTailLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_TAIL_LABELS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showEdgeTailLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_TAIL_LABELS)
End Sub

' ===========================================================================
' Callbacks for showPorts

'@Ignore ParameterNotUsed
Public Sub showPorts_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EDGE_PORTS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub showPorts_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EDGE_PORTS)
End Sub

' ===========================================================================
' Callbacks for strict

'@Ignore ParameterNotUsed
Public Sub strict_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_STRICT).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub strict_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_STRICT)
End Sub

' ===========================================================================
' Callbacks for transparent

'@Ignore ParameterNotUsed
Public Sub transparent_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_TRANSPARENT).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub transparent_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_TRANSPARENT)
End Sub

' ===========================================================================
' Callbacks for center

'@Ignore ParameterNotUsed
Public Sub center_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_CENTER).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub center_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_CENTER)
End Sub

' ===========================================================================
' Callbacks for compound

'@Ignore ParameterNotUsed
Public Sub compound_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_COMPOUND).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub compound_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_COMPOUND)
End Sub

'@Ignore ParameterNotUsed
Public Sub compound_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for concentrate

'@Ignore ParameterNotUsed
Public Sub concentrate_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_CONCENTRATE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub concentrate_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_CONCENTRATE)
End Sub

' ===========================================================================
' Callbacks for forceLabels

'@Ignore ParameterNotUsed
Public Sub forceLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_FORCE_LABELS).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub forceLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_FORCE_LABELS)
End Sub

' ===========================================================================
' Callbacks for newrank

'@Ignore ParameterNotUsed
Public Sub newrank_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_NEWRANK).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub newrank_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_NEWRANK)
End Sub

'@Ignore ParameterNotUsed
Public Sub newrank_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for rotate

'@Ignore ParameterNotUsed
Public Sub rotate_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPH_ORIENTATION).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub rotate_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPH_ORIENTATION)
End Sub

Public Function getPressed(ByVal worksheetName As String, ByVal keyword As String, ByVal matchValue As String) As Boolean
    getPressed = UCase$(GetCellString(worksheetName, keyword)) = UCase$(matchValue)
End Function

' ===========================================================================
' Callbacks for overlap

'@Ignore ParameterNotUsed
Public Sub overlap_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = False
        Case LAYOUT_DOT
            visible = False
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_OSAGE
            visible = False
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = True
        Case LAYOUT_TWOPI
            visible = False
        Case Else
            visible = False
    End Select
End Sub

'@Ignore ParameterNotUsed
Public Sub overlap_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value = Mid$(controlId, Len("overlap_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub overlap_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value = control.id
End Sub


'@Ignore ParameterNotUsed
Public Sub overlap_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "overlap_" & SettingsSheet.Range(SETTINGS_GRAPH_OVERLAP).value
End Sub

' ===========================================================================
' Callbacks for toggleDebugLabels

'@Ignore ParameterNotUsed
Public Sub toggleDebugLabels_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_DEBUG).value = Toggle(pressed, TOGGLE_ON, TOGGLE_OFF)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleDebugLabels_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_DEBUG)
End Sub

' ===========================================================================
' Callbacks for toggleLogToConsole

'@Ignore ParameterNotUsed
Public Sub toggleDebugLogToConsole_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_LOG_TO_CONSOLE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleDebugLogToConsole_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub toggleDebugLogToConsole_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = enableConsole()
End Sub

' ===========================================================================
' Callbacks for toggleGraphvizVerbose

'@Ignore ParameterNotUsed
Public Sub toggleGraphvizVerbose_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_GRAPHVIZ_VERBOSE).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleGraphvizVerbose_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_GRAPHVIZ_VERBOSE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub toggleGraphvizVerbose_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = enableConsole()
End Sub

' ===========================================================================
' Callbacks for useDefinedStyles

'@Ignore ParameterNotUsed
Public Sub useDefinedStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_INCLUDE_STYLE_FORMAT).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub useDefinedStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_INCLUDE_STYLE_FORMAT)
End Sub

' ===========================================================================
' Callbacks for useExtraStyles

'@Ignore ParameterNotUsed
Public Sub useExtraStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_INCLUDE_EXTRA_ATTRIBUTES).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub useExtraStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_INCLUDE_EXTRA_ATTRIBUTES)
End Sub

' ===========================================================================
' Callbacks for yesView

'Callback for yesNoView onAction
'@Ignore ParameterNotUsed
Public Sub yesNoView_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    Dim columnName As String
    columnName = ConvertColumnNumberToLetters(index + GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW))
    SettingsSheet.Range(SETTINGS_YES_NO_SWITCH_COLUMN).value = columnName
    AutoDraw
End Sub

'Callback for yesNoView getItemCount
'@Ignore ParameterNotUsed
Public Sub yesNoView_getItemCount(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim itemCount As Long
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim columnName As String
    
    row = CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING))
    itemCount = 0
    
    ' Count the non-empty cells beginning at the first view column
    For col = GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW) To GetLastColumn(StylesSheet.name, row)
        columnName = StylesSheet.Cells.item(row, col)
        If columnName <> vbNullString Then
            itemCount = itemCount + 1
        End If
    Next col

    ' If the last view column is the currently selected column, and the user deletes the column then it
    ' is necessary to change the selection to the last column which will be present after the deletion occurs.
    lastCol = itemCount + GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW) - 1
    
    If lastCol < GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE) Then
        SettingsSheet.Range(SETTINGS_YES_NO_SWITCH_COLUMN).value = ConvertColumnNumberToLetters(lastCol)
        RefreshRibbon
    End If
    
    returnedVal = itemCount
End Sub

' Callback for yesNoView getItemLabel
'@Ignore ParameterNotUsed
Public Sub yesNoView_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef itemLabel As Variant)
    itemLabel = StylesSheet.Cells.item(CLng(SettingsSheet.Range(SETTINGS_STYLES_ROW_HEADING)), _
                            index + GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW))
End Sub

'@Ignore ParameterNotUsed
Public Sub yesNoView_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef itemIndex As Variant)
    Dim indx As Long
    indx = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE) - GetSettingColNum(SETTINGS_STYLES_COL_FIRST_YES_NO_VIEW)
    itemIndex = indx
End Sub

' Utility routines

Public Function IsAViewSpecified() As Boolean
    IsAViewSpecified = Not (SettingsSheet.Range(SETTINGS_VIEW_NAME).value = "0")
End Function



'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sql_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub mac_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = True
#Else
    returnedVal = False
#End If
End Sub

' ===========================================================================
' Callbacks for graphZoomLevel

Public Sub graphZoomLevel_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    label = GetCurrentZoom() & "%"
End Sub

' ===========================================================================
' Callbacks for graphZoomOut

'@Ignore ParameterNotUsed
Public Sub graphZoomOut_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    enabled = GetCurrentZoom() > MIN_ZOOM
End Sub

'@Ignore ParameterNotUsed
Public Sub GraphZoomOut_OnAction(ByVal control As IRibbonControl)
    Dim zoom As Long
    zoom = SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value - ZOOM_STEP
    
    UpdateZoom zoom
End Sub

' ===========================================================================
' Callbacks for graphZoomIn

'@Ignore ParameterNotUsed
Public Sub graphZoomIn_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    enabled = GetCurrentZoom() < MAX_ZOOM
End Sub

Public Sub GraphZoomIn_OnAction(ByVal control As IRibbonControl)
    Dim zoom As Long
    zoom = SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value + ZOOM_STEP
    
    UpdateZoom zoom
End Sub

' Helper function to update zoom and refresh UI
Private Sub UpdateZoom(zoom As Long)
    ' Clamp zoom value within bounds
    If zoom < MIN_ZOOM Then zoom = MIN_ZOOM
    If zoom > MAX_ZOOM Then zoom = MAX_ZOOM
    
    ' Update zoom value in settings
    SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value = zoom
    
    ' Invalidate ribbon controls
    InvalidateRibbonControl "graphZoomIn"
    InvalidateRibbonControl "graphZoomOut"
    InvalidateRibbonControl "graphZoomLevel"
    
    ' Refresh graph
    CreateGraphWorksheetQuickly
End Sub

' Helper function to get current zoom value
Private Function GetCurrentZoom() As Long
    GetCurrentZoom = SettingsSheet.Range(SETTINGS_SCALE_IMAGE).value
End Function

' ===========================================================================
' Callbacks for dim

'@Ignore ParameterNotUsed
Public Sub dim_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_DIM).value = Mid$(controlId, Len("dim_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub dim_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "dim_" & SettingsSheet.Range(SETTINGS_GRAPH_DIM).value
End Sub

'@Ignore ParameterNotUsed
Public Sub dim_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for dimen

'@Ignore ParameterNotUsed
Public Sub dimen_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_DIMEN).value = Mid$(controlId, Len("dimen_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub dimen_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "dimen_" & SettingsSheet.Range(SETTINGS_GRAPH_DIMEN).value
End Sub

'@Ignore ParameterNotUsed
Public Sub dimen_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for mode

'@Ignore ParameterNotUsed
Public Sub mode_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_MODE).value = Mid$(controlId, Len("mode_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub mode_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "mode_" & SettingsSheet.Range(SETTINGS_GRAPH_MODE).value
End Sub

'@Ignore ParameterNotUsed
Public Sub mode_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for model

'@Ignore ParameterNotUsed
Public Sub model_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_MODEL).value = Mid$(controlId, Len("model_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub model_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "model_" & SettingsSheet.Range(SETTINGS_GRAPH_MODEL).value
End Sub

'@Ignore ParameterNotUsed
Public Sub model_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_NEATO
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for smoothing

'@Ignore ParameterNotUsed
Public Sub smoothing_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_SMOOTHING).value = Mid$(controlId, Len("smoothing_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub smoothing_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "smoothing_" & SettingsSheet.Range(SETTINGS_GRAPH_SMOOTHING).value
End Sub

'@Ignore ParameterNotUsed
Public Sub smoothing_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_SFDP
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for clusterrank

'@Ignore ParameterNotUsed
Public Sub clusterrank_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_GRAPH_CLUSTER_RANK).value = Mid$(controlId, Len("clusterrank_") + 1)
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub clusterrank_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "clusterrank_" & SettingsSheet.Range(SETTINGS_GRAPH_CLUSTER_RANK).value
End Sub

'@Ignore ParameterNotUsed
Public Sub clusterrank_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for ordering

Private Sub RefreshOrderingGroup()
    InvalidateRibbonControl RIBBON_CTL_ORDERING_GROUP
    InvalidateRibbonControl RIBBON_CTL_ORDERING_IN
    InvalidateRibbonControl RIBBON_CTL_ORDERING_OUT
    InvalidateRibbonControl RIBBON_CTL_ORDERING_DUMMY1
End Sub

'@Ignore ParameterNotUsed
Public Sub ordering_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_DOT
            visible = True
        Case Else
            visible = False
    End Select
End Sub

Public Sub ordering_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value = LCase$(Mid$(control.id, Len("ordering") + 1))
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value = vbNullString
    End If
    RefreshOrderingGroup
    AutoDraw
End Sub

Public Sub ordering_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If LCase$(SettingsSheet.Range(SETTINGS_GRAPH_ORDERING).value) = LCase$(Mid$(control.id, Len("ordering") + 1)) Then
        pressed = True
    Else
        pressed = False
    End If
End Sub

' ===========================================================================
' Callbacks for outputorder

Private Sub RefreshOutputorderGroup()
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_GROUP
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_NODES_FIRST
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_EDGES_FIRST
    InvalidateRibbonControl RIBBON_CTL_OUTPUTORDER_BREADTH_FIRST
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorderBreadthFirst_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    If SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString Then
        pressed = True
    Else
        pressed = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "breadthfirst"
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorderBreadthFirst_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "breadthfirst"
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString
    End If
    RefreshOutputorderGroup
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorderEdgesFirst_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "edgesfirst"
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorderEdgesFirst_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "edgesfirst"
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString
    End If
    RefreshOutputorderGroup
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorderNodesFirst_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "nodesfirst"
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorderNodesFirst_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = "nodesfirst"
    Else
        SettingsSheet.Range(SETTINGS_GRAPH_OUTPUT_ORDER).value = vbNullString
    End If
    RefreshOutputorderGroup
    AutoDraw
End Sub

'@Ignore ParameterNotUsed
Public Sub outputorder_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = True
        Case LAYOUT_DOT
            visible = True
        Case LAYOUT_FDP
            visible = True
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_OSAGE
            visible = True
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = True
        Case LAYOUT_TWOPI
            visible = True
        Case Else
            visible = False
    End Select
End Sub

'@Ignore ParameterNotUsed
Public Sub algsep1_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = False
        Case LAYOUT_DOT
            visible = True
        Case LAYOUT_FDP
            visible = False
        Case LAYOUT_NEATO
            visible = False
        Case LAYOUT_OSAGE
            visible = False
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = False
        Case LAYOUT_TWOPI
            visible = False
        Case Else
            visible = False
    End Select
End Sub

'@Ignore ParameterNotUsed
Public Sub algsep2_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
        Case LAYOUT_CIRCO
             visible = False
        Case LAYOUT_DOT
            visible = True
        Case LAYOUT_FDP
            visible = False
        Case LAYOUT_NEATO
            visible = True
        Case LAYOUT_OSAGE
            visible = False
        Case LAYOUT_PATCHWORK
            visible = False
        Case LAYOUT_SFDP
            visible = True
        Case LAYOUT_TWOPI
            visible = False
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub graphvizHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLGraphvizTab").value, NewWindow:=True
End Sub

