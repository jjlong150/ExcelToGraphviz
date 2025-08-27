Attribute VB_Name = "modRibbonTabSQL"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ProcedureNotUsed

Option Explicit

' We have to cache the dropdown values for the callback functions
Private filterValues As Dictionary
Private excelFiles As Collection

' ===========================================================================
' Ribbon callbacks for SQL Tab
' ===========================================================================

'@Ignore ParameterNotUsed
Public Sub sqlRun_onAction(ByVal control As IRibbonControl)
    RunSQLAsExtension
    AutoDraw
End Sub

Public Sub RunSQLAsExtension()
    Application.Cursor = xlWait
    OptimizeCode_Begin
    RunSQL
    InvalidateRibbonControl RIBBON_CTL_SQL_CONN_POOL_RESET
    OptimizeCode_End
    Application.Cursor = xlDefault
End Sub

'@Ignore ParameterNotUsed
Public Sub sqlClearStatus_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    ClearSQLStatus
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for sqlFilterColumn

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterColumn_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    If SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value = vbNullString Then
        Set filterValues = New Dictionary
    Else
        Set filterValues = GetFilterValues(GetSettingColNum(SETTINGS_SQL_COL_FILTER))
    End If
    listSize = 23   ' Represents columns E-Z
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterColumn_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = vbNullString
    Else
        label = ConvertColumnNumberToLetters(index + 4)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterColumn_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value = vbNullString
    Else
        SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value = ConvertColumnNumberToLetters(index + 4)
    End If
    
    SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value = vbNullString
    
    InvalidateRibbonControl RIBBON_CTL_SQL_FILTER_VALUE
    InvalidateRibbonControl RIBBON_CTL_SQL_FILTER_REFRESH
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterColumn_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value) = vbNullString Then
        returnedVal = 0
    Else
        returnedVal = GetSettingColNum(SETTINGS_SQL_COL_FILTER) - 4
    End If
End Sub

' ===========================================================================
' Callbacks for sqlFilterValues

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value) = vbNullString Then
        Set filterValues = New Dictionary
        listSize = 1
    Else
        Set filterValues = GetFilterValues(GetSettingColNum(SETTINGS_SQL_COL_FILTER))
        listSize = filterValues.count + 1
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = vbNullString
    Else
        label = filterValues.Keys()(index - 1)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value = vbNullString
    Else
        SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value = filterValues.Keys()(index - 1)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value) = vbNullString Or Trim$(SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value) = vbNullString Then
        returnedVal = 0
    Else
        Dim key As Variant
        Dim itemIndex As Long
        itemIndex = 1
        returnedVal = 0
        For Each key In filterValues.Keys()
            If CStr(key) = Trim$(SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).value) Then
                returnedVal = itemIndex
            End If
            itemIndex = itemIndex + 1
        Next
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for sqlFilterRefresh

'@Ignore ParameterNotUsed
Public Sub sqlFilterRefresh_onAction(ByVal control As IRibbonControl)
    InvalidateRibbonControl RIBBON_CTL_SQL_FILTER_VALUE
    InvalidateRibbonControl RIBBON_CTL_SQL_FILTER_REFRESH
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterRefresh_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not (SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).value = vbNullString)
End Sub


Private Function GetFilterValues(ByVal filterColumn As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary

    ' Determine the last row with data
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    Dim sqlRow As Long
    Dim cellValue As String
    For sqlRow = 2 To lastRow
        cellValue = Trim$(SqlSheet.Cells.item(sqlRow, filterColumn).value)
        If cellValue <> vbNullString Then
            If Not dictionaryObj.Exists(cellValue) Then
                dictionaryObj.Add cellValue, sqlRow
            End If
        End If
    Next
    
    Set GetFilterValues = dictionaryObj
End Function

' ===========================================================================
' Callbacks for sqlEditCell

'@Ignore ParameterNotUsed
Public Sub sqlEditCell_onAction(ByVal control As IRibbonControl)
    CellValueEditForm.show
End Sub

' Used by the "Edit" button that appears on the SQL cell
Public Sub ShowSQLEditForm()
    CellValueEditForm.show
End Sub

'@Ignore ParameterNotUsed
Public Sub sqlEditCell_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)

    enabled = False
    
    If ActiveSheet.name <> SqlSheet.name Then
        Exit Sub
    End If
    
    If Selection.Cells.count > 1 Then
        Exit Sub
    End If
    
    If ActiveCell.HasFormula Then
        Exit Sub
    End If
    
    If ActiveCell.Column = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT) Then
        enabled = True
    End If
    
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub sqlHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSqlTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for Copy to Clipboard

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlClipboard_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

'@Ignore ParameterNotUsed
Public Sub sqlClipboard_onAction(ByVal control As IRibbonControl)
#If Not Mac Then
    
    If ClipBoard_SetData(ActiveCell.value) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySqlSuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySqlFailed"), 5
    End If
    
#End If
End Sub

' ===========================================================================
' Callbacks for sqlConnPoolReset

'@Ignore ParameterNotUsed
Public Sub sqlConnPoolReset_onAction(ByVal control As IRibbonControl)
    CleanupConnectionPool
    InvalidateRibbonControl RIBBON_CTL_SQL_CONN_POOL_RESET
End Sub

'@Ignore ParameterNotUsed
Public Sub sqlConnPoolReset_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    If GetConnectionCount() > 0 Then
        enabled = True
    Else
        enabled = False
    End If
End Sub

Public Sub sqlConnPoolReset_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetLabel(control.ID) & " (" & GetConnectionCount() & ")"
End Sub

' ===========================================================================
' Callbacks for sqlConnPoolDevMode

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlConnPoolDevMode_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_SQL_CLOSE_CONNECTIONS).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlConnPoolDevMode_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_SQL_CLOSE_CONNECTIONS)
End Sub


' ===========================================================================
' Callbacks for datasourceDir

'@Ignore ParameterNotUsed
Public Sub datasourceDir_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = True
    
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY))
    If dirName = vbNullString Then Exit Sub

    ' Validate that the directory exists. Clear the fields if it does not exist
    ' to prevent problems when the spreadsheet is shared.
    If Not DirectoryExists(dirName) Then
        SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY).value = vbNullString
        SettingsSheet.Range(SETTINGS_DATASOURCE_FILE).value = vbNullString
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub datasourceDir_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY))
    If dirName = vbNullString Then
        label = GetLabel("getDatasourceDir")
    Else
        label = vbNullString
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub datasourceDir_onAction(ByVal control As IRibbonControl)
    ' Save the current directory
    Dim dirName As String
    dirName = Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY))
    
    ' Display the folder picking dialog and save the chosen folder
    SelectDirectoryToCell SettingsSheet.name, SETTINGS_DATASOURCE_DIRECTORY
    
    ' Check if the directory has changed
    If dirName <> Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY)) Then
        DatasourceFileReset
        DatasourceRibbonRefresh
    End If
End Sub

Private Sub DatasourceDirectoryReset()
    SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY).value = vbNullString
End Sub

Private Sub DatasourceFileReset()
    SettingsSheet.Range(SETTINGS_DATASOURCE_FILE).value = vbNullString
    Set excelFiles = Nothing
End Sub

Private Sub DatasourceRibbonRefresh()
    InvalidateRibbonControl "datasourceDir"
    InvalidateRibbonControl "datasourceDirLabel"
    InvalidateRibbonControl "datasourceFile"
    InvalidateRibbonControl "datasourceFileRefresh"
    InvalidateRibbonControl "datasourceReset"
End Sub

' ===========================================================================
' Callbacks for datasourceDirLabel

'@Ignore ParameterNotUsed
Public Sub datasourceDirLabel_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    label = Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY))
End Sub

' ===========================================================================
' Callbacks for datasourceFile

'@Ignore ParameterNotUsed
Public Sub datasourceFile_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY)) = vbNullString Then
        visible = False
    Else
        visible = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFile_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    Dim datasourceFolder As String
    datasourceFolder = Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY))
    
    If datasourceFolder <> vbNullString Then
        GetDatasources datasourceFolder
    End If
    
    If excelFiles Is Nothing Then
        listSize = 1
    Else
        listSize = excelFiles.count + 1
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFile_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = GetLabel("datasourceSelectAFile")
    Else
        label = excelFiles(index)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFile_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        SettingsSheet.Range(SETTINGS_DATASOURCE_FILE).value = vbNullString
    Else
        SettingsSheet.Range(SETTINGS_DATASOURCE_FILE).value = excelFiles(index)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFile_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef itemIndex As Variant)
    Dim targetFile As String
    targetFile = Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_FILE).value)
    
    If targetFile = vbNullString Then
        itemIndex = 0
        Exit Sub
    End If
    
    If excelFiles Is Nothing Then
        itemIndex = 0
        Exit Sub
    End If
    
    If excelFiles.count = 0 Then
        itemIndex = 0
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To excelFiles.count
        If StrComp(excelFiles(i), targetFile, vbTextCompare) = 0 Then
            itemIndex = i
            Exit Sub
        End If
    Next i
    
    ' If no match is found, clear the selection
    SettingsSheet.Range(SETTINGS_DATASOURCE_FILE).value = vbNullString
    itemIndex = 0
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFile_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY)) = vbNullString Then
        enabled = False
    Else
        enabled = True
    End If
End Sub

Public Sub GetDatasources(ByVal folderPath As String)
    Set excelFiles = New Collection

    Dim fileName As String
    fileName = Dir(folderPath & "\*.xls*") ' Broad match
    Do While fileName <> ""
        Select Case LCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))
            Case "xls", "xlsx", "xlsm", "xlsb"
                excelFiles.Add fileName
        End Select
        fileName = Dir()
    Loop
End Sub

' ===========================================================================
' Callbacks for datasourceFileRefresh

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFileRefresh_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY)) = vbNullString Then
        visible = False
    Else
        visible = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceFileRefresh_onAction(ByVal control As IRibbonControl)
    InvalidateRibbonControl "datasourceFile"
End Sub

' ===========================================================================
' Callbacks for datasourceFileRefresh

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceReset_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_DATASOURCE_DIRECTORY)) = vbNullString Then
        visible = False
    Else
        visible = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub datasourceReset_onAction(ByVal control As IRibbonControl)
    DatasourceDirectoryReset
    DatasourceFileReset
    DatasourceRibbonRefresh
End Sub

