Attribute VB_Name = "modRibbonTabSQL"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ProcedureNotUsed

Option Explicit

' We have to cache the dropdown values for the callback functions
Private filterValues As Dictionary

' ===========================================================================
' Ribbon callbacks for SQL Tab
' ===========================================================================

'@Ignore ParameterNotUsed
Public Sub sqlRun_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    RunSQL
    OptimizeCode_End
    AutoDraw
End Sub

Public Sub RunSQLAsExtension()
    OptimizeCode_Begin
    RunSQL
    OptimizeCode_End
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
    
    InvalidateRibbonControl "sqlFilterValue"
    InvalidateRibbonControl "sqlFilterRefresh"
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
    InvalidateRibbonControl "sqlFilterValue"
    InvalidateRibbonControl "sqlFilterRefresh"
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
        lastRow = .Cells.Item(.Cells.count).row
    End With

    Dim sqlRow As Long
    Dim cellValue As String
    For sqlRow = 2 To lastRow
        cellValue = Trim$(SqlSheet.Cells.Item(sqlRow, filterColumn).value)
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
