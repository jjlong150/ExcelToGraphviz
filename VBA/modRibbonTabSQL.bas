Attribute VB_Name = "modRibbonTabSQL"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

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
    If SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value = vbNullString Then
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
        SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value = vbNullString
    Else
        SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value = ConvertColumnNumberToLetters(index + 4)
    End If
    
    SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).Value = vbNullString
    
    InvalidateRibbonControl "sqlFilterValue"
    InvalidateRibbonControl "sqlFilterRefresh"
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterColumn_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value) = vbNullString Then
        returnedVal = 0
    Else
        returnedVal = GetSettingColNum(SETTINGS_SQL_COL_FILTER) - 4
    End If
End Sub

' ===========================================================================
' Callbacks for sqlFilterValues

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value) = vbNullString Then
        Set filterValues = New Dictionary
        listSize = 1
    Else
        Set filterValues = GetFilterValues(GetSettingColNum(SETTINGS_SQL_COL_FILTER))
        listSize = filterValues.Count + 1
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
        SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).Value = vbNullString
    Else
        SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).Value = filterValues.Keys()(index - 1)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sqlFilterValue_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If Trim$(SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value) = vbNullString Or Trim$(SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).Value) = vbNullString Then
        returnedVal = 0
    Else
        Dim key As Variant
        Dim itemIndex As Long
        itemIndex = 1
        returnedVal = 0
        For Each key In filterValues.Keys()
            If CStr(key) = Trim$(SettingsSheet.Range(SETTINGS_SQL_FILTER_VALUE).Value) Then
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
    returnedVal = Not (SettingsSheet.Range(SETTINGS_SQL_COL_FILTER).Value = vbNullString)
End Sub


Private Function GetFilterValues(ByVal filterColumn As Long) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary

    ' Determine the last row with data
    Dim lastRow As Long
    With SqlSheet.UsedRange
        lastRow = .Cells.Item(.Cells.Count).row
    End With

    Dim sqlRow As Long
    Dim cellValue As String
    For sqlRow = 2 To lastRow
        cellValue = Trim$(SqlSheet.Cells.Item(sqlRow, filterColumn).Value)
        If cellValue <> vbNullString Then
            If Not dictionaryObj.Exists(cellValue) Then
                dictionaryObj.Add cellValue, sqlRow
            End If
        End If
    Next
    
    Set GetFilterValues = dictionaryObj
End Function

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub sqlHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSqlTab").Value, NewWindow:=True
End Sub

