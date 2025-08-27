Attribute VB_Name = "modRibbonTabStyles"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule AssignmentNotUsed, UseMeaningfulName, UnassignedVariableUsage, ProcedureNotUsed, ParameterNotUsed, ImplicitByRefModifier

Option Explicit

Public Sub stylesClear_onAction(ByVal control As IRibbonControl)
    ClearStylesPreview
End Sub

Public Sub stylesPreview_onAction(ByVal control As IRibbonControl)
    StylesSheet.Activate
    GenerateStylesPreview ActiveCell.row
    ClearStatusBar
End Sub

Public Sub stylesPreviewAll_onAction(ByVal control As IRibbonControl)
    StylesSheet.Activate
    ClearStylesPreview
    GenerateStylesPreviewAll
    ClearStatusBar
End Sub

' ===========================================================================
' Callbacks for stylesSuffixBegin

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixBegin_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range("StylesSuffixOpen").value = Text
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixBegin_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range("StylesSuffixOpen"))
End Sub

' ===========================================================================
' Callbacks for stylesSuffixEnd

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixEnd_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range("StylesSuffixClose").value = Text
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub stylesSuffixEnd_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range("StylesSuffixClose"))
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub stylesHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLStylesTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for stylesEdit

'@Ignore ParameterNotUsed
Public Sub stylesEdit_onAction(ByVal control As IRibbonControl)
    RestoreStyleDesigner
End Sub

'@Ignore ParameterNotUsed
Public Sub stylesEdit_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    enabled = False

    If ActiveSheet.name <> StylesSheet.name Then Exit Sub
    If Not TypeOf Selection Is Range Then Exit Sub
    If Selection.rows.count <> 1 Then Exit Sub

    Dim row As Long
    row = Selection.row

    Dim typeCol As Long
    typeCol = GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)

    Dim styleType As String
    styleType = StylesSheet.Cells(row, typeCol).value

    If Not (styleType = TYPE_NODE Or styleType = TYPE_EDGE Or styleType = TYPE_SUBGRAPH_OPEN) Then Exit Sub

    enabled = True
End Sub
