Attribute VB_Name = "modRibbonTabSvg"
'@IgnoreModule ProcedureNotUsed
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")

Option Explicit

' ===========================================================================
' Callbacks for svgPostprocess

'@Ignore ParameterNotUsed
Public Sub svgPostprocess_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_POST_PROCESS_SVG).value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
End Sub

'@Ignore ParameterNotUsed
Public Sub svgPostprocess_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_POST_PROCESS_SVG)
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub svgHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSvgTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for svgEditCell

'@Ignore ParameterNotUsed
Public Sub svgEditCell_onAction(ByVal control As IRibbonControl)
    CellValueEditForm.show
End Sub

'@Ignore ParameterNotUsed
Public Sub svgEditCell_getEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    If ActiveSheet.name <> SvgSheet.name Then
        enabled = False
    ElseIf Selection.Cells.count > 1 Then
        enabled = False
    ElseIf ActiveCell.HasFormula Then
        enabled = False
    Else
        enabled = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub svgClipboard_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

'@Ignore ParameterNotUsed
Public Sub svgClipboard_onAction(ByVal control As IRibbonControl)
#If Not Mac Then
    
    If ClipBoard_SetData(ActiveCell.value) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySvgSuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySvgFailed"), 5
    End If
    
#End If
End Sub


