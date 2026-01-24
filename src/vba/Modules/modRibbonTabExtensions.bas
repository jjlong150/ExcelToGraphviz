Attribute VB_Name = "modRibbonTabExtensions"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ProcedureNotUsed

Option Explicit

' ===========================================================================
' Extensions Tab

'@Ignore ParameterNotUsed
Public Sub extTab_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_EXT_TAB_NAME).value
End Sub

'@Ignore ParameterNotUsed
Public Sub extTab_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = extTabGroup_getVisible(BUTTON_PREFIX_EXT_CODE, 6)
    If visible = False Then
        visible = extTabGroup_getVisible(BUTTON_PREFIX_EXT_WEB, 6)
    End If
End Sub

' ===========================================================================
' Custom Code Group

' Group visibility
'@Ignore ParameterNotUsed
Public Sub extTab_codeGroup_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = extTabGroup_getVisible(BUTTON_PREFIX_EXT_CODE, 6)
End Sub

' Group label
'@Ignore ParameterNotUsed
Public Sub extTab_codeGroup_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_EXT_TAB_GROUP_NAME_CODE).value
End Sub

' Buttons which invoke subroutines
'@Ignore ParameterNotUsed
Public Sub extCode_onAction(ByVal control As IRibbonControl)
    Dim subroutine As String
    subroutine = SettingsSheet.Range(control.id & BUTTON_SUFFIX_SUB).value
    Application.Run subroutine
End Sub

'@Ignore ParameterNotUsed
Public Sub extCode_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not (SettingsSheet.Range(control.id & BUTTON_SUFFIX_SUB).value = vbNullString)
End Sub

' ===========================================================================
' Web Resources Group

' Group visibility
'@Ignore ParameterNotUsed
Public Sub extTab_webGroup_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = extTabGroup_getVisible(BUTTON_PREFIX_EXT_WEB, 6)
End Sub

' Group label
'@Ignore ParameterNotUsed
Public Sub extTab_webGroup_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(SETTINGS_EXT_TAB_GROUP_NAME_WEB).value
End Sub

' Buttons which invoke web hyperlinks
'@Ignore ParameterNotUsed
Public Sub extWeb_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range(control.id & BUTTON_SUFFIX_URL).value, NewWindow:=True
End Sub

'@Ignore ParameterNotUsed
Public Sub extWeb_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not (SettingsSheet.Range(control.id & BUTTON_SUFFIX_URL).value = vbNullString)
End Sub

' ===========================================================================
' Utility routines

Public Function extTabGroup_getVisible(ByVal prefix As String, ByVal numButtons As Long) As Boolean
    
    extTabGroup_getVisible = False
    
    ' Determines if the group has any enabled settings. If so,
    ' show the group. If not, hide the group.
    Dim buttonCount As Long
    For buttonCount = 1 To numButtons
        If GetSettingBoolean(prefix & buttonCount & BUTTON_SUFFIX_VISIBLE) Then
            extTabGroup_getVisible = True
            Exit For
        End If
    Next buttonCount

End Function
