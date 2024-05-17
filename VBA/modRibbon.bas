Attribute VB_Name = "modRibbon"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Loader")
'@IgnoreModule ProcedureNotUsed

Option Explicit

Private internalMyRibbon As IRibbonUI
Private internalMyTag As String

' Get/Let to encapsulate myTag
Public Static Property Get myTag() As String
    myTag = internalMyTag
End Property

Public Static Property Let myTag(ByVal tag As String)
    internalMyTag = tag
End Property

' Get/Let to encapsulate myRibbon
Public Static Property Get myRibbon() As IRibbonUI
    Set myRibbon = internalMyRibbon
End Property

Public Static Property Let myRibbon(ByVal ribbon As IRibbonUI)
    Set internalMyRibbon = ribbon
End Property

' Load the ribbon
Public Sub ribbon_onLoad(ByVal ribbon As IRibbonUI)
    '@Ignore ValueRequired
    myRibbon = ribbon
    Application.OnTime Now + TimeValue("00:00:01"), "ribbon_activateTab"
End Sub

Public Sub ribbon_activateTab()
    On Error GoTo ErrorHandler

    ' Defer initialization of 'settings' worksheet until workbook
    ' startup is complete
    TabSelectGraphOptions
    
    ' Show the appropriate tab for the worksheet displayed
    If ActiveSheet.name = StyleDesignerSheet.name Then
        ActivateTab (RIBBON_TAB_STYLE_DESIGNER)
    ElseIf ActiveSheet.name = SourceSheet.name Then
        ActivateTab (RIBBON_TAB_SOURCE)
    ElseIf ActiveSheet.name = SqlSheet.name Then
        ActivateTab (RIBBON_TAB_SQL)
    ElseIf ActiveSheet.name = SvgSheet.name Then
        ActivateTab (RIBBON_TAB_SVG)
    Else
        ActivateTab (RIBBON_TAB_GRAPHVIZ)
    End If
    
    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub ribbon_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)

    If control.ID = RIBBON_TAB_SQL Then
        visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SQL)
    ElseIf control.ID = RIBBON_TAB_SOURCE Then
        visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE)
    ElseIf control.ID = RIBBON_TAB_SVG Then
        visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SVG)
    Else
        visible = True
    End If

End Sub

Public Sub RefreshRibbon(ByVal tag As String)
    On Error GoTo ErrorHandler
    myTag = tag
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        MsgBox "Error refreshing the ribbon. Save and reopen this file."
    Else
        myRibbon.Invalidate
        If Err.number <> 0 Then
            ' This message cannot be localized due to error state.
            MsgBox "Lost the Ribbon object. Save this file, close worksbook, and reopen."
        End If
    End If

    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub InvalidateRibbonControl(ByVal controlName As String)
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        UpdateStatusBar replace("Error updating the ribbon for control named '{controlName}'. Save and reopen this file.", "{controlName}", controlName)
    Else
        myRibbon.InvalidateControl controlName
    End If
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub ActivateTab(ByVal tabName As String)
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        UpdateStatusBar replace("Error activating a ribbon tab named '{tabName}'. Save and reopen this file.", "{tabName}", tabName)
    Else
        myRibbon.ActivateTab tabName
    End If
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub SyncHelpToggleButtons()

    ' Graphviz tab
    InvalidateRibbonControl RIBBON_CTL_HELP_SHAPES
    InvalidateRibbonControl RIBBON_CTL_HELP_COLORS
    InvalidateRibbonControl RIBBON_CTL_HELP_ATTRIBUTES
    
    ' Style Designer tab
    InvalidateRibbonControl RIBBON_CTL_HELP_DESIGN_SHAPES
    InvalidateRibbonControl RIBBON_CTL_HELP_DESIGN_COLORS
    
    ' Tools tab
    InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_SHAPES
    InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_COLORS
    InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_ATTRIBUTES

End Sub

' ===========================================================================
' Ribbon Callbacks for prefixed ribbon buttons

Public Sub button_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.ID & BUTTON_SUFFIX_VISIBLE)
End Sub

Public Sub button_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.ID & BUTTON_SUFFIX_TEXT).Value
End Sub

Public Sub button_getScreentip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.ID & BUTTON_SUFFIX_SCREENTIP).Value
End Sub

Public Sub button_getSupertip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.ID & BUTTON_SUFFIX_SUPERTIP).Value
End Sub

