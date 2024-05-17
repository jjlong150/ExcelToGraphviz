VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Settings")

Option Explicit

Private Sub Worksheet_Activate()
    SyncHelpToggleButtons
    ActivateTab (RIBBON_TAB_GRAPHVIZ)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Whoa

    ' Detect changes in the settings dropdown lists
    If Not Intersect(Target, SettingsSheet.Range(SETTINGS_HELP_ATTRIBUTES)) Is Nothing Then
        Application.EnableEvents = False
        Application.screenUpdating = False
        
        SyncHelpToggleButtons
        
        If SettingsSheet.Range(SETTINGS_HELP_ATTRIBUTES).Value = TOGGLE_SHOW Then
            HelpAttributesSheet.visible = True
            HelpAttributesSheet.Activate
        Else
            HelpAttributesSheet.visible = False
            DataSheet.Activate
        End If
    
        Application.screenUpdating = True
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS)) Is Nothing Then
        Application.EnableEvents = False
        Application.screenUpdating = False
        
        InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_SETTINGS
        
        If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).Value = TOGGLE_SHOW Then
            SettingsSheet.visible = True
            SettingsSheet.Activate
        Else
            SettingsSheet.visible = False
            DataSheet.Activate
        End If
    
        Application.screenUpdating = True
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SOURCE)) Is Nothing Then
        Application.EnableEvents = False
        Application.screenUpdating = False
        
        InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_SOURCE
        
        If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SOURCE).Value = TOGGLE_SHOW Then
            CreateGraphSource
            SourceSheet.visible = True
            SourceSheet.Activate
            ActivateTab RIBBON_TAB_SOURCE
        Else
            SourceSheet.visible = False
            ClearSourceWorksheet
            DataSheet.Activate
            ActivateTab RIBBON_TAB_GRAPHVIZ
        End If
    
        Application.screenUpdating = True
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SQL)) Is Nothing Then
        Application.EnableEvents = False
        
        InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_SQL
        
        If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SQL).Value = TOGGLE_SHOW Then
            SqlSheet.visible = True
            SqlSheet.Activate
            ActivateTab RIBBON_TAB_SQL
        Else
            SqlSheet.visible = False
            DataSheet.Activate
            ActivateTab RIBBON_TAB_GRAPHVIZ
        End If
    
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SVG)) Is Nothing Then
        Application.EnableEvents = False
        
        InvalidateRibbonControl RIBBON_CTL_TOOLS_TOGGLE_SVG
        
        If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SVG).Value = TOGGLE_SHOW Then
            SvgSheet.visible = True
            SvgSheet.Activate
            ActivateTab RIBBON_TAB_SVG
        Else
            SvgSheet.visible = False
            DataSheet.Activate
            ActivateTab RIBBON_TAB_GRAPHVIZ
        End If
    
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_HELP_COLORS)) Is Nothing Then
        Application.EnableEvents = False
        Application.screenUpdating = False
        
        SyncHelpToggleButtons
        
        If SettingsSheet.Range(SETTINGS_HELP_COLORS).Value = TOGGLE_SHOW Then
            HelpColorsSheet.visible = True
            HelpColorsSheet.Activate
        Else
            HelpColorsSheet.visible = False
            DataSheet.Activate
        End If
    
        Application.screenUpdating = True
        
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_HELP_SHAPES)) Is Nothing Then
        Application.EnableEvents = False
        Application.screenUpdating = False
        
        SyncHelpToggleButtons
        
        If SettingsSheet.Range(SETTINGS_HELP_SHAPES).Value = TOGGLE_SHOW Then
            HelpShapesSheet.visible = True
            HelpShapesSheet.Activate
        Else
            HelpShapesSheet.visible = False
            DataSheet.Activate
        End If
    
        Application.screenUpdating = True
        
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY)) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_GRAPHVIZ
        
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_FILE_NAME)) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_GRAPHVIZ
    
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_SOURCE_INDENT)) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE

    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb1Text,SourceWeb1Visible,SourceWeb1URL,SourceWeb1Screentip,SourceWeb1Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb2Text,SourceWeb2Visible,SourceWeb2URL,SourceWeb2Screentip,SourceWeb2Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb3Text,SourceWeb3Visible,SourceWeb3URL,SourceWeb3Screentip,SourceWeb3Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb4Text,SourceWeb4Visible,SourceWeb4URL,SourceWeb4Screentip,SourceWeb4Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb5Text,SourceWeb5Visible,SourceWeb5URL,SourceWeb5Screentip,SourceWeb5Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE

    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb6Text,SourceWeb6Visible,SourceWeb6URL,SourceWeb6Screentip,SourceWeb6Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_SOURCE

    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode1Text,ExtCode1Visible,ExtCode1Sub,ExtCode1Screentip,ExtCode1Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode2Text,ExtCode2Visible,ExtCode2Sub,ExtCode2Screentip,ExtCode2Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode3Text,ExtCode3Visible,ExtCode3Sub,ExtCode3Screentip,ExtCode3Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode4Text,ExtCode4Visible,ExtCode4Sub,ExtCode4Screentip,ExtCode4Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode5Text,ExtCode5Visible,ExtCode5Sub,ExtCode5Screentip,ExtCode5Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode6Text,ExtCode6Visible,ExtCode6Sub,ExtCode6Screentip,ExtCode6Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb1Text,ExtWeb1Visible,ExtWeb1URL,ExtWeb1Screentip,ExtWeb1Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb2Text,ExtWeb2Visible,ExtWeb2URL,ExtWeb2Screentip,ExtWeb2Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb3Text,ExtWeb3Visible,ExtWeb3URL,ExtWeb3Screentip,ExtWeb3Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb4Text,ExtWeb4Visible,ExtWeb4URL,ExtWeb4Screentip,ExtWeb4Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb5Text,ExtWeb5Visible,ExtWeb5URL,ExtWeb5Screentip,ExtWeb5Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb6Text,ExtWeb6Visible,ExtWeb6URL,ExtWeb6Screentip,ExtWeb6Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
   
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtTabName,ExttabGroupNameCode,ExtTabGroupNameWeb")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon tag:=RIBBON_TAB_EXTENSIONS
    End If
    
    
LetsContinue:
    Application.EnableEvents = True
    Exit Sub
Whoa:
    MsgBox Err.number & " " & Err.Description, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    
    Resume LetsContinue

End Sub

