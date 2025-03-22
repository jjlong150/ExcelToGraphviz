VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Settings")

Option Explicit

Private Sub Worksheet_Activate()
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Whoa

    ' Detect changes in the settings dropdown lists
    If Not Intersect(Target, SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY)) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
        
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_FILE_NAME)) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_SOURCE_INDENT)) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
        
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb1Text,SourceWeb1Visible,SourceWeb1URL,SourceWeb1Screentip,SourceWeb1Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb2Text,SourceWeb2Visible,SourceWeb2URL,SourceWeb2Screentip,SourceWeb2Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb3Text,SourceWeb3Visible,SourceWeb3URL,SourceWeb3Screentip,SourceWeb3Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb4Text,SourceWeb4Visible,SourceWeb4URL,SourceWeb4Screentip,SourceWeb4Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb5Text,SourceWeb5Visible,SourceWeb5URL,SourceWeb5Screentip,SourceWeb5Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon

    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb6Text,SourceWeb6Visible,SourceWeb6URL,SourceWeb6Screentip,SourceWeb6Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon

    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode1Text,ExtCode1Visible,ExtCode1Sub,ExtCode1Screentip,ExtCode1Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode2Text,ExtCode2Visible,ExtCode2Sub,ExtCode2Screentip,ExtCode2Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode3Text,ExtCode3Visible,ExtCode3Sub,ExtCode3Screentip,ExtCode3Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode4Text,ExtCode4Visible,ExtCode4Sub,ExtCode4Screentip,ExtCode4Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode5Text,ExtCode5Visible,ExtCode5Sub,ExtCode5Screentip,ExtCode5Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode6Text,ExtCode6Visible,ExtCode6Sub,ExtCode6Screentip,ExtCode6Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb1Text,ExtWeb1Visible,ExtWeb1URL,ExtWeb1Screentip,ExtWeb1Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb2Text,ExtWeb2Visible,ExtWeb2URL,ExtWeb2Screentip,ExtWeb2Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb3Text,ExtWeb3Visible,ExtWeb3URL,ExtWeb3Screentip,ExtWeb3Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb4Text,ExtWeb4Visible,ExtWeb4URL,ExtWeb4Screentip,ExtWeb4Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb5Text,ExtWeb5Visible,ExtWeb5URL,ExtWeb5Screentip,ExtWeb5Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb6Text,ExtWeb6Visible,ExtWeb6URL,ExtWeb6Screentip,ExtWeb6Supertip")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
   
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtTabName,ExttabGroupNameCode,ExtTabGroupNameWeb")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    End If
    
LetsContinue:
    Application.EnableEvents = True
    Exit Sub
Whoa:
    MsgBox Err.number & " " & Err.Description, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    
    Resume LetsContinue

End Sub

