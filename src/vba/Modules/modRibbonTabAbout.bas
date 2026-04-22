Attribute VB_Name = "modRibbonTabAbout"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabAbout
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Callback bridge for the "Info" Ribbon Tab, providing hyperlinks to
'   project resources, community pages, and author information.
'
' RESPONSIBILITIES:
'   - Dispatch IRibbonControl callbacks for all Info tab controls.
'   - Open external URLs (GitHub, SourceForge, LinkedIn, Website, Email).
'   - Maintain cross-platform hyperlink behavior.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Named Ranges: HelpURLAboutTab.
'   - Worksheets: SettingsSheet.
'
' CROSS-PLATFORM NOTES:
'   - Fully supported on Windows and macOS.
'
' ERROR HANDLING:
'   - Minimal; relies on Excel hyperlink engine.
'
' RELATED WIKI PAGES:
'   - Web Presence & Community
'   - User Interface: Ribbon Tabs
' =============================================================================

Option Explicit

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub aboutHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLAboutTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for aboutE2G

'@Ignore ParameterNotUsed
Public Sub aboutE2G_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:="https://exceltographviz.com/", NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for aboutSourceForge

'@Ignore ParameterNotUsed
Public Sub aboutSourceForge_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:="https://sourceforge.net/projects/relationship-visualizer/", NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for aboutGitHub

'@Ignore ParameterNotUsed
Public Sub aboutGithub_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:="https://github.com/jjlong150/ExcelToGraphviz", NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for aboutLinkedIn

'@Ignore ParameterNotUsed
Public Sub aboutAuthorLinkedIn_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:="https://www.linkedin.com/in/jeffreyjlong/", NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for aboutAuthorEmail

'@Ignore ParameterNotUsed
Public Sub aboutAuthorEmail_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:="mailto:Jeffrey Long <relationship.visualizer@gmail.com>", NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for aboutAuthorEmail

'@Ignore ParameterNotUsed
Public Sub aboutBuyMeACoffee_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:="https://buymeacoffee.com/exceltographviz", NewWindow:=True
End Sub


