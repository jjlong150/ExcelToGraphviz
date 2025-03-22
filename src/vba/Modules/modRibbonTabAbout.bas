Attribute VB_Name = "modRibbonTabAbout"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
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


