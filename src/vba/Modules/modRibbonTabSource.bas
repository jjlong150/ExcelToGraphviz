Attribute VB_Name = "modRibbonTabSource"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ProcedureNotUsed

Option Explicit

' ===========================================================================
' Ribbon callbacks for source worksheet
' ===========================================================================

' ===========================================================================
' Callbacks for sourceCreate

'@Ignore ParameterNotUsed
Public Sub sourceCreate_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    ClearSource
    ShowSource CreateGraphSource()
    OptimizeCode_End
End Sub

'@Ignore ParameterNotUsed
Public Sub sourceCreate_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for debugSource

'@Ignore ParameterNotUsed
Public Sub debugSource_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    ShowSourceForm
    ClearSource
    ShowSource CreateGraphSource()
    OptimizeCode_End
End Sub


' ===========================================================================
' Callbacks for sourceIndent

'@Ignore ParameterNotUsed
Public Sub sourceIndent_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    OptimizeCode_Begin
    SettingsSheet.Range(SETTINGS_SOURCE_INDENT).value = Mid$(controlId, Len("source_") + 1)
    ClearSource
    ShowSource CreateGraphSource()
    OptimizeCode_End
End Sub

'@Ignore ParameterNotUsed
Public Sub sourceIndent_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = CLng(SettingsSheet.Range(SETTINGS_SOURCE_INDENT))
End Sub

' ===========================================================================
' Callbacks for sourceCopy

'@Ignore ParameterNotUsed
Public Sub sourceCopy_onAction(ByVal control As IRibbonControl)
    CopySourceCodeToClipboard
End Sub

'@Ignore ParameterNotUsed
Public Sub sourceCopy_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub sourceCopy_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = True
#End If
End Sub

' ===========================================================================
' Callbacks for sourceClear

'@Ignore ParameterNotUsed
Public Sub sourceClear_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    ClearSourceWorksheet
    ClearSourceForm
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for sourceSave

'@Ignore ParameterNotUsed
Public Sub sourceSave_onAction(ByVal control As IRibbonControl)
    Dim fileName As String
    
#If Mac Then
    fileName = RunAppleScriptTask("getSaveAsFileName", GRAPHVIZ_EXTENSION)
#Else
    fileName = GetSaveAsFilename("Graphviz Files (*.gv), *gv")
#End If

    If fileName <> vbNullString Then
        SourceWorksheetToFile (fileName)
        MsgBox GetMessage("msgboxSavedToFile") & vbNewLine & fileName, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub sourceSave_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for sourceNumber

'@Ignore ParameterNotUsed
Public Sub sourceNumber_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    UpdateSourceWorksheetLineNumbers
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for sourceGraphToWorksheet

'@Ignore ParameterNotUsed
Public Sub sourceGraphToWorksheet_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    CreateGraphFromSourceToWorksheet
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for sourceGraphToFile

'@Ignore ParameterNotUsed
Public Sub sourceGraphToFile_onAction(ByVal control As IRibbonControl)
    OptimizeCode_Begin
    CreateGraphFromSourceToFile
    OptimizeCode_End
End Sub

' ===========================================================================
' Callbacks for launchGVEdit

'@Ignore ParameterNotUsed
Public Sub launchGVEdit_onAction(ByVal control As IRibbonControl)
    LaunchGVEdit
End Sub

'@Ignore ParameterNotUsed
Public Sub launchGVEdit_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SearchPathForFile("gvedit.exe")
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub launchGVEdit_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
#If Mac Then
    visible = False
#Else
    visible = SearchPathForFile("gvedit.exe")
#End If
End Sub

' ===========================================================================
' Callbacks for Web References

'@Ignore ParameterNotUsed
Public Sub source_web_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = extTabGroup_getVisible("SourceWeb", 5)
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub sourceHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSourceTab").value, NewWindow:=True
End Sub


