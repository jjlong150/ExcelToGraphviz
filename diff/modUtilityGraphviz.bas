Attribute VB_Name = "modUtilityGraphviz"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Data")
'@IgnoreModule ProcedureNotUsed

Option Explicit
'MR Change


Public Sub ConvertFile(ByVal diagramFile As String, ByVal outputFormat As String)
    Dim wsh As Worksheet
    Dim shp As Shape
    Dim fil As Variant
    Dim cho As ChartObject
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    'quatsch
    Call fso.CopyFile(diagramFile, diagramFile & ".svg", True)
    Set shp = InsertPicture(diagramFile & ".svg", ActiveSheet.Range("BA1"), False, True)
    Set wsh = ActiveSheet
    Set cho = wsh.ChartObjects.Add(Left:=shp.Left, Top:=shp.Top, Width:=shp.Width, height:=shp.height)
    shp.Copy
    cho.Select
    ActiveChart.Paste
    ActiveChart.Export filename:=diagramFile, FilterName:=outputFormat
    cho.Delete
    shp.Delete
    Set shp = Nothing
End Sub

Public Function CreateGraphDiagram(ByVal filenameGraphviz As String, _
                                    ByVal diagramFile As String, _
                                    ByVal outputFormat As String, _
                                    ByVal graphEngine As String, _
                                    ByVal commandLineParameters As String, _
                                    ByVal timeout As Long) As Long
    
    CreateGraphDiagram = 0
    Dim graphvizCommand As String
    
#If Mac Then
     ' Generate the diagram which corresponds to the Graphviz file
    graphvizCommand = AddQuotes(filenameGraphviz) & " -T" & outputFormat & " -o " & AddQuotes(diagramFile) & " " & commandLineParameters
    
    ' Execute the command
    Dim applescriptResult As String
    applescriptResult = RunAppleScriptTask("runDot", graphvizCommand)
#Else
    On Error GoTo EndCreatePicture:
    
    ' Assume success
    CreateGraphDiagram = ShellAndWaitResult.success
    
     ' Generate the diagram which corresponds to the Graphviz file
' MR Change
    graphvizCommand = AddQuotes(ThisWorkbook.path & "\" & SettingsSheet.Range(SETTINGS_GV_PATH).Value & "dot-wasm.cmd") & " -K " & graphEngine & " " & AddQuotes(filenameGraphviz) & " -T svg > " & AddQuotes(diagramFile) & " " & commandLineParameters
    
    ' Execute the command in syncronous fashion for up to "timeout" seconds.
    CreateGraphDiagram = ShellAndWait(graphvizCommand, timeout, vbHide, PromptUser)
    
    If outputFormat <> "svg" Then
    ConvertFile diagramFile, outputFormat
    End If
    
    
    
    
    
EndCreatePicture:
    On Error GoTo 0
#End If
    
End Function

Public Sub AlertGraphvizNotFound(ByVal graphEngine As String)
#If Mac Then
    'TODO Port
#Else
    MsgBox replace(GetMessage("msgboxGraphvizNotFound"), "{graphEngine}", graphEngine), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
#End If
End Sub


