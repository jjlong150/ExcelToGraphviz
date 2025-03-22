Attribute VB_Name = "modWorksheetStyles"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Styles")

Option Explicit

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub GenerateStylesPreviewAll()
    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    Dim styleCount As Long
    styleCount = styles.lastRow - styles.firstRow + 1
    
    ShowProgressIndicator GetLabel("stylesProgressIndicator")

    ' Loop through the rows, generating preview images from the format strings
    Dim row As Long
    For row = styles.firstRow To styles.lastRow
        UpdateProgressIndicator (((row - 1) * 100) / styleCount)
        GenerateStylesPreview row
    Next row
    
    HideProgressIndicator
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub GenerateStylesPreview(ByRef row As Long)

    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    If StylesSheet.Cells.Item(row, styles.flagColumn) = FLAG_COMMENT Then
        Exit Sub
    End If
    
    If StylesSheet.Cells.Item(row, styles.nameColumn).value = vbNullString Then
        Exit Sub
    End If

    ' Determine the last column of view switches. Allow for a blank column, followed by the preview column
    Dim lastCol As Long
    lastCol = GetLastColumn(StylesSheet.name, row) + 2
    
    ' Convert column number to letter
    Dim previewColumn As String
    previewColumn = ConvertColumnNumberToLetters(lastCol)
    
    ' Generating preview images from the format strings
    Dim styleName As String
    styleName = StylesSheet.Cells.Item(row, styles.nameColumn).value
    Application.StatusBar = styleName
    
    Dim styleType As String
    styleType = StylesSheet.Cells.Item(row, styles.typeColumn).value
    
    Dim graphvizSource As String
    Select Case styleType
        Case TYPE_NODE
            graphvizSource = "digraph " & AddQuotes("Node Preview") & " { bgcolor=transparent " & AddQuotes(styleName) & " [label=" & AddQuotes(replace(styleName, " ", "\n")) & " " & StylesSheet.Cells.Item(row, styles.formatColumn).value & "] }"
        Case TYPE_EDGE
            graphvizSource = "digraph " & AddQuotes("Edge Preview") & " { bgcolor=transparent layout=dot rankdir=LR tail[shape=point color=invis]; head[shape=point color=invis]; tail->head[label=" & AddQuotes(styleName) & " " & StylesSheet.Cells.Item(row, styles.formatColumn).value & "] }"
        Case TYPE_SUBGRAPH_OPEN
            graphvizSource = "digraph " & AddQuotes("Cluster Preview") & " { bgcolor=transparent layout=dot rankdir=LR subgraph cluster_1 { label=" & AddQuotes(styleName) & " " & StylesSheet.Cells.Item(row, styles.formatColumn).value & " node[style=filled fillcolor=white]; A->Z; } }"
        Case Else
    End Select
    
    If graphvizSource <> vbNullString Then
        PreviewStyleAndAutosize styleName, graphvizSource, previewColumn, row
    End If
    
    ' Repaint the screen
    DoEvents
End Sub

Public Sub ClearStylesPreview()
    ' Delete the images
    DeleteAllPictures StylesSheet.name
    
    ' Get the 'styles' sheet layout
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()

    ' Reset the row height
    Dim row As Long
    For row = styles.firstRow To styles.lastRow
        StylesSheet.rows.Item(row).AutoFit
    Next row
End Sub


Public Sub PreviewStyleAndAutosize(ByVal styleName As String, ByVal graphvizSource As String, ByVal targetCol As String, ByRef targetRow As Long)
    
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    ' Instantiate a Graphviz object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz
    
    ' Prepare the file names
    graphvizObj.OutputDirectory = GetTempDirectory()
    graphvizObj.FilenameBase = styleName
    graphvizObj.GraphFormat = "png"

    ' Determine where to place the preview image
    Dim targetCell As String
    targetCell = "$" & targetCol & "$" & targetRow
    
    ' Remove any image from a previous run of the macro
    DeleteCellPictures StylesSheet.name, targetCell
      
    ' Write the Graphviz data to a file so it can be sent to a rendering engine
    graphvizObj.graphvizSource = graphvizSource
    graphvizObj.SourceToFile
    
    ' Display source if debugging
    ShowSource graphvizSource
    
    ' Generate an image using graphviz
    graphvizObj.CaptureMessages = GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
    graphvizObj.Verbose = RunGraphvizInVerboseMode()
    graphvizObj.CommandLineParameters = SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).value
    graphvizObj.GraphLayout = GetGraphvizEngine()
    graphvizObj.GraphvizPath = SettingsSheet.Range(SETTINGS_GV_PATH).value
    
    graphvizObj.RenderGraph

    ' Display any console output first
    DisplayTextOnConsoleWorksheet graphvizObj.GraphvizCommand, graphvizObj.GraphvizMessages
    
    ' Display the generated image
    '@Ignore VariableNotUsed
    Dim shapeObject As Shape
    '@Ignore AssignmentNotUsed
    Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range(targetCell), False, True)
    Set shapeObject = Nothing
              
    ' Resize the row height to hold the image
#If Mac Then
    ActiveSheet.rows(targetRow).rowHeight = 126 ' Unable to get image size easily, so default to 1.75"
#Else
    Dim wia As Object
    On Error GoTo bypassResize
    Set wia = CreateObject("WIA.ImageFile")
    If Not wia Is Nothing Then
        wia.LoadFile graphvizObj.DiagramFilename
        ActiveSheet.rows(targetRow).rowHeight = determineRowHeight(targetRow, wia.height)
    End If
    
bypassResize:
    On Error GoTo 0
    Set wia = Nothing
#End If

    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Release the Graphviz object
    Set graphvizObj = Nothing
End Sub

Private Function determineRowHeight(ByRef targetRow As Long, ByVal imageHeight As Long) As Long
    ' Convert pixels to points, assuming screen is 96 DPI, and there are 72 points to one inch.
    Dim rowHeight As Long
    rowHeight = ((imageHeight * 72) / 96)
    
    ' Lower bound - Never set the row height less than 20 points
    If rowHeight <= 20 Then
        rowHeight = 20
    End If
    
    ' Upper bound - Excel does not permit a row height greater than 546 points
    If rowHeight >= 546 Then
        rowHeight = 546
    End If
    
    ' If the calculated row height is less than the current row height
    ' then return the larger current row height.
    If rowHeight <= ActiveSheet.rows(targetRow).rowHeight Then
        rowHeight = ActiveSheet.rows(targetRow).rowHeight
    End If
    
    determineRowHeight = rowHeight
End Function
