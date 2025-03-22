Attribute VB_Name = "modWorksheetSource"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Source")

Option Explicit

Public Sub LaunchGVEdit()
    If SearchPathForFile("gvedit.exe") Then
        '@Ignore VariableNotUsed
        Dim taskId As Variant
        '@Ignore AssignmentNotUsed
        taskId = Shell("gvedit.exe", 1)
    End If
End Sub

Public Sub DisplaySourceInWorksheet(ByVal dotSource As String)

    Dim row As Long
    Dim parsedFileData As Variant
    
    ' Remove any existing content
    ClearSourceWorksheet
        
    ' Get the layout of the "source" worksheet
    Dim source As sourceWorksheet
    source = GetSettingsForSourceWorksheet()
    
    ' Initialize row counters
    row = source.firstRow

    ' Create column headings
    SourceSheet.Cells.Item(source.headingRow, source.lineNumberColumn).value = GetLabel("worksheetSourceLine")
    SourceSheet.Cells.Item(source.headingRow, source.sourceColumn).value = GetLabel("worksheetSourceGraphvizSource")
    
    ' Split entire file into array - lines delimited by LF
    parsedFileData = split(dotSource, vbLf)
    
    ' Transfer the array of lines to the worksheet in one swift action
    Dim sourceCol As String
    sourceCol = ConvertColumnNumberToLetters(source.sourceColumn)
    
    Dim writeToRange As String
    writeToRange = sourceCol & row & ":" & sourceCol & (row + (UBound(parsedFileData) - LBound(parsedFileData)))
    SourceSheet.Range(writeToRange).value = Application.Transpose(parsedFileData)

    ' Update the line numbers
    UpdateSourceWorksheetLineNumbers

End Sub

Public Sub ClearSourceWorksheet()
    
    ' The sheet to be cleared
    Dim worksheetName As String
    worksheetName = SourceSheet.name
    
    ' Get the layout of the 'data' worksheet
    Dim sourceLayout As sourceWorksheet
    sourceLayout = GetSettingsForSourceWorksheet()

    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With SourceSheet.UsedRange
        lastRow = .Cells.Item(.Cells.count).row
    End With

    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < sourceLayout.firstRow Then
        lastRow = sourceLayout.firstRow
    End If
    
    ' Determine the columns to clear
    Dim lastColumn As Long
    lastColumn = GetLastColumn(worksheetName, sourceLayout.headingRow)

    ' Remove any existing content
    Dim cellRange As String
    cellRange = "A" & sourceLayout.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    SourceSheet.Range(cellRange).ClearContents
    
End Sub

Public Sub SourceWorksheetToFile(ByVal filename As String)
    ' Get the layout of the "source" worksheet
    Dim source As sourceWorksheet
    source = GetSettingsForSourceWorksheet()

    Dim rowNumber As Long
    
#If Mac Then

    Dim gvSource As String
    gvSource = vbNullString
    
    Dim lastRow As Long
    With SourceSheet.UsedRange
        lastRow = .Cells(.Cells.count).row
    End With

    For rowNumber = source.firstRow To lastRow
        gvSource = gvSource & SourceSheet.Cells(rowNumber, source.sourceColumn).value & vbLf
    Next rowNumber
    
    WriteTextToFile gvSource, filename
    
#Else
    
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    If utf8Stream Is Nothing Then GoTo EndMacro
    
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")
    If binaryStream Is Nothing Then GoTo EndMacro
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    Dim lastRow As Long
    With SourceSheet.UsedRange
        lastRow = .Cells.Item(.Cells.count).row
    End With

    For rowNumber = source.firstRow To lastRow
        utf8Stream.WriteText SourceSheet.Cells.Item(rowNumber, source.sourceColumn).value & vbLf
    Next rowNumber
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    utf8Stream.position = 3
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile filename, SaveOptionsEnum.adSaveCreateOverWrite

EndMacro:
    ' Clean up our objects so we don't get a memory leak
    If Not (utf8Stream Is Nothing) Then
        If (utf8Stream.state And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then utf8Stream.Close
        Set utf8Stream = Nothing
    End If
    
    If Not (binaryStream Is Nothing) Then
        If (binaryStream.state And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then binaryStream.Close
        Set binaryStream = Nothing
    End If
#End If
End Sub

Public Sub CopySourceCodeToClipboard()
#If Not Mac Then

    ' Get the layout of the "source" worksheet
    Dim source As sourceWorksheet
    source = GetSettingsForSourceWorksheet()
    
    ' Pull all the rows into a single string
    Dim dotSource As String
    dotSource = vbNullString
    
    Dim lastRow As Long
    With SourceSheet.UsedRange
        lastRow = .Cells.Item(.Cells.count).row
    End With

    Dim i As Long
    For i = source.firstRow To lastRow
        dotSource = dotSource & SourceSheet.Cells.Item(i, source.sourceColumn).value
    Next i
    
    If ClipBoard_SetData(dotSource) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyFailed"), 5
    End If
    
#End If
End Sub

Public Sub CreateGraphFromSourceToWorksheet()
    ' Clear the status bar
    ClearStatusBar

    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(DataSheet.name)

    ' Remove any existing graph image from the target worksheet
    Dim displayDataSheetName As String
    displayDataSheetName = GraphSheet.name
            
    ActiveWorkbook.Sheets.[_Default](displayDataSheetName).Activate
    DeleteAllPictures displayDataSheetName

    ' Instantiate a new Graphviz object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz

    ' Build file names
    graphvizObj.OutputDirectory = GetTempDirectory()
    graphvizObj.FilenameBase = "RelationshipVisualizer"
    graphvizObj.GraphFormat = ini.graph.imageTypeWorksheet

    ' Create the '.gv' Graphviz source code file from the source worksheet
    SourceWorksheetToFile graphvizObj.GraphvizFilename
   
    ' Convert the Graphviz source code into a diagram
    graphvizObj.CaptureMessages = ini.console.logToConsole
    graphvizObj.Verbose = ini.console.graphvizVerbose
    graphvizObj.CommandLineParameters = ini.CommandLine.parameters
    graphvizObj.GraphLayout = ini.graph.engine
    graphvizObj.GraphvizPath = ini.CommandLine.GraphvizPath
    
    graphvizObj.RenderGraph
    
    ' Display any console output first
    DisplayTextOnConsoleWorksheet graphvizObj.GraphvizCommand, graphvizObj.GraphvizMessages
    
    ' Display the image
    If FileExists(graphvizObj.DiagramFilename) Then
        '@Ignore VariableNotUsed
        Dim shapeObject As Shape
        '@Ignore AssignmentNotUsed
        Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range("B2"), False, True)
        Set shapeObject = Nothing
    Else
        MsgBox GetMessage("msgboxNoGraphCreated"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If

    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Clean up objects
    Set graphvizObj = Nothing
End Sub

Public Sub CreateGraphFromSourceToFile()
    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(DataSheet.name)

    ' Determine output directory, and build file names
    If ini.output.directory = vbNullString Then
        MsgBox GetMessage("msgboxNoDirectorySpecified"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        SettingsSheet.Activate
        ActiveSheet.Range("OutputDirectory").Activate
        Exit Sub
    End If

    ' Get the file name, minus the file extension
    Dim styleColumn As Long
    styleColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    ' Compose the filename
    If Not FileLocationProvided(ini) Then
        Exit Sub
    End If

    ' Instantiate a new Graphviz object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz

    ' Create the filenames
    graphvizObj.OutputDirectory = ini.output.directory
    graphvizObj.FilenameBase = GetFilenameBase(ini, styleColumn)
    graphvizObj.GraphFormat = ini.graph.imageTypeFile
    
    ' Create the '.gv' Graphviz source code file from the source worksheet
    SourceWorksheetToFile graphvizObj.GraphvizFilename
    If Not FileExists(graphvizObj.GraphvizFilename) Then
        MsgBox GetMessage("msgboxSourceFileNotFound"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        Set graphvizObj = Nothing
        Exit Sub
    End If

    ' Render the graph from the Graphviz source
    graphvizObj.CaptureMessages = ini.console.logToConsole
    graphvizObj.Verbose = ini.console.graphvizVerbose
    graphvizObj.CommandLineParameters = ini.CommandLine.parameters
    graphvizObj.GraphLayout = ini.graph.engine
    graphvizObj.GraphvizPath = ini.CommandLine.GraphvizPath
    
    graphvizObj.RenderGraph
    
    ' Display any console output first
    DisplayTextOnConsoleWorksheet graphvizObj.GraphvizCommand, graphvizObj.GraphvizMessages
    
    ' If the diagram file is not there, then Graphviz failed
    If FileExists(graphvizObj.DiagramFilename) Then
        MsgBox GetMessage("msgboxGraphFilenameIs") & vbNewLine & graphvizObj.DiagramFilename, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    Else
        MsgBox GetMessage("msgboxNoGraphCreated"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If

    ' Delete the command file if disposition is 'delete'
    If ini.graph.fileDisposition = "delete" Then
        DeleteFile graphvizObj.GraphvizFilename
    End If

    ' Clean up objects
    Set graphvizObj = Nothing
End Sub

Public Sub UpdateSourceWorksheetLineNumbers()
    Dim cellRange As String
    Dim rowLast As Long
    Dim sourceLayout As sourceWorksheet
    Dim lineNumCol As String
    
    ' Get the layout of the 'data' worksheet
    sourceLayout = GetSettingsForSourceWorksheet()
    
    ' Determine the range of the cells which need to be cleared
    With SourceSheet.UsedRange
        rowLast = .Cells.Item(.Cells.count).row
    End With

    ' If the worksheet is already empty we do not want to wipe out the heading row
    If rowLast < sourceLayout.firstRow Then
        rowLast = sourceLayout.firstRow
    End If
    
    ' Determine the columns to clear
    lineNumCol = ConvertColumnNumberToLetters(sourceLayout.lineNumberColumn)

    ' Remove any existing content
    cellRange = lineNumCol & sourceLayout.firstRow & ":" & lineNumCol & rowLast
    SourceSheet.Range(cellRange).ClearContents
    
    ' Renumber the rows
    Dim rowNumber As Long
    Dim lineCnt As Long
    lineCnt = 1
    
    Dim lastRow As Long
    With SourceSheet.UsedRange
        lastRow = .Cells.Item(.Cells.count).row
    End With

    For rowNumber = sourceLayout.firstRow To lastRow
        SourceSheet.Cells.Item(rowNumber, sourceLayout.lineNumberColumn).value = lineCnt
        lineCnt = lineCnt + 1
    Next rowNumber

End Sub

Public Sub ClearSource()
    ClearSourceForm
    ClearSourceWorksheet
End Sub

Public Sub ShowSource(ByVal dotSource As String)
    DisplaySourceInForm dotSource
    
    If GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE) Then
        DisplaySourceInWorksheet dotSource
    End If
End Sub
