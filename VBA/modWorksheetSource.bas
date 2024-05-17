Attribute VB_Name = "modWorksheetSource"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

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

Public Sub DisplayFileOnSourceWorksheet(ByVal filename As String)

    Dim row As Long
    Dim line As Long
    Dim parsedFileData As Variant
    
    ' Remove any existing content
    ClearSourceWorksheet
        
    ' Get the layout of the "source" worksheet
    Dim source As sourceWorksheet
    source = GetSettingsForSourceWorksheet()
    
    ' Initialize row counters
    row = source.firstRow
    line = 1

    ' Create column headings
    SourceSheet.Cells.Item(source.headingRow, source.lineNumberColumn).Value = GetLabel("worksheetSourceLine")
    SourceSheet.Cells.Item(source.headingRow, source.sourceColumn).Value = GetLabel("worksheetSourceGraphvizSource")
    
#If Mac Then
    Dim textLine As String
    Dim fileNum As Integer
    fileNum = FreeFile()
    Open filename For Input As #fileNum
    
    While Not EOF(fileNum)
        Line Input #fileNum, textLine ' read in data 1 line at a time
        SourceSheet.Cells(row, source.lineNumberColumn).Value = line
        SourceSheet.Cells(row, source.sourceColumn).Value = textLine
        row = row + 1
        line = line + 1
    Wend
    Close #fileNum
#Else    ' Read the file into a stream object
    Dim textStream As Object
    Set textStream = CreateObject("ADODB.Stream")
    
    textStream.Charset = UTF8_CHARSET
    textStream.Open
    textStream.LoadFromFile filename
    
    parsedFileData = Split(textStream.ReadText, vbLf) 'split entire file into array - lines delimited by LF
    
    Dim textLine As Variant
    For Each textLine In parsedFileData
        SourceSheet.Cells.Item(row, source.lineNumberColumn).Value = line
        SourceSheet.Cells.Item(row, source.sourceColumn).Value = textLine
        row = row + 1
        line = line + 1
    Next textLine
        
    ' Clean up our objects so we don't get a memory leak
    textStream.Close
    Set textStream = Nothing
#End If
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
        lastRow = .Cells.Item(.Cells.Count).row
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
        lastRow = .Cells(.Cells.Count).row
    End With

    For rowNumber = source.firstRow To lastRow
        gvSource = gvSource & SourceSheet.Cells(rowNumber, source.sourceColumn).Value & vbLf
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
        lastRow = .Cells.Item(.Cells.Count).row
    End With

    For rowNumber = source.firstRow To lastRow
        utf8Stream.WriteText SourceSheet.Cells.Item(rowNumber, source.sourceColumn).Value & vbLf
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
        lastRow = .Cells.Item(.Cells.Count).row
    End With

    Dim i As Long
    For i = source.firstRow To lastRow
        dotSource = dotSource & SourceSheet.Cells.Item(i, source.sourceColumn).Value
    Next i
    
    'Cast dotSource to variant for 64-bit VBA support
    Dim vDotSource As Variant
    vDotSource = dotSource
    
    'Write source code to the clipboard
    On Error Resume Next
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", vDotSource
            If .GetData("text") <> vDotSource Then
                UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyFailed"), 5
            Else
                UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySuccess"), 5
            End If
        End With
    End With
    On Error GoTo 0
    
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
    Dim targetCell As String

    displayDataSheetName = GraphSheet.name
    targetCell = "B2"
            
    ActiveWorkbook.Sheets.[_Default](displayDataSheetName).Activate
    DeleteAllPictures displayDataSheetName

    ' Determine output directory, and build file names
    Dim outputDirectory As String
    outputDirectory = GetTempDirectory()

    Dim filenameBase As String
    Dim graphvizFile As String
    Dim diagramFile As String

    ' Get the file name, minus the file extension
    filenameBase = outputDirectory & Application.pathSeparator & "RelationshipVisualizer"

    ' Add the file extensions
    graphvizFile = filenameBase & GRAPHVIZ_EXTENSION
    diagramFile = filenameBase & "." & ini.graph.imageTypeWorksheet

    ' Create the '.gv' Graphviz source code file from the source worksheet
    SourceWorksheetToFile graphvizFile

    ' Convert the Graphviz source code into a diagram
    Dim ret As Long

    ret = CreateGraphDiagram(graphvizFile, diagramFile, _
                             ini.graph.imageTypeWorksheet, ini.graph.engine, _
                             ini.commandLine.parameters, CLng(ini.graph.maxSeconds) * 1000)
    
    If ret = ShellAndWaitResult.success Then    ' Show the graph image
        If FileExists(diagramFile) Then
            '@Ignore VariableNotUsed
            Dim shapeObject As Shape
            '@Ignore AssignmentNotUsed
            Set shapeObject = InsertPicture(diagramFile, ActiveSheet.Range(targetCell), False, True)
            Set shapeObject = Nothing
        Else
            MsgBox GetMessage("msgboxNoGraphCreated"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        End If
    Else                                        ' ShellAndWait failed
        ShellAndWaitMessage ret
    End If

    ' Delete the temporary files
    DeleteFile graphvizFile
    DeleteFile diagramFile
    
End Sub

Public Sub CreateGraphFromSourceToFile()
    Dim graphvizFile As String
    Dim ret As Long
    Dim diagramFile As String
    Dim styleColumn As Long
    
    Dim filenameBase As String

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
    styleColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    ' Compose the filename
    If FileLocationProvided(ini) Then
        filenameBase = GetFilenameBase(ini, styleColumn)
    Else
        Exit Sub
    End If

    ' Create the filenames
    graphvizFile = filenameBase & GRAPHVIZ_EXTENSION      ' Input (Graphviz) source code filename
    diagramFile = filenameBase & "." & ini.graph.imageTypeFile ' Output (diagram) filename

    ' Create the '.gv' Graphviz source code file from the source worksheet
    SourceWorksheetToFile graphvizFile
    If Not FileExists(graphvizFile) Then
        MsgBox GetMessage("msgboxSourceFileNotFound"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        Exit Sub
    End If

    ' Convert source code into a graph diagram
    ret = CreateGraphDiagram(graphvizFile, diagramFile, ini.graph.imageTypeFile, _
                             ini.graph.engine, ini.commandLine.parameters, CLng(ini.graph.maxSeconds) * 1000)
    
    If ret <> ShellAndWaitResult.success Then   ' Inform user of failure
        ShellAndWaitMessage ret
    End If

    ' If the diagram file is not there, then Graphviz failed
    If FileExists(diagramFile) Then
        MsgBox GetMessage("msgboxGraphFilenameIs") & vbNewLine & diagramFile, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    Else
        MsgBox GetMessage("msgboxNoGraphCreated"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If

    ' Delete the command file if disposition is 'delete'
    If ini.graph.fileDisposition = "delete" Then
        DeleteFile graphvizFile
    End If

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
        rowLast = .Cells.Item(.Cells.Count).row
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
        lastRow = .Cells.Item(.Cells.Count).row
    End With

    For rowNumber = sourceLayout.firstRow To lastRow
        SourceSheet.Cells.Item(rowNumber, sourceLayout.lineNumberColumn).Value = lineCnt
        lineCnt = lineCnt + 1
    Next rowNumber

End Sub


