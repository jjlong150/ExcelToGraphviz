Attribute VB_Name = "modWorksheetSource"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetSource
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Source
'
' ROLE:
'   Manage the DOT Source Viewer subsystem, providing a synchronized,
'   editor-like interface for inspecting, editing, exporting, and rendering
'   Graphviz DOT code. Acts as the bridge between worksheet-based diagnostics,
'   the clipboard, external editors, and the Graphviz CLI.
'
' RESPONSIBILITIES:
'   - Source display and editing:
'       o DisplaySourceInWorksheet: populate the Source sheet with DOT code
'         using high-performance array transfers
'       o UpdateSourceWorksheetLineNumbers: maintain synchronized line numbers
'       o ClearSourceWorksheet: purge prior content while preserving headers
'
'   - External editor integration:
'       o LaunchGVEdit: attempt to launch gvedit.exe when available on PATH
'
'   - File export:
'       o SourceWorksheetToFile: export worksheet DOT to UTF-8 (BOM-free on Windows)
'       o StringToFile: export raw DOT strings with cross-platform encoding rules
'
'   - Clipboard integration:
'       o CopySourceCodeToClipboard: aggregate DOT source and copy to clipboard
'         (Windows-only)
'
'   - Direct rendering pipelines:
'       o CreateGraphFromSourceToWorksheet: render DOT from worksheet into Graph sheet
'       o VisualizeGraph: render DOT from a raw string
'       o CreateGraphFromSourceToFile: publish rendered diagrams to disk
'
' ARCHITECTURAL NOTES:
'   - Uses sourceWorksheet UDT + Named Range API to remain independent of
'     physical worksheet geometry.
'   - All DOT export routines enforce UTF-8 encoding; Windows paths use
'     ADODB.Stream with explicit BOM stripping for Graphviz compatibility.
'   - Bulk array writes (Application.Transpose) avoid cell-by-cell iteration
'     and preserve editor responsiveness.
'   - Integrates tightly with the Console subsystem for stdout/stderr capture.
'   - Supports manual override workflows by allowing direct DOT editing and
'     bypassing the Data-sheet parser.
'
' USAGE:
'   - Ideal for debugging DOT output, performing manual edits, exporting
'     source files, or rendering ad-hoc diagrams directly from DOT text.
'
' RELATED WIKI PAGES:
'   - DOT Source Viewer & Console Architecture
'   - File System Handshake (UTF-8 / BOM Rules)
'   - Transformation Pipeline (Manual DOT Overrides)
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: LaunchGVEdit
'
' PURPOSE:
'   Attempts to launch the external Graphviz "GVEdit" executable to allow
'   manual editing of the generated DOT source.
'
' TECHNICAL WORKFLOW:
'   1. SYSTEM DISCOVERY: Uses 'SearchPathForFile' to verify 'gvedit.exe'
'      exists within the system's PATH.
'   2. SHELL EXECUTION (Windows): Invokes 'Shell' to start the binary with
'      a normal focus window (1).
'
' TECHNICAL NOTES:
'   - Platform: Windows (Shell). This feature is typically restricted on
'     macOS due to the lack of a native 'gvedit' binary in most installs.
'   - DeepWiki Context: Part of the "DOT Source Viewer & Console" architecture.
' ==========================================================================
Public Sub LaunchGVEdit()
    If SearchPathForFile("gvedit.exe") Then
        '@Ignore VariableNotUsed
        Dim taskId As Variant
        '@Ignore AssignmentNotUsed
        taskId = Shell("gvedit.exe", 1)
    End If
End Sub

' ==========================================================================
' PROCEDURE: DisplaySourceInWorksheet
'
' PURPOSE:
'   Populates the 'Source' worksheet with the raw Graphviz DOT code, providing
'   a dedicated UI layer for inspection and debugging of the generated graph.
'
' TECHNICAL WORKFLOW:
'   1. WORKSPACE RESET: Invokes 'ClearSourceWorksheet' to purge previous data.
'   2. SCHEMA DISCOVERY: Retrieves the 'sourceWorksheet' UDT via
'      'GetSettingsForSourceWorksheet' to respect the Named Range "Contract."
'   3. HEADERS: Injects localized column headings for line numbers and source
'      text using the 'GetLabel' utility.
'   4. DATA PARSING: Splits the 'dotSource' string into a variant array
'      using the Line Feed (vbLf) delimiter.
'   5. BULK TRANSFER:
'      - Calculates a target 'writeToRange' based on array bounds.
'      - Uses 'Application.Transpose' to write the entire array to the
'        worksheet in a single operation, maximizing performance.
'   6. UI SYNCHRONIZATION: Triggers 'UpdateSourceWorksheetLineNumbers' to
'      finalize the visual display.
'
' TECHNICAL NOTES:
'   - Layer: UI / Presentation (Source Viewer).
'   - Performance: Avoids cell-by-cell iteration in favor of range-based
'     array writing.
' ==========================================================================
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
    SourceSheet.Cells.item(source.headingRow, source.lineNumberColumn).value = GetLabel("worksheetSourceLine")
    SourceSheet.Cells.item(source.headingRow, source.sourceColumn).value = GetLabel("worksheetSourceGraphvizSource")
    
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

' ==========================================================================
' PROCEDURE: ClearSourceWorksheet
'
' PURPOSE:
'   Cleans the 'Source' worksheet by removing all existing DOT code and
'   line numbers while preserving the protected header row.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA RESOLUTION: Obtains the 'sourceWorksheet' UDT to identify
'      the 'firstRow' and 'headingRow' markers (Named Range API Contract).
'   2. RANGE CALCULATION:
'      - Determines the vertical boundary using the worksheet's 'UsedRange'.
'      - Determines the horizontal boundary using 'GetLastColumn' to
'        ensure all diagnostic data is captured.
'   3. DATA PURGE: Constructs a dynamic 'cellRange' string and executes
'      '.ClearContents' to wipe the data without destroying cell formatting.
'   4. VALIDATION: Includes a safety check to ensure 'lastRow' never
'      encroaches upon the 'headingRow' if the sheet is already empty.
'
' TECHNICAL NOTES:
'   - Layer: UI / Presentation (Source Viewer).
'   - Strategy: Performance-optimized range clearing instead of row deletion.
' ==========================================================================
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
        lastRow = .Cells.item(.Cells.count).row
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

' ==========================================================================
' PROCEDURE: SourceWorksheetToFile
'
' PURPOSE:
'   Exports the Graphviz DOT source currently displayed in the worksheet to
'   a physical file, ensuring cross-platform compatibility and correct
'   UTF-8 encoding.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Obtains 'sourceWorksheet' UDT to resolve data columns
'      per the Named Range API "Contract."
'   2. MAC EXECUTION (#If Mac):
'      - Concatenates worksheet rows into a single string with Line Feed (vbLf)
'        delimiters.
'      - Invokes 'WriteTextToFile' to handle the macOS file system handshake.
'   3. WINDOWS EXECUTION (#Else):
'      - ADODB.STREAM ENCODING: Uses ADODB.Stream to force UTF-8 character
'        encoding, essential for international character support in Graphviz.
'      - BOM REMOVAL: Implements a specific logic to strip the "Byte Order Mark"
'        (BOM) by skipping the first 3 bytes before saving, ensuring compatibility
'        with some versions of the Graphviz parser that fail on BOM headers.
'   4. RESOURCE HYGIENE: Force-closes streams and releases objects to
'      prevent memory leaks or file locks.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (Windows ADO / Mac String Concatenation).
'   - DeepWiki Context: Crucial for the "Graphviz Syntax" and "File System
'     Handshake" architectural pages.
' ==========================================================================
Public Sub SourceWorksheetToFile(ByVal fileName As String)
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
    
    WriteTextToFile gvSource, fileName
    
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
        lastRow = .Cells.item(.Cells.count).row
    End With

    For rowNumber = source.firstRow To lastRow
        utf8Stream.WriteText SourceSheet.Cells.item(rowNumber, source.sourceColumn).value & vbLf
    Next rowNumber
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    If utf8Stream.Size >= 3 Then
        utf8Stream.position = 3   ' skip BOM
    Else
        utf8Stream.position = 0   ' no BOM present
    End If
    
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile fileName, SaveOptionsEnum.adSaveCreateOverWrite

EndMacro:
    ' Clean up our objects so we don't get a memory leak
    If Not (utf8Stream Is Nothing) Then
        If (utf8Stream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then utf8Stream.Close
        Set utf8Stream = Nothing
    End If
    
    If Not (binaryStream Is Nothing) Then
        If (binaryStream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then binaryStream.Close
        Set binaryStream = Nothing
    End If
#End If
End Sub

' ==========================================================================
' PROCEDURE: StringToFile
'
' PURPOSE:
'   Exports a raw text string to a physical file with explicit UTF-8
'   encoding and BOM removal to ensure Graphviz compatibility.
'
' TECHNICAL WORKFLOW:
'   1. MAC EXECUTION (#If Mac): Delegates to 'WriteTextToFile' for native
'      macOS file handling.
'   2. WINDOWS EXECUTION (#Else):
'      - ENCODING: Uses ADODB.Stream to encode the string as UTF-8.
'      - BOM STRIPPING: Positions the stream at the 3rd byte to exclude the
'        Byte Order Mark (BOM) before copying to a binary stream.
'      - PERSISTENCE: Saves the binary stream to disk using 'adSaveCreateOverWrite'.
'   3. RESOURCE HYGIENE: Implements a 'CleanUp' block to close handles and
'      nullify objects, preventing memory leaks or file locks.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (ADODB on Windows / native Mac call).
'   - DeepWiki Context: Central to the "File System Handshake" logic.
'   - Dependencies: Late-bound ADO (modUtilityADODBConstants.bas).
' ==========================================================================
Public Sub StringToFile(ByVal textString As String, ByVal fileName As String)
#If Mac Then
    WriteTextToFile textString, fileName
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
    
    utf8Stream.WriteText textString
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    utf8Stream.position = 3
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile fileName, SaveOptionsEnum.adSaveCreateOverWrite

EndMacro:
    ' Clean up our objects so we don't get a memory leak
    If Not (utf8Stream Is Nothing) Then
        If (utf8Stream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then utf8Stream.Close
        Set utf8Stream = Nothing
    End If
    
    If Not (binaryStream Is Nothing) Then
        If (binaryStream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then binaryStream.Close
        Set binaryStream = Nothing
    End If
#End If
End Sub

' ==========================================================================
' PROCEDURE: CopySourceCodeToClipboard
'
' PURPOSE:
'   Aggregates the Graphviz DOT source from the 'Source' worksheet into a
'   single string and copies it to the system clipboard for external use.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Resolves the 'sourceWorksheet' UDT to identify the
'      data boundaries per the Named Range API "Contract."
'   2. STRING AGGREGATION: Iterates from 'source.firstRow' to the last used
'      row, concatenating the 'sourceColumn' values into 'dotSource'.
'   3. CLIPBOARD TRANSFER (Windows): Invokes 'ClipBoard_SetData' to place
'      the string into the Windows clipboard buffer.
'   4. UX FEEDBACK: Updates the Excel StatusBar with a localized success
'      or failure message for 5 seconds using 'UpdateStatusBarForNSeconds'.
'
' TECHNICAL NOTES:
'   - Platform: Windows Only (#If Not Mac). Clipboard APIs in VBA typically
'     rely on Windows-specific libraries/APIs.
'   - Layer: UI / Logic.
' ==========================================================================
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
        lastRow = .Cells.item(.Cells.count).row
    End With

    Dim i As Long
    For i = source.firstRow To lastRow
        dotSource = dotSource & SourceSheet.Cells.item(i, source.sourceColumn).value
    Next i
    
    If ClipBoard_SetData(dotSource) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopySuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyFailed"), 5
    End If
    
#End If
End Sub

' ==========================================================================
' PROCEDURE: CreateGraphFromSourceToWorksheet
'
' PURPOSE:
'   Directly renders a Graphviz diagram using the manually edited DOT source
'   from the 'Source' worksheet, bypassing the standard data-parsing engine.
'
' TECHNICAL WORKFLOW:
'   1. STATE RETRIEVAL: Loads the 'settings' UDT via 'GetSettings' to
'      access the cached "Contract" of system configurations.
'   2. WORKSPACE RESET: Clears existing images from the 'Graph' sheet
'      using 'DeleteAllPictures'.
'   3. FILE SYSTEM HANDSHAKE:
'      - Invokes 'SourceWorksheetToFile' to commit the on-sheet source to
'         a temporary '.gv' file.
'   4. EXTERNAL RENDERING:
'      - Configures a 'Graphviz' class instance with engine parameters
'        (Layout, Path, Verbose Mode).
'      - Calls 'RenderGraph' to execute the CLI binary (Windows Shell).
'   5. DIAGNOSTICS: Captures stdout/stderr and redirects to the
'      'Console' worksheet via 'DisplayTextOnConsoleWorksheet'.
'   6. UI INJECTION: Uses 'InsertPicture' to embed the final diagram
'      at anchor cell "B2" with accessibility Alt-Text.
'   7. CLEANUP: Deletes temporary DOT and image assets.
'
' TECHNICAL NOTES:
'   - Layer: External Layer (Graphviz) / UI Layer.
'   - DeepWiki Context: Documents the "DOT Source Inspection" pipeline.
' ==========================================================================
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
        Dim shapeObject As shape
        '@Ignore AssignmentNotUsed
        Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range("B2"), False, True, "Graph image created from source worksheet data.")
        Set shapeObject = Nothing
    Else
        EmitMessage GetMessage("msgboxNoGraphCreated")
    End If

    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Clean up objects
    Set graphvizObj = Nothing
End Sub

' ==========================================================================
' PROCEDURE: VisualizeGraph
'
' PURPOSE:
'   Provides an atomic rendering pipeline that accepts a raw DOT string and
'   displays the resulting diagram on the 'Graph' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOADING: Retrieves the 'settings' UDT via 'GetSettings' to
'      apply the user's active engine and format preferences.
'   2. WORKSPACE RESET: Clears previous renders from the target sheet
'      using 'DeleteAllPictures'.
'   3. FILE SYSTEM HANDSHAKE:
'      - Commits the 'dotSource' parameter to a physical file via 'StringToFile'.
'      - Strips BOM (on Windows) to ensure Graphviz binary compatibility.
'   4. EXTERNAL RENDERING: Configures a 'Graphviz' class instance and
'      invokes 'RenderGraph' to call the external 'dot.exe' binary.
'   5. DIAGNOSTICS: Redirects CLI stdout/stderr messages to the
'      'Console' worksheet for troubleshooting.
'   6. UI INJECTION: Inserts the generated PNG/EMF at cell "B2" using
'      'InsertPicture' with automated Alt-Text labeling.
'   7. RESOURCE HYGIENE: Force-deletes temporary assets and releases memory.
'
' TECHNICAL NOTES:
'   - Layer: Logic Layer / External Layer (Graphviz).
'   - DeepWiki Context: Implements the "Transformation Pipeline" described
'     in the Architecture documentation.
' ==========================================================================
Public Sub VisualizeGraph(dotSource As String)
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
    StringToFile dotSource, graphvizObj.GraphvizFilename
   
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
        Dim shapeObject As shape
        '@Ignore AssignmentNotUsed
        Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range("B2"), False, True, "Graph image created from source worksheet data.")
        Set shapeObject = Nothing
    Else
        EmitMessage GetMessage("msgboxNoGraphCreated")
    End If

    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Clean up objects
    Set graphvizObj = Nothing
End Sub

' ==========================================================================
' PROCEDURE: CreateGraphFromSourceToFile
'
' PURPOSE:
'   Exports a Graphviz diagram to a physical file on disk using the DOT
'   source manually edited or inspected in the 'Source' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. PRE-FLIGHT VALIDATION:
'      - Loads 'settings' UDT via 'GetSettings' (Contract API).
'      - Verifies 'OutputDirectory' and 'FileLocationProvided'; aborts and
'        redirects user to the 'Settings' sheet if missing.
'   2. FILE SYSTEM HANDSHAKE:
'      - Invokes 'SourceWorksheetToFile' to commit the on-sheet text to a
'        temporary '.gv' file.
'   3. EXTERNAL RENDERING:
'      - Configures a 'Graphviz' class instance with engine parameters.
'      - Executes 'RenderGraph' via the external CLI binary.
'   4. DIAGNOSTICS: Captures stdout/stderr for the 'Console' worksheet log.
'   5. DISPOSITION MANAGEMENT:
'      - If rendering succeeds, alerts the user with the final file path.
'      - Conditionally deletes the '.gv' source file based on the
'        'fileDisposition' setting ("delete" vs. "keep").
'
' TECHNICAL NOTES:
'   - Layer: External Layer (Graphviz) / File System.
'   - DeepWiki Context: Part of the "Publishing & Post-Processing" pipeline.
' ==========================================================================
Public Sub CreateGraphFromSourceToFile()
    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(DataSheet.name)

    ' Get file output settings
    Dim output As FileOutput
    output = GetSettingsForFileOutput()
    
    ' Determine output directory, and build file names
    If output.directory = vbNullString Then
        EmitMessage GetMessage("msgboxNoDirectorySpecified")
        SettingsSheet.Activate
        ActiveSheet.Range("OutputDirectory").Activate
        Exit Sub
    End If

    ' Get the file name, minus the file extension
    Dim styleColumn As Long
    styleColumn = GetSettingColNum(SETTINGS_STYLES_COL_SHOW_STYLE)
    
    ' Compose the filename
    If Not FileLocationProvided(output) Then
        Exit Sub
    End If

    ' Instantiate a new Graphviz object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz

    ' Create the filenames
    graphvizObj.OutputDirectory = output.directory
    graphvizObj.FilenameBase = GetFilenameBase(ini, styleColumn)
    graphvizObj.GraphFormat = ini.graph.imageTypeFile
    
    ' Create the '.gv' Graphviz source code file from the source worksheet
    SourceWorksheetToFile graphvizObj.GraphvizFilename
    If Not FileExists(graphvizObj.GraphvizFilename) Then
        EmitMessage GetMessage("msgboxSourceFileNotFound")
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
        EmitMessage GetMessage("msgboxGraphFilenameIs") & vbNewLine & graphvizObj.DiagramFilename
    Else
        EmitMessage GetMessage("msgboxNoGraphCreated")
    End If

    ' Delete the command file if disposition is 'delete'
    If ini.graph.fileDisposition = "delete" Then
        DeleteFile graphvizObj.GraphvizFilename
    End If

    ' Clean up objects
    Set graphvizObj = Nothing
End Sub

' ==========================================================================
' PROCEDURE: UpdateSourceWorksheetLineNumbers
'
' PURPOSE:
'   Synchronizes the "Line Number" column with the actual content of the
'   Source worksheet to ensure the code editor UI remains accurate.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA RESOLUTION: References the 'sourceWorksheet' UDT to identify
'      the 'firstRow' and 'lineNumberColumn' per the Named Range API.
'   2. STALE DATA PURGE: Identifies the 'UsedRange' and clears any existing
'      integers in the line number column to prevent ghost numbers.
'   3. ITERATIVE RENUMBERING: Loops from the 'firstRow' to the 'lastRow' of
'      active content, injecting a sequential counter (1..n).
'   4. BOUNDARY PROTECTION: Ensures the logic never overwrites the
'      header row, even if the sheet is visually empty.
'
' TECHNICAL NOTES:
'   - Layer: UI / Presentation (Source Viewer).
'   - DeepWiki Context: Maintains the "DOT Source Inspection" UI state.
' ==========================================================================
Public Sub UpdateSourceWorksheetLineNumbers()
    Dim cellRange As String
    Dim rowLast As Long
    Dim sourceLayout As sourceWorksheet
    Dim lineNumCol As String
    
    ' Get the layout of the 'data' worksheet
    sourceLayout = GetSettingsForSourceWorksheet()
    
    ' Determine the range of the cells which need to be cleared
    With SourceSheet.UsedRange
        rowLast = .Cells.item(.Cells.count).row
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
        lastRow = .Cells.item(.Cells.count).row
    End With

    For rowNumber = sourceLayout.firstRow To lastRow
        SourceSheet.Cells.item(rowNumber, sourceLayout.lineNumberColumn).value = lineCnt
        lineCnt = lineCnt + 1
    Next rowNumber

End Sub

' ==========================================================================
' PROCEDURE: ClearSource
'
' PURPOSE:
'   A high-level orchestrator that purges DOT source data from both the
'   worksheet and the persistent UserForm.
'
' TECHNICAL WORKFLOW:
'   1. FORM RESET: Invokes 'ClearSourceForm' to reset the 'DotSourceForm'
'      dialog state.
'   2. WORKSHEET RESET: Invokes 'ClearSourceWorksheet' to wipe the
'      'Source' sheet while preserving headers (Named Range API Contract).
'
' USAGE:
'   - Typically linked to "Clear" or "Reset" actions within the
'     DOT Source Viewer subsystem.
' ==========================================================================
Public Sub ClearSource()
    ClearSourceForm
    ClearSourceWorksheet
End Sub

' ==========================================================================
' PROCEDURE: ShowSource
'
' PURPOSE:
'   Directs the Graphviz DOT source to the appropriate user interface
'   components based on the system's current diagnostic settings.
'
' TECHNICAL WORKFLOW:
'   1. FORM UPDATING: Always invokes 'DisplaySourceInForm' to update the
'      persistent 'DotSourceForm' dialog with the latest 'dotSource'.
'   2. CONDITIONAL VISUALIZATION:
'      - Queries the 'settings' UDT (specifically 'SETTINGS_TOOLS_TOGGLE_SOURCE')
'        to check if the user has enabled worksheet-level source inspection.
'      - If TRUE: Invokes 'DisplaySourceInWorksheet' to populate the
'        physical sheet per the Named Range API "Contract."
'
' USAGE:
'   - Called globally whenever new DOT source is generated (e.g., during
'     'RenderElement' or 'CreateGraph').
' ==========================================================================
Public Sub ShowSource(ByVal dotSource As String)
    DisplaySourceInForm dotSource
    
    If GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE) Then
        DisplaySourceInWorksheet dotSource
    End If
End Sub
