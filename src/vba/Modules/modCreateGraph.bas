Attribute VB_Name = "modCreateGraph"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modCreateGraph
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Logic / Transformation Pipeline
'
' ROLE:
'   Central Graphviz orchestration engine. Converts structured worksheet data
'   into DOT source, executes the external Graphviz binary, and injects the
'   resulting diagram back into Excel. Coordinates the full end-to-end
'   rendering lifecycle for both interactive (AutoDraw) and batch-export
'   workflows.
'
' RESPONSIBILITIES:
'   - Manage the complete graph-generation pipeline:
'       • Worksheet parsing and validation
'       • Style and view resolution
'       • DOT synthesis (ConvertDataWorksheetToGvSource)
'       • Temporary file creation and cleanup
'       • Graphviz execution (Graphviz.cls)
'       • Image insertion, scaling, and naming
'   - Provide AutoDraw reactivity for live preview during data entry.
'   - Support batch export across multiple Views, including filename token
'     substitution (%D, %T, %V, %W, %E, %S) and timestamp/option appending.
'   - Handle SVG post-processing (animations, replacements) when enabled.
'   - Maintain cross-platform parity (Windows stopwatch vs macOS sandbox
'     routing, path separators, temp-directory handling).
'   - Expose CreateGraphSource for DOT-only workflows (Source Viewer, debugging).
'
' INTERACTIONS:
'   - Graphviz.cls: External engine wrapper (RenderGraph, SourceToFile).
'   - modDataTypes: settings, dataWorksheet, and style UDTs.
'   - modUtilityString: label scrubbing, token substitution, HTML-label handling.
'   - modUtilityFileSystem: temp directories, file existence, deletion.
'   - modUtilityStatusBar: progress and timing feedback.
'   - Ribbon Tabs: Graphviz, Source, SVG, Styles, Launchpad.
'
' CROSS-PLATFORM NOTES:
'   - Windows: Stopwatch timing, native file access, direct DOT execution.
'   - macOS: AppleScript-mediated file dialogs, sandbox-safe temp routing,
'            conditional filename overrides for "delete" disposition.
'
' ERROR HANDLING:
'   - Defensive validation of worksheet existence, view selection, and
'     output-directory prerequisites.
'   - DOT generation failures surface via the errorMessageColumn and abort
'     rendering cleanly.
'   - Graphviz execution errors routed to the Console worksheet.
'
' RELATED WIKI PAGES:
'   - Rendering Pipeline Overview
'   - Working with the Data Worksheet
'   - Batch Export & View Iteration
'   - DOT Source Generation & Validation
'   - Image Path Resolution
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: AutoDraw
'
' PURPOSE:
'   Executes a graph refresh safely and atomically, ensuring that worksheet
'   events and screen repaints do not interfere with the rendering pipeline.
'   This routine assumes that the caller has already validated run mode
'   (e.g., Auto vs Manual) and that a redraw is appropriate.
'
' TECHNICAL WORKFLOW:
'   1. EVENT & UI SUSPENSION:
'      - Disables ScreenUpdating to prevent flicker and mid-render repaints.
'      - Disables EnableEvents to prevent recursive Worksheet_Change triggers.
'
'   2. RENDER EXECUTION:
'      - Calls 'CreateGraphWorksheet' to rebuild the DOT source, invoke
'        Graphviz, and insert the updated image into the GraphSheet.
'
'   3. RESTORATION:
'      - Re-enables events and screen updates, restoring normal Excel behavior.
'
' TECHNICAL NOTES:
'   - This routine is intentionally minimal: it provides a safe execution
'     boundary around the rendering pipeline without introducing UI latency
'     (e.g., cursor changes or DoEvents).
'   - Triggered indirectly via Worksheet_Change ? AutoDrawDebounced ? AutoDraw.
'   - DeepWiki Context: Represents the "safe execution wrapper" for the
'     AutoDraw reactivity model described in the Data Worksheet documentation.
' ==========================================================================
Public Sub AutoDraw()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    CreateGraphWorksheet
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: ClearWorksheetGraphs
'
' PURPOSE:
'   Purges all Graphviz-generated imagery from the primary workspace to
'   reset the visual state or prepare for a fresh rendering cycle.
'
' TECHNICAL WORKFLOW:
'   1. DATA SHEET RESET: Invokes 'DeleteAllPictures' on the active Data
'      worksheet (resolved via 'GetDataWorksheetName').
'   2. GRAPH SHEET RESET: Invokes 'DeleteAllPictures' on the dedicated
'      'Graph' worksheet.
'
' TECHNICAL NOTES:
'   - Layer: UI / Presentation Layer.
'   - Strategy: Prevents image "stacking" where new renders might be
'     hidden behind stale OLE objects.
' ==========================================================================
Public Sub ClearWorksheetGraphs()
    ' Delete pictures from 'data' worksheet
    DeleteAllPictures GetDataWorksheetName()
    ' Delete pictures from the 'graph' worksheet
    DeleteAllPictures GraphSheet.name
End Sub

' ==========================================================================
' PROCEDURE: ClearErrors
'
' PURPOSE:
'   Resets the error state of the active Data worksheet by removing error
'   flags and localized diagnostic messages from individual rows.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Retrieves the 'dataWorksheet' UDT to resolve
'      the 'flag' and 'errorMessage' column indices per the Named Range API.
'   2. ITERATIVE SCAN: Loops through the worksheet from 'firstRow' to
'      'lastRow' (as defined by the system's "Contract").
'   3. CONDITIONAL RESET: Checks the 'flagColumn' for the 'FLAG_ERROR'
'      constant; if found, it invokes 'ClearCell' to wipe both the
'      visual indicator and the descriptive error text.
'
' TECHNICAL NOTES:
'   - Layer: Data Management / Logic.
'   - Strategy: Pre-flight maintenance used to ensure previous validation
'     runs do not pollute new rendering attempts.
' ==========================================================================
Public Sub ClearErrors()

    ' Data worksheet variables
    Dim data As dataWorksheet
    data = GetSettingsForDataWorksheet(GetDataWorksheetName())
    
    ' Iterate through the rows
    Dim row As Long
    For row = data.firstRow To data.lastRow
        If GetCell(data.worksheetName, row, data.flagColumn) = FLAG_ERROR Then
            ClearCell data.worksheetName, row, data.flagColumn
            ClearCell data.worksheetName, row, data.errorMessageColumn
        End If
    Next row

End Sub

' ==========================================================================
' PROCEDURE: CreateGraphWorksheetQuickly
'
' PURPOSE:
'   Provides a high-performance, hotkey-accessible entry point for generating
'   a graph from the active worksheet data.
'
' TECHNICAL WORKFLOW:
'   1. UI FEEDBACK: Sets the 'xlWait' cursor and executes 'DoEvents' to
'      ensure the UI remains responsive during the initial handshake.
'   2. PERFORMANCE OPTIMIZATION: Invokes 'OptimizeCode_Begin' to suspend
'      calculations, events, and screen updates.
'   3. EXECUTION: Calls 'CreateGraphWorksheet' to run the full parsing and
'      rendering pipeline.
'   4. STATE RESTORATION: Re-enables Excel features via 'OptimizeCode_End'
'      and restores the default cursor.
'
' TECHNICAL NOTES:
'   - Access: Mapped to Ctrl+Shift+Q (@ExcelHotkey q).
'   - Strategy: Minimizes overhead for power users who frequently
'     regenerate graphs during data entry.
' ==========================================================================
'@ExcelHotkey q
'
Public Sub CreateGraphWorksheetQuickly()
Attribute CreateGraphWorksheetQuickly.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents
    
    OptimizeCode_Begin
    CreateGraphWorksheet
    OptimizeCode_End
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
End Sub

' ==========================================================================
' PROCEDURE: CreateGraphWorksheet
'
' PURPOSE:
'   THE CENTRAL RENDERING ORCHESTRATOR. Transforms structured worksheet data
'   into a visual diagram by managing the end-to-end Graphviz lifecycle.
'
' TECHNICAL WORKFLOW:
'   1. ENVIRONMENT INIT: Retrieves the 'settings' UDT and clears previous
'      visual assets.
'   2. PERFORMANCE MONITORING: Starts a Windows-specific 'Stopwatch' to
'      track rendering latency.
'   3. DOT GENERATION: Invokes 'ConvertDataWorksheetToGvSource' to translate
'      Excel rows into DOT language. If validation fails, it reveals the
'      'errorMessageColumn' and aborts.
'   4. DIAGNOSTIC HOOKS: Passes the generated string to 'ShowSource' to
'      update the Source Viewer/Form if debugging is enabled.
'   5. ENGINE EXECUTION:
'      - Commits the DOT string to a temporary physical file.
'      - Configures the 'Graphviz' class with engine and CLI parameters.
'      - Invokes 'RenderGraph' to trigger the external binary.
'   6. UI INJECTION: Inserts the resulting image at the 'targetCell'
'      (either B2 on the Graph sheet or a specific cell on the Data sheet).
'   7. POST-PROCESSING: Applies user-defined scaling (Zoom) and renames
'      the picture object for downstream reference.
'   8. RESOURCE HYGIENE: Deletes temporary files and releases class instances.
'
' TECHNICAL NOTES:
'   - Cross-Platform: Conditional logic manages the Windows 'Stopwatch' vs.
'     macOS execution.
'   - Layer: The primary bridge between the Logic Layer (VBA) and the
'     External Layer (Graphviz Engine).
' ==========================================================================
'@Ignore MissingMemberAnnotation
Public Sub CreateGraphWorksheet()
Attribute CreateGraphWorksheet.VB_ProcData.VB_Invoke_Func = " \n14"

    On Error Resume Next
    
#If Mac Then
    ' For some reason, my Mac fails when I code it as "#If Not Mac Then"
#Else
    ' Stopwatch is only available on Windows OS
    Dim timex As Stopwatch
    Set timex = New Stopwatch
    timex.start
#End If
    
    ' Clear the status bar
    ClearStatusBar

    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(GetDataWorksheetName())

    If Not WorksheetExists(ini.data.worksheetName) Then
        EmitMessage GetMessage("msgboxNoDataToGraph")
        Exit Sub
    End If

    ' Remove any existing graph image from the target worksheet
    Dim displayDataSheetName As String
    Dim targetCell As String

    If ini.graph.imageWorksheet = "data" Then
        displayDataSheetName = ini.data.worksheetName
        targetCell = ini.data.graphDisplayColumnAsAlpha & ini.data.firstRow
    Else
        displayDataSheetName = GraphSheet.name
        targetCell = "B2"
    End If
            
    ActiveWorkbook.Sheets.[_Default](displayDataSheetName).Activate
    DeleteAllPictures displayDataSheetName

    ' Instantiate a Graphviz Object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz
    
    ' Build the file names
    graphvizObj.OutputDirectory = GetTempDirectory()
    graphvizObj.FilenameBase = "RelationshipVisualizer"
    graphvizObj.GraphFormat = ini.graph.imageTypeWorksheet
    
    ' Clear any source code being displayed
    ClearSource

    ' Expose the view name so it can be used as data in the graph
    SettingsSheet.Range("ViewNameLabel").value = StylesSheet.Cells.item(ini.styles.headingRow, ini.styles.selectedViewColumn).value

    ' View name might be referenced in the graph options, so refresh the value
    ini.graph.options = Trim$(SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).value)

    ' Create the '.gv' Graphviz source code file from the relationships in the
    ' data worksheet
    Dim graphvizSource As String
    If Not ConvertDataWorksheetToGvSource(ini, ini.styles.selectedViewColumn, graphvizSource) Then
        ' Report errors to the user
        ShowColumn ini.data.worksheetName, ini.data.errorMessageColumn, True
        Exit Sub
    End If
    
    ' Display source if debugging
    ShowSource graphvizSource

    ' Write the graphviz source to a file
    graphvizObj.graphvizSource = graphvizSource
    graphvizObj.SourceToFile
    
    ' Display source if debugging
    ShowSource graphvizSource

    ' Hide the messages column
    ShowColumn ini.data.worksheetName, ini.data.errorMessageColumn, False

    ' Convert the Graphviz source code into a diagram
    graphvizObj.CaptureMessages = ini.console.logToConsole
    graphvizObj.Verbose = ini.console.graphvizVerbose
    graphvizObj.CommandLineParameters = ini.CommandLine.parameters
    graphvizObj.GraphLayout = ini.graph.engine
    graphvizObj.GraphvizPath = ini.CommandLine.GraphvizPath
    
    graphvizObj.RenderGraph

    ' Display any console output first
    DisplayTextOnConsoleWorksheet graphvizObj.GraphvizCommand, graphvizObj.GraphvizMessages
        
    '@Ignore VariableNotUsed
    Dim shapeObject As shape
    Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range(targetCell), False, True, "Graph image created from data worksheet data.")
    
    ' Scale the graph to the zoom percentage specified
    Dim scaleFactor As Double
    scaleFactor = ini.graph.scaleImage / 100
    ActiveSheet.Pictures(ActiveSheet.Pictures.count).ShapeRange.ScaleHeight scaleFactor, msoFalse, msoScaleFromTopLeft
    
    If ini.graph.pictureName <> vbNullString Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.count).name = ini.graph.pictureName
    End If
    Set shapeObject = Nothing

    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Clean up
    Set graphvizObj = Nothing
    
#If Mac Then
    ' For some reason, my Mac fails when I code it as "#If Not Mac Then"
#Else
    timex.stop_it
    Application.StatusBar = timex.Elapsed_sec & " seconds"
#End If
    
    On Error GoTo 0
End Sub

' ==========================================================================
' SECTION: EXTERNAL FILE EXPORT & BATCH PROCESSING
' ==========================================================================

' ==========================================================================
' PROCEDURE: CreateGraphFile
'
' PURPOSE:
'   The primary batch-export engine. Iterates through defined "View" columns
'   to generate and save multiple diagram files to disk.
'
' TECHNICAL WORKFLOW:
'   1. PRE-FLIGHT: Loads 'settings' UDT and validates that the output
'      directory and filename prefixes are established.
'   2. VIEW ITERATION: Loops from 'firstViewColumn' to 'lastViewColumn'.
'   3. DYNAMIC LABELING: Updates the 'ViewNameLabel' named range for every
'      iteration, allowing the graph title or filename to react to the
'      active View name.
'   4. MAC SANDBOX MANAGEMENT (#If Mac): If disposition is set to 'delete',
'      it routes the DOT source through the system Temp directory to avoid
'      repetitive permission prompts.
'   5. RENDERING: Orchestrates DOT generation via 'ConvertDataWorksheetToGvSource'
'      and invokes the 'Graphviz' class for CLI execution.
'   6. SVG POST-PROCESSING: If 'postProcessSVG' is enabled, it triggers
'      'FindAndReplaceSVG' to inject animations or XML modifications.
'   7. CLEANUP: Deletes temporary source files based on 'fileDisposition'.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Implements the "Batch Export Process" and "File System
'     Handshake" architectural flows.
'   - Strategy: Decouples the rendering loop from the UI, enabling high-
'     volume production of graph variants.
' ==========================================================================
Public Sub CreateGraphFile(ByVal firstViewColumn As Long, ByVal lastViewColumn As Long)
    ' Clear the status bar
    ClearStatusBar
    
    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(GetDataWorksheetName())

    If Not WorksheetExists(ini.data.worksheetName) Then
        EmitMessage GetMessage("msgboxNoDataToGraph")
        Exit Sub
    End If

    ' Determine output directory, and build file names
    If ini.output.directory = vbNullString Then
        ini.output.directory = vbNullString = ActiveWorkbook.path
    End If

    ' Validate filename info
    If Not FileLocationProvided(ini) Then
        Exit Sub
    End If

    ' Hide the messages column
    ShowColumn ini.data.worksheetName, ini.data.errorMessageColumn, False
    
    Dim viewColumn As Long
    For viewColumn = firstViewColumn To lastViewColumn
    
        ' Expose the view name so it can be used as data in the graph
        SettingsSheet.Range("ViewNameLabel").value = StylesSheet.Cells.item(ini.styles.headingRow, viewColumn).value
        
        ' View name might be referenced in the graph options, so refresh the value
        ini.graph.options = Trim$(SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).value)

        ' Create new Graphviz object
        Dim graphvizObj As Graphviz
        Set graphvizObj = New Graphviz
        
        ' Build the file names
        graphvizObj.OutputDirectory = ini.output.directory
        graphvizObj.FilenameBase = GetFilenameBase(ini, viewColumn)
        graphvizObj.GraphFormat = ini.graph.imageTypeFile
#If Mac Then
        ' If we are running on a Mac, and we are not going to keep the source file, use a filename within
        ' the sandbox which the user will not have to grant permission to use. If keeping the file, they
        ' will just have to grant permission.
        If ini.graph.fileDisposition = "delete" Then
            graphvizObj.GraphvizFilename = GetTempDirectory() & Application.pathSeparator & "RelationshipVisualizer.gv"
        End If
#End If
        ' Clear any source code being displayed
        ClearSource

        ' Create Graphviz graph source code
        Dim graphvizSource As String
        If Not ConvertDataWorksheetToGvSource(ini, viewColumn, graphvizSource) Then
            Exit Sub
        End If
        
        ' Display source if debugging
        ShowSource graphvizSource

        ' Write the Graphviz source to a file
        graphvizObj.graphvizSource = graphvizSource
        graphvizObj.SourceToFile
        
        ' Convert the Graphviz source code into a diagram
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
            ' Post-process SVG files to add things like animations
            If ini.graph.imageTypeFile = FILETYPE_SVG And ini.graph.postProcessSVG Then
                FindAndReplaceSVG graphvizObj.DiagramFilename, graphvizObj.DiagramFilename
            End If
            
            UpdateStatusBarForNSeconds GetMessage("statusbarGraphFilenameIs") & " " & graphvizObj.DiagramFilename, 10
        Else
            EmitMessage GetMessage("msgboxNoGraphCreated")
        End If

        ' Delete the graph source code file if disposition is 'delete'
        If ini.graph.fileDisposition = "delete" Then
             DeleteFile graphvizObj.GraphvizFilename
        End If
        
        ' Cleanup objects
        Set graphvizObj = Nothing
    Next viewColumn

    ' Sync up settings with dropdown choice
    SettingsSheet.Range("ViewNameLabel").value = SettingsSheet.Range("ViewName").value

End Sub

' ==========================================================================
' FUNCTION: CreateGraphSource
'
' PURPOSE:
'   Generates a raw Graphviz DOT source string from the active data sheet
'   without initiating an external rendering process.
'
' TECHNICAL WORKFLOW:
'   1. CONTEXT INITIALIZATION: Loads the 'settings' UDT for the active
'      data worksheet to resolve layout and view constraints.
'   2. SOURCE SYNTHESIS: Invokes 'ConvertDataWorksheetToGvSource' using the
'      currently selected 'viewColumn' defined in the Style Gallery settings.
'   3. ERROR MANAGEMENT: If DOT generation fails, returns a null string;
'      otherwise, returns the complete Graphviz markup.
'   4. UI CLEANUP: Force-hides the 'errorMessageColumn' to ensure the
'      data sheet remains clean after the operation.
'
' USAGE:
'   - Primary data provider for the "DOT Source Viewer" and "Source Form."
'   - Allows for structural validation of the graph logic without the
'     overhead of file I/O or binary execution.
' ==========================================================================
Public Function CreateGraphSource() As String

    ' Read in the runtime settings
    Dim ini As settings
    ini = GetSettings(GetDataWorksheetName())

    If Not WorksheetExists(ini.data.worksheetName) Then
        EmitMessage GetMessage("msgboxNoDataToGraph")
        Exit Function
    End If

    Dim graphvizSource As String
    If ConvertDataWorksheetToGvSource(ini, ini.styles.selectedViewColumn, graphvizSource) Then
        CreateGraphSource = graphvizSource
    Else
        CreateGraphSource = vbNullString
    End If

    ' Hide the messages column
    ShowColumn ini.data.worksheetName, ini.data.errorMessageColumn, False
End Function

' ==========================================================================
' FUNCTION: FileLocationProvided
'
' PURPOSE:
'   Ensures all file system prerequisites are met before the rendering engine
'   attempts to write a diagram to disk.
'
' TECHNICAL WORKFLOW:
'   1. DIRECTORY VALIDATION: Checks the existence of the 'output.directory'
'      using 'DirectoryExists'. If missing, alerts the user with a localized
'      error message.
'   2. FILENAME VALIDATION: Verifies that 'output.fileNamePrefix' is not
'      empty, ensuring the "Publishing" pipeline has a valid target name.
'   3. STATE RETURN: Returns FALSE if either check fails, acting as a
'      critical safety gate for file-export operations.
'
' TECHNICAL NOTES:
'   - Layer: File System / Logic Layer.
'   - Strategy: Prevents VBA runtime errors during binary execution by
'     validating paths at the UI/Logic boundary.
' ==========================================================================
Public Function FileLocationProvided(ByRef ini As settings) As Boolean
    FileLocationProvided = True
    
    ' Validate that the output directory exists
    If Not DirectoryExists(ini.output.directory) Then
        EmitMessage replace(GetMessage("msgboxDirDoesNotExist"), "{dir}", ini.output.directory), buttons:=vbCritical
        FileLocationProvided = False
    End If

    ' Get the base value of the file name
    If ini.output.fileNamePrefix = vbNullString Then
        EmitMessage GetMessage("msgboxPrefixNotSpecified"), buttons:=vbCritical
        FileLocationProvided = False
    End If

End Function

' ==========================================================================
' FUNCTION: GetFilenameBase
'
' PURPOSE:
'   Constructs a highly-customizable filename by resolving dynamic tokens
'   and metadata into a sanitized string for file system operations.
'
' TECHNICAL WORKFLOW:
'   1. TOKEN RESOLUTION: Parses the user-defined prefix for specific tokens:
'      - %D / %T: Injects localized Date and Time stamps.
'      - %V: Injects the current View Name from the Style Gallery.
'      - %W: Injects the name of the active Data Worksheet.
'      - %E / %S: Injects the Graphviz Engine and Splines configuration.
'   2. HEURISTIC APPENDING: If tokens aren't present but "Append" toggles are
'      enabled, the function automatically appends Date/Time or Engine
'      options to the end of the string.
'   3. SYNTAX NORMALIZATION: Formats appended options within brackets [ ]
'      to maintain consistent file naming conventions.
'   4. SANITIZATION: Returns a trimmed string ready for path concatenation.
'
' TECHNICAL NOTES:
'   - Layer: Logic / File System.
'   - Strategy: Empowers users to create descriptive, unique filenames for
'     batch exports without manual renaming.
' ==========================================================================
Public Function GetFilenameBase(ByRef ini As settings, ByVal showStyleColumn As Long) As String

    Dim fileBase As String

    ' Build up the file name from the user-specified prefix
    fileBase = ini.output.fileNamePrefix
    
    ' Include Timestamp if desired
    If ini.output.appendTimeStamp Then
        If InStr(fileBase, "%D") Or InStr(fileBase, "%T") Then
            ' Substitute date for %D
            If InStr(fileBase, "%D") Then
                fileBase = replace(fileBase, "%D", ini.output.date)
            End If
            
            ' Substitute time for %D
            If InStr(fileBase, "%T") Then
                fileBase = replace(fileBase, "%T", ini.output.time)
            End If
        Else
            fileBase = fileBase & " " & ini.output.date & " " & ini.output.time
        End If
    End If

    ' Include the view name
    If InStr(fileBase, "%V") Then
        ' Substitute View name for %V
        fileBase = replace(fileBase, "%V", StylesSheet.Cells.item(ini.styles.headingRow, showStyleColumn).value)
    Else
        fileBase = fileBase & " " & StylesSheet.Cells.item(ini.styles.headingRow, showStyleColumn).value
    End If

    ' Include the worksheet name
    If InStr(fileBase, "%W") Then
        ' Substitute data worksheet name for %W
        fileBase = replace(fileBase, "%W", ini.data.worksheetName)
    End If
    
    ' Include Graphing Options if desired
    If ini.output.appendOptions Then
        If InStr(fileBase, "%E") Or InStr(fileBase, "%S") Then
            ' Substitute Graph engine for %E
            If InStr(fileBase, "%E") Then
                fileBase = replace(fileBase, "%E", SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value)
            End If
        
            ' Substitute Splines engine for %S
            If InStr(fileBase, "%S") Then
                fileBase = replace(fileBase, "%S", ini.graph.splines)
            End If
        Else
            fileBase = fileBase & " [" & SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
            If ini.graph.splines <> vbNullString Then
                fileBase = fileBase & COMMA & ini.graph.splines
            End If
            fileBase = fileBase & "]"
        End If
    End If

    GetFilenameBase = Trim$(fileBase)

End Function

' ==========================================================================
' FUNCTION: GetExcelToGraphvizImageDirectory
'
' PURPOSE:
'   Retrieves the absolute path stored in the 'ExcelToGraphvizImages'
'   system environment variable.
'
' TECHNICAL WORKFLOW:
'   1. SYSTEM QUERY: Uses the VBA 'Environ$' function to poll the host OS
'      for the project-specific variable.
'   2. NORMALIZATION: Trims any leading or trailing whitespace to ensure
'      the path string is valid for downstream file I/O operations.
'
' TECHNICAL NOTES:
'   - Layer: File System / Logic Layer.
'   - DeepWiki Context: Documents the "Image Path Resolution" logic used
'     to provide a standardized directory for icons and backgrounds
'     independent of the Workbook's physical location.
' ==========================================================================
Public Function GetExcelToGraphvizImageDirectory() As String
    GetExcelToGraphvizImageDirectory = Trim$(Environ$("ExcelToGraphvizImages"))
End Function

' ==========================================================================
' SECTION: PARSING LOGIC & ELEMENT CLASSIFICATION
' ==========================================================================

' ==========================================================================
' FUNCTION: GetImagePath
'
' PURPOSE:
'   Aggregates multiple directory paths into a single delimited string to
'   inform the Graphviz 'imagepath' attribute where to find visual assets.
'
' TECHNICAL WORKFLOW:
'   1. BASE RESOLUTION: Retrieves the user-defined path from the
'      'SETTINGS_IMAGE_PATH' named range.
'   2. PLATFORM DELIMITERS: Selects the correct path separator based on OS
'      standards (Colon for macOS, Semicolon for Windows).
'   3. HIERARCHICAL MERGE:
'      - Prepends the 'ActiveWorkbook.path' to ensure relative assets
'        are prioritized.
'      - Appends the 'ExcelToGraphvizImages' environment variable path
'        if it exists.
'   4. CONCATENATION: Joins all valid paths into a single string for DOT
'      attribute injection.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (Conditional separators).
'   - DeepWiki Context: Implements the "Image Path Resolution" logic,
'     ensuring the Graphviz engine can resolve external icons/backgrounds.
' ==========================================================================
Public Function GetImagePath() As String

    Dim imagePath As String
    imagePath = SettingsSheet.Range(SETTINGS_IMAGE_PATH).value
    
    Dim pathSeparator As String
#If Mac Then
    pathSeparator = COLON
#Else
    pathSeparator = SEMICOLON
#End If

    ' Include current directory on the image path
    If imagePath = vbNullString Then
        imagePath = Application.ActiveWorkbook.path
    Else
        imagePath = Application.ActiveWorkbook.path & pathSeparator & imagePath
    End If

    ' Append the directory associated with the environment variable
    ' to the image path, if a path has been specified
    Dim envImagePath As String
    envImagePath = GetExcelToGraphvizImageDirectory()
    If envImagePath <> vbNullString Then
        imagePath = imagePath & pathSeparator & envImagePath
    End If

    GetImagePath = imagePath
    
End Function

' ==========================================================================
' FUNCTION: DetermineStyleName
'
' PURPOSE:
'   Acts as the primary "Classifier" for the parsing engine, determining
'   how a worksheet row should be translated into Graphviz DOT syntax.
'
' TECHNICAL WORKFLOW:
'   1. STRUCTURAL DETECTION:
'      - Detects Subgraph boundaries by checking for '{' (Open) or '}' (Close).
'      - Detects Native DOT passthrough when the Item column starts with '>'.
'   2. RELATIONSHIP HEURISTICS:
'      - If a 'Related Item' is present, the row is classified as an EDGE.
'      - If no 'Related Item' is present, it is classified as a NODE.
'   3. KEYWORD OVERRIDE:
'      - Recognizes global Graphviz keywords (node, edge, graph) to apply
'        broad attribute settings.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Implements the "Row Classification Logic" detailed
'     in the Graph Generation Pipeline documentation.
'   - Strategy: Centralizes the transformation logic that maps Excel rows
'     to Graphviz object types (TYPE_NODE, TYPE_EDGE, etc.).
' ==========================================================================
Private Function DetermineStyleName(ByRef ini As settings, ByVal row As Long) As String

    Dim styleName As String
    
    Dim dataItem As String
    dataItem = GetCell(ini.data.worksheetName, row, ini.data.itemColumn)

    If dataItem <> vbNullString Then
        If EndsWith(dataItem, OPEN_BRACE) Then
            styleName = TYPE_SUBGRAPH_OPEN
        
        ElseIf dataItem = CLOSE_BRACE Then
            styleName = TYPE_SUBGRAPH_CLOSE
        
        ElseIf dataItem = GREATER_THAN Then
            styleName = TYPE_NATIVE
        
        Else
            Dim dataIsRelatedtoItem As String
            dataIsRelatedtoItem = GetCell(ini.data.worksheetName, row, ini.data.isRelatedToItemColumn)
            
            If dataIsRelatedtoItem = vbNullString Then
                If dataItem = KEYWORD_NODE Or dataItem = KEYWORD_EDGE Or dataItem = KEYWORD_GRAPH Then
                    styleName = TYPE_KEYWORD
                Else
                    styleName = TYPE_NODE
                End If
            Else
                styleName = TYPE_EDGE
            End If
        End If
    End If

    DetermineStyleName = styleName
    
End Function

' ==========================================================================
' FUNCTION: RemovePort
'
' PURPOSE:
'   Extracts the base Node ID from a string that potentially contains
'   Graphviz port or compass point notation (e.g., "Node:port:sw").
'
' TECHNICAL WORKFLOW:
'   1. DELIMITER DETECTION: Scans the 'nodeId' string for the colon (:)
'      separator used by Graphviz for port addressing.
'   2. TOKEN EXTRACTION: If a colon is present, it invokes
'      'GetStringTokenAtPosition' to retrieve only the first segment.
'   3. FALLBACK: Returns the original string if no port syntax is detected.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Essential for the "Defining Nodes & Edges" page,
'     ensuring the parser can identify parent nodes even when specific
'     connection ports are defined.
'   - Syntax: Supports standard DOT notation (node:port).
' ==========================================================================
Private Function RemovePort(ByVal nodeId As String) As String
    
    ' Strip off the port (if specified)
    If InStr(nodeId, ":") > 0 Then
        RemovePort = GetStringTokenAtPosition(nodeId, ":", 1)
    Else
        RemovePort = nodeId
    End If

End Function

' ==========================================================================
' FUNCTION: ConvertDataWorksheetToGvSource
'
' PURPOSE:
'   The primary structural logic controller. Orchestrates data validation,
'   style caching, and relationship filtering before final DOT generation.
'
' TECHNICAL WORKFLOW:
'   1. STYLE CACHING: Invokes 'CacheEnabledStyles' to load formatting rules
'      into a Dictionary, enabling high-performance attribute lookups.
'   2. PRE-FLIGHT CLEANUP: Iterates through the data worksheet to purge
'      'FLAG_ERROR' indicators and stale messages from previous runs.
'   3. CONNECTIVITY ANALYSIS: If orphan filtering is enabled:
'      - 'ConfirmItemStyleIsValid': Verifies objects have mapped styles.
'      - 'DetermineWhatGraphShouldInclude': Evaluates Node-Edge-Node
'        integrity to filter out disconnected elements.
'   4. VALIDATION GATE: Calls 'ValidateData'; only proceeds to source
'      generation if 'errorCount' is zero.
'   5. SOURCE SYNTHESIS: Triggers 'CreateGraphvizSource' to assemble the
'      final DOT markup string.
'   6. RESOURCE HYGIENE: Force-clears all Dictionary objects to prevent
'      memory leaks.
'
' TECHNICAL NOTES:
'   - Strategy: Prevents invalid or "broken" graphs by enforcing a
'     strict validation-before-rendering pipeline.
'   - Layer: Logic Layer / Data Management.
' ==========================================================================
Private Function ConvertDataWorksheetToGvSource(ByRef ini As settings, _
                                                ByVal showStyleColumn As Long, _
                                                ByRef graphvizSource As String) As Boolean
    ' Assume conversion is not successful
    ConvertDataWorksheetToGvSource = False

    ' Dictionaries to determine what data is referenced
    Dim nodeIds As Dictionary
    Set nodeIds = New Dictionary
    
    Dim edgeIds As Dictionary
    Set edgeIds = New Dictionary
    
    Dim nodeIdsInRelationships As Dictionary
    Set nodeIdsInRelationships = New Dictionary

    ' Cache the style definitions in the 'styles' worksheet
    Dim styles As Dictionary
    Set styles = CacheEnabledStyles(ini, showStyleColumn)
    
    ' Remove any error messages from a previous run
    Dim row As Long
    For row = ini.data.firstRow To ini.data.lastRow
        If GetCell(ini.data.worksheetName, row, ini.data.flagColumn) = FLAG_ERROR Then
            ClearCell ini.data.worksheetName, row, ini.data.flagColumn
            ClearCell ini.data.worksheetName, row, ini.data.errorMessageColumn
        End If
    Next row
    
    ' Inspect the data if we are to filter out orphan types
    If Not ini.graph.includeOrphanNodes Or Not ini.graph.includeOrphanEdges Then
        ' Iterate through the rows to determine what nodes and edges have valid
        ' style definitions, and collect this information in lists.
        ConfirmItemStyleIsValid ini, styles, nodeIds, edgeIds
        
        ' Determine if both the tail and head of the included relationships refer
        ' to nodes which have been included, and have style definitions
        DetermineWhatGraphShouldInclude ini, styles, nodeIds, nodeIdsInRelationships
    End If

    ' Generate the dot language Graphviz file
    Dim errorCount As Long
    errorCount = ValidateData(ini, styles)
                                
    If errorCount = 0 Then
        CreateGraphvizSource ini, styles, nodeIds, nodeIdsInRelationships, graphvizSource
        ConvertDataWorksheetToGvSource = True
    End If
    
    ' Clean up so we don't have a memory leak
    Set styles = Nothing
    Set nodeIds = Nothing
    Set edgeIds = Nothing
    Set nodeIdsInRelationships = Nothing
    
End Function

' ==========================================================================
' SECTION: STRUCTURAL ANALYSIS & ORPHAN FILTERING
' ==========================================================================

' ==========================================================================
' PROCEDURE: ConfirmItemStyleIsValid
'
' PURPOSE:
'   The primary structural scanner. Performs an initial pass of the Data
'   worksheet to catalog every Node and Edge endpoint that possesses a
'   valid, enabled style definition.
'
' TECHNICAL WORKFLOW:
'   1. COMMENT FILTERING: Skips any rows explicitly marked with 'FLAG_COMMENT'.
'   2. STYLE RESOLUTION: Retrieves the 'styleName'; if blank, it invokes
'      'DetermineStyleName' to infer the type (Node/Edge/Subgraph/Native).
'   3. REGISTRY ENROLLMENT:
'      - TYPE_NODE: Parses the 'Item' column (handling comma-delimited lists
'        and stripping ports) to populate the 'nodeIds' Dictionary.
'      - TYPE_EDGE: Parses both 'Item' (Tail) and 'Related Item' (Head) columns
'        to catalog all referenced endpoints in the 'edgeIds' Dictionary.
'   4. DATA NORMALIZATION: Uses 'UCase$' for style name lookups and 'RemovePort'
'      to ensure base ID consistency.
'
' TECHNICAL NOTES:
'   - Layer: Logic Layer / Pre-processing.
'   - Strategy: Builds the "Source of Truth" for valid IDs, serving as the
'     input for the subsequent Orphan Filtering logic.
' ==========================================================================
Private Sub ConfirmItemStyleIsValid(ByRef ini As settings, _
                                   ByVal styles As Dictionary, _
                                   ByVal nodeIds As Dictionary, _
                                   ByVal edgeIds As Dictionary)
    Dim row As Long
    Dim data As dataRow
    
    Dim nodeId As String
    Dim itemIdArray() As String
    
    Dim arrayIndex As Long
    
    For row = ini.data.firstRow To ini.data.lastRow
        If GetCell(ini.data.worksheetName, row, ini.data.flagColumn) <> FLAG_COMMENT Then ' line is not commented out
            data.styleName = GetCell(ini.data.worksheetName, row, ini.data.styleNameColumn)

            ' Try to determine the style if not supplied
            If data.styleName = vbNullString Then
                data.styleName = DetermineStyleName(ini, row)
            End If

            ' Get the style names in a consistent case
            data.styleName = UCase$(data.styleName)
            
            If data.styleName <> vbNullString Then ' a style was specified
                If styles.Exists(data.styleName) Then ' show this in the diagram

                    ' We want data of this style in the output file
                    data.item = GetCell(ini.data.worksheetName, row, ini.data.itemColumn)
                    data.relatedItem = GetCell(ini.data.worksheetName, row, ini.data.isRelatedToItemColumn)
                        
                    ' What type of row is it?
                    data.styleType = styles.item(data.styleName).styleType

                    If data.styleType = TYPE_NODE Then

                        If data.item <> vbNullString And UCase$(data.item) <> KEYWORD_NODE And data.relatedItem = vbNullString Then
                        
                            ' There are potentially multiple item IDs, so parse them from the data.item string
                            itemIdArray = split(data.item, COMMA)
                            For arrayIndex = LBound(itemIdArray) To UBound(itemIdArray)
                                nodeId = RemovePort(itemIdArray(arrayIndex))
                                If Not nodeIds.Exists(nodeId) Then
                                    nodeIds.Add nodeId, True
                                End If
                            Next
                        End If

                    ElseIf data.styleType = TYPE_EDGE Then

                        If data.item <> vbNullString And UCase$(data.item) <> KEYWORD_EDGE And data.relatedItem <> vbNullString Then
                            ' There are potentially multiple item IDs, so parse them from the data.item string
                            itemIdArray = split(data.item, COMMA)
                            For arrayIndex = LBound(itemIdArray) To UBound(itemIdArray)
                                nodeId = RemovePort(itemIdArray(arrayIndex))
                                If Not edgeIds.Exists(nodeId) Then
                                    edgeIds.Add nodeId, True
                                End If
                            Next
                            
                            ' There are potentially multiple related item IDs, so parse them from the data.relatedItem string
                            itemIdArray = split(data.relatedItem, COMMA)
                            For arrayIndex = LBound(itemIdArray) To UBound(itemIdArray)
                                nodeId = RemovePort(itemIdArray(arrayIndex))

                                If Not edgeIds.Exists(nodeId) Then
                                    edgeIds.Add nodeId, True
                                End If
                            Next
                        End If                   ' if tail and head are non-blank
                    End If                       ' if NODE elseif EDGE
                End If                           ' style is to be included in output diagram
            End If                               ' style was specified
        End If                                   ' not a comment line
    Next row

End Sub

' ==========================================================================
' PROCEDURE: DetermineWhatGraphShouldInclude
'
' PURPOSE:
'   Performs a connectivity audit to identify "Island" (orphan) nodes by
'   tracking which IDs participate in valid, stylable relationships.
'
' TECHNICAL WORKFLOW:
'   1. EDGE SCAN: Iterates through the data worksheet specifically looking
'      for rows classified as TYPE_EDGE.
'   2. MULTI-TARGET RESOLUTION: Splits comma-delimited 'Item' and
'      'Related Item' strings into individual ID arrays.
'   3. RELATIONSHIP VALIDATION:
'      - Cross-references both the Tail (Item) and Head (Related Item)
'        against the 'nodeIds' dictionary (Nodes with valid styles).
'      - Only if BOTH endpoints are stylable is the connection deemed valid.
'   4. CONNECTIVITY MAPPING: Populates 'nodeIdsInRelationships' with the
'      base IDs (ports removed) of every node that has at least one degree.
'
' TECHNICAL NOTES:
'   - Strategy: This is the logic engine for the "Nodes without Relationships"
'     suppression setting.
'   - Complexity: Handles Cartesian product relationships when multiple
'     items are related to multiple targets in a single row.
' ==========================================================================
Private Sub DetermineWhatGraphShouldInclude(ByRef ini As settings, _
                                           ByVal styles As Dictionary, _
                                           ByVal nodeIds As Dictionary, _
                                           ByVal nodeIdsInRelationships As Dictionary)
    Dim data As dataRow

    Dim itemId As String
    Dim relatedItemId As String
    
    Dim Items() As String
    Dim itemIndex As Long
    
    Dim relatedItems() As String
    Dim relatedItemIndex As Long
    
    Dim row As Long
    For row = ini.data.firstRow To ini.data.lastRow
        If GetCell(ini.data.worksheetName, row, ini.data.flagColumn) <> FLAG_COMMENT Then ' row is not a comment
            ' Get the style of the item
            data.styleName = GetCell(ini.data.worksheetName, row, ini.data.styleNameColumn)

            ' Try to determine the style if not supplied
            If data.styleName = vbNullString Then
                data.styleName = DetermineStyleName(ini, row)
            End If

            ' Get the style names in a consistent case
            data.styleName = UCase$(data.styleName)
            
            If data.styleName <> vbNullString Then ' this is not a blank line
                If styles.Exists(data.styleName) Then ' this style should be shown in diagram

                    ' We want data of this style in the output file
                    data.item = GetCell(ini.data.worksheetName, row, ini.data.itemColumn)
                    data.relatedItem = GetCell(ini.data.worksheetName, row, ini.data.isRelatedToItemColumn)

                    If styles.item(data.styleName).styleType = TYPE_EDGE Then ' this line is a relationship

                        If data.item <> vbNullString And UCase$(data.item) <> KEYWORD_EDGE And data.relatedItem <> vbNullString Then ' a tail and head are present

                            Items = split(data.item, COMMA)
                            relatedItems = split(data.relatedItem, COMMA)
                            
                            For itemIndex = LBound(Items) To UBound(Items)
                                For relatedItemIndex = LBound(relatedItems) To UBound(relatedItems)
                                    ' If both the tail and the head in the relationship refer
                                    ' to included nodes having style definitions, track the nodes
                                    ' as "Is Used" so that we later determine island nodes to exclude
                                    ' from the graph.
                                
                                    itemId = RemovePort(Items(itemIndex))
                                    relatedItemId = RemovePort(relatedItems(relatedItemIndex))

                                    If nodeIds.Exists(itemId) And nodeIds.Exists(relatedItemId) Then
                                        If Not nodeIdsInRelationships.Exists(itemId) Then
                                            nodeIdsInRelationships.Add itemId, True
                                        End If
                                
                                        If Not nodeIdsInRelationships.Exists(relatedItemId) Then
                                            nodeIdsInRelationships.Add relatedItemId, True
                                        End If
                                    End If       ' tail and head relate to included nodes
                                Next
                            Next
                        End If                   ' tail and head are non-blank
                    End If                       ' data.styleName = EDGE
                End If                           ' show item = YES
            End If                               ' not a blank line
        End If                                   ' not commented out
    Next row

End Sub

' ==========================================================================
' FUNCTION: ValidateData
'
' PURPOSE:
'   THE SEMANTIC AUDITOR. Performs a structural integrity pass on the Data
'   worksheet to ensure logic is sound before the DOT generation phase.
'
' TECHNICAL WORKFLOW:
'   1. DATA EXTRACTION: Uses 'GetDataRow' to pull row attributes into a UDT
'      for high-speed evaluation.
'   2. STYLE RESOLUTION: Normalizes style names and verifies their existence
'      within the cached 'styles' Dictionary.
'   3. TYPE-SPECIFIC RULES:
'      - TYPE_NODE: Flags errors if the 'Item' is missing or if a
'        'Related Item' is present (which would imply an Edge).
'      - TYPE_EDGE: Flags errors if either 'Item' (Tail) or 'Related Item'
'        (Head) are missing.
'      - SUBGRAPHS: Tracks the 'openSubgraphs' stack. Increments on '{'
'        and decrements on '}'.
'   4. STACK VALIDATION: Flags immediate errors for excess closing braces
'      and a final error if the stack isn't zero at the end of the sheet.
'   5. LOGGING: Invokes 'LogError' to write diagnostic messages back to the
'      worksheet for user correction.
'
' TECHNICAL NOTES:
'   - Layer: Logic Layer / Validation.
'   - DeepWiki Context: Implements the "Error Handling Philosophy" to prevent
'     VBA state loss by catching issues before external execution.
' ==========================================================================
Private Function ValidateData(ByRef ini As settings, ByVal styles As Dictionary) As Long

    Dim data As dataRow
    
    Dim row As Long
    Dim openSubgraphs As Long
    Dim errCnt As Long

    ' Initializations
    openSubgraphs = 0
    errCnt = 0
    
    ' Iterate through the rows of data
    For row = ini.data.firstRow To ini.data.lastRow

        data = GetDataRow(ini, ini.data.worksheetName, row)

        If data.comment <> FLAG_COMMENT Then   ' Don't process the row if it has been commented out
            ' Try to determine the style if not supplied
            If data.styleName = vbNullString Then
                data.styleName = DetermineStyleName(ini, row)
            End If

            ' Get the style names in a consistent case
            data.styleName = UCase$(data.styleName)
            
            ' See if the row has data
            If data.styleName <> vbNullString Then
                ' Determine if this item should be shown in the diagram
                If styles.Exists(data.styleName) Then ' We want data of this style in the output file
                    
                    ' Look up processing attributes from cached stylesheet information
                    data.styleType = styles.item(data.styleName).styleType
                    
                    ' Validate the rows according to object type
                    If data.styleType = TYPE_NODE Then
                        If data.item = vbNullString Then
                            LogError ini, row, GetMessage("errormsgNodeNoItemFound"), errCnt
                        
                        ElseIf data.relatedItem <> vbNullString Then
                            LogError ini, row, GetMessage("errormsgImpliedEdgeType"), errCnt
                        End If
                       
                    ElseIf data.styleType = TYPE_EDGE Then
                        '@Ignore EmptyIfBlock
                        If UCase$(data.item) = KEYWORD_EDGE Then
                            ' No error
                        ElseIf data.item = vbNullString Then
                            LogError ini, row, GetMessage("errormsgEdgeNoItemFound"), errCnt
                        
                        ElseIf data.relatedItem = vbNullString Then
                            LogError ini, row, GetMessage("errormsgEdgeNoRelatedItemFound"), errCnt
                        End If
                        
                    ElseIf data.styleType = TYPE_SUBGRAPH_OPEN Then
                        openSubgraphs = openSubgraphs + 1
                                                
                    ElseIf data.styleType = TYPE_SUBGRAPH_CLOSE Then
                        openSubgraphs = openSubgraphs - 1
    
                        If openSubgraphs < 0 Then
                            LogError ini, row, GetMessage("errormsgBracesExcessClose"), errCnt
                        End If
                    End If
                End If
            End If
        End If
    Next row

    ' Alert the user if it appears that the subgraphs open and close braces are out of balance
    If openSubgraphs > 0 Then
        LogError ini, row, replace(GetMessage("errormsgBracesExcessOpen"), "{openSubgraphs}", openSubgraphs), errCnt
    End If

    ' Return count of errors encountered
    ValidateData = errCnt
    
End Function

' ==========================================================================
' SECTION: DOT SOURCE GENERATION & SYNTAX ASSEMBLY
' ==========================================================================

' ==========================================================================
' FUNCTION: isKeyword
'
' PURPOSE:
'   Identifies if a worksheet entry represents a global Graphviz
'   configuration scope rather than a specific unique entity.
'
' TECHNICAL WORKFLOW:
'   1. NORMALIZATION: Converts the 'item' string to uppercase.
'   2. COMPARISON: Evaluates against core DOT keywords: 'NODE', 'EDGE',
'      or 'GRAPH'.
'   3. LOGICAL RETURN: Returns TRUE if the entry matches any of the
'      global scope triggers.
'
' TECHNICAL NOTES:
'   - Strategy: Prevents the parser from treating global attribute blocks
'     as individual nodes or edges.
'   - Layer: Logic Layer / Parser.
' ==========================================================================
Private Function isKeyword(ByVal item As String) As Boolean
    isKeyword = (UCase$(item) = KEYWORD_NODE) Or (UCase$(item) = KEYWORD_EDGE) Or (UCase$(item) = KEYWORD_GRAPH)
End Function

' ==========================================================================
' PROCEDURE: CreateGraphvizSource
'
' PURPOSE:
'   THE DOT ASSEMBLER. Orchestrates the construction of the final .gv
'   source string by synthesizing worksheet data into structured DOT syntax.
'
' TECHNICAL WORKFLOW:
'   1. HEADER INITIALIZATION: Establishes the graph's fundamental signature
'      (Strict status, directed vs. undirected command) and opens the
'      primary Graphviz block with '{'.
'   2. GLOBAL DIRECTIVES: Invokes 'ProcessGraphOptions' to inject
'      workbook-wide settings (rankdir, splines, imagepath, etc.).
'   3. STATE MANAGEMENT: Initializes the 'clusterCnt' for subgraph naming
'      and a dynamic 'indent' counter for human-readable code formatting.
'   4. MAIN PARSING LOOP: Iterates through the data worksheet, routing
'      rows to specialized handlers:
'      - 'ProcessNode' / 'ProcessEdge': Standard graph entities.
'      - 'ProcessSubgraphOpen' / 'Close': Handles cluster naming and
'        recursive indentation shifts.
'      - 'ProcessKeyword' / 'ProcessNative': Global overrides and raw
'        code passthrough ('>').
'   5. DEBUG ENHANCEMENT: If 'debug' mode is enabled, it automatically
'      injects row metadata into labels via 'FormatDebugLabel'.
'   6. CLOSURE: Finalizes the buffer with a closing brace '}'.
'
' TECHNICAL NOTES:
'   - Performance: Uses 'Join(Array(...))' for efficient string concatenation.
'   - DeepWiki Context: Implements the "Transformation Pipeline" and
'     "Stack-based Parsing" logic for subgraphs.
' ==========================================================================
Private Sub CreateGraphvizSource(ByRef ini As settings, _
                                    ByVal styles As Dictionary, _
                                    ByVal nodeIds As Dictionary, _
                                    ByVal relationshipIds As Dictionary, _
                                    ByRef graphvizSource As String)
    ' Subgraph cluster counter
    Dim clusterCnt As Long
    clusterCnt = 0
    
    ' Set the  Graphviz 'strict' directive
    Dim graphStrict As String
    If ini.graph.addStrict Then
        graphStrict = "strict"
    End If
    
    ' Create the first lines of the dot graph program
    graphvizSource = Trim$(graphStrict & " " & ini.graph.command & " " & AddQuotes(Mid$(ActiveWorkbook.name, 1, InStr(1, ActiveWorkbook.name, ".") - 1))) & vbNewLine
    graphvizSource = graphvizSource & OPEN_BRACE & vbNewLine
    
    ' Establish source indentation
    Dim indent As Long
    indent = IncreaseIndent(0)
    
    ' Write out the graph directives before processing the rows of data
    ProcessGraphOptions graphvizSource, ini, indent
    
    ' Iterate through the rows of data
    Dim row As Long
    Dim data As dataRow
    For row = ini.data.firstRow To ini.data.lastRow

        data = GetDataRow(ini, ini.data.worksheetName, row)

        ' Don't process the row if it has been commented out
        If data.comment <> FLAG_COMMENT Then
        
            ' Try to determine the style if not supplied
            If data.styleName = vbNullString Then
                data.styleName = DetermineStyleName(ini, row)
            End If

            ' Treat all style names as uppercase for consistency
            data.styleName = UCase$(data.styleName)
            
            ' See if the row has data
            '@Ignore EmptyIfBlock
            If data.styleName = vbNullString Then
                ' No style was specified, assume the row is blank and skip it.
            Else
                ' Determine if this item should be shown in the diagram
                Dim showStyle As Boolean
                showStyle = styles.Exists(data.styleName)
                
                Dim boolKeyword As Boolean
                boolKeyword = isKeyword(data.item)
                
                If showStyle Or boolKeyword Then ' We want data of this style in the output file
                    
                    ' Look up processing attributes from cached stylesheet information
                    data.styleType = styles.item(data.styleName).styleType
                    
                    If ini.graph.includeStyleFormat And showStyle Then
                        data.format = styles.item(data.styleName).styleFormat
                    Else
                        data.format = vbNullString
                    End If
                    
                    ' Append information to the label if debugging is enabled
                    If ini.graph.debug Then
                        data.label = FormatDebugLabel(row, data)
                        data.xLabel = FormatDebugXLabel(row, data)
                    End If
                    
                    ' Process the rows according to object type
                    If boolKeyword Then
                        graphvizSource = Join(Array(graphvizSource, ProcessKeyword(ini, data, indent)), vbNullString)

                    ElseIf data.styleType = TYPE_NODE Then
                        graphvizSource = Join(Array(graphvizSource, ProcessNode(ini, data, indent, relationshipIds)), vbNullString)

                    ElseIf data.styleType = TYPE_EDGE Then
                        graphvizSource = Join(Array(graphvizSource, ProcessEdge(ini, data, indent, nodeIds)), vbNullString)

                    ElseIf data.styleType = TYPE_SUBGRAPH_OPEN Then
                        graphvizSource = Join(Array(graphvizSource, ProcessSubgraphOpen(ini, data, indent, clusterCnt)), vbNullString)
                        indent = IncreaseIndent(indent)
                        
                    ElseIf data.styleType = TYPE_SUBGRAPH_CLOSE Then
                        indent = DecreaseIndent(indent)
                        graphvizSource = Join(Array(graphvizSource, ProcessSubgraphClose(ini, data, indent)), vbNullString)

                    ElseIf data.styleType = TYPE_KEYWORD Then
                        graphvizSource = graphvizSource & ProcessKeyword(ini, data, indent)

                    ElseIf data.styleType = TYPE_NATIVE Then
                        graphvizSource = Join(Array(graphvizSource, ProcessNative(ini, data, indent)), vbNullString)

                    '@Ignore EmptyElseBlock
                    Else
                        ' Not recognized, ignore it
                    End If
                End If
            End If
        End If
    Next row

    ' Write the last dot statement to terminate the dot source file
    indent = DecreaseIndent(indent)
    graphvizSource = Join(Array(graphvizSource, Space(indent * ini.source.indent), CLOSE_BRACE, vbNewLine), vbNullString)

End Sub

' ==========================================================================
' SECTION: GLOBAL GRAPH DIRECTIVES & ENGINE-SPECIFIC OPTIONS
' ==========================================================================

' ==========================================================================
' PROCEDURE: ProcessGraphOptions
'
' PURPOSE:
'   THE GLOBAL CONFIGURATOR. Translates high-level project settings into
'   valid DOT graph-level attribute statements.
'
' TECHNICAL WORKFLOW:
'   1. CORE VISUALS: Applies global attributes like 'splines', 'bgcolor'
'      (transparency), 'center', and 'concentrate' using 'AddAttributeLine'.
'   2. ASSET RESOLUTION: Injects the 'imagepath' directory list to ensure
'      Graphviz can find external icons/backgrounds.
'   3. ENGINE-SPECIFIC PARAMETERS: Uses a 'Select Case' structure to apply
'      parameters tailored to the active layout engine:
'      - DOT: 'rankdir', 'compound', 'newrank', 'clusterrank'.
'      - NEATO/FDP/SFDP: 'overlap', 'dim/dimen', 'mode', 'model', 'smoothing'.
'      - CIRCO/TWOPI/OSAGE: 'outputorder'.
'   4. ORIENTATION: Handles the 'Rotate 90' flag for landscape renderings.
'   5. POWER-USER OVERRIDE: Appends the 'ini.graph.options' string at the
'      very end, allowing manual DOT code from the 'Settings' worksheet to
'      supersede any automated assignments.
'
' TECHNICAL NOTES:
'   - Strategy: Decouples the rendering engine's vast attribute set from
'     the Excel UI via the 'settings' UDT and 'AddAttributeLine' helper.
'   - DeepWiki Context: Directly implements the engine logic described in
'     the "Graphviz Ribbon Tab" documentation.
' ==========================================================================

'
Private Sub ProcessGraphOptions(ByRef graphvizSource As String, ByRef ini As settings, ByVal indent As Long)

    Dim spaces As String
    
    ' Create the indentation string
    spaces = Space(indent * ini.source.indent)
    
    ' Latest Windows version requires you to use DOT.EXE with layout specified as a graph option.
    If ini.graph.layout <> "dot" Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_LAYOUT, ini.graph.layout
    End If
    
    ' Specify how the edges should be drawn and include as the "spline" parameter
    If Trim$(ini.graph.splines) <> vbNullString Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_SPLINES, ini.graph.splines
    End If
    
    ' Make the background transparent if desired
    If ini.graph.transparentBackground Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_BGCOLOR, "transparent"
    End If
    
    If ini.graph.center Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_CENTER, TOGGLE_TRUE
    End If
       
    If ini.graph.concentrate Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_CONCENTRATE, TOGGLE_TRUE
    End If
    
    If ini.graph.forceLabels Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_FORCELABELS, TOGGLE_TRUE
    End If
    
    ' Specify the directory path where images are located
    If ini.graph.includeGraphImagePath Then
        If ini.graph.imagePath <> vbNullString Then
            AddAttributeLine graphvizSource, spaces, GRAPHVIZ_IMAGEPATH, AddQuotes(ini.graph.imagePath)
        End If
    End If
    
    ' Process the graph options which are specific to layout engines
    Select Case ini.graph.layout
        Case LAYOUT_CIRCO
            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case LAYOUT_DOT
            If ini.graph.rankdir <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_RANKDIR, ini.graph.rankdir
            End If

            If ini.graph.clusterrank <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_CLUSTERRANK, ini.graph.clusterrank
            End If

            If ini.graph.compound Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_COMPOUND, TOGGLE_TRUE
            End If

            If ini.graph.ordering <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_ORDERING, ini.graph.ordering
            End If

            If ini.graph.newrank Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_NEWRANK, TOGGLE_TRUE
            End If
    
            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case LAYOUT_FDP
            If ini.graph.layoutDim <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_DIM, ini.graph.layoutDim
            End If

            If ini.graph.layoutDimen <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_DIMEN, ini.graph.layoutDimen
            End If

            If ini.graph.overlap <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OVERLAP, ini.graph.overlap
            End If

            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case LAYOUT_NEATO
            If ini.graph.layoutDim <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_DIM, ini.graph.layoutDim
            End If

            If ini.graph.layoutDimen <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_DIMEN, ini.graph.layoutDimen
            End If
            
            If ini.graph.overlap <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OVERLAP, ini.graph.overlap
            End If

            If ini.graph.mode <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_MODE, ini.graph.mode
            End If

            If ini.graph.model <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_MODEL, ini.graph.model
            End If

            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case LAYOUT_OSAGE
            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case LAYOUT_SFDP
            If ini.graph.layoutDim <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_DIM, ini.graph.layoutDim
            End If

            If ini.graph.layoutDimen <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_DIMEN, ini.graph.layoutDimen
            End If
            
            If ini.graph.mode <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_MODE, ini.graph.mode
            End If

            If ini.graph.overlap <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OVERLAP, ini.graph.overlap
            End If

            If ini.graph.smoothing <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_SMOOTHING, ini.graph.smoothing
            End If

            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case LAYOUT_TWOPI
            If Trim$(ini.graph.outputOrder) <> vbNullString Then
                AddAttributeLine graphvizSource, spaces, GRAPHVIZ_OUTPUTORDER, ini.graph.outputOrder
            End If
            
        Case Else
    End Select

    If ini.graph.orientation Then
        AddAttributeLine graphvizSource, spaces, GRAPHVIZ_ROTATE, "90"
    End If
    
    ' Graph options from the settings worksheet come last to give the ability to override anything above
    If ini.graph.options <> vbNullString Then
        graphvizSource = graphvizSource & spaces & ini.graph.options & vbNewLine
    End If
End Sub

' ==========================================================================
' PROCEDURE: AddAttributeLine
'
' PURPOSE:
'   A low-level string-assembly utility that appends a single, valid
'   Graphviz attribute statement to the source buffer.
'
' TECHNICAL WORKFLOW:
'   1. CONCATENATION: Combines the current indentation, attribute name,
'      assignment operator (=), and value.
'   2. TERMINATION: Appends a semicolon (SEMICOLON) and a newline (vbNewLine)
'      to ensure strict adherence to DOT language syntax.
'   3. PERFORMANCE: Uses 'Join(Array(...))' to minimize memory allocation
'      overhead during large-scale graph generation.
'
' TECHNICAL NOTES:
'   - Strategy: Centralizes the "semicolon-terminated" pattern to prevent
'     syntax errors across all object handlers (Node, Edge, Graph).
'   - Constraint: Assumes 'attributeValue' is already properly formatted
'     (e.g., quoted or numeric).
' ==========================================================================
Private Sub AddAttributeLine(ByRef graphvizSource As String, ByVal spaces As String, ByVal attributeName As String, ByVal attributeValue As String)
    graphvizSource = Join(Array(graphvizSource, spaces, Trim$(attributeName), "=", attributeValue, SEMICOLON, vbNewLine), vbNullString)
End Sub

' ==========================================================================
' SECTION: INDENTATION & NESTING LOGIC
' ==========================================================================

' ==========================================================================
' FUNCTION: IncreaseIndent
'
' PURPOSE:
'   Increments the indentation depth tracker used to generate human-readable
'   and structurally organized DOT source code.
'
' TECHNICAL WORKFLOW:
'   1. STACK ADVANCE: Adds a value of 1 to the current 'indent' level.
'
' USAGE:
'   - Triggered by 'CreateGraphvizSource' immediately after processing
'     a 'TYPE_SUBGRAPH_OPEN' ({) row.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Essential for the "Stack-based Parsing" logic used
'     to maintain correct nesting hierarchy in complex diagrams.
' ==========================================================================
Private Function IncreaseIndent(ByVal indent As Long) As Long
    IncreaseIndent = indent + 1
End Function

' ==========================================================================
' FUNCTION: DecreaseIndent
'
' PURPOSE:
'   Decrements the indentation depth tracker when exiting a nested
'   Graphviz scope (e.g., closing a Subgraph or Cluster).
'
' TECHNICAL WORKFLOW:
'   1. STACK RETREAT: Subtracts 1 from the current 'indent' level.
'   2. BOUNDARY PROTECTION: Implements a safety floor to ensure the indent
'      never drops below 0, preventing string-generation errors.
'
' USAGE:
'   - Invoked by 'CreateGraphvizSource' after processing a
'     'TYPE_SUBGRAPH_CLOSE' (}) row or before closing the main graph block.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Works in tandem with 'IncreaseIndent' to support
'     the "Stack-based Parsing" architecture for hierarchical grouping.
' ==========================================================================
Private Function DecreaseIndent(ByVal indent As Long) As Long
    DecreaseIndent = indent - 1
    If DecreaseIndent < 0 Then
        DecreaseIndent = 0
    End If
End Function

' ==========================================================================
' SECTION: DATA MAPPING & SYNTACTIC HELPERS
' ==========================================================================

' ==========================================================================
' FUNCTION: GetDataRow
'
' PURPOSE:
'   THE DATA MAPPER. Extracts and structures raw worksheet data from a
'   single row into a 'dataRow' UDT for high-speed internal processing.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN RESOLUTION: Maps logical Graphviz properties (Labels,
'      Tooltips, Ports) to their physical worksheet coordinates using the
'      'ini.data' settings contract.
'   2. ATTRIBUTE GATHERING: Captures core entity data:
'      - IDENTIFIERS: 'item' (Node ID/Tail) and 'relatedItem' (Head).
'      - LABELS: Standard, XLabel (external), TailLabel, and HeadLabel.
'      - METADATA: 'Tooltip' and 'extraAttrs' for DOT passthrough.
'   3. STATE CAPTURE: Records the 'comment' flag and 'styleName' to inform
'      the subsequent classification and validation stages.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Implements the "GetDataRow internal structure"
'     specified in the Defining Nodes & Edges architecture page.
'   - Strategy: Centralizes all worksheet-to-VBA field mapping to isolate
'     the core logic from changes in the spreadsheet layout.
' ==========================================================================
Public Function GetDataRow(ByRef ini As settings, ByVal worksheetName As String, ByVal row As Long) As dataRow

    GetDataRow.comment = GetCell(worksheetName, row, ini.data.flagColumn)
    GetDataRow.item = GetCell(worksheetName, row, ini.data.itemColumn)
    GetDataRow.label = GetCell(worksheetName, row, ini.data.labelColumn)
    GetDataRow.xLabel = GetCell(worksheetName, row, ini.data.xLabelColumn)
    GetDataRow.tailLabel = GetCell(worksheetName, row, ini.data.tailLabelColumn)
    GetDataRow.headLabel = GetCell(worksheetName, row, ini.data.headLabelColumn)
    GetDataRow.Tooltip = GetCell(worksheetName, row, ini.data.tooltipColumn)
    GetDataRow.relatedItem = GetCell(worksheetName, row, ini.data.isRelatedToItemColumn)
    GetDataRow.styleName = GetCell(worksheetName, row, ini.data.styleNameColumn)
    GetDataRow.extraAttrs = GetCell(worksheetName, row, ini.data.extraAttributesColumn)
    GetDataRow.errorMessage = GetCell(worksheetName, row, ini.data.errorMessageColumn)

End Function

''
' STYLE CACHE ENGINE: Loads all 'Yes' flagged styles into a high-speed Dictionary.
' 1. Skips commented rows in the Style sheet.
' 2. Filters for the currently active View (column).
' 3. Instantiates Style class objects for every enabled style.
' Used to prevent redundant worksheet lookups during the main generation loop.
' @param showStyleColumn [Long]: The column index of the current View.
'
Private Function CacheEnabledStyles(ByRef ini As settings, ByVal showStyleColumn As Long) As Dictionary

    ' Dictionary to hold the key and associated values
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    ' Loop through the specified range
    Dim row As Long
    Dim styleName As String
    
    For row = ini.styles.firstRow To ini.styles.lastRow
        '@Ignore EmptyIfBlock
        If StylesSheet.Cells.item(row, ini.styles.flagColumn).value = FLAG_COMMENT Then
            ' Comment row, ignore it
        ElseIf StylesSheet.Cells.item(row, showStyleColumn).value = TOGGLE_YES Then
            ' Retrieve the style name
            styleName = UCase$(StylesSheet.Cells.item(row, ini.styles.nameColumn).value)

            If styleName <> vbNullString Then    ' a style name is present
                If Not dictionaryObj.Exists(styleName) Then ' ignore duplicate style names
                    Set dictionaryObj.item(styleName) = GetStyle(StylesSheet.Cells.item(row, ini.styles.typeColumn), _
                                                              StylesSheet.Cells.item(row, ini.styles.formatColumn))
                End If
            End If
        End If
    Next row

    Set CacheEnabledStyles = dictionaryObj
    
End Function

' ==========================================================================
' SECTION: OBJECT FACTORIES & ERROR REPORTING
' ==========================================================================

' ==========================================================================
' FUNCTION: CacheEnabledStyles
'
' PURPOSE:
'   THE STYLE CACHE ENGINE. Loads all active style definitions for the
'   selected View into a high-speed Dictionary to optimize rendering performance.
'
' TECHNICAL WORKFLOW:
'   1. DICTIONARY INIT: Instantiates a new 'Dictionary' to serve as the
'      in-memory style registry.
'   2. VIEW-BASED FILTERING: Scans the Styles sheet and only processes rows
'      where the 'showStyleColumn' (the active View) is set to 'TOGGLE_YES'.
'   3. DUPLICATE PROTECTION: Identifies the 'styleName' (normalized to UCase)
'      and ensures only the first occurrence of a unique name is cached.
'   4. OBJECT HYDRATION: Invokes the 'GetStyle' factory function to create
'      'Style' class instances, populating them with 'ObjectType' and
'      'Format' attributes.
'
' TECHNICAL NOTES:
'   - Layer: Logic Layer / Caching.
'   - DeepWiki Context: Implements the "Multi-layered Caching" strategy
'     noted in the Style Designer documentation to prevent redundant
'     worksheet I/O during the main DOT generation loop.
' ==========================================================================
Public Function GetStyle(ByVal styleType As String, ByVal styleFormat As String) As style

    Dim value As style
    Set value = New style
        
    value.styleType = styleType
    value.styleFormat = styleFormat
    
    Set GetStyle = value

End Function

' ==========================================================================
' PROCEDURE: LogError
'
' PURPOSE:
'   THE IN-SHEET ERROR LOGGER. Provides visual and textual feedback to the
'   user when a row fails structural or semantic validation.
'
' TECHNICAL WORKFLOW:
'   1. VISUAL FLAGGING: Updates the 'flagColumn' with the 'FLAG_ERROR'
'      constant, typically triggering Excel conditional formatting (e.g.,
'      red background).
'   2. MESSAGE INJECTION: Writes the descriptive 'errorMessage' string
'      directly into the 'errorMessageColumn' for the specific failing row.
'   3. STATE ACCUMULATION: Increments the 'errCnt' by reference, which
'      serves as the primary "Kill Switch" for the rendering pipeline.
'
' TECHNICAL NOTES:
'   - Strategy: Implements the "ValidateData" pattern to prevent passing
'     malformed DOT code to the external Graphviz engine.
'   - Layer: UI / Data Management.
' ==========================================================================
Private Sub LogError(ByRef ini As settings, ByVal row As Long, ByVal errorMessage As String, ByRef errCnt As Long)

    SetCell ini.data.worksheetName, row, ini.data.flagColumn, FLAG_ERROR
    SetCell ini.data.worksheetName, row, ini.data.errorMessageColumn, errorMessage

    errCnt = errCnt + 1
    
End Sub

' ==========================================================================
' FUNCTION: FormatId
'
' PURPOSE:
'   THE ID FORMATTER. Sanitizes Node IDs for the DOT engine by applying
'   correct quoting and handling specialized port/compass point syntax.
'
' TECHNICAL WORKFLOW:
'   1. PORT DETECTION: Scans the 'nodeId' for the colon (:) delimiter.
'   2. CONDITIONAL PORT HANDLING:
'      - If 'includePorts' is TRUE: Individually quotes the Node ID and
'        conditionally quotes the port/compass segment (e.g., "Node":"port").
'      - If 'includePorts' is FALSE: Discards the port segment and returns
'        only the quoted base Node ID.
'   3. STANDARD QUOTING: For IDs without ports, applies 'AddQuotes' to
'      ensure strings with spaces or special characters are valid DOT tokens.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Directly supports the "Defining Nodes & Edges"
'     architecture by enabling advanced port syntax (node:port).
'   - Strategy: Prevents syntax errors in the Graphviz parser caused by
'     unquoted reserved characters or whitespace.
' ==========================================================================
Private Function FormatId(ByVal nodeId As String, ByVal includePorts As Boolean) As String

    Dim formattedId As String
    
    ' Build the id, taking ports into consideration
    If InStr(nodeId, ":") > 0 Then  ' nodeId specifies a port.
        If includePorts Then        ' wrap both sides of the id in quotes
            formattedId = AddQuotes(GetStringTokenAtPosition(nodeId, ":", 1)) & ":" & AddQuotesConditionally(GetStringTokenAtPosition(nodeId, ":", 2))
        Else    ' strip the port off
            formattedId = AddQuotes(GetStringTokenAtPosition(nodeId, ":", 1))
        End If
    Else        ' no port was specified
        formattedId = AddQuotes(nodeId)
    End If

    FormatId = formattedId
    
End Function

' ==========================================================================
' FUNCTION: FormatDebugLabel
'
' PURPOSE:
'   THE DEBUG OVERLAY. Injects row numbers and connectivity metadata into
'   object labels to facilitate visual auditing of the graph structure.
'
' TECHNICAL WORKFLOW:
'   1. HTML SAFETY CHECK: Invokes 'IsLabelHTMLLike' to detect labels wrapped
'      in angle brackets (< >). Skips debugging for these to avoid corrupting
'      Graphviz HTML-table syntax.
'   2. METADATA COMPOSITION:
'      - TYPE_EDGE: Appends "(Row: # Tail->Head)" to the label.
'      - TYPE_NODE: Appends "(Row: # ID)" to the label.
'      - TYPE_SUBGRAPH: Appends "(Row: #)" for cluster identification.
'   3. STRING INJECTION: Concatenates the debug string with a newline (NEWLINE)
'      if a label already exists, or replaces a null label with the metadata.
'
' TECHNICAL NOTES:
'   - Layer: Logic / Diagnostics.
'   - Strategy: Allows developers to trace rendered shapes back to their
'     exact source row in the Excel Data worksheet.
' ==========================================================================
Private Function FormatDebugLabel(ByVal row As Long, ByRef data As dataRow) As String
                        
    Dim debugstr As String

    FormatDebugLabel = data.label
    
    If Not IsLabelHTMLLike(data.label) Then
        If data.styleType = TYPE_EDGE Then
            debugstr = "(Row: " & row & " " & FormatId(data.item, True) & "->" & FormatId(data.relatedItem, True) & ")"
                        
            If data.label = vbNullString Then
                FormatDebugLabel = debugstr
            Else
                FormatDebugLabel = data.label & NEWLINE & debugstr
            End If
                        
        ElseIf data.styleType = TYPE_NODE Then
            debugstr = "(Row: " & row & " " & AddQuotes(data.item) & ")"
                            
            If data.label = vbNullString Then
                FormatDebugLabel = debugstr
            Else
                FormatDebugLabel = data.label & NEWLINE & debugstr
            End If
                        
        ElseIf data.styleType = TYPE_SUBGRAPH_OPEN Then
            debugstr = "(Row: " & row & ")"
                            
            If data.label = vbNullString Then
                FormatDebugLabel = debugstr
            Else
                FormatDebugLabel = data.label & NEWLINE & debugstr
            End If
        End If
    End If
    
End Function

' ==========================================================================
' SECTION: DEBUGGING & EXTERNAL LABEL FORMATTING
' ==========================================================================

' ==========================================================================
' FUNCTION: FormatDebugXLabel
'
' PURPOSE:
'   THE XLABEL DEBUG OVERLAY. Injects source row numbers and entity mapping
'   into the external 'xLabel' field when debugging is enabled.
'
' TECHNICAL WORKFLOW:
'   1. HTML SAFETY CHECK: Uses 'IsLabelHTMLLike' to detect angle-bracket
'      syntax (<...>); if present, the debug string is suppressed to
'      prevent breaking Graphviz HTML-label parsing.
'   2. METADATA COMPOSITION:
'      - TYPE_EDGE: Formats a string containing the row index and the
'        Tail->Head relationship mapping.
'      - TYPE_NODE: Formats a string containing the row index and Node ID.
'   3. STRING INJECTION: Appends the 'debugstr' to the existing 'xLabel'
'      using a 'NEWLINE' constant, provided the xLabel is not null.
'
' TECHNICAL NOTES:
'   - Layer: Logic / Diagnostics.
'   - Usage: Complements 'FormatDebugLabel' by providing traceability
'     for external (floating) labels in complex layouts.
' ==========================================================================
Private Function FormatDebugXLabel(ByVal row As Long, ByRef data As dataRow) As String
                        
    Dim debugstr As String

    FormatDebugXLabel = data.xLabel

    If Not IsLabelHTMLLike(data.label) Then
        If data.styleType = TYPE_EDGE Then
            debugstr = "(Row: " & row & " " & AddQuotes(data.item) & "->" & AddQuotes(data.relatedItem) & ")"
            
            If data.xLabel <> vbNullString Then
                FormatDebugXLabel = data.xLabel & NEWLINE & debugstr
            End If
            
        ElseIf data.styleType = TYPE_NODE Then
            debugstr = "(Row: " & row & " " & AddQuotes(data.item) & ")"
                            
            If data.xLabel <> vbNullString Then
                FormatDebugXLabel = data.xLabel & NEWLINE & debugstr
            End If
        End If
    End If
    
End Function

' ==========================================================================
' SECTION: EDGE LABEL ASSEMBLY
' ==========================================================================

' ==========================================================================
' FUNCTION: FormatEdgeLabels
'
' PURPOSE:
'   Builds the complete Graphviz edge-label attribute string, combining
'   template-driven placeholders with explicit label fields from the data row.
'   Supports all four Graphviz label positions: label, xlabel, taillabel,
'   and headlabel.
'
' TECHNICAL WORKFLOW:
'   1. TEMPLATE EXPANSION:
'        - Begins with 'styleAttributes' (the style-layer template).
'        - For each supported placeholder ({label}, {xlabel}, {taillabel},
'          {headlabel}), replaces it with the corresponding data value.
'
'   2. FALLBACK ATTRIBUTE EMISSION:
'        - If a placeholder is *not* present in the template, appends the
'          appropriate attribute (e.g., " label=", " xlabel=") when the
'          corresponding data field is non-blank.
'
'   3. BLANK-LABEL OVERRIDE:
'        - When edge labels are enabled and the main label is blank:
'             • If 'blankEdgeLabels' = TRUE, emits the Graphviz "\E" token.
'             • Otherwise, emits an empty formatted label.
'
'   4. SANITIZATION:
'        - All emitted label values pass through 'FormatLabel' (or AddQuotes
'          for "\E") to ensure correct quoting and HTML-label handling.
'
' TECHNICAL NOTES:
'   - This function merges style-layer templates with data-layer values,
'     enabling both declarative styling and dynamic label substitution.
'   - DeepWiki Context: Implements the multi-label synthesis rules described
'     in the "Defining Nodes & Edges" and "Styles" documentation.
' ==========================================================================
Private Function FormatEdgeLabels(ByRef ini As settings, ByRef data As dataRow, ByRef styleAttributes) As String

    Dim edgeLabel As String
    edgeLabel = styleAttributes
    
    ' Handle label= attribute
    If ini.graph.includeEdgeLabels Then
        If InStr(1, edgeLabel, "{label}", vbTextCompare) Then
            ' Expand the {label} placeholder
            If data.label = vbNullString And ini.graph.blankEdgeLabels Then
                edgeLabel = replace(edgeLabel, "{label}", "\E", 1, -1, vbTextCompare)
            Else
                edgeLabel = replace(edgeLabel, "{label}", data.label, 1, -1, vbTextCompare)
            End If
        Else
            ' Append the label
            If data.label <> vbNullString Then
                edgeLabel = edgeLabel & " label=" & FormatLabel(data.label)
            ElseIf ini.graph.blankEdgeLabels Then
                edgeLabel = edgeLabel & " label=" & AddQuotes("\E")
            End If
        End If
    End If

    ' Handle xlabel= attribute
    If InStr(1, edgeLabel, "{xlabel}", vbTextCompare) Then
        ' Expand the {xlabel} placeholder
        edgeLabel = replace(edgeLabel, "{xlabel}", data.xLabel, 1, -1, vbTextCompare)
    Else
        ' Append the label
        If data.xLabel <> vbNullString Then
            edgeLabel = edgeLabel & " xlabel=" & FormatLabel(data.xLabel)
        End If
    End If
    
    ' Handle taillabel= attribute
    If InStr(1, edgeLabel, "{taillabel}", vbTextCompare) Then
        ' Expand the {taillabel} placeholder
        edgeLabel = replace(edgeLabel, "{taillabel}", data.tailLabel, 1, -1, vbTextCompare)
    Else
        ' Append the taillabel
        If data.tailLabel <> vbNullString Then
            edgeLabel = edgeLabel & " taillabel=" & FormatLabel(data.tailLabel)
        End If
    End If
   
    ' Handle headlabel= attribute
    If InStr(1, edgeLabel, "{headlabel}", vbTextCompare) Then
        ' Expand the {taillabel} placeholder
        edgeLabel = replace(edgeLabel, "{headlabel}", data.headLabel, 1, -1, vbTextCompare)
    Else
        ' Append the headlabel
        If data.headLabel <> vbNullString Then
            edgeLabel = edgeLabel & " headlabel=" & FormatLabel(data.headLabel)
        End If
    End If
    
    FormatEdgeLabels = edgeLabel
    
End Function

' ==========================================================================
' FUNCTION: FormatGraphLabels
'
' PURPOSE:
'   Produces the Graphviz graph-level label attribute by merging a style
'   template with the data row's primary label. Supports placeholder-based
'   substitution as well as fallback attribute emission.
'
' TECHNICAL WORKFLOW:
'   1. TEMPLATE EXPANSION:
'        - Starts with 'styleAttributes' (the style-layer template).
'        - If the template contains the {label} token, replaces it with the
'          data row's label value.
'
'   2. FALLBACK ATTRIBUTE EMISSION:
'        - If no {label} placeholder is present, appends a standard
'          "label=" attribute using the formatted label text.
'
'   3. SANITIZATION:
'        - All emitted label values pass through 'FormatLabel' to ensure
'          correct quoting and HTML-label handling.
'
' TECHNICAL NOTES:
'   - Graph-level labels differ from node/edge labels: no blank-label override
'     logic is applied here; the label is always emitted or substituted.
'   - DeepWiki Context: Implements the graph-label synthesis rules described
'     in the "Graph Attributes" and "Styles" documentation.
' ==========================================================================
Private Function FormatGraphLabels(ByRef ini As settings, ByRef data As dataRow, ByRef styleAttributes) As String

    Dim graphLabel As String
    graphLabel = styleAttributes
    
    ' label=
    If InStr(1, graphLabel, "{label}", vbTextCompare) Then
        ' Expand the {label} placeholder
        graphLabel = replace(graphLabel, "{label}", data.label, 1, -1, vbTextCompare)
    Else
        ' Append the label
        graphLabel = graphLabel & " label=" & FormatLabel(data.label)
    End If

    FormatGraphLabels = graphLabel
    
End Function

' ==========================================================================
' SECTION: NODE LABEL ASSEMBLY
' ==========================================================================

' ==========================================================================
' FUNCTION: FormatNodeLabels
'
' PURPOSE:
'   Builds the complete Graphviz node-label attribute string by merging
'   style-layer templates with data-layer values. Supports both primary
'   labels ("label=") and external labels ("xlabel="), with placeholder
'   expansion when template tokens are present.
'
' TECHNICAL WORKFLOW:
'   1. TEMPLATE EXPANSION:
'        - Begins with 'styleAttributes' (the style-layer template).
'        - Replaces the {label} and {xlabel} placeholders when present.
'
'   2. FALLBACK ATTRIBUTE EMISSION:
'        - If a placeholder is *not* present:
'             • Emits "label=" when node labels are enabled and the data
'               label is non-blank.
'             • Emits an explicit empty label ("") when labels are enabled
'               but 'blankNodeLabels' is FALSE.
'             • Emits "xlabel=" when external labels are enabled and data
'               exists in the xLabel field.
'
'   3. SANITIZATION:
'        - All emitted label values pass through 'FormatLabel' to ensure
'          correct quoting and HTML-label handling.
'
' TECHNICAL NOTES:
'   - Supports both "ID-as-Label" and "Clean Node" aesthetics depending on
'     the 'blankNodeLabels' setting.
'   - DeepWiki Context: Implements the node-label synthesis rules described
'     in the "Defining Nodes & Edges" and "Styles" documentation.
' ==========================================================================
Private Function FormatNodeLabels(ByRef ini As settings, ByRef data As dataRow, ByVal styleAttributes As String) As String

    Dim nodeLabel As String
    nodeLabel = styleAttributes
    
    ' Handle label= attribute
    If ini.graph.includeNodeLabels Then
        If InStr(1, nodeLabel, "{label}", vbTextCompare) Then
            ' Expand the {label} placeholder
            nodeLabel = replace(nodeLabel, "{label}", data.label, 1, -1, vbTextCompare)
        Else
            ' Append the label
            If data.label <> vbNullString Then
                nodeLabel = nodeLabel & " label=" & FormatLabel(data.label)
            ElseIf Not ini.graph.blankNodeLabels Then
                nodeLabel = nodeLabel & " label=" & FormatLabel(vbNullString)
            End If
        End If
    End If
    
    ' Handle xlabel= attribute
    If ini.graph.includeNodeXLabels Then
        If InStr(1, nodeLabel, "{xlabel}", vbTextCompare) Then
            ' Expand the {xlabel} placeholder
            nodeLabel = replace(nodeLabel, "{xlabel}", data.xLabel, 1, -1, vbTextCompare)
        Else
            ' Append the xlabel
            If data.xLabel <> vbNullString Then
                nodeLabel = nodeLabel & " xlabel=" & FormatLabel(data.xLabel)
            End If
        End If
    End If

    FormatNodeLabels = nodeLabel
    
End Function

' ==========================================================================
' SECTION: SUBGRAPH & CLUSTER INITIALIZATION
' ==========================================================================

' ==========================================================================
' FUNCTION: ProcessSubgraphOpen
'
' PURPOSE:
'   THE HIERARCHY HANDLER. Generates the opening statement for a Graphviz
'   subgraph or cluster, managing automatic naming and attribute merging.
'
' TECHNICAL WORKFLOW:
'   1. NAME RESOLUTION:
'      - Extracts the name from the 'item' column (text before '{').
'      - If blank, it auto-increments 'clusterCnt' and assigns a "cluster_"
'        prefix to ensure Graphviz renders a bounding box.
'   2. ATTRIBUTE INJECTION:
'      - Appends the base 'format' from the cached Style definition.
'      - Merges 'extraAttrs' if 'includeExtraAttributes' is enabled.
'   3. LABEL HANDLING:
'      - Checks for the "{label}" placeholder in the format string for
'        dynamic injection.
'      - If no placeholder exists, it appends a standard 'label=' attribute
'        sanitized via 'FormatLabel'.
'   4. SVG ENHANCEMENT: Appends 'tooltip=' attributes if the output is set
'      to SVG and data is present.
'   5. INDENTATION: Prepends leading spaces based on the current nesting
'      depth for clean, human-readable source code.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Foundational for the "Subgraphs & Clusters"
'     architecture, enabling recursive grouping of nodes.
'   - Strategy: Centralizes the "cluster" vs "subgraph" naming logic to
'     ensure consistent visual grouping.
' ==========================================================================
Private Function ProcessSubgraphOpen(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long, ByRef clusterCnt As Long) As String

    Dim subgraphName As String
    subgraphName = Trim$(GetStringBetweenDelimiters(data.item, vbNullString, OPEN_BRACE))
                        
    If subgraphName = vbNullString Then          ' No subgraph name supplied
        ' Increment the cluster counter, and use it in the cluster name
        clusterCnt = clusterCnt + 1
        subgraphName = "cluster_" & clusterCnt
    End If

    Dim subgraphDirective As String
    subgraphDirective = Space(indent * ini.source.indent) & "subgraph " & AddQuotesConditionally(subgraphName) & " {" & " " & Trim$(data.format)

    ' Inclduing the extra style attributes can be turned on/off in the settings
    If data.extraAttrs <> vbNullString Then
        If ini.graph.includeExtraAttributes Then
            subgraphDirective = subgraphDirective & " " & data.extraAttrs
        End If
    End If

    ' The subgraph can have an optional label. Include it if specified
    If data.label <> vbNullString Then
        If InStr(1, data.format, "{label}", vbTextCompare) Then
            subgraphDirective = replace(subgraphDirective, "{label}", data.label, 1, -1, vbTextCompare)
        Else
            subgraphDirective = subgraphDirective & " label=" & FormatLabel(data.label)
        End If
    End If
                            
    ' If output format is SVG, then include the tooltip data
    Dim Tooltip As String
    If ini.graph.includeTooltip Then
        If data.Tooltip <> vbNullString Then
            Tooltip = " tooltip=" & AddQuotes(ScrubText(data.Tooltip))
        End If
    End If
    
    ProcessSubgraphOpen = subgraphDirective & Tooltip & vbNewLine

End Function

' ==========================================================================
' SECTION: NODE ENTITY PROCESSING
' ==========================================================================

' ==========================================================================
' FUNCTION: ProcessNode
'
' PURPOSE:
'   THE NODE DISPATCHER. Orchestrates the translation of worksheet node
'   definitions into DOT syntax, supporting multi-node batching and
'   connectivity filtering.
'
' TECHNICAL WORKFLOW:
'   1. BATCH PROCESSING: Splits the 'Item' column by commas to handle
'      multiple Node IDs defined in a single Excel row.
'   2. ORPHAN SUPPRESSION:
'      - If 'includeOrphanNodes' is FALSE, it cross-references each ID
'        (ports removed) against the 'nodesUsedInRelationships' dictionary.
'      - Only "connected" nodes are passed to the next stage.
'   3. SYNTAX GENERATION: Invokes 'WriteNode' for every validated ID to
'      construct the specific DOT attribute string.
'   4. CONCATENATION: Merges individual node strings into a single buffer
'      using 'Join(Array(...))' for optimal performance.
'
' TECHNICAL NOTES:
'   - Strategy: Implements the "Single Row, Multiple Nodes" efficiency
'     pattern while enforcing graph-theory constraints like orphan removal.
'   - Layer: Logic Layer / Parser.
' ==========================================================================
Private Function ProcessNode(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long, ByVal nodesUsedInRelationships As Dictionary) As String
                        
    Dim item As String
    Dim Items() As String
    
    Dim graphvizSource As String
    
    Dim arrayIndex As Long
    
    item = data.item
    Items = split(item, COMMA)
    
    For arrayIndex = LBound(Items) To UBound(Items)
        data.item = Trim$(Items(arrayIndex))
                        
        ' Filter out nodes without node relationships
        If Not ini.graph.includeOrphanNodes Then
            If nodesUsedInRelationships.Exists(RemovePort(data.item)) Then
                graphvizSource = Join(Array(graphvizSource, WriteNode(ini, data, indent)), vbNullString)
            End If
        Else
            graphvizSource = Join(Array(graphvizSource, WriteNode(ini, data, indent)), vbNullString)
        End If
    Next

    ProcessNode = graphvizSource
End Function

' ==========================================================================
' SECTION: EDGE ENTITY PROCESSING & MATRIX EXPANSION
' ==========================================================================

' ==========================================================================
' PROCEDURE: ProcessEdge
'
' PURPOSE:
'   THE EDGE DISPATCHER. Translates worksheet relationship rows into DOT
'   syntax, supporting the expansion of many-to-many "matrix" relationships.
'
' TECHNICAL WORKFLOW:
'   1. MATRIX EXPANSION: Splits both 'item' (Tails) and 'relatedItem' (Heads)
'      by commas. It then performs a nested loop to generate a cross-product
'      of all possible connections from a single row.
'   2. ORPHAN INTEGRITY CHECK:
'      - If 'includeOrphanEdges' is FALSE: It verifies that both endpoints
'        (ports removed) exist in the 'definedNodes' registry.
'      - Connections to non-existent or unstyled nodes are suppressed.
'   3. SYNTAX GENERATION: Invokes 'WriteEdge' for every validated Tail-Head
'      pair to construct the specific DOT relationship string.
'   4. CONCATENATION: Aggregates all expanded edge strings into a single
'      buffer for return to the main assembly loop.
'
' TECHNICAL NOTES:
'   - Complexity: An $O(N \times M)$ expansion where $N$ is the number of
'     Tails and $M$ is the number of Heads in a single Excel cell.
'   - DeepWiki Context: Implements the "Relationship Expansion" logic
'     specified in the Defining Nodes & Edges architecture.
' ==========================================================================
Private Function ProcessEdge(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long, ByVal definedNodes As Dictionary) As String
                        
    Dim item As String
    Dim relatedItem As String
    Dim Items() As String
    Dim relatedItems() As String
    
    Dim graphvizSource As String
    
    Dim itemIndex As Long
    Dim relatedItemIndex As Long
    
    item = data.item
    Items = split(item, COMMA)
    
    relatedItem = data.relatedItem
    relatedItems = split(relatedItem, COMMA)
    
    For itemIndex = LBound(Items) To UBound(Items)
        For relatedItemIndex = LBound(relatedItems) To UBound(relatedItems)
            data.item = Trim$(Items(itemIndex))
            data.relatedItem = Trim$(relatedItems(relatedItemIndex))
            
            ' Filter out relationships without node definitions
            If Not ini.graph.includeOrphanEdges Then
                If definedNodes.Exists(RemovePort(data.item)) And definedNodes.Exists(RemovePort(data.relatedItem)) Then
                    graphvizSource = graphvizSource & WriteEdge(ini, data, indent)
                End If
            Else
                graphvizSource = graphvizSource & WriteEdge(ini, data, indent)
            End If
        Next
    Next

    ProcessEdge = graphvizSource
End Function

' ==========================================================================
' SECTION: SUBGRAPH & CLUSTER TERMINATION
' ==========================================================================

' ==========================================================================
' FUNCTION: ProcessSubgraphClose
'
' PURPOSE:
'   THE HIERARCHY TERMINATOR. Generates the closing brace for a Graphviz
'   subgraph or cluster, ensuring structural and visual alignment.
'
' TECHNICAL WORKFLOW:
'   1. INDENTATION: Prepends leading spaces based on the *restored* parent
'      nesting level to align the closing brace with its opening 'subgraph'
'      statement.
'   2. SYNTAX GENERATION: Appends the 'data.item' (typically the '}' character)
'      followed by a newline to cleanly terminate the block scope.
'
' TECHNICAL NOTES:
'   - DeepWiki Context: Works in tandem with 'ProcessSubgraphOpen' to manage
'     the "Stack-based Parsing" logic for nested groups.
'   - Strategy: Maintains human-readable DOT source code within the
'     Source Viewer by reflecting the logical nesting in the visual layout.
' ==========================================================================
Private Function ProcessSubgraphClose(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long) As String
    ProcessSubgraphClose = Space(indent * ini.source.indent) & data.item & vbNewLine
End Function

' ==========================================================================
' SECTION: ENTITY WRITERS (FINAL DOT ASSEMBLY)
' ==========================================================================

' ==========================================================================
' PROCEDURE: WriteNode
'
' PURPOSE:
'   THE NODE ASSEMBLER. Translates a specific node instance into its final
'   Graphviz DOT representation, merging styles, labels, and metadata.
'
' TECHNICAL WORKFLOW:
'   1. ID PURIFICATION: Strips port syntax for the base declaration to
'      ensure the node is correctly identified in the Graphviz symbol table.
'   2. HTML ADAPTATION: Detects HTML-like labels (<...>); if no other style
'      is provided, it automatically injects 'shape=plaintext' to prevent
'      Graphviz from wrapping the table in a default box.
'   3. ATTRIBUTE MERGING:
'      - Combines the Style Gallery 'format' with user-defined 'extraAttrs'.
'      - Appends SVG tooltips if the rendering format supports them.
'   4. LABEL INTEGRATION: Invokes 'FormatNodeLabels' to handle standard
'      labels and external xLabels.
'   5. OPTIMIZED EMISSION:
'      - If no attributes exist: Outputs a compact 'ID;' declaration.
'      - If attributes exist: Outputs a structured 'ID [ attributes ];' block
'        with appropriate indentation.
'
' TECHNICAL NOTES:
'   - Performance: Uses 'Join(Array(...))' to handle string concatenation
'     efficiently during high-volume node generation.
'   - DeepWiki Context: Implements the "Defining Nodes" logic where Excel
'     data meets DOT syntax requirements.
' ==========================================================================
Private Function WriteNode(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long) As String

    Dim styleAttributes As String
    
    Dim nodeId As String
    nodeId = data.item
    
    ' Strip off the port (if specified)
    If InStr(nodeId, ":") > 0 Then
        nodeId = GetStringTokenAtPosition(nodeId, ":", 1)
    End If

    ' If output format is SVG, then include the tooltip data
    Dim Tooltip As String
    If ini.graph.includeTooltip Then
        If data.Tooltip <> vbNullString Then
            Tooltip = " tooltip=" & AddQuotes(ScrubText(data.Tooltip))
        End If
    End If
    
    styleAttributes = Trim$(data.format)
    
    ' Include the extra style attributes if enabled in the settings
    If ini.graph.includeExtraAttributes Then
        styleAttributes = Trim$(styleAttributes & " " & data.extraAttrs)
    End If

    ' If no style has been specified, assume the user wants the shape to be what the
    ' HTML will render. For this situation Graphviz has to be told the shape is "plaintext"
    If (IsLabelHTMLLike(data.label)) And styleAttributes = vbNullString Then
        styleAttributes = "shape=plaintext "
    End If

    ' Collect the label, and xlabel labels into name value pairs
    styleAttributes = FormatNodeLabels(ini, data, styleAttributes)
    
    If Trim$(styleAttributes & Tooltip) = vbNullString Then
        WriteNode = Join(Array(Space(indent * ini.source.indent), AddQuotesConditionally(nodeId), SEMICOLON, vbNewLine), vbNullString)
    Else
        WriteNode = Join(Array(Space(indent * ini.source.indent), AddQuotesConditionally(nodeId), " [ ", Trim$(styleAttributes) & Tooltip & " ];", vbNewLine), vbNullString)
    End If

End Function

' ==========================================================================
' PROCEDURE: WriteEdge
'
' PURPOSE:
'   THE EDGE ASSEMBLER. Translates a relationship row into the final Graphviz
'   DOT connection string, managing directionality, ports, and multi-position labels.
'
' TECHNICAL WORKFLOW:
'   1. ATTRIBUTE SYNTHESIS:
'      - Merges the Style 'format' with 'extraAttrs' (if enabled).
'      - Appends SVG tooltips for interactive metadata support.
'   2. ID PREPARATION: Invokes 'FormatId' for both Tail and Head, conditionally
'      including or stripping port notation based on 'includeEdgePorts'.
'   3. LABEL AGGREGATION: Calls 'FormatEdgeLabels' to bundle standard,
'      external (xLabel), head, and tail labels into attribute pairs.
'   4. OPERATOR SELECTION: Injects 'ini.graph.edgeOperator' (-> or --) to
'      match the graph type (Digraph vs. Graph).
'   5. OPTIMIZED EMISSION:
'      - If no attributes: Outputs a simple 'A -> B;' declaration.
'      - If attributes exist: Outputs 'A -> B [ attributes ];' with
'        correct hierarchical indentation.
'
' TECHNICAL NOTES:
'   - Performance: Uses 'Join(Array(...))' to handle string concatenation
'     efficiently during high-volume edge generation.
'   - DeepWiki Context: Implements the "Defining Edges" logic where Excel
'     links are converted to DOT relationships.
' ==========================================================================
Private Function WriteEdge(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long) As String

    Dim styleAttributes As String
    styleAttributes = data.format

    ' Include the extra style attributes if enabled in the settings
    If ini.graph.includeExtraAttributes Then
        styleAttributes = Join(Array(styleAttributes, " ", data.extraAttrs), vbNullString)
    End If

    ' If output format is SVG, then include the tooltip data
    Dim Tooltip As String
    If ini.graph.includeTooltip Then
        If data.Tooltip <> vbNullString Then
            Tooltip = Join(Array(" tooltip=", AddQuotes(ScrubText(data.Tooltip))), vbNullString)
        End If
    End If
    
    ' Collect the label, xlabel, taillabel, and headlabel labels into name value pairs
    styleAttributes = Trim$(styleAttributes)
    styleAttributes = FormatEdgeLabels(ini, data, styleAttributes)

    ' Add the quotes to the id and (optional) port for the item, and the "is related to" item
    Dim tailId As String
    tailId = FormatId(data.item, ini.graph.includeEdgePorts)
    
    Dim headId As String
    headId = FormatId(data.relatedItem, ini.graph.includeEdgePorts)
    
    ' Write out the edge command
    If Trim$(styleAttributes & Tooltip) = vbNullString Then
        WriteEdge = Join(Array(Space(indent * ini.source.indent), tailId, " ", ini.graph.edgeOperator, " ", headId, SEMICOLON, vbNewLine), vbNullString)
    Else
        WriteEdge = Join(Array(Space(indent * ini.source.indent), tailId, " ", ini.graph.edgeOperator, " ", headId, "[ ", Trim$(styleAttributes) & Tooltip & " ];", vbNewLine), vbNullString)
    End If
    
End Function

' ==========================================================================
' SECTION: NATIVE PASSTHROUGH & GLOBAL OVERRIDES
' ==========================================================================

' ==========================================================================
' FUNCTION: ProcessNative
'
' PURPOSE:
'   THE NATIVE PASSTHROUGH. Allows power users to inject raw, unparsed DOT
'   code directly into the generation stream, bypassing the project's
'   standard data-mapping logic.
'
' TECHNICAL WORKFLOW:
'   1. TRIGGER: Executed when a row is classified as 'TYPE_NATIVE'
'      (typically identified by the '>' character in the Item column).
'   2. INJECTION: Retrieves the 'label' field—which contains the raw DOT
'      syntax—and prepends the current level of indentation.
'   3. TERMINATION: Appends a newline to ensure the next DOT statement
'      starts on a fresh line in the Source Viewer.
'
' TECHNICAL NOTES:
'   - Strategy: Provides an "Escape Hatch" for advanced Graphviz features
'     not natively supported by the Excel UI (e.g., custom rank blocks
'     or complex multi-line attribute strings).
'   - Layer: Logic Layer / Native Passthrough.
' ==========================================================================
Private Function ProcessNative(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long) As String
    ProcessNative = Space(indent * ini.source.indent) & data.label & vbNewLine
End Function

' ==========================================================================
' FUNCTION: ProcessKeyword
'
' PURPOSE:
'   Generates a fully-formed DOT keyword block (node, edge, or graph) by
'   merging style-layer attributes with data-layer values. Acts as the
'   global-scope attribute injector, establishing default properties for
'   all subsequent declarations in the DOT stream.
'
' TECHNICAL WORKFLOW:
'   1. BASE ATTRIBUTE ASSEMBLY:
'        - Starts with the style 'format' string.
'        - Optionally appends 'extraAttrs' when enabled, producing the
'          complete style-layer attribute template.
'
'   2. CONTEXT-SENSITIVE LABEL SYNTHESIS:
'        - KEYWORD_NODE:
'             • Passes the assembled template to FormatNodeLabels, which
'               expands placeholders and emits node-label attributes.
'        - KEYWORD_EDGE:
'             • Passes the template to FormatEdgeLabels for multi-positional
'               edge-label synthesis.
'        - KEYWORD_GRAPH:
'             • If a graph-level label exists, delegates to FormatGraphLabels
'               for placeholder expansion or fallback label emission.
'
'   3. DOT SYNTAX GENERATION:
'        - Emits a standard DOT keyword block of the form:
'              <keyword> [ <attributes> ];
'          with indentation controlled by the caller.
'
' TECHNICAL NOTES:
'   - Implements Graphviz's cascading "state machine" behavior: once a keyword
'     block is emitted, its attributes become defaults for all following nodes,
'     edges, or subgraphs within the same scope. Graphviz does not support
'     clearing defaults via an empty keyword block (e.g., node []; has no
'     reset effect). To disable previously established defaults, you must
'     either:
'         • explicitly override each attribute with new values, or
'         • open a new subgraph { ... } to create a fresh attribute scope.
'     These are the only mechanisms Graphviz provides for neutralizing
'     inherited keyword defaults.
' ==========================================================================

Private Function ProcessKeyword(ByRef ini As settings, ByRef data As dataRow, ByVal indent As Long) As String

    Dim styleAttributes As String
    styleAttributes = Trim$(data.format)

    If ini.graph.includeExtraAttributes Then
        styleAttributes = Trim$(Join(Array(styleAttributes, " ", data.extraAttrs), vbNullString))
    End If

    If UCase$(data.item) = KEYWORD_NODE Then
        styleAttributes = FormatNodeLabels(ini, data, styleAttributes)
    
    ElseIf UCase$(data.item) = KEYWORD_EDGE Then
        styleAttributes = FormatEdgeLabels(ini, data, styleAttributes)
    
    ElseIf UCase$(data.item) = KEYWORD_GRAPH Then
        If data.label <> vbNullString Then
            styleAttributes = FormatGraphLabels(ini, data, styleAttributes)
        End If
    End If
        
    ProcessKeyword = Join(Array(Space(indent * ini.source.indent), data.item, "[ ", Trim$(styleAttributes), " ];", vbNewLine), vbNullString)
    
End Function

' ==========================================================================
' SECTION: LABEL SANITIZATION & SYNTAX SAFETY
' ==========================================================================

' ==========================================================================
' FUNCTION: FormatLabel
'
' PURPOSE:
'   THE LABEL GATEKEEPER. Standardizes label values for DOT output by
'   distinguishing between raw text and Graphviz HTML-like markup.
'
' TECHNICAL WORKFLOW:
'   1. HTML DETECTION: Uses 'IsLabelHTMLLike' to identify if the string is
'      wrapped in angle brackets (<...>); if TRUE, the string is returned
'      untouched to allow Graphviz to parse the internal XML/HTML tags.
'   2. TEXT SANITIZATION: If not HTML, the value is passed through:
'      - 'ScrubText': Handles escape characters and reserved DOT sequences.
'      - 'AddQuotes': Wraps the sanitized string in double quotes for
'        standard attribute assignment.
'
' TECHNICAL NOTES:
'   - Strategy: Prevents Graphviz syntax crashes by ensuring reserved
'     characters in labels are either escaped or correctly identified as
'     HTML code.
' ==========================================================================
Private Function FormatLabel(ByVal labelValue As String) As String

    If IsLabelHTMLLike(labelValue) Then          ' just return it intact
        FormatLabel = labelValue
    Else
        FormatLabel = AddQuotes(ScrubText(labelValue))
    End If

End Function

' ==========================================================================
' SECTION: HTML-LIKE LABEL DETECTION
' ==========================================================================

' ==========================================================================
' FUNCTION: IsLabelHTMLLike
'
' PURPOSE:
'   THE SYNTAX CLASSIFIER. Detects if a label string contains Graphviz
'   HTML-like markup (XML based) to determine if standard DOT quoting
'   should be bypassed.
'
' TECHNICAL WORKFLOW:
'   1. PRE-PROCESSING: Normalizes the input by stripping Line Feed characters
'      (Chr 10) to facilitate reliable boundary checking.
'   2. BOUNDARY VALIDATION: Checks if the string starts with '<' and ends
'      with '>', which is the Graphviz requirement for HTML-like labels.
'   3. HEURISTIC INSPECTION: Scans the internal content for terminal XML
'      markers ("</" or "/>"). This validates intent and distinguishes
'      actual markup from simple inequality comparisons.
'   4. LOGICAL RETURN: Returns TRUE if the string satisfies the structural
'      requirements, signaling the parser to emit the string unquoted.
'
' TECHNICAL NOTES:
'   - Performance: Uses a "process of elimination" structure to minimize
'     string evaluations.
'   - Strategy: Prioritizes speed over exhaustive XML validation, deferring
'     syntax correction to the external Graphviz engine.
' ==========================================================================
Public Function IsLabelHTMLLike(ByVal label As String) As Boolean
     
     IsLabelHTMLLike = False
    
    ' Remove newline characters to create a single line
    Dim singleLineLabel As String
    singleLineLabel = replace(label, Chr$(10), vbNullString)

    ' HTML-like labels have to be wrapped in '<' and '>' characters
    ' Use process of elimination instead of 'and' conditions to improve performance
    If StartsWith(singleLineLabel, LESS_THAN) Then
        If EndsWith(singleLineLabel, GREATER_THAN) Then   ' Label is wrapped in '<' and '>'
        
            ' Interrogate the string between the HTML-like indicators to see if
            ' a portion of an HTML termination element is present. This test is not a
            ' fool-proof determination that the label text contains valid HTML elements,
            ' but it is a fast assessment. If the HTML is not valid it will show up in
            ' the diagram, and the user can correct their label data.
            
            ' Pluck the label out from between the '<' and '>' characters
            singleLineLabel = Trim$(GetStringBetweenDelimiters(singleLineLabel, LESS_THAN, GREATER_THAN))
            If (InStr(singleLineLabel, "</") > 0) Or (InStr(singleLineLabel, "/>") > 0) Then ' At least one HTML close element is present.
                IsLabelHTMLLike = True   ' label likely contains HTML-like content
            End If
        End If
    End If
    
End Function

' ==========================================================================
' SECTION: DATA SOURCE RESOLUTION & VALIDATION
' ==========================================================================

' ==========================================================================
' FUNCTION: GetDataWorksheetName
'
' PURPOSE:
'   THE CONTEXT RESOLVER. Dynamically identifies the correct worksheet to use
'   as the data source, enabling the rendering engine to work on custom sheets
'   while protecting system-critical worksheets.
'
' TECHNICAL WORKFLOW:
'   1. SYSTEM BLACKLIST: Checks the 'ActiveSheet' name against a hard-coded
'      list of protected system worksheets (Settings, Styles, Help, etc.).
'   2. SCHEMA VALIDATION: If the active sheet is not on the blacklist, it
'      retrieves the 'dataWorksheet' UDT and verifies the worksheet's
'      integrity by comparing header values (Item, Label, Related Item)
'      against the 'DataSheet' master template.
'   3. SAFE FALLBACK: If the active sheet is a protected system sheet or
'      fails the schema validation, the function defaults to the standard
'      'DataSheet.name'.
'   4. IDENTITY RETURN: Returns the validated 'worksheetName' to the caller
'      to anchor the rest of the parsing pipeline.
'
' TECHNICAL NOTES:
'   - Strategy: Empowers "Multi-Sheet" projects by allowing users to create
'     alternate data views that still adhere to the global Data Model.
'   - Layer: Logic Layer / Context Management.
' ==========================================================================
Public Function GetDataWorksheetName() As String

    Dim worksheetName As String
    worksheetName = ActiveSheet.name
    
    ' Worksheets which are not allowed to hold graph data
    If worksheetName = DataSheet.name _
       Or worksheetName = GraphSheet.name _
       Or worksheetName = StylesSheet.name _
       Or worksheetName = StyleDesignerSheet.name _
       Or worksheetName = SettingsSheet.name _
       Or worksheetName = HelpShapesSheet.name _
       Or worksheetName = HelpColorsSheet.name _
       Or worksheetName = HelpAttributesSheet.name _
       Or worksheetName = AboutSheet.name _
       Or worksheetName = SourceSheet.name _
       Or worksheetName = SqlSheet.name _
       Or worksheetName = ChoicesSheet.name _
       Or worksheetName = DiagnosticsSheet.name _
       Or worksheetName = ListsSheet.name _
    Then
        worksheetName = DataSheet.name
    Else
        ' Ensure the worksheet has the same layout of the 'data' worksheet by comparing a few of the key headings
        Dim data As dataWorksheet
        data = GetSettingsForDataWorksheet(worksheetName)

        If GetCell(worksheetName, data.headingRow, data.itemColumn) <> DataSheet.Cells.item(data.headingRow, data.itemColumn).value Then
            worksheetName = DataSheet.name
        ElseIf GetCell(worksheetName, data.headingRow, data.labelColumn) <> DataSheet.Cells.item(data.headingRow, data.labelColumn).value Then
            worksheetName = DataSheet.name
        ElseIf GetCell(worksheetName, data.headingRow, data.isRelatedToItemColumn) <> DataSheet.Cells.item(data.headingRow, data.isRelatedToItemColumn).value Then
            worksheetName = DataSheet.name
        End If
    End If
    
    GetDataWorksheetName = worksheetName
End Function
