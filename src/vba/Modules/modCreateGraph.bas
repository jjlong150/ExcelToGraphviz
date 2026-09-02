Attribute VB_Name = "modCreateGraph"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modCreateGraph
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Logic / Transformation Pipeline
'
' ROLE:
'   The central Graphviz orchestration engine. Transforms structured worksheet
'   data into DOT source, executes the external Graphviz binary, and injects
'   the resulting diagram back into Excel. Coordinates the full end-to-end
'   rendering lifecycle for both interactive (AutoDraw) and batch-export
'   workflows, serving as the bridge between the VBA logic layer and the
'   external Graphviz engine.
'
' RESPONSIBILITIES:
'   o Rendering Pipeline Management:
'       - Parse and validate worksheet data.
'       - Resolve styles, views, and graph options.
'       - Synthesize DOT source (ConvertDataWorksheetToGvSource).
'       - Create and manage temporary files.
'       - Execute Graphviz via Graphviz.cls and capture console output.
'       - Insert, scale, and name rendered images in Excel.
'
'   o Interactive Rendering (AutoDraw):
'       - Provide safe, atomic redraws during worksheet editing.
'       - Suspend UI events and screen updates to prevent flicker or recursion.
'
'   o Batch Export:
'       - Iterate across multiple Views to generate multiple diagrams.
'       - Support filename token substitution (%D, %T, %V, %W, %E, %S).
'       - Handle macOS sandbox routing and conditional source-file deletion.
'       - Apply optional SVG post-processing (animations, replacements).
'
'   o DOT-Only Workflows:
'       - Expose CreateGraphSource for debugging, validation, and Source Viewer.
'
' INTERACTIONS:
'   o Graphviz.cls:
'       - External engine wrapper (RenderGraph, SourceToFile).
'
'   o modCreateCommon:
'       - Shared parsing, normalization, style caching, label overrides,
'         placeholder expansion, HTML-label detection, and filesystem helpers.
'
'   o modDataTypes:
'       - settings, dataWorksheet, style UDTs.
'
'   o Utility Modules:
'       - modUtilityString: label scrubbing, token substitution.
'       - modUtilityFileSystem: temp directories, file existence, deletion.
'       - modUtilityStatusBar: progress and timing feedback.
'
'   o Ribbon Tabs:
'       - Graphviz, Source, SVG, Styles, Launchpad.
'
' CROSS-PLATFORM NOTES:
'   o Windows:
'       - Stopwatch timing, native file access, direct DOT execution.
'
'   o macOS:
'       - AppleScript-mediated file dialogs.
'       - Sandbox-safe temp routing and conditional filename overrides.
'
' ERROR HANDLING:
'   o Defensive validation of worksheet existence, view selection, and
'     output-directory prerequisites.
'   o DOT generation failures surface via LogError and abort cleanly.
'   o Graphviz execution errors routed to the Console worksheet.
'
' RELATED WIKI PAGES:
'   o Rendering Pipeline Overview
'   o Working with the Data Worksheet
'   o Batch Export & View Iteration
'   o DOT Source Generation & Validation
'   o Image Path Resolution
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
'   - Triggered indirectly via Worksheet_Change -> AutoDrawDebounced -> AutoDraw.
'   - DeepWiki Context: Represents the "safe execution wrapper" for the
'     AutoDraw reactivity model described in the Data Worksheet documentation.
' ==========================================================================
Public Sub AutoDraw()
    Application.screenUpdating = False
    Application.enableEvents = False
    CreateGraphWorksheet
    Application.enableEvents = True
    Application.screenUpdating = True
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
'      Excel rows into DOT language. If validation fails, it aborts.
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
Public Sub CreateGraphWorksheet()
Attribute CreateGraphWorksheet.VB_ProcData.VB_Invoke_Func = " \n14"

    On Error GoTo ErrorHandler
    
#If Mac Then
    ' For some reason, my Mac fails when I code it as "#If Not Mac Then"
#Else
    ' Stopwatch is only available on Windows OS
    Dim timex As Stopwatch
    Set timex = New Stopwatch
    timex.start
#End If
    
    ' Objects needed for control flow or cleanup
    Dim ini As settings
    Dim graphvizObj As Graphviz
    Dim shapeObject As shape

    ' Clear the status bar
    ClearStatusBar

    ' Read in the runtime settings
    ini = GetSettings(GetDataWorksheetName())

    If Not WorksheetExists(ini.data.worksheetName) Then
        EmitMessage GetMessage("msgboxNoDataToGraph")
        GoTo Cleanup
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
    Set graphvizObj = New Graphviz
    
    ' Build the file names
    graphvizObj.outputDirectory = GetTempDirectory()
    graphvizObj.filenameBase = "RelationshipVisualizer"
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
        GoTo Cleanup
    End If
    
    ' Display source if debugging
    ShowSource graphvizSource

    ' Write the graphviz source to a file
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
        
    Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, _
                                    ActiveSheet.Range(targetCell), _
                                    False, _
                                    True, _
                                    GetMessage("InsertPictureAltText"))
    
    ' Scale the graph to the zoom percentage specified
    Dim scaleFactor As Double
    scaleFactor = ini.graph.scaleImage / 100
    ActiveSheet.Pictures(ActiveSheet.Pictures.Count).ShapeRange.ScaleHeight scaleFactor, msoFalse, msoScaleFromTopLeft
    
    If ini.graph.pictureName <> vbNullString Then
        ActiveSheet.Pictures(ActiveSheet.Pictures.Count).name = ini.graph.pictureName
    End If
    
Cleanup:
    On Error Resume Next
    If Not graphvizObj Is Nothing Then
        DeleteFile graphvizObj.GraphvizFilename
        DeleteFile graphvizObj.DiagramFilename
    End If
    Set graphvizObj = Nothing
    Set shapeObject = Nothing
    On Error GoTo 0
    
#If Mac Then
    ' For some reason, my Mac fails when I code it as "#If Not Mac Then"
#Else
    If Not timex Is Nothing Then
        timex.stop_it
        Application.StatusBar = timex.Elapsed_sec & " seconds"
    End If
#End If

    Exit Sub
    
ErrorHandler:
    On Error GoTo 0
    
    Dim msg As String
    msg = "An error occurred while creating the graph:" & vbCrLf & vbCrLf & _
          Err.Description & vbCrLf & vbCrLf & _
          "(Error #" & Err.number & ")"
    
    EmitMessage msg
    
    Resume Cleanup   ' Ensure temp files are cleaned up
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

    ' Get file output settings
    Dim output As FileOutput
    output = GetSettingsForFileOutput()

    ' Determine output directory, and build file names
    If output.directory = vbNullString Then
        output.directory = ActiveWorkbook.path
    End If

    ' Validate filename info
    If Not FileLocationProvided(output) Then
        Exit Sub
    End If

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
        graphvizObj.outputDirectory = output.directory
        graphvizObj.filenameBase = GetFilenameBase(ini, viewColumn)
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
        graphvizObj.Renderer = ini.graph.Renderer
        
        graphvizObj.RenderGraph
        
        ' Display any console output first
        DisplayTextOnConsoleWorksheet graphvizObj.GraphvizCommand, graphvizObj.GraphvizMessages
                
        ' If the diagram file is not there, then Graphviz failed
        If FileExists(graphvizObj.DiagramFilename) Then
            ' Post-process SVG files to add things like animations
            If ini.graph.imageTypeFile = FILETYPE_SVG And ini.graph.postProcessSVG Then
                FindAndReplaceSVG graphvizObj.DiagramFilename, graphvizObj.DiagramFilename
            End If
            
            ' Display the published graph?
            If SettingsSheet.Range("openAfterPublish").value = TOGGLE_YES Then
                SafeFollowHyperlink graphvizObj.DiagramFilename
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

Public Sub SafeFollowHyperlink(ByVal target As String)
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink target

    If Err.number <> 0 Then
        ' User likely clicked Cancel on the security prompt.
        Err.Clear
    End If

    On Error GoTo 0
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
Public Function ConvertDataWorksheetToGvSource(ByRef ini As settings, _
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
    
    Dim items() As String
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

                            items = split(data.item, COMMA)
                            relatedItems = split(data.relatedItem, COMMA)
                            
                            For itemIndex = LBound(items) To UBound(items)
                                For relatedItemIndex = LBound(relatedItems) To UBound(relatedItems)
                                    ' If both the tail and the head in the relationship refer
                                    ' to included nodes having style definitions, track the nodes
                                    ' as "Is Used" so that we later determine island nodes to exclude
                                    ' from the graph.
                                
                                    itemId = RemovePort(items(itemIndex))
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
    'Dim openSubgraphs As Long
    Dim errCnt As Long

    ' Initializations
    'openSubgraphs = 0
    errCnt = 0
    
    Dim clusters As Stack
    Set clusters = New Stack
    
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
                        If UCase$(data.item) = KEYWORD_EDGE Then
                            ' No error
                        ElseIf data.item = vbNullString Then
                            LogError ini, row, GetMessage("errormsgEdgeNoItemFound"), errCnt
                        
                        ElseIf data.relatedItem = vbNullString Then
                            LogError ini, row, GetMessage("errormsgEdgeNoRelatedItemFound"), errCnt
                        End If
                        
                    ElseIf data.styleType = TYPE_SUBGRAPH_OPEN Then
                        clusters.Push "{"
                                                
                    ElseIf data.styleType = TYPE_SUBGRAPH_CLOSE Then
                        If clusters.IsEmpty Then
                            LogError ini, row, GetMessage("errormsgBracesExcessClose"), errCnt
                        Else
                            clusters.Pop
                        End If
                    End If
                End If
            End If
        End If
    Next row

    ' Alert the user if it appears that the subgraphs open and close braces are out of balance
    'If openSubgraphs > 0 Then
    If Not clusters.IsEmpty Then
        LogError ini, row, replace(GetMessage("errormsgBracesExcessOpen"), "{openSubgraphs}", clusters.Count), errCnt
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
                        data.Format = styles.item(data.styleName).styleFormat
                    Else
                        data.Format = vbNullString
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
' SECTION: OBJECT FACTORIES & ERROR REPORTING
' ==========================================================================

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
'      directly into TODO.
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
    
    ' Localize the full error message
    Dim fullMessage As String
    fullMessage = GetMessage("errormsgRow")
    fullMessage = replace(fullMessage, "{worksheet}", ini.data.worksheetName, 1, 1, vbTextCompare)
    fullMessage = replace(fullMessage, "{row}", CStr(row), 1, 1, vbTextCompare)
    fullMessage = replace(fullMessage, "{errorMessage}", errorMessage, 1, 1, vbTextCompare)
    
    EmitMessageSilent fullMessage, esError
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
' ROUTINE: FormatDebugLabel
'
' PURPOSE:
'   Generates and applies a diagnostic debug label for the current data row.
'   Builds a type-appropriate debug string, then integrates it into the
'   resolved label using ApplyDebugLabel, which handles plain-text, blank,
'   and HTML-like labels safely and consistently.
'
' FUNCTIONAL WORKFLOW:
'   1. DEBUG STRING CONSTRUCTION:
'        - Invokes BuildDebugLabel(row, data) to produce a concise,
'          type-specific diagnostic string:
'             o Edge rows: "row: <row>, <tailId>-><headId>"
'             o Node rows: "row: <row>, "<nodeId>""
'             o Subgraph/keyword rows: "row: <row>"
'             o Others: empty string
'
'   2. LABEL AUGMENTATION:
'        - Passes the original label and the debug string to ApplyDebugLabel,
'          which:
'             o Leaves HTML-like labels intact except for inserting
'               "<BR/>debugStr>" before the closing ">".
'             o Replaces blank labels with the debug string.
'             o Appends debugStr on a new line for plain-text labels.
'
'   3. OUTPUT:
'        - Returns the fully augmented label, ready for normalization and
'          Graphviz emission.
'
' TECHNICAL NOTES:
'   - HTML-label augmentation preserves structural validity by removing the
'     trailing ">", inserting "<BR/>", appending debugStr, and restoring ">".
'   - NEWLINE is used only for plain-text labels.
'   - This routine does not perform NormalizeLabelText; callers may normalize
'     afterward if required.
'   - DeepWiki Context: Implements the debug-label formatting rules described
'     in the "Debugging", "Labels", and "Serialization" sections.
' ==========================================================================
Private Function FormatDebugLabel(ByRef ini As settings, _
                                 ByRef data As dataRow, _
                                 ByVal label As String) As String

    Dim debugStr As String
    debugStr = BuildDebugLabel(ini, data)

    FormatDebugLabel = ApplyDebugLabel(label, debugStr)
End Function

' ==========================================================================
' ROUTINE: BuildDebugLabel
'
' PURPOSE:
'   Constructs a diagnostic debug label for the current data row. Produces a
'   concise, type-specific string that identifies the row number and the
'   relevant node or edge identifiers. Used by ApplyDebugLabel to append
'   debugging metadata to Graphviz labels without altering core semantics.
'
' FUNCTIONAL WORKFLOW:
'   1. EDGE ROWS (TYPE_EDGE):
'        - Emits: "row: <row>, <tailId>-><headId>"
'        - tailId  = FormatId(data.item, True)
'        - headId  = FormatId(data.relatedItem, True)
'        - Includes port information when present.
'
'   2. NODE ROWS (TYPE_NODE):
'        - Emits: "row: <row>, "<nodeId>""
'        - nodeId is quoted via AddQuotes to ensure Graphviz compatibility.
'
'   3. SUBGRAPH OPEN / KEYWORD ROWS (TYPE_SUBGRAPH_OPEN, TYPE_KEYWORD):
'        - Emits: "row: <row>"
'        - No identifier is included.
'
'   4. OTHER TYPES:
'        - Emits an empty string.
'        - Ensures safe fallback behavior for unrecognized styleType values.
'
' TECHNICAL NOTES:
'   - This routine does not perform HTML or newline formatting; callers must
'     use ApplyDebugLabel to integrate the debug string into the final label.
'   - FormatId ensures consistent quoting and port handling across all debug
'     output.
'   - DeepWiki Context: Implements the debug-label construction rules
'     described in the "Debugging", "Labels", and "Serialization" sections.
' ==========================================================================
Private Function BuildDebugLabel(ByRef ini As settings, ByRef data As dataRow) As String
    Select Case data.styleType
        Case TYPE_EDGE
            BuildDebugLabel = "row: " & data.row & ", " & _
                              FormatId(data.item, ini.graph.includeEdgePorts) & "-&gt;" & _
                              FormatId(data.relatedItem, ini.graph.includeEdgePorts)
        Case TYPE_NODE
            BuildDebugLabel = "row: " & data.row & ", " & AddQuotes(data.item)

        Case TYPE_SUBGRAPH_OPEN, TYPE_KEYWORD
            BuildDebugLabel = "row: " & data.row

        Case Else
            BuildDebugLabel = vbNullString
    End Select
End Function

' ==========================================================================
' ROUTINE: ApplyDebugLabel
'
' PURPOSE:
'   Appends a debug string to a resolved label using rules that preserve
'   Graphviz compatibility and protect HTML-like labels. Handles plain-text,
'   blank, and HTML-formatted labels differently to ensure readable and
'   structurally valid output.
'
' FUNCTIONAL WORKFLOW:
'   1. EMPTY DEBUG STRING:
'        - When debugStr = "", returns the original label unchanged.
'
'   2. HTML-LIKE LABEL HANDLING:
'        - When IsLabelHTMLLike(label) = True:
'             o Removes a trailing ">" if present.
'             o Appends "<BR/>" followed by debugStr.
'             o Re-adds the closing ">".
'        - Produces: <...> ? <...<BR/>debugStr>
'        - Ensures the label remains valid HTML for Graphviz.
'
'   3. BLANK LABEL HANDLING:
'        - When label = "":
'             o Returns debugStr as the entire label.
'             o Allows debug output to stand alone when no label text exists.
'
'   4. PLAIN-TEXT LABEL HANDLING:
'        - For non-blank, non-HTML labels:
'             o Returns label & NEWLINE & debugStr.
'             o Produces a readable, multi-line diagnostic label.
'
' TECHNICAL NOTES:
'   - NEWLINE is injected only for plain-text labels.
'   - HTML-label augmentation uses "<BR/>" to ensure valid HTML line breaks.
'   - This routine does not perform normalization; callers may apply
'     NormalizeLabelText afterward if Graphviz-safe output is required.
'   - DeepWiki Context: Implements the debug-label rules described in the
'     "Labels", "Debugging", and "Serialization" sections.
' ==========================================================================
Private Function ApplyDebugLabel(ByVal label As String, _
                                 ByVal debugStr As String) As String

    ' No debug string -> return original label
    If Len(debugStr) = 0 Then
        ApplyDebugLabel = label
        Exit Function
    End If

    ' HTML-like labels get special handling:
    '   <...>  ?  <...<BR/>debugStr>
    If IsLabelHTMLLike(label) Then
        Dim core As String

        ' Remove the trailing ">" if present
        If Right$(label, 1) = ">" Then
            core = Left$(label, Len(label) - 1)
        Else
            core = label
        End If

        ' Append HTML break + debug string + closing ">"
        ApplyDebugLabel = core & "<BR/>" & debugStr & ">"
        Exit Function
    End If

    ' Blank label -> debug string becomes the entire label
    If label = vbNullString Then
        ApplyDebugLabel = debugStr
        Exit Function
    End If

    ' Plain-text label -> append debug string on a new line
    ApplyDebugLabel = label & NEWLINE & debugStr
End Function

' ==========================================================================
' SECTION: DEBUGGING & EXTERNAL LABEL FORMATTING
' ==========================================================================

' ==========================================================================
' SECTION: EDGE LABEL ASSEMBLY
' ==========================================================================

' ==========================================================================
' ROUTINE: FormatEdgeLabels
'
' PURPOSE:
'   Applies edge-specific label formatting rules to the attribute dictionary.
'   Processes all supported edge label types (label, xlabel, taillabel,
'   headlabel, tooltip) according to inclusion switches, inheritance rules,
'   placeholder substitution, and blank-value handling. Produces a fully
'   resolved attribute set ready for Graphviz emission.
'
' FUNCTIONAL WORKFLOW:
'   1. PRIMARY EDGE LABEL ("label"):
'        - Uses data.label as the source value.
'        - Controlled by ini.graph.includeEdgeLabels.
'        - Blank handling uses ini.graph.blankEdgeLabels.
'        - Blank token: "\E".
'        - Delegates full processing to ProcessEdgeLabelAttribute.
'
'   2. SECONDARY LABEL ("xlabel"):
'        - Uses data.xlabel.
'        - Controlled by ini.graph.includeEdgeXLabels.
'        - Blank values omitted (allowBlank = False).
'        - Delegates processing to ProcessEdgeLabelAttribute.
'
'   3. TAIL LABEL ("taillabel"):
'        - Uses data.taillabel.
'        - Controlled by ini.graph.includeEdgeTailLabels.
'        - Blank values omitted.
'        - Delegates processing to ProcessEdgeLabelAttribute.
'
'   4. HEAD LABEL ("headlabel"):
'        - Uses data.headlabel.
'        - Controlled by ini.graph.includeEdgeHeadLabels.
'        - Blank values omitted.
'        - Delegates processing to ProcessEdgeLabelAttribute.
'
'   5. TOOLTIP LABEL ("tooltip"):
'        - Uses data.tooltip.
'        - Controlled by ini.graph.includeEdgeTooltips.
'        - Blank values allowed (allowBlank = True).
'        - Delegates processing to ProcessEdgeLabelAttribute.
'
'   6. ATTRIBUTE DICTIONARY UPDATE:
'        - Each call to ProcessEdgeLabelAttribute may:
'             o Resolve inheritance and placeholder substitution
'             o Apply debug formatting when enabled
'             o Normalize text for Graphviz compatibility
'             o Add, update, or remove attributes based on inclusion rules
'
' TECHNICAL NOTES:
'   - This routine performs no dictionary reconstruction; callers must invoke
'     RebuildStyleAttributeString afterward.
'   - Ensures consistent edge-label behavior across all edge-emission paths.
'   - DeepWiki Context: Implements the edge-label rules described in the
'     "Edges", "Labels", "Overrides", and "Serialization" sections.
' ==========================================================================
Private Sub FormatEdgeLabels(ByRef ini As settings, ByRef data As dataRow, ByRef d As Dictionary)
    
    ' label
    ProcessEdgeLabelAttribute ini, _
                        data, _
                        d, _
                        "label", _
                        data.label, _
                        ini.graph.includeEdgeLabels, _
                        ini.graph.blankEdgeLabels, _
                        "\E"
    ' xlabel
    ProcessEdgeLabelAttribute ini, _
                        data, _
                        d, _
                        "xlabel", _
                        data.xlabel, _
                        ini.graph.includeEdgeXLabels, _
                        False, _
                        vbNullString
    ' taillabel
    ProcessEdgeLabelAttribute ini, _
                        data, _
                        d, _
                        "taillabel", _
                        data.taillabel, _
                        ini.graph.includeEdgeTailLabels, _
                        False, _
                        vbNullString
    ' headlabel
    ProcessEdgeLabelAttribute ini, _
                        data, _
                        d, _
                        "headlabel", _
                        data.headlabel, _
                        ini.graph.includeEdgeHeadLabels, _
                        False, _
                        vbNullString
    ' tooltip
    ProcessEdgeLabelAttribute ini, _
                        data, _
                        d, _
                        "tooltip", _
                        data.tooltip, _
                        ini.graph.includeEdgeTooltips, _
                        Not ini.graph.blankEdgeTooltips, _
                        vbNullString
End Sub

' ==========================================================================
' ROUTINE: ProcessEdgeLabelAttribute
'
' PURPOSE:
'   Processes a single edge-level attribute and updates the attribute
'   dictionary accordingly. Applies inclusion rules, resolves inheritance and
'   placeholder overrides, appends optional debugging information, normalizes
'   the final label text, and emits the correct Graphviz-compatible value.
'
' FUNCTIONAL WORKFLOW:
'   1. INCLUSION GATING:
'        - When includeAttr = False:
'             o Removes attrName from the dictionary if present.
'             o Skips all inheritance, placeholder, and normalization logic.
'
'   2. INHERITANCE & PLACEHOLDER RESOLUTION:
'        - Calls ApplyLabelOverrides to:
'             o Resolve inherited values from graph/node/edge scope.
'             o Substitute template placeholders.
'             o Produce the working labelValue.
'
'   3. DEBUG AUGMENTATION (OPTIONAL):
'        - When ini.graph.debug = True:
'             o For "label": applies FormatDebugLabel.
'        - Embeds row-level diagnostics directly into the emitted label.
'
'   4. NORMALIZATION:
'        - Scrubs the final labelValue using NormalizeLabelText to ensure
'          Graphviz-safe output (escaping, trimming, and
'          control-character removal).
'
'   5. DICTIONARY EMISSION:
'        - If attrName already exists:
'             o Updates the value according to blank-handling rules:
'                   - If labelValue = "" and allowBlank = True:
'                         d(attrName) = blankToken
'                   - If labelValue = "" and allowBlank = False:
'                         d(attrName) = ""
'                   - Otherwise:
'                         d(attrName) = labelValue
'
'        - If attrName does not exist:
'             o Adds attrName only when:
'                   - labelValue is non-blank, OR
'                   - labelValue is blank AND allowBlank = True.
'             o Otherwise omits the attribute entirely, allowing Graphviz to
'               apply its default behavior.
'
' TECHNICAL NOTES:
'   - This routine is edge-specific; node and graph attributes follow their
'     own handlers.
'   - Blank-token emission (e.g., "/E") is used to force explicit Graphviz
'     logic to display an Edge ID.
'   - Debug formatting is applied *after* inheritance and placeholder
'     resolution to ensure diagnostics reflect the final resolved value.
'   - DeepWiki Context: Implements the edge-attribute rules described in the
'     "Edges", "Labels", "Overrides", and "Serialization" documentation.
' ==========================================================================
Private Sub ProcessEdgeLabelAttribute( _
        ByRef ini As settings, _
        ByRef data As dataRow, _
        ByRef d As Dictionary, _
        ByVal attrName As String, _
        ByRef attrValue As String, _
        ByVal includeAttr As Boolean, _
        ByVal allowBlank As Boolean, _
        ByVal blankToken As String)

    Dim labelValue As String
    labelValue = attrValue
    
    ' If Attribute should not be in the dictionary, remove it. No need
    ' to resolve inheritance or placeholders.
    If Not includeAttr Then
        If d.Exists(attrName) Then d.Remove attrName
        Exit Sub
    End If
    
    ' Attribute is desired in the result set. Start by resolving
    ' inheritance and substituting placeholders.
    ApplyLabelOverrides ini, data, attrName, attrValue, labelValue
    
    ' Append debugging information if requested
    If ini.graph.debug And (attrName = "label") Then
        labelValue = FormatDebugLabel(ini, data, labelValue)
    End If
    
    ' Address special characters in the label
    If IsLabelHTMLLike(labelValue) Then
        ' Don't touch the label, use it as given
    Else
        ' Normalize the string for Graphviz use
        labelValue = NormalizeLabelText(labelValue)
    End If
        
    ' Revise the attribute dictionary with the final label
    If d.Exists(attrName) Then  ' Update the dictionary
        If labelValue = vbNullString Then
            If allowBlank Then
                d(attrName) = blankToken        ' /E
            Else
                d(attrName) = labelValue        ' ""
            End If
        Else
            d(attrName) = labelValue            ' String value
        End If
    Else ' Add attribute to the dictionary
        If labelValue = vbNullString Then
            If allowBlank Then
                d.Add attrName, blankToken
            Else
                ' Omit the attribute, use Graphviz default behavior
            End If
        Else
            d.Add attrName, labelValue
        End If
    End If
End Sub

' ==========================================================================
' ROUTINE: FormatGraphLabels
'
' PURPOSE:
'   Applies graph-level label formatting rules to the attribute dictionary.
'   Processes the graph's primary label according to inclusion switches,
'   inheritance rules, placeholder substitution, and blank-value handling.
'   Produces a fully resolved attribute set ready for Graphviz emission.
'
' FUNCTIONAL WORKFLOW:
'   1. GRAPH LABEL ("label"):
'        - Uses data.label as the source value.
'        - Graph labels are always subject to inheritance and placeholder
'          substitution.
'        - Blank values are allowed (allowBlank = True).
'        - Delegates full processing to ProcessGraphLabelAttribute.
'
'   2. ATTRIBUTE DICTIONARY UPDATE:
'        - ProcessGraphLabelAttribute may:
'             o Resolve inheritance and placeholder substitution
'             o Apply debug formatting when enabled
'             o Normalize text for Graphviz compatibility
'             o Add, update, or remove the "label" attribute based on
'               inclusion rules
'
' TECHNICAL NOTES:
'   - This routine performs no dictionary reconstruction; callers must invoke
'     RebuildStyleAttributeString afterward.
'   - Ensures consistent graph-label behavior across all graph-emission paths.
'   - DeepWiki Context: Implements the graph-label rules described in the
'     "Graphs", "Labels", "Overrides", and "Serialization" sections.
' ==========================================================================
Private Sub FormatGraphLabels(ByRef ini As settings, ByRef data As dataRow, ByRef d As Dictionary)
    ProcessGraphLabelAttribute ini, _
                               data, _
                               d, _
                               "label", _
                               data.label, _
                               False
End Sub

' ==========================================================================
' ROUTINE: BuildStyleAttributeDictionary
'
' PURPOSE:
'   Constructs the effective attribute dictionary for a synthesized row using
'   the row's resolved style-format string (data.format) and extra-attribute
'   string (data.extraAttrs). Centralizes the attribute-source selection logic
'   so node and edge pipelines behave consistently and predictably.
'
' FUNCTIONAL WORKFLOW:
'   1. STYLE-FORMAT INCLUSION:
'        - When ini.graph.includeStyleFormat = True:
'             o If ini.graph.includeExtraAttributes = True:
'                   - Merges data.format and data.extraAttrs via
'                     MergeAttributeSets.
'             o Otherwise:
'                   - Parses data.format only.
'
'   2. EXTRA-ATTRIBUTE INCLUSION (NO STYLE FORMAT):
'        - When ini.graph.includeStyleFormat = False:
'             o If ini.graph.includeExtraAttributes = True:
'                   - Parses data.extraAttrs into a dictionary.
'             o Otherwise:
'                   - Returns an empty dictionary.
'
'   3. OUTPUT:
'        - Returns a Dictionary containing the final attribute set, ready for
'          inheritance, placeholder substitution, debugging augmentation, and
'          normalization by downstream handlers.
'
' TECHNICAL NOTES:
'   - This routine does not apply label inheritance or overrides; callers must
'     invoke ApplyLabelOverrides or Handle*Attribute afterward.
'   - Ensures consistent attribute-source behavior across node, edge, and
'     graph pipelines.
'   - DeepWiki Context: Implements the attribute-source selection rules
'     described in the "Styles", "Attributes", and "Serialization" sections.
' ==========================================================================
Private Function BuildStyleAttributeDictionary(ByRef ini As settings, _
                                               ByRef data As dataRow) As Dictionary

    Dim d As Dictionary

    If ini.graph.includeStyleFormat Then
        If ini.graph.includeExtraAttributes Then
            Set d = MergeAttributeSets(data.Format, data.extraAttrs)
        Else
            Set d = ParseAttributeString(data.Format)
        End If
    Else
        If ini.graph.includeExtraAttributes Then
            Set d = ParseAttributeString(data.extraAttrs)
        Else
            Set d = New Dictionary
        End If
    End If

    Set BuildStyleAttributeDictionary = d
End Function

' ==========================================================================
' FUNCTION: HandleGraphAttribute
'
' PURPOSE:
'   Applies graph-level label synthesis rules to a single Graphviz graph
'   attribute ("label") within the dictionary-based style pipeline. Supports
'   placeholder-based substitution and fallback attribute emission using
'   standard Graphviz label formatting.
'
' TECHNICAL WORKFLOW:
'   1. ATTRIBUTE PRESENCE CHECK:
'        - If the attribute already exists in the normalized style dictionary,
'          retrieves its current value for placeholder expansion or override
'          preservation.
'
'   2. PLACEHOLDER EXPANSION:
'        - If the attribute value contains the {label} placeholder token,
'          replaces it with the data row's label value.
'        - Static label values in the template are preserved as-is.
'
'   3. FALLBACK ATTRIBUTE EMISSION:
'        - If the attribute is not present in the dictionary and no placeholder
'          exists in the template, emits a standard "label=" attribute using
'          'FormatLabel' to ensure correct quoting and HTML-label handling.
'
'   4. ATTRIBUTE INTEGRATION:
'        - Updates the dictionary in-place, allowing the final attribute string
'          to be rebuilt by 'RebuildStyleAttributeString' with consistent
'          formatting across all graph-level attributes.
'
' TECHNICAL NOTES:
'   - Graph-level labels do not participate in blank-label suppression logic;
'     a label is always emitted or substituted.
'   - This routine is intentionally graph-specific; node and edge attributes
'     use their own dedicated handlers with additional rules.
'   - DeepWiki Context: Implements the graph-label synthesis rules described
'     in the "Graph Attributes" and "Styles" documentation.
' ==========================================================================
Private Sub ProcessGraphLabelAttribute( _
        ByRef ini As settings, _
        ByRef data As dataRow, _
        ByRef d As Dictionary, _
        ByVal attrName As String, _
        ByRef attrValue As String, _
        ByVal allowEmpty As Boolean)

    Dim labelValue As String
    labelValue = attrValue
    
    ' Attribute is desired in the result set. Start by resolving
    ' inheritance and substituting placeholders.
    ApplyLabelOverrides ini, data, attrName, attrValue, labelValue
    
    ' Append debugging information if requested
    If ini.graph.debug Then
        If Len(labelValue) > 0 Then
            labelValue = FormatDebugLabel(ini, data, labelValue)
        End If
    End If
    
    ' Address special characters in the label
    If IsLabelHTMLLike(labelValue) Then
        ' Don't touch the label, use it as given
    Else
        ' Normalize the string for Graphviz use
        labelValue = NormalizeLabelText(labelValue)
    End If
        
    ' Revise the attribute dictionary with the final label
    If d.Exists(attrName) Then  ' Update the dictionary
        If labelValue = vbNullString Then
            If allowEmpty Then
                d(attrName) = ""
            Else
                d.Remove attrName
            End If
        Else
            d(attrName) = labelValue
        End If
    Else ' Add attribute to the dictionary
        If labelValue = vbNullString Then
            If allowEmpty Then
                d.Add attrName, ""
            Else
                ' Nothing to add
            End If
        Else
            d.Add attrName, labelValue
        End If
    End If
End Sub

' ==========================================================================
' SECTION: NODE LABEL ASSEMBLY
' ==========================================================================

' ==========================================================================
' ROUTINE: FormatNodeLabels
'
' PURPOSE:
'   Applies node-specific label formatting rules to the attribute dictionary.
'   Processes each supported node label (label, xlabel, tooltip) according to
'   inclusion switches, inheritance rules, placeholder substitution, and
'   blank-value handling. Produces a fully resolved attribute set ready for
'   Graphviz emission.
'
' FUNCTIONAL WORKFLOW:
'   1. PRIMARY NODE LABEL ("label"):
'        - Uses data.label as the source value.
'        - Controlled by ini.graph.includeNodeLabels.
'        - Blank handling is inverted: allowBlank = Not ini.graph.blankNodeLabels.
'        - Delegates full processing to ProcessNodeLabelAttribute.
'
'   2. SECONDARY LABEL ("xlabel"):
'        - Uses data.xlabel.
'        - Controlled by ini.graph.includeNodeXLabels.
'        - Blank values are omitted (allowBlank = False).
'        - Delegates processing to ProcessNodeLabelAttribute.
'
'   3. TOOLTIP LABEL ("tooltip"):
'        - Uses data.tooltip.
'        - Controlled by ini.graph.includeNodeTooltips.
'        - Blank values are allowed (allowBlank = True).
'        - Delegates processing to ProcessNodeLabelAttribute.
'
'   4. ATTRIBUTE DICTIONARY UPDATE:
'        - Each call to ProcessNodeLabelAttribute may:
'             o Resolve inheritance and placeholder substitution
'             o Apply debug formatting when enabled
'             o Normalize text for Graphviz compatibility
'             o Add, update, or remove attributes based on inclusion rules
'
' TECHNICAL NOTES:
'   - This routine performs no dictionary reconstruction; callers must invoke
'     RebuildStyleAttributeString afterward.
'   - Ensures consistent node-label behavior across all node-emission paths.
'   - DeepWiki Context: Implements the node-label rules described in the
'     "Nodes", "Labels", "Overrides", and "Serialization" sections.
' ==========================================================================

Private Sub FormatNodeLabels(ByRef ini As settings, ByRef data As dataRow, ByRef d As Dictionary)

    ' label
    ProcessNodeLabelAttribute ini, data, d, "label", data.label, ini.graph.includeNodeLabels, Not ini.graph.blankNodeLabels
    
    ' xlabel
    ProcessNodeLabelAttribute ini, data, d, "xlabel", data.xlabel, ini.graph.includeNodeXLabels, False
    
    ' tooltip
    ProcessNodeLabelAttribute ini, data, d, "tooltip", data.tooltip, ini.graph.includeNodeTooltips, Not ini.graph.blankNodeTooltips
End Sub

' ==========================================================================
' ROUTINE: ProcessNodeLabelAttribute
'
' PURPOSE:
'   Synthesizes a single node attribute ("label" or "xlabel") for the
'   dictionary-based Graphviz style pipeline. Applies override rules,
'   placeholder expansion, optional debug decoration, conditional emission,
'   and controlled empty-value handling.
'
' FUNCTIONAL WORKFLOW:
'   1. ATTRIBUTE INCLUSION:
'        - If the attribute is disabled (includeAttr = False), it is removed
'          from the dictionary and no further processing occurs.
'
'   2. OVERRIDE & PLACEHOLDER RESOLUTION:
'        - Applies inherited or template-based overrides via ApplyLabelOverrides.
'        - Expands any placeholders using the current data-row context.
'
'   3. DEBUG AUGMENTATION:
'        - When debugging is enabled, appends a formatted debug label
'          (FormatDebugLabel or FormatDebugXLabel) depending on attribute type.
'
'   4. FINAL EMISSION RULES:
'        - If the attribute already exists in the dictionary:
'             o Updates it when the resolved value is non-empty.
'             o Writes an explicit empty string when allowed (allowEmpty = True).
'        - If the attribute does not exist:
'             o Adds the resolved value when non-empty.
'             o Adds an explicit empty string when allowed.
'
' TECHNICAL NOTES:
'   - Empty-label emission is controlled by allowEmpty and used primarily to
'     suppress Graphviz's default node-label fallback (\N).
'   - XLabels do not participate in blank-label suppression logic.
'   - This routine is node-specific; edge and graph attributes use separate
'     handlers with their own synthesis rules.
'   - DeepWiki Context: Implements the node-attribute synthesis rules described
'     in the "Node Attributes" and "Styles" documentation.
' ==========================================================================
Private Sub ProcessNodeLabelAttribute( _
        ByRef ini As settings, _
        ByRef data As dataRow, _
        ByRef d As Dictionary, _
        ByVal attrName As String, _
        ByRef attrValue As String, _
        ByVal includeAttr As Boolean, _
        ByVal allowEmpty As Boolean)

    Dim labelValue As String
    labelValue = attrValue
    
    ' If Attribute should not be in the dictionary, remove it. No need
    ' to resolve inheritance or placeholders.
    If Not includeAttr Then
        If d.Exists(attrName) Then d.Remove attrName
        Exit Sub
    End If
    
    ' Attribute is desired in the result set. Start by resolving
    ' inheritance and substituting placeholders.
    ApplyLabelOverrides ini, data, attrName, attrValue, labelValue
    
    ' Append debugging information if requested
    If ini.graph.debug And (attrName = "label" Or attrName = "xlabel") Then
        labelValue = FormatDebugLabel(ini, data, labelValue)
    End If
    
    ' Address special characters in the label
    If IsLabelHTMLLike(labelValue) Then
        ' Don't touch the label, use it as given
    Else
        ' Normalize the string for Graphviz use
        labelValue = NormalizeLabelText(labelValue)
    End If
        
    ' Revise the attribute dictionary with the final label
    If d.Exists(attrName) Then  ' Update the dictionary
        If labelValue <> vbNullString Then
            d(attrName) = labelValue
        ElseIf allowEmpty Then
            d(attrName) = ""
        End If
    Else ' Add attribute to the dictionary
        If labelValue <> vbNullString Then
            d.Add attrName, labelValue
        ElseIf allowEmpty Then
            d.Add attrName, ""
        End If
    End If
End Sub

Private Function NormalizeLabelText(ByVal rawData As String) As String
    If rawData = Chr$(34) & Chr$(34) Then   ' Special case: "" to blank a label
        NormalizeLabelText = rawData
    Else
        NormalizeLabelText = replace(rawData, Chr$(10), NEWLINE)             ' Chr(10) 0x0a LF  Line Feed
        NormalizeLabelText = replace(NormalizeLabelText, "\" & Chr$(34), Chr$(34))    ' In case they already escaped the double quote
        NormalizeLabelText = replace(NormalizeLabelText, Chr$(34), "\" & Chr$(34))    ' Chr(34)      " Double quotes (or speech marks)
    End If
End Function

' ==========================================================================
' ROUTINE: ProcessClusterLabelAttribute
'
' PURPOSE:
'   Processes a single cluster-level label attribute and updates the attribute
'   dictionary accordingly. Applies inclusion switches, inheritance rules,
'   placeholder substitution, debug augmentation, normalization, and
'   empty-value handling to produce a fully resolved attribute suitable for
'   Graphviz emission.
'
' FUNCTIONAL WORKFLOW:
'   1. INCLUSION SWITCH:
'        - When includeAttr = False:
'             o Removes attrName from the dictionary if present.
'             o Skips inheritance, placeholder substitution, debugging, and
'               normalization.
'             o Exits immediately.
'
'   2. INHERITANCE & PLACEHOLDER SUBSTITUTION:
'        - Invokes ApplyLabelOverrides to resolve:
'             o Inherited label values
'             o Placeholder tokens
'             o Context-dependent substitutions
'        - Produces labelValue as the working label text.
'
'   3. DEBUG AUGMENTATION:
'        - When ini.graph.debug = True AND attrName = "label":
'             o Invokes FormatDebugLabel to append row-level diagnostics.
'             o HTML-like labels receive "<BR/>debugStr>" before the closing
'               ">".
'
'   4. NORMALIZATION:
'        - Invokes NormalizeLabelText to ensure Graphviz-safe output:
'             o Escapes special characters
'             o Normalizes whitespace
'             o Removes illegal control characters
'
'   5. DICTIONARY UPDATE:
'        - If attrName already exists:
'             o Non-blank labelValue ? update the entry.
'             o Blank labelValue:
'                   - allowEmpty = True  ? set to "".
'                   - allowEmpty = False ? remove the entry.
'
'        - If attrName does not exist:
'             o Non-blank labelValue ? add the entry.
'             o Blank labelValue:
'                   - allowEmpty = True  ? add "".
'                   - allowEmpty = False ? do nothing.
'
' TECHNICAL NOTES:
'   - This routine handles only cluster-level label attributes; callers must
'     invoke RebuildStyleAttributeString afterward to serialize the dictionary.
'   - Debug augmentation uses the same HTML-aware rules as ApplyDebugLabel.
'   - DeepWiki Context: Implements the cluster-label rules described in the
'     "Clusters", "Labels", "Overrides", and "Serialization" sections.
' ==========================================================================
Private Sub ProcessClusterLabelAttribute( _
        ByRef ini As settings, _
        ByRef data As dataRow, _
        ByRef d As Dictionary, _
        ByVal attrName As String, _
        ByRef attrValue As String, _
        ByVal includeAttr As Boolean, _
        ByVal allowEmpty As Boolean)

    Dim labelValue As String
    labelValue = attrValue
    
    ' If Attribute should not be in the dictionary, remove it. No need
    ' to resolve inheritance or placeholders.
    If Not includeAttr Then
        If d.Exists(attrName) Then d.Remove attrName
        Exit Sub
    End If
    
    ' Attribute is desired in the result set. Start by resolving
    ' inheritance and substituting placeholders.
    ApplyLabelOverrides ini, data, attrName, attrValue, labelValue
    
    ' Append debugging information if requested
    If ini.graph.debug And (attrName = "label") Then
        labelValue = FormatDebugLabel(ini, data, labelValue)
    End If
    
    ' Address special characters in the label
    If IsLabelHTMLLike(labelValue) Then
        ' Don't touch the label, use it as given
    Else
        ' Normalize the string for Graphviz use
        labelValue = NormalizeLabelText(labelValue)
    End If
        
    ' Revise the attribute dictionary with the final label
    If d.Exists(attrName) Then  ' Update the dictionary
        If labelValue <> vbNullString Then
            d(attrName) = labelValue
        Else
            If allowEmpty Then
                d(attrName) = ""
            Else
                d.Remove attrName
            End If
        End If
    Else ' Add attribute to the dictionary
        If labelValue <> vbNullString Then
            d.Add attrName, labelValue
        Else
            If allowEmpty Then
                d.Add attrName, ""
            End If
        End If
    End If
End Sub

' ==========================================================================
' ROUTINE: RebuildStyleAttributeString
'
' PURPOSE:
'   Serializes an attribute dictionary into a Graphviz-compatible attribute
'   string. Applies label-specific formatting rules, quotes non-label values,
'   and emits a normalized, space-prefixed attribute list suitable for node,
'   edge, and graph statements.
'
' FUNCTIONAL WORKFLOW:
'   1. EMPTY DICTIONARY HANDLING:
'        - When d Is Nothing or d.Count = 0:
'             o Returns an empty string.
'             o Caller omits the attribute block entirely.
'
'   2. DICTIONARY TRAVERSAL:
'        - Iterates through all keys in the dictionary.
'        - Converts each key to lowercase for consistent Graphviz output.
'
'   3. LABEL-SPECIFIC FORMATTING:
'        - When IsLabelAttribute(k) = True:
'             o Formats the value using FormatLabel(d(key)).
'             o Ensures correct quoting, HTML handling, and newline behavior.
'
'   4. NON-LABEL ATTRIBUTE QUOTING:
'        - For all other attributes:
'             o Quotes the value using AddQuotes(d(key)).
'             o Ensures Graphviz-safe emission of strings containing spaces,
'               punctuation, or special characters.
'
'   5. STRING ASSEMBLY:
'        - Appends each attribute as:
'             " k=value"
'        - Produces a space-prefixed list that trims cleanly at the end.
'
'   6. OUTPUT:
'        - Returns Trim$(result) to remove leading/trailing whitespace.
'        - Caller inserts the string inside "[ ... ]" when emitting statements.
'
' TECHNICAL NOTES:
'   - Attribute order follows dictionary enumeration; Graphviz does not
'     require stable ordering.
'   - FormatLabel handles HTML-like labels, newline normalization, and debug
'     augmentation (when enabled).
'   - DeepWiki Context: Implements the attribute-serialization rules described
'     in the "Attributes", "Labels", and "Serialization" sections.
' ==========================================================================
Private Function RebuildStyleAttributeString(ByVal d As Dictionary) As String
    If d Is Nothing Or d.Count = 0 Then
        RebuildStyleAttributeString = ""
        Exit Function
    End If

    Dim result As String
    Dim key As Variant
    Dim k As String

    For Each key In d.keys
        k = LCase$(key)

        If IsLabelAttribute(k) Then
            result = result & " " & k & "=" & FormatGraphvizLabel(d(key))
        Else
            result = result & " " & k & "=" & AddQuotes(d(key))
        End If
    Next key

    RebuildStyleAttributeString = Trim$(result)
End Function

' ==========================================================================
' ROUTINE: IsLabelAttribute
'
' PURPOSE:
'   Determines whether a given attribute name represents a Graphviz label
'   attribute. Used during attribute-string reconstruction to decide whether
'   the value should be formatted via FormatLabel or quoted via AddQuotes.
'
' FUNCTIONAL WORKFLOW:
'   1. ATTRIBUTE CLASSIFICATION:
'        - Returns True when k matches one of the recognized label attributes:
'             o "label"
'             o "xlabel"
'             o "taillabel"
'             o "headlabel"
'             o "tooltip"
'
'   2. NON-LABEL ATTRIBUTES:
'        - Returns False for all other attribute names.
'        - Caller treats the attribute as a standard key/value pair requiring
'          AddQuotes during serialization.
'
' TECHNICAL NOTES:
'   - Attribute names are expected to be lowercase before calling this
'     routine; callers typically apply LCase$ during dictionary traversal.
'   - DeepWiki Context: Supports the attribute-serialization rules described
'     in the "Attributes", "Labels", and "Serialization" sections.
' ==========================================================================
Private Function IsLabelAttribute(ByVal k As String) As Boolean
    Select Case k
        Case "label", "xlabel", "taillabel", "headlabel", "tooltip"
            IsLabelAttribute = True
        Case Else
            IsLabelAttribute = False
    End Select
End Function

' ==========================================================================
' FUNCTION: MergeAttributeSets
'
' PURPOSE:
'   Produces a unified attribute dictionary by combining row-level style
'   attributes with template-level style attributes. Ensures consistent
'   key normalization and override precedence for all Graphviz attribute
'   pipelines (node, edge, and graph).
'
' TECHNICAL WORKFLOW:
'   1. ATTRIBUTE PARSING:
'        - Converts both input strings ('baseStr' and 'overrideStr') into
'          Dictionaries using 'ParseAttributeString'.
'        - Each dictionary is passed through 'NormalizeKeys' to enforce
'          lowercase, whitespace-trimmed keys for reliable comparison.
'
'   2. OVERRIDE MERGE:
'        - Combines the normalized dictionaries using 'MergeDictionaries'.
'        - Attributes from 'overrideStr' take precedence over those from
'          'baseStr', matching Graphviz's "last attribute wins" semantics.
'
'   3. PIPELINE INTEGRATION:
'        - The merged dictionary is returned for downstream processing by
'          node, edge, or graph label handlers.
'        - Ensures consistent behavior across all style-attribute workflows.
'
' TECHNICAL NOTES:
'   - Key normalization prevents case-sensitivity mismatches between
'     user-supplied templates and row-level formats.
'   - This routine is foundational: all label-formatting functions rely on
'     its override semantics and normalized key handling.
'   - DeepWiki Context: Implements the attribute-merge rules described in
'     the "Style Layers" and "Attribute Normalization" documentation.
' ==========================================================================
Private Function MergeAttributeSets(ByVal baseStr As String, ByVal overrideStr As String) As Dictionary
    Dim baseDict As Dictionary
    Dim overrideDict As Dictionary

    Set baseDict = NormalizeKeys(ParseAttributeString(baseStr))
    Set overrideDict = NormalizeKeys(ParseAttributeString(overrideStr))

    Set MergeAttributeSets = MergeDictionaries(baseDict, overrideDict)   ' override wins
End Function

' ==========================================================================
' FUNCTION: NormalizeKeys
'
' PURPOSE:
'   Produces a sanitized, case-normalized attribute dictionary to ensure
'   consistent key handling across all Graphviz style pipelines. Converts
'   incoming attribute keys to lowercase and trims incidental whitespace,
'   preventing mismatches between user-supplied templates and row-level
'   formats.
'
' TECHNICAL WORKFLOW:
'   1. KEY ENUMERATION:
'        - Iterates through all keys in the source Dictionary.
'        - Extracts each key as provided by 'ParseAttributeString' or
'          upstream merge operations.
'
'   2. NORMALIZATION:
'        - Converts each key to lowercase using 'LCase$'.
'        - Trims incidental whitespace to avoid accidental key duplication.
'        - Preserves the original attribute values without modification.
'
'   3. DICTIONARY REBUILD:
'        - Constructs a new Dictionary containing only normalized keys.
'        - Ensures reliable key comparison for downstream handlers such as
'          'ProcessNodeLabelAttribute', 'ProcessEdgeLabelAttribute', and
'          'HandleGraphAttribute'.
'
' TECHNICAL NOTES:
'   - Graphviz treats attribute names case-insensitively; this function
'     enforces that behavior within the VBA pipeline.
'   - Normalization prevents subtle bugs caused by mixed-case keys or
'     inconsistent CompareMode settings in Tim Hall's Dictionary class.
'   - DeepWiki Context: Implements the attribute-normalization rules
'     described in the "Style Layers" and "Attribute Normalization"
'     documentation.
' ==========================================================================
Private Function NormalizeKeys(ByVal d As Dictionary) As Dictionary
    Dim result As New Dictionary
    Dim key As Variant

    For Each key In d.keys
        result.Add LCase$(key), d(key)
    Next key

    Set NormalizeKeys = result
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
Private Function ProcessSubgraphOpen(ByRef ini As settings, _
                                     ByRef data As dataRow, _
                                     ByVal indent As Long, _
                                     ByRef clusterCnt As Long) As String
    Dim subgraphName As String
    subgraphName = Trim$(GetStringBetweenDelimiters(data.item, vbNullString, OPEN_BRACE))
                        
    If subgraphName = vbNullString Then          ' No subgraph name supplied
        ' Increment the cluster counter, and use it in the cluster name
        clusterCnt = clusterCnt + 1
        subgraphName = "cluster_" & clusterCnt
    End If

    Dim subgraphDirective As String
    subgraphDirective = Space(indent * ini.source.indent) & "subgraph " & AddQuotesConditionally(subgraphName) & " {" & " "

    ' Apply attribute inheritance
    Dim d As Dictionary
    Set d = BuildStyleAttributeDictionary(ini, data)

    ' The subgraph can have an optional label. Include it if specified
    ProcessClusterLabelAttribute ini, data, d, "label", data.label, ini.graph.includeClusterLabels, Not ini.graph.blankClusterLabels

    ' If output format is SVG, then include the tooltip data
    ProcessClusterLabelAttribute ini, data, d, "tooltip", data.tooltip, ini.graph.includeClusterTooltips, Not ini.graph.blankClusterTooltips
    
    ' Convert dictionary to a string of attributes
    Dim styleAttributes As String
    styleAttributes = RebuildStyleAttributeString(d)
    
    ProcessSubgraphOpen = subgraphDirective & styleAttributes & vbNewLine
    Set d = Nothing
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
Private Function ProcessNode(ByRef ini As settings, _
                             ByRef data As dataRow, _
                             ByVal indent As Long, _
                             ByVal nodesUsedInRelationships As Dictionary) As String
    Dim item As String
    Dim items() As String
    
    Dim graphvizSource As String
    
    Dim arrayIndex As Long
    
    item = data.item
    items = split(item, COMMA)
    
    For arrayIndex = LBound(items) To UBound(items)
        data.item = Trim$(items(arrayIndex))
                        
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
Private Function ProcessEdge(ByRef ini As settings, _
                             ByRef data As dataRow, _
                             ByVal indent As Long, _
                             ByVal definedNodes As Dictionary) As String
    Dim item As String
    Dim relatedItem As String
    Dim items() As String
    Dim relatedItems() As String
    
    Dim graphvizSource As String
    
    Dim itemIndex As Long
    Dim relatedItemIndex As Long
    
    item = data.item
    items = split(item, COMMA)
    
    relatedItem = data.relatedItem
    relatedItems = split(relatedItem, COMMA)
    
    For itemIndex = LBound(items) To UBound(items)
        For relatedItemIndex = LBound(relatedItems) To UBound(relatedItems)
            data.item = Trim$(items(itemIndex))
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
Private Function ProcessSubgraphClose(ByRef ini As settings, _
                                      ByRef data As dataRow, _
                                      ByVal indent As Long) As String
                                     
    ProcessSubgraphClose = Space(indent * ini.source.indent) & data.item & vbNewLine
End Function

' ==========================================================================
' SECTION: ENTITY WRITERS (FINAL DOT ASSEMBLY)
' ==========================================================================

' ==========================================================================
' ROUTINE: WriteNode
'
' PURPOSE:
'   Emits a fully formatted Graphviz node statement using the row's resolved
'   identifier, style-format attributes, and label metadata. Builds the
'   attribute dictionary, applies node-label formatting rules, reconstructs
'   the attribute string, and returns the final node command with proper
'   indentation.
'
' FUNCTIONAL WORKFLOW:
'   1. ATTRIBUTE DICTIONARY CONSTRUCTION:
'        - Builds the initial attribute dictionary via
'          BuildStyleAttributeDictionary(ini, data).
'        - Incorporates style-format attributes and extra attributes according
'          to graph-level switches.
'
'   2. NODE IDENTIFIER NORMALIZATION:
'        - Retrieves the node ID from data.item.
'        - Removes any port suffix (e.g., "node:port") by extracting the
'          portion before the colon.
'        - Ensures the final ID is suitable for Graphviz emission.
'
'   3. NODE LABEL FORMATTING:
'        - Applies FormatNodeLabels to:
'             o Resolve inheritance and placeholder substitution
'             o Apply inclusion switches for label/xlabel and other node
'               attributes
'             o Normalize text for Graphviz compatibility
'             o Insert debugging metadata when enabled
'
'   4. ATTRIBUTE STRING RECONSTRUCTION:
'        - Converts the updated dictionary back into a Graphviz attribute
'          string via RebuildStyleAttributeString.
'
'   5. NODE STATEMENT EMISSION:
'        - Writes the final node command using the configured indent level:
'             <indent><nodeId>;
'          or, when attributes exist:
'             <indent><nodeId> [ <attributes> ];
'        - Returns the completed line, including trailing newline.
'
' TECHNICAL NOTES:
'   - Node IDs are quoted conditionally via AddQuotesConditionally to ensure
'     compatibility with Graphviz rules for identifiers containing spaces or
'     special characters.
'   - Attribute omission follows Graphviz defaults when the dictionary is
'     empty.
'   - This routine does not modify context-level structures; it only formats
'     the node line for output.
'   - DeepWiki Context: Implements the node-emission rules described in the
'     "Nodes", "Labels", "Identifiers", and "Serialization" sections.
' ==========================================================================
Private Function WriteNode(ByRef ini As settings, _
                           ByRef data As dataRow, _
                           ByVal indent As Long) As String

    ' Convert the data row to a dictionary of attributes, applying styles and
    ' extra attribute overrides as dictated by switches
    Dim d As Dictionary
    Set d = BuildStyleAttributeDictionary(ini, data)

    ' Get the node ID
    Dim nodeId As String
    nodeId = data.item
    
    ' Strip off the port (if specified)
    If InStr(nodeId, ":") > 0 Then
        nodeId = GetStringTokenAtPosition(nodeId, ":", 1)
    End If

    ' Resolve inheritance perform placeholder substitution
    FormatNodeLabels ini, data, d
    
    ' Convert the dictionary back into an attribute string
    Dim attributes As String
    attributes = RebuildStyleAttributeString(d)
    
    If Len(attributes) = 0 Then
        WriteNode = Join(Array(Space(indent * ini.source.indent), AddQuotesConditionally(nodeId), SEMICOLON, vbNewLine), vbNullString)
    Else
        WriteNode = Join(Array(Space(indent * ini.source.indent), AddQuotesConditionally(nodeId), " [ ", attributes & " ];", vbNewLine), vbNullString)
    End If

    ' Release resources
    Set d = Nothing
End Function

' ==========================================================================
' ROUTINE: WriteEdge
'
' PURPOSE:
'   Emits a fully formatted Graphviz edge statement using the row's resolved
'   identifiers, style-format attributes, and label metadata. Builds the
'   attribute dictionary, applies edge-label formatting rules, reconstructs
'   the attribute string, and returns the final edge command with proper
'   indentation and operator selection.
'
' FUNCTIONAL WORKFLOW:
'   1. ATTRIBUTE DICTIONARY CONSTRUCTION:
'        - Builds the initial attribute dictionary via
'          BuildStyleAttributeDictionary(ini, data).
'        - Incorporates style-format attributes and extra attributes according
'          to graph-level switches.
'
'   2. EDGE LABEL FORMATTING:
'        - Applies FormatEdgeLabels to:
'             o Resolve inheritance and placeholder substitution
'             o Apply inclusion switches for label/xlabel/tail/head labels
'             o Normalize text for Graphviz compatibility
'             o Insert debugging metadata when enabled
'
'   3. ATTRIBUTE STRING RECONSTRUCTION:
'        - Converts the updated dictionary back into a Graphviz attribute
'          string via RebuildStyleAttributeString.
'
'   4. IDENTIFIER & PORT HANDLING:
'        - Formats the tail and head identifiers using FormatId, including
'          optional port suffixes when ini.graph.includeEdgePorts = True.
'
'   5. EDGE STATEMENT EMISSION:
'        - Writes the final edge command using the configured indent level:
'             <indent><tailId> <operator> <headId>;
'          or, when attributes exist:
'             <indent><tailId> <operator> <headId>[ <attributes> ];
'        - Returns the completed line, including trailing newline.
'
' TECHNICAL NOTES:
'   - Operator is selected from ini.graph.edgeOperator (e.g., "--" or "->").
'   - Attribute omission follows Graphviz defaults when the dictionary is
'     empty.
'   - This routine does not modify context-level structures; it only formats
'     the edge line for output.
'   - DeepWiki Context: Implements the edge-emission rules described in the
'     "Edges", "Labels", "Identifiers", and "Serialization" sections.
' ==========================================================================
Private Function WriteEdge(ByRef ini As settings, _
                           ByRef data As dataRow, _
                           ByVal indent As Long) As String

    ' Convert the data row to a dictionary of attributes, applying styles and
    ' extra attribute overrides as dictated by switches
    Dim d As Dictionary
    Set d = BuildStyleAttributeDictionary(ini, data)

    ' Collect the label, xlabel, taillabel, and headlabel labels into name value pairs
    FormatEdgeLabels ini, data, d

    ' Convert the dictionary back into an attribute string
    Dim attributes As String
    attributes = RebuildStyleAttributeString(d)
    
    ' Add the quotes to the id and (optional) port for the item, and the "is related to" item
    Dim tailId As String
    tailId = FormatId(data.item, ini.graph.includeEdgePorts)
    
    Dim headId As String
    headId = FormatId(data.relatedItem, ini.graph.includeEdgePorts)
    
    ' Write out the edge command
    If Len(attributes) = 0 Then
        WriteEdge = Join(Array(Space(indent * ini.source.indent), tailId, " ", ini.graph.edgeOperator, " ", headId, SEMICOLON, vbNewLine), vbNullString)
    Else
        WriteEdge = Join(Array(Space(indent * ini.source.indent), tailId, " ", ini.graph.edgeOperator, " ", headId, "[ ", attributes, " ];", vbNewLine), vbNullString)
    End If

    ' Release resources
    Set d = Nothing
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
'   2. INJECTION: Retrieves the 'label' field-which contains the raw DOT
'      syntax-and prepends the current level of indentation.
'   3. TERMINATION: Appends a newline to ensure the next DOT statement
'      starts on a fresh line in the Source Viewer.
'
' TECHNICAL NOTES:
'   - Strategy: Provides an "Escape Hatch" for advanced Graphviz features
'     not natively supported by the Excel UI (e.g., custom rank blocks
'     or complex multi-line attribute strings).
'   - Layer: Logic Layer / Native Passthrough.
' ==========================================================================
Private Function ProcessNative(ByRef ini As settings, _
                               ByRef data As dataRow, _
                               ByVal indent As Long) As String
                              
    ProcessNative = Space(indent * ini.source.indent) & data.label & vbNewLine
End Function

' ==========================================================================
' ROUTINE: ProcessKeyword
'
' PURPOSE:
'   Emits a fully formatted keyword statement (node, edge, or graph) using the
'   row's resolved attributes and indentation level. Builds the attribute
'   dictionary, applies label-formatting rules appropriate to the keyword
'   type, reconstructs the attribute string, and returns the final Graphviz
'   statement for inclusion in the Knowledge Graph output.
'
' FUNCTIONAL WORKFLOW:
'   1. ATTRIBUTE DICTIONARY CONSTRUCTION:
'        - Builds the initial attribute dictionary via
'          BuildStyleAttributeDictionary(ini, data).
'        - Provides the working set of attributes for label inheritance,
'          overrides, and formatting.
'
'   2. KEYWORD DISPATCH:
'        - Selects the correct label-formatting routine based on data.item:
'             o KEYWORD_NODE  -> FormatNodeLabels
'             o KEYWORD_EDGE  -> FormatEdgeLabels
'             o KEYWORD_GRAPH -> FormatGraphLabels
'        - Each formatter resolves inclusion switches, applies placeholder
'          substitution, performs inheritance, and normalizes label text.
'
'   3. ATTRIBUTE STRING RECONSTRUCTION:
'        - Converts the modified dictionary back into a Graphviz-compatible
'          attribute string via RebuildStyleAttributeString.
'
'   4. STATEMENT EMISSION:
'        - Produces the final keyword statement using the configured indent
'          level:
'             <indent><keyword>[ <attributes> ];
'        - Returns the completed line, including trailing newline.
'
' TECHNICAL NOTES:
'   - This routine does not modify context-level structures; it only formats
'     the keyword line for output.
'   - Indentation uses ini.source.indent to ensure consistent formatting
'     across all emitted statements.
'   - DeepWiki Context: Implements the keyword-emission rules described in
'     the "Keywords", "Labels", and "Serialization" sections.
' ==========================================================================
Private Function ProcessKeyword(ByRef ini As settings, _
                                ByRef data As dataRow, _
                                ByVal indent As Long) As String
    
    ' Handle attribute overrides and placeholder expansiongs
    Dim d As Dictionary
    Set d = BuildStyleAttributeDictionary(ini, data)
    
    ' Resolve inclusion switches and expand placeholders
    Select Case UCase$(data.item)
        Case KEYWORD_NODE
            FormatNodeLabels ini, data, d

        Case KEYWORD_EDGE
            FormatEdgeLabels ini, data, d

        Case KEYWORD_GRAPH
            FormatGraphLabels ini, data, d
    End Select

    ' Convert the modified dictionary back into an attribute string
    Dim attributes As String
    attributes = RebuildStyleAttributeString(d)
    
    ProcessKeyword = _
        Space(indent * ini.source.indent) & _
        data.item & "[ " & attributes & " ];" & vbNewLine

    ' Release resources
    Set d = Nothing
End Function

' ==========================================================================
' SECTION: LABEL SANITIZATION & SYNTAX SAFETY
' ==========================================================================

' ==========================================================================
' FUNCTION: FormatGraphvizLabel
'
' PURPOSE:
'   Produces a Graphviz-compliant label string by applying quoting rules,
'   HTML-label detection, and placeholder preservation. Ensures that all
'   node, edge, and graph labels are emitted in a form accepted by Graphviz
'   after placeholder expansion and attribute synthesis.
'
' TECHNICAL WORKFLOW:
'   1. HTML-LABEL DETECTION:
'        - Checks whether the label begins with '<' and ends with '>'.
'        - If so, returns the label unchanged, as Graphviz requires HTML
'          labels to remain unquoted and structurally intact.
'
'   2. QUOTING RULES:
'        - For non-HTML labels, wraps the label in double quotes.
'        - Escapes embedded double quotes when necessary to preserve
'          Graphviz syntax correctness.
'
'   3. PLACEHOLDER PRESERVATION:
'        - Leaves placeholder tokens (e.g., {label}, {xlabel}) untouched
'          when they appear in static template values; expansion occurs
'          earlier in the pipeline.
'
'   4. PIPELINE INTEGRATION:
'        - Returns a sanitized label string ready for inclusion in the
'          final attribute string produced by 'RebuildStyleAttributeString'.
'
' TECHNICAL NOTES:
'   - HTML-label detection is intentionally minimal: only outer-angle-bracket
'     framing is required for Graphviz to treat the label as HTML.
'   - Quoting rules ensure compatibility with both DOT and HTML-like label
'     constructs used throughout the style pipeline.
'   - DeepWiki Context: Implements the label-formatting rules described in
'     the "Label Syntax" and "Styles" documentation.
' ==========================================================================
Private Function FormatGraphvizLabel(ByVal labelValue As String) As String

    ' Case: ""
    If labelValue = Chr$(34) & Chr$(34) Then
        FormatGraphvizLabel = labelValue
        Exit Function
    End If
    
    ' Case: HTML-like, e.g. <<b>Bold Text</b>>
    If IsLabelHTMLLike(labelValue) Then
        FormatGraphvizLabel = labelValue
        Exit Function
    End If
    
    ' Case: Ordinary text
    FormatGraphvizLabel = AddQuotes(labelValue)
End Function

' ==========================================================================
' SECTION: HTML-LIKE LABEL DETECTION
' ==========================================================================

