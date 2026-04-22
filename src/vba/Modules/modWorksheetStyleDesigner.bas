Attribute VB_Name = "modWorksheetStyleDesigner"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetStyleDesigner
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Style Designer
'
' ROLE:
'   Core logic for the interactive Style Designer worksheet. Provides the
'   live-preview engine, attribute synthesis pipeline, and persistence
'   mechanisms that convert Ribbon UI selections into Graphviz-ready
'   style definitions for nodes, edges, and clusters.
'
' RESPONSIBILITIES:
'   - Live preview engine:
'       • RenderElement orchestrates DOT generation and Graphviz rendering
'       • GeneratePreviewGraph builds complete preview graphs with captions
'       • PreviewStyle executes the full render pipeline and injects images
'
'   - Attribute synthesis:
'       • GetNodeStyle / GetEdgeStyle / GetClusterStyle aggregate UI inputs
'       • AddAttribute / AddAttributeGroup / AddStyleAttribute construct
'         syntactically valid Graphviz attribute strings
'
'   - UI integration:
'       • Manages label fields, dropdowns, gradient controls, shape galleries
'       • Handles dynamic visibility of controls (e.g., gradient, image options)
'
'   - Rendering environment:
'       • GetRenderInfo summarizes active Graphviz engine settings
'       • Supports HTML-like labels, image paths, rankdir, splines, and layout
'
'   - Persistence:
'       • Writes composed styles back to the Styles worksheet
'       • Supports preview-image caching and cleanup
'
' ARCHITECTURAL NOTES:
'   - Integrates tightly with Graphviz (dot) via the Graphviz class wrapper
'   - Uses Ribbon galleries for high-speed color, font, and shape selection
'   - Supports Windows and macOS, with platform-specific image and font logic
'   - Designed for real-time feedback with minimal worksheet flicker
'
' VERSION NOTES:
'   - v5.5.00–v5.8.00 (2022–2023):
'       • Added font preview images and dynamic font pruning
'       • Added metric measurement support (mm -> inches)
'       • Added solid-fill gradient support (Weight %)
'       • Improved color/shape galleries and preview rendering
'
'   - v6.0.00–v6.1.01 (2023–2024):
'       • Major performance improvements to Style Designer tab loading
'       • Added caching of color and font preview images
'       • Added progress indicators during large gallery loads
'       • Added Mrecord to the list of supported shapes
'
'   - v8.0.0 (Aug 27, 2025):
'       • Replaced dropdowns with Ribbon galleries for colors, fonts, shapes
'       • Added RGB color picker (Windows + macOS)
'       • Improved preview rendering and iconography
'       • Added style-saving enhancements and auto-refresh of previews
'       • Added relative image-path extraction for portability
'
'   - v9.0.0–v9.1.0 (Dec 2025–Jan 2026):
'       • Updated font gallery logic and expanded font list
'       • Improved font preview rendering and deduplication
'
'   - v10.1.0 (Feb 9, 2026):
'       • Added floating action buttons for editing and refreshing styles
'       • Restored Image Zoom dropdown list with 5%–150% range
'
' USAGE:
'   - Called whenever a Style Designer control changes or when the user
'     presses the Render button to preview a Node, Edge, or Cluster style.
'
' RELATED WIKI PAGES:
'   - Style Designer Overview
'   - Graphviz Attribute Reference
'   - Live Preview Rendering Pipeline
' =============================================================================

Option Explicit

' Uncomment code below if encontering "Runtime Error 49, Bad DLL calling convention"
' Refer to: https://stackoverflow.com/questions/15758834/runtime-error-49-bad-dll-calling-convention
'Private Enum Something
'    member = 1
'End Enum

' ==========================================================================
' PROCEDURE: RenderElement
'
' PURPOSE:
'   The primary controller for the Style Designer's "Live Preview" feature.
'   It synchronizes UI inputs, generates the DOT attribute string, and
'   triggers the Graphviz rendering pipeline for a single design element.
'
' TECHNICAL WORKFLOW:
'   1. LABEL AGGREGATION: Collects label data (standard, xLabel, head/tail)
'      into a 'LabelSet' object, filtered by 'elementType' (Node, Edge, Cluster).
'   2. ATTRIBUTE GENERATION (Conditional):
'      - If 'createFormat' is TRUE: Rebuilds the DOT string by polling the
'        designer's dropdown settings (via 'GetNodeStyle', etc.).
'      - If 'createFormat' is FALSE: Reads the existing manual edits
'        directly from the 'formatCellName' range.
'   3. DOT COMPOSITION: Invokes 'GeneratePreviewGraph' to wrap attributes
'      into a valid Graphviz source string, optionally including captions.
'   4. OUTPUT EXECUTION:
'      - Updates the 'ShowSource' debug view.
'      - Calls 'PreviewStyle' to render the PNG and place it in 'previewCell'.
'
' USAGE:
'   - Called whenever a change is detected in the Style Designer interface
'     or when the "Render" button is pressed.
' ==========================================================================
Public Sub RenderElement(ByVal formatCellName As String, ByVal previewCellName As String, ByVal elementType As String, ByVal createFormat As Boolean)

    Dim styleAttributes As String
    Dim previewCell As String
    Dim dotSource As String
    Dim addCaption As Boolean
    
    ' Nodes, edges, and clusters all support label attribute
    Dim labels As LabelSet
    labels.label = Trim$(StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT).value)
    Select Case elementType
        Case KEYWORD_NODE
            labels.xLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).value)
            labels.headLabel = vbNullString
            labels.tailLabel = vbNullString
        Case KEYWORD_EDGE
            labels.xLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).value)
            labels.headLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_HEAD_LABEL_TEXT).value)
            labels.tailLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_TAIL_LABEL_TEXT).value)
        Case KEYWORD_CLUSTER
            labels.xLabel = vbNullString
            labels.headLabel = vbNullString
            labels.tailLabel = vbNullString
    End Select

    If createFormat Then
        ' Clear the Style cell (can't use .ClearContents on merged cells)
        StyleDesignerSheet.Range(formatCellName).value = vbNullString
        
        ' Generate the Style Definition from the dropdown lists
        Select Case elementType
            Case KEYWORD_NODE
                styleAttributes = GetNodeStyle()
            Case KEYWORD_EDGE
                styleAttributes = GetEdgeStyle()
            Case KEYWORD_CLUSTER
                styleAttributes = GetClusterStyle()
        End Select
        
        ' Display the style definition which was created
        StyleDesignerSheet.Range(formatCellName).value = styleAttributes
    Else
        ' The user has composed/edited the format. Use the value in the format cell
        styleAttributes = Trim$(StyleDesignerSheet.Range(formatCellName).value)
    End If
    
    ' Get the user-specified cell where the preview image should be displayed
    previewCell = Trim$(StyleDesignerSheet.Range(previewCellName).value)
    If previewCell <> vbNullString Then
        
        ' Find out if the user wants the graph options included in the preview
        If StyleDesignerSheet.Range(DESIGNER_ADD_CAPTION).value = TOGGLE_YES Then
            addCaption = True
        End If
        
        ' Create the Graphviz statements which can preview the style
        dotSource = GeneratePreviewGraph(elementType, labels, styleAttributes, addCaption)
        
        ' Display the source
        ShowSource dotSource
        
        ' Generate the image, and display it at the location specified
        PreviewStyle dotSource, previewCell
    End If

End Sub

' ==========================================================================
' FUNCTION: GeneratePreviewGraph
'
' PURPOSE:
'   Constructs a complete, valid Graphviz DOT source string tailored for
'   rendering style previews. It ensures that global settings (layout,
'   splines, rankdir) are correctly integrated with local style attributes.
'
' TECHNICAL WORKFLOW:
'   1. OPTION AGGREGATION: Pulls global engine settings from 'SettingsSheet',
'      including Layout, Splines, and ImagePath.
'   2. LAYOUT LOGIC: Specifically injects 'rankdir' (direction) if using
'      the "dot" engine, while suppressing spline logic for Node previews.
'   3. TEMPLATE INJECTION: Builds the DOT body based on 'elementType':
'      - NODES: Renders a single node with optional xLabel.
'      - EDGES: Renders a standard connection with head/tail label support.
'      - CLUSTERS: Creates a 'subgraph cluster_1' containing dummy nodes
'        to visualize interior container styling.
'   4. CAPTIONING: Optionally appends a technical caption using 'GetPreviewCaption'.
'   5. TOKEN REPLACEMENT: Swaps internal placeholders (%N, %H, %T) with
'      localized labels from the workbook to finalize the source.
'
' RETURN:
'   A String containing the full 'digraph main { ... }' Graphviz source.
' ==========================================================================
Public Function GeneratePreviewGraph(ByVal elementType As String, _
                                     ByRef labels As LabelSet, _
                                     ByVal styleAttributes As String, _
                                     ByVal addCaption As Boolean) As String

    Dim graphOptions As String
    
    Dim layout As String
    layout = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
    If layout <> vbNullString Then
        AddNameValue graphOptions, GRAPHVIZ_LAYOUT, layout
    End If

    ' Node previews do not use splines
    If elementType <> KEYWORD_NODE Then
        Dim splines As String
        splines = SettingsSheet.Range(SETTINGS_SPLINES).value
        If splines <> vbNullString Then
            AddNameValue graphOptions, GRAPHVIZ_SPLINES, splines
        End If
    End If
    
    ' Tweak the graph options to give the previews a tiny border
    AddNameValue graphOptions, GRAPHVIZ_PAD, AddQuotes("0.0625,0.0625")

    ' If the graphing layout is "dot" add in the direction specification
    If layout = LAYOUT_DOT And elementType <> KEYWORD_NODE Then
        Dim direction As String
        direction = SettingsSheet.Range(SETTINGS_RANKDIR).value
        If direction <> vbNullString Then
            AddNameValue graphOptions, GRAPHVIZ_RANKDIR, direction
        End If
    End If

    ' HTML-like labels can specify <img>, so inclue the image path
    AddNameValue graphOptions, GRAPHVIZ_IMAGEPATH, AddQuotes(GetImagePath())
    
    graphOptions = graphOptions & " " & SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).value
    
    ' =====================================================================
    ' Convert the data to graphviz format
    ' =====================================================================
    
    Dim dotSource As String
    dotSource = "digraph main {" & graphOptions & vbNewLine
   
    If addCaption Then
        dotSource = dotSource & " " & GetPreviewCaption(elementType, layout, SettingsSheet.Range(SETTINGS_SPLINES).value, direction) & vbNewLine
    End If

    If elementType = KEYWORD_NODE Then
        dotSource = dotSource & "  %N1 [" & FormatLabel(GRAPHVIZ_LABEL, labels.label) & FormatOptionalLabel(GRAPHVIZ_XLABEL, labels.xLabel) & " " & styleAttributes & " ];" & vbNewLine
        
    ElseIf elementType = KEYWORD_EDGE Then
        dotSource = dotSource & GetPreviewNodeEdge(GetPreviewNodeStyle("gray", "gray"))
        dotSource = dotSource & " [" & FormatLabel(GRAPHVIZ_LABEL, labels.label) & FormatOptionalLabel(GRAPHVIZ_XLABEL, labels.xLabel) & FormatOptionalLabel("headlabel", labels.headLabel) & FormatOptionalLabel("taillabel", labels.tailLabel) & " " & styleAttributes & " ];" & vbNewLine
        
    ElseIf elementType = KEYWORD_CLUSTER Then
        dotSource = dotSource & "  subgraph cluster_1 { "
        dotSource = dotSource & styleAttributes & FormatLabel(GRAPHVIZ_LABEL, labels.label) & " " & vbNewLine
        
        dotSource = dotSource & "    node[ shape=rect style=filled fillcolor=white pencolor=black fixedsize=true ];" & vbNewLine
        dotSource = dotSource & "    1[sortv=1 height=0.25  width=0.25];" & vbNewLine
        dotSource = dotSource & "    2[sortv=2 height=0.5   width=0.5];" & vbNewLine
        dotSource = dotSource & "    3[sortv=3 height=0.375 width=0.375];" & vbNewLine
        dotSource = dotSource & "    4[sortv=4 height=0.5   width=0.25];" & vbNewLine
        dotSource = dotSource & "    5[sortv=5 height=0.25  width=0.25];" & vbNewLine
        dotSource = dotSource & "    6[sortv=6 height=0.25  width=0.5];" & vbNewLine
        dotSource = dotSource & "    7[sortv=7 height=0.25  width=0.25];" & vbNewLine
        dotSource = dotSource & "  " & CLOSE_BRACE & vbNewLine
    End If
    
    dotSource = dotSource & CLOSE_BRACE

    dotSource = replace(dotSource, "%N", GetLabel("PreviewNode"))
    dotSource = replace(dotSource, "%H", GetLabel("PreviewHead"))
    dotSource = replace(dotSource, "%T", GetLabel("PreviewTail"))
    
    GeneratePreviewGraph = dotSource
    
End Function

' ==========================================================================
' FUNCTION: GetRenderInfo
'
' PURPOSE:
'   Generates a human-readable summary of the current Graphviz engine
'   configuration and rendering parameters.
'
' TECHNICAL WORKFLOW:
'   1. CLI FLAG ASSEMBLY: Extracts the target image format (-T) and
'      layout engine (-K) from the 'SettingsSheet'.
'   2. ATTRIBUTE MAPPING: Appends global 'splines' and 'rankdir' settings
'      to the string to reflect the active rendering environment.
'   3. COMPATIBILITY VALIDATION: Checks if the current 'layout' engine
'      supports 'cluster' (subgraph) rendering.
'   4. WARNING INJECTION: Appends a warning message if the user selects
'      a layout engine (e.g., Circo, Neato, Patchwork) known to be
'      incompatible with cluster definitions.
'
' RETURN:
'   A String representing the active render profile and any compatibility
'   conflicts (e.g., "-Tpng -Kdot splines=ortho rankdir=LR").
' ==========================================================================
Public Function GetRenderInfo() As String
    Dim label As String
    
    Dim format As String
    format = SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value
    label = "-T" & format
    
    Dim layout As String
    layout = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value
    label = label & " -K" & layout
    
    Dim splines As String
    splines = SettingsSheet.Range(SETTINGS_SPLINES).value
    If splines <> vbNullString Then
        label = label & " splines=" & splines
    End If
    
    If layout = LAYOUT_DOT Then
        Dim direction As String
        direction = SettingsSheet.Range(JSON_SETTINGS_RANKDIR).value
        If direction <> vbNullString Then
            label = label & " rankdir=" & direction
        End If
    End If
    
    Dim mode As String
    mode = StyleDesignerSheet.Range(DESIGNER_MODE).value
    If mode = KEYWORD_CLUSTER Then
        If layout = LAYOUT_CIRCO Or layout = LAYOUT_NEATO Or layout = LAYOUT_PATCHWORK Or layout = LAYOUT_SFDP Or layout = LAYOUT_TWOPI Then
            label = label & " | " & layout & " does not support clusters"
        End If
    End If
    
    
    GetRenderInfo = label
End Function

' ==========================================================================
' FUNCTION: FormatLabel
'
' PURPOSE:
'   Constructs a syntactically correct Graphviz label attribute,
'   automatically detecting and handling the difference between
'   standard strings and HTML-like labels.
'
' TECHNICAL WORKFLOW:
'   1. TYPE DETECTION: Invokes 'IsLabelHTMLLike' to check if the
'      'labelValue' contains HTML/XML-like delimiters (<...>).
'   2. BRANCHED FORMATTING:
'      - HTML LABELS: Appends the value directly without quotes or scrubbing,
'        as required by Graphviz for angled-bracket labels.
'      - STANDARD LABELS: Applies 'ScrubText' to handle illegal characters
'        and wraps the result in quotes via 'AddQuotes'.
'   3. STRING ASSEMBLY: Returns a space-prefixed attribute assignment
'      (e.g., ' label="Value"' or ' label=<HTML_Value>').
' ==========================================================================
Private Function FormatLabel(ByVal labelName As String, ByVal labelValue As String) As String
    If IsLabelHTMLLike(labelValue) Then
        FormatLabel = " " & labelName & "=" & labelValue
    Else
        FormatLabel = " " & labelName & "=" & AddQuotes(ScrubText(labelValue))
    End If
End Function

' ==========================================================================
' FUNCTION: FormatOptionalLabel
'
' PURPOSE:
'   A defensive wrapper for label generation that suppresses attribute
'   creation if the provided value is empty.
'
' TECHNICAL WORKFLOW:
'   1. NULL CHECK: Trims the 'labelValue' and evaluates if it is an empty
'      string or null.
'   2. CONDITIONAL EXECUTION:
'      - If empty: Returns 'vbNullString' to prevent cluttering the DOT
'        source with empty attributes (e.g., label="").
'      - If populated: Delegating the heavy lifting to 'FormatLabel' to
'        ensure correct quoting and HTML-like detection.
'
' USAGE:
'   - Used for secondary labels like 'xlabel', 'headlabel', or 'taillabel'
'     that may not be defined for every style.
' ==========================================================================
Private Function FormatOptionalLabel(ByVal labelName As String, ByVal labelValue As String) As String
    If Trim$(labelValue) = vbNullString Then
        FormatOptionalLabel = vbNullString
    Else
        FormatOptionalLabel = FormatLabel(labelName, labelValue)
    End If
End Function

' ==========================================================================
' FUNCTION: GetPreviewNodeEdge
'
' PURPOSE:
'   Generates the DOT structural skeleton required to preview an edge style,
'   including its terminal "Head" and "Tail" nodes.
'
' TECHNICAL WORKFLOW:
'   1. TAIL DEFINITION: Declares the starting node placeholder (%T) using
'      the provided 'nodeStyle' attributes.
'   2. HEAD DEFINITION: Declares the destination node placeholder (%H) using
'      the same consistent styling.
'   3. EDGE CONNECTION: Establishes the directional relationship (%T->%H),
'      leaving the string open for trailing edge-specific attributes.
'   4. TOKENIZATION: Utilizes placeholders (%T, %H) to be resolved later
'      by the main rendering engine for localization support.
'
' USAGE:
'   - Called by 'GeneratePreviewGraph' specifically when the 'elementType'
'     is set to 'KEYWORD_EDGE'.
' ==========================================================================
Public Function GetPreviewNodeEdge(ByVal nodeStyle As String) As String
    GetPreviewNodeEdge = GetPreviewNodeEdge & "    %T [" & nodeStyle & "];" & vbNewLine
    GetPreviewNodeEdge = GetPreviewNodeEdge & "    %H [" & nodeStyle & "];" & vbNewLine
    GetPreviewNodeEdge = GetPreviewNodeEdge & "    %T->%H"
End Function

' ==========================================================================
' FUNCTION: GetPreviewCaption
'
' PURPOSE:
'   Creates a technical "legend" node within the Graphviz DOT source to
'   display rendering metadata alongside the visual preview.
'
' TECHNICAL WORKFLOW:
'   1. METADATA ASSEMBLY: Constructs a multi-line string containing the
'      element type, engine layout, and spline settings using Graphviz
'      left-aligned line breaks (\l).
'   2. CONDITIONAL CONTEXT: Appends engine-specific data (rankdir) and
'      compatibility warnings (e.g., engines that ignore clusters).
'   3. NODE STYLING: Defines a 'plaintext' node named "legend" with
'      specific font properties (Arial, Size 10) to ensure the caption
'      is legible but non-intrusive.
'   4. OUTPUT: Returns a fully formed DOT node declaration.
'
' USAGE:
'   - Called by 'GeneratePreviewGraph' when the 'addCaption' flag is TRUE.
'   - Useful for debugging layout behavior during style development.
' ==========================================================================
Public Function GetPreviewCaption(ByVal elementType As String, ByVal layout As String, ByVal graphSplines As String, ByVal direction As String) As String

    Dim label As String
    label = elementType & "\l\llayout: " & layout & " \lsplines: " & graphSplines & "\l"
    
    If layout = LAYOUT_DOT Then
        label = label & "rankdir: " & direction & "\l"
    End If
    
    If elementType = KEYWORD_CLUSTER Then
        If layout = LAYOUT_CIRCO Or layout = LAYOUT_NEATO Or layout = LAYOUT_PATCHWORK Or layout = LAYOUT_SFDP Or layout = LAYOUT_TWOPI Then
            label = label & "\lNOTE: '" & layout & "' layout does not support clusters.\l"
        End If
    End If

    Dim caption As String
    caption = AddQuotes("legend") & "["
    AddNameValue caption, GRAPHVIZ_SHAPE, "plaintext"
    AddNameValue caption, GRAPHVIZ_FONTNAME, "Arial"
    AddNameValue caption, GRAPHVIZ_FONTSIZE, "10"
    AddNameValue caption, GRAPHVIZ_LABEL, AddQuotes(label)
    GetPreviewCaption = caption & "];"
    
End Function

' ==========================================================================
' FUNCTION: GetPreviewNodeStyle
'
' PURPOSE:
'   Returns a hard-coded DOT attribute string for the terminal nodes used in
'   edge previews (the "Head" and "Tail" nodes).
'
' TECHNICAL WORKFLOW:
'   1. GEOMETRY DEFINITION: Configures a 0.5" octagonal polygon (sides=8)
'      with a 22.5-degree orientation to provide a distinct "anchor" look.
'   2. TYPOGRAPHY: Sets a consistent font (Arial, 10pt) and maps the
'      'fontColor' parameter to the 'fontcolor' attribute.
'   3. VISUAL STYLING: Enforces a 'filled' white background with a
'      customizable 'pencolor' to ensure the nodes don't distract from
'      the edge being previewed.
'   4. ATTRIBUTE AGGREGATION: Uses 'AddNameValue' to safely concatenate
'      the DOT key-value pairs.
'
' USAGE:
'   - Primary utility for 'GetPreviewNodeEdge' to maintain a uniform
'     appearance for structural "dummy" nodes.
' ==========================================================================
Public Function GetPreviewNodeStyle(ByVal pencolor As String, ByVal fontColor As String) As String

    Dim styleAttributes As String
    
    AddNameValue styleAttributes, GRAPHVIZ_SHAPE, "polygon"
    AddNameValue styleAttributes, GRAPHVIZ_SIDES, "8"
    AddNameValue styleAttributes, GRAPHVIZ_COLOR, pencolor
    AddNameValue styleAttributes, GRAPHVIZ_FIXEDSIZE, "true"
    AddNameValue styleAttributes, GRAPHVIZ_FONTNAME, "Arial"
    AddNameValue styleAttributes, GRAPHVIZ_FONTSIZE, "10"
    AddNameValue styleAttributes, GRAPHVIZ_FONTCOLOR, fontColor
    AddNameValue styleAttributes, GRAPHVIZ_HEIGHT, "0.50"
    AddNameValue styleAttributes, GRAPHVIZ_WIDTH, "0.50"
    AddNameValue styleAttributes, GRAPHVIZ_STYLE, "filled"
    AddNameValue styleAttributes, GRAPHVIZ_FILLCOLOR, "white"
    AddNameValue styleAttributes, GRAPHVIZ_ORIENTATION, "22.5"

    GetPreviewNodeStyle = styleAttributes
End Function

' ==========================================================================
' PROCEDURE: PreviewStyle
'
' PURPOSE:
'   Executes the end-to-end rendering pipeline to transform a DOT source
'   string into a physical image displayed within the Style Designer.
'
' TECHNICAL WORKFLOW:
'   1. ENGINE INITIALIZATION: Instantiates a new 'Graphviz' object and
'      configures environmental paths (Temp directory, binary paths).
'   2. WORKSPACE PREPARATION: Purges existing preview images from the
'      'StyleDesignerSheet' to prevent visual stacking/clutter.
'   3. RENDERING EXECUTION:
'      - Commits 'graphvizSource' to a temporary physical file.
'      - Invokes 'RenderGraph' using user-specified engine and CLI parameters.
'   4. DIAGNOSTICS: Captures and redirects Graphviz console feedback to
'      the 'Console' worksheet for troubleshooting.
'   5. IMAGE INJECTION: Calls 'InsertPicture' to embed the resulting PNG
'      at the 'targetCell' location with a descriptive Alt-Text tag.
'   6. RESOURCE HYGIENE: Force-deletes temporary DOT and PNG files and
'      clears the object from memory to prevent file system bloat.
'
' USAGE:
'   - Triggered by 'RenderElement' to finalize the "Live Preview" loop.
' ==========================================================================
Public Sub PreviewStyle(ByVal graphvizSource As String, ByVal targetCell As String)
    
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    ' Instantiate a Graphviz object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz
    
    ' Prepare the file names
    graphvizObj.OutputDirectory = GetTempDirectory()
    graphvizObj.FilenameBase = "PreviewStyle"
    graphvizObj.GraphFormat = SettingsSheet.Range(SETTINGS_IMAGE_TYPE).value

    ' Remove any image from a previous run of the macro
    DeleteAllPictures StyleDesignerSheet.name
      
    ' Write the Graphviz data to a file so it can be sent to a rendering engine
    graphvizObj.graphvizSource = graphvizSource
    graphvizObj.SourceToFile
    
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
    Dim shapeObject As shape
    '@Ignore AssignmentNotUsed
    Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range(targetCell), False, True, "Style designer preview image.")
    Set shapeObject = Nothing
                    
    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Release the Graphviz object
    Set graphvizObj = Nothing
End Sub

' ==========================================================================
' PROCEDURE: AddAttribute
'
' PURPOSE:
'   A helper utility that conditionally appends a Graphviz attribute to a
'   style string based on the presence of a value in a specific Designer cell.
'
' TECHNICAL WORKFLOW:
'   1. VALUE EXTRACTION: Retrieves and trims the content from the Excel
'      range specified by 'cellName'.
'   2. EXISTENCE CHECK: Exits immediately if the cell is empty to prevent
'      generating malformed or unnecessary DOT attributes.
'   3. STRING COMPOSITION: Uses 'Join' and 'Array' to concatenate the
'      current 'styleAttributes' with the new key-value pair.
'   4. SMART QUOTING: Invokes 'AddQuotesConditionally' to ensure the
'      'cellValue' is properly escaped for the Graphviz parser.
'
' USAGE:
'   - Frequently used within 'GetNodeStyle', 'GetEdgeStyle', and
'     'GetClusterStyle' to build complex DOT strings from UI inputs.
' ==========================================================================
Private Sub AddAttribute(ByRef styleAttributes As String, _
                            ByVal attrName As String, _
                            ByVal cellName As String)
    ' Get the cell value
    Dim cellValue As String
    cellValue = Trim$(StyleDesignerSheet.Range(cellName).value)
    
    If cellValue = vbNullString Then Exit Sub
    
    styleAttributes = Join(Array(styleAttributes, " ", attrName, "=", AddQuotesConditionally(cellValue)), vbNullString)

End Sub

'@Ignore UseMeaningfulName
Public Sub AddAttributeGroup(ByRef styleAttributes As String, _
                             ByVal attrName As String, _
                             ByVal cellName1 As String, _
                             ByVal cellName2 As String, _
                             ByVal cellName3 As String, _
                             ByVal separator As String)

    Dim cellValue As String

    ' Get first group attribute. If blank, ignore the others in the group
    cellValue = Trim$(StyleDesignerSheet.Range(cellName1).value)
    
    If cellValue <> vbNullString Then
        ' Start building the group attribute
        styleAttributes = styleAttributes & " " & attrName & "=" & Chr$(34) & cellValue
    
        ' Get the second attribute of the group
        cellValue = Trim$(StyleDesignerSheet.Range(cellName2).value)
        
        ' Add to set of attributes if not blank
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & separator & cellValue
    
            ' Get the third group attribute
            cellValue = Trim$(StyleDesignerSheet.Range(cellName3).value)
            
            ' Add to set of attributes if not blank
            If cellValue <> vbNullString Then
                styleAttributes = styleAttributes & separator & cellValue
            End If
        End If
    
        ' Close the double quotes around the set of attributes
        styleAttributes = styleAttributes & Chr$(34)
    End If
End Sub

' ==========================================================================
' PROCEDURE: AddAttributeGroup
'
' PURPOSE:
'   Constructs multi-part Graphviz attributes (like 'margin' or 'peripheries')
'   where multiple Excel cell values are combined into a single quoted string.
'
' TECHNICAL WORKFLOW:
'   1. PRIMARY VALIDATION: Checks the first cell ('cellName1'); if empty, the
'      entire group is skipped to maintain valid DOT syntax.
'   2. NESTED CONCATENATION: Iteratively appends values from the second and
'      third cells only if they contain data, using a custom 'separator'
'      (usually a comma or space).
'   3. DELIMITER MANAGEMENT: Manually wraps the entire combined sequence in
'      double quotes (Chr$(34)) to treat the group as a single attribute value.
'
' USAGE:
'   - Used for attributes that accept coordinate pairs or complex lists
'     (e.g., width/height pairs or margin settings).
' ==========================================================================
'@Ignore UseMeaningfulName
Public Sub AddStyleAttribute(ByRef styleAttributes As String, _
                             ByVal cellName1 As String, _
                             ByVal cellName2 As String, _
                             ByVal cellName3 As String, _
                             ByVal gradientType As String)
    Dim cellValue As String
    Dim styleList As String
    
    ' Get first style attribute. If blank, ignore the others
    cellValue = Trim$(StyleDesignerSheet.Range(cellName1).value)
    
    If cellValue <> vbNullString Then
        ' Start building the style attribute
        styleList = cellValue
    
        ' Get the second style attribute
        cellValue = Trim$(StyleDesignerSheet.Range(cellName2).value)
        
        ' Add to set of styles if not blank
        If cellValue <> vbNullString Then
            styleList = styleList & COMMA & cellValue
    
            ' Get the third style attribute
            cellValue = Trim$(StyleDesignerSheet.Range(cellName3).value)
            
            ' Add to set of styles if not blank
            If cellValue <> vbNullString Then
                styleList = styleList & COMMA & cellValue
            End If
        End If
        
        ' If a fill color attribute was specified, a value of "filled" or "radial" must be included
        ' as one of the values in the 'style' attribute.
        If gradientType <> vbNullString Then
            styleList = styleList & COMMA & gradientType
        End If
    
        ' Close the double quotes around the style attributes
        If InStr(styleList, COMMA) Then
            styleAttributes = styleAttributes & " style=" & AddQuotes(styleList)
        Else
             styleAttributes = styleAttributes & " style=" & styleList
        End If
    
        ' Even though the style attributes are blank, we still need to return a style attribute if a
        ' fill color was specified elsewhere. gradientType will tell us if this is required.
    ElseIf gradientType <> vbNullString Then
        styleAttributes = styleAttributes & " style=" & gradientType
    End If

End Sub

' ==========================================================================
' PROCEDURE: AddFillColorAttribute
'
' PURPOSE:
'   Constructs a Graphviz 'fillcolor' attribute, supporting both solid
'   colors and complex linear gradient definitions.
'
' TECHNICAL WORKFLOW:
'   1. BASE VALIDATION: Resolves the primary fill color; if empty, the
'      procedure exits to avoid malformed attributes.
'   2. MODE DETECTION:
'      - SOLID: If no secondary color is provided, it applies a simple
'        quoted 'fillcolor'.
'      - GRADIENT: If a secondary color exists, it combines the colors
'        using the Graphviz colon (:) syntax.
'   3. WEIGHT CALCULATION: If a gradient weight is specified, it formats
'      the value as a decimal fraction (e.g., ";0.50") and appends it
'      to the primary color segment.
'   4. ANGLE INTEGRATION: Separately invokes 'AddAttribute' to handle
'      the 'gradientangle' setting if present in the designer.
'
' USAGE:
'   - Central logic for rendering advanced background styles in the
'     Style Designer.
' ==========================================================================
'@Ignore UseMeaningfulName
Public Sub AddFillColorAttribute(ByRef styleAttributes As String, _
                                      ByVal cellNameFillColor1 As String, _
                                      ByVal cellNameFillColor2 As String, _
                                      ByVal cellNameGradientAngle As String)
    Dim fillColor As String
    Dim gradientColor As String
       
    ' Get Fill Color first
    fillColor = Trim$(StyleDesignerSheet.Range(cellNameFillColor1).value)
    If fillColor = vbNullString Then
        Exit Sub
    End If
    
    ' Since we have a fill color, check for a gradient color
    gradientColor = Trim$(StyleDesignerSheet.Range(cellNameFillColor2).value)
    If gradientColor = vbNullString Then
        styleAttributes = styleAttributes & " fillcolor=" & AddQuotes(fillColor)
    Else
        Dim gradientWeight As String
        gradientWeight = Trim$(StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_WEIGHT).value)
        If gradientWeight <> vbNullString Then
            gradientWeight = ";0." & Right$("00" & gradientWeight, 2)
        End If
        fillColor = fillColor & gradientWeight & ":" & gradientColor
        
        ' Gradient angle attribute
        AddAttribute styleAttributes, GRAPHVIZ_GRADIENTANGLE, cellNameGradientAngle
        styleAttributes = styleAttributes & " fillcolor=" & AddQuotes(fillColor)
    End If
End Sub

' ==========================================================================
' FUNCTION: GetGradientType
'
' PURPOSE:
'   Determines the appropriate Graphviz 'style' attribute value (e.g., "filled")
'   based on the presence of color data in the Style Designer.
'
' TECHNICAL WORKFLOW:
'   1. DIRECT LOOKUP: Checks the specific 'Gradient Type' cell first; if
'      populated, this value takes precedence.
'   2. HEURISTIC FALLBACK: If the type cell is empty, the function
'      infers the need for a "filled" style by checking:
'      - Secondary Fill Color: If present, a gradient is implied.
'      - Primary Fill Color: If present, a solid fill is implied.
'   3. DEFAULTING: Returns 'vbNullString' only if no colors or styles are
'      defined, preventing unnecessary 'style=' attributes in the DOT source.
'
' USAGE:
'   - Helper function used to ensure the 'style' attribute is correctly
'     synchronized with the selected color inputs.
' ==========================================================================
'@Ignore UseMeaningfulName
Public Function GetGradientType(ByVal cellNameFillColor1 As String, _
                                ByVal cellNameFillColor2 As String, _
                                ByVal cellNameGradientType As String) As String

    GetGradientType = vbNullString
    
    ' Determine gradient type by process of elimination. First see if the gradient
    ' type cell has a value. If so, return that value
    Dim cellValue As String
    cellValue = Trim$(StyleDesignerSheet.Range(cellNameGradientType).value)
    
    If cellValue <> vbNullString Then
        GetGradientType = cellValue
    
        ' Gradient type cell is empty. If a gradient fill color has been specified, return "filled"
    ElseIf Trim$(StyleDesignerSheet.Range(cellNameFillColor2).value) <> vbNullString Then
        GetGradientType = GRAPHVIZ_STYLE_GRADIENT_FILLED
    
        ' Gradient type cell is empty. If a fill color has been specified, return "filled"
    ElseIf Trim$(StyleDesignerSheet.Range(cellNameFillColor1).value) <> vbNullString Then
        GetGradientType = GRAPHVIZ_STYLE_GRADIENT_FILLED
    End If
    
End Function

' ==========================================================================
' FUNCTION: ConvertMillimetersToInches
'
' PURPOSE:
'   Converts a metric measurement (mm) into the imperial units (inches)
'   required by the Graphviz engine for specific attributes like node size.
'
' TECHNICAL WORKFLOW:
'   1. DATA CONVERSION: Casts the input 'mm' string to a Double and
'      divides by 25.4.
'   2. PRECISION FORMATTING: Formats the result to four decimal places
'      using the '#0.0000' mask to ensure high-fidelity rendering.
'   3. RETURN: Passes back a string-formatted imperial value compatible
'      with DOT attribute syntax.
'
' USAGE:
'   - Vital for reconciling non-US metric-based measurements with the
'     Graphviz engine's unit expectations.
' ==========================================================================
Private Function ConvertMillimetersToInches(ByVal mm As String) As String
    Dim inches As Double
    inches = CDbl(mm) / 25.4
    ConvertMillimetersToInches = CStr(format(inches, "#0.0000"))
End Function

' ==========================================================================
' FUNCTION: GetNodeStyle
'
' PURPOSE:
'   Constructs a comprehensive Graphviz attribute string for a Node element
'   by aggregating settings from the Style Designer worksheet.
'
' TECHNICAL WORKFLOW:
'   1. SHAPE LOGIC: Maps basic shape settings and conditionally retrieves
'      polygon-specific attributes (sides, skew, distortion) if applicable.
'   2. UNIT NORMALIZATION: Checks for 'METRIC' toggle; if enabled, invokes
'      'ConvertMillimetersToInches' to scale height and width for Graphviz.
'   3. VISUAL LAYERING: Sequentially appends Color Schemes, Fill (including
'      gradients), Borders (penwidth, peripheries), and Typography settings.
'   4. ASSET MANAGEMENT: Incorporates image references, scaling, and
'      positioning if a node image is specified.
'   5. STYLE SYNC: Calls 'AddStyleAttribute' and 'GetGradientType' to ensure
'      border styles and fill types are logically consistent.
'   6. LABEL INTEGRATION: Conditionally appends 'label' and 'xlabel' text
'      only if the respective "Include" checkboxes are enabled in the UI.
'
' RETURN:
'   A trimmed String containing the full DOT attribute list for a Node.
' ==========================================================================
Private Function GetNodeStyle() As String
    
    Dim styleAttributes As String
    Dim cellValue As String
    
    ' Label attributes
    AddAttribute styleAttributes, GRAPHVIZ_LABELLOC, DESIGNER_LABEL_LOCATION
    
    ' Color Scheme
    AddAttribute styleAttributes, GRAPHVIZ_COLORSCHEME, DESIGNER_COLOR_SCHEME
    
    ' Shape attributes
    AddAttribute styleAttributes, GRAPHVIZ_SHAPE, DESIGNER_NODE_SHAPE

    ' If the shape is 'polygon', get the number of polygon sides
    If Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_SHAPE).value) = "polygon" Then
        AddAttribute styleAttributes, GRAPHVIZ_SIDES, DESIGNER_NODE_SIDES
        AddAttribute styleAttributes, GRAPHVIZ_SKEW, DESIGNER_NODE_SKEW
        AddAttribute styleAttributes, GRAPHVIZ_DISTORTION, DESIGNER_NODE_DISTORTION
        AddAttribute styleAttributes, GRAPHVIZ_REGULAR, DESIGNER_NODE_REGULAR
    End If
    
    ' If metric units were specified, they have to be converted to inches as that is what Graphviz expects
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).value = TOGGLE_YES Then
        cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_HEIGHT).value)
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & " height=" & ConvertMillimetersToInches(cellValue)
        End If

        cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_WIDTH).value)
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & " width=" & ConvertMillimetersToInches(cellValue)
        End If
    Else
        AddAttribute styleAttributes, GRAPHVIZ_HEIGHT, DESIGNER_NODE_HEIGHT
        AddAttribute styleAttributes, GRAPHVIZ_WIDTH, DESIGNER_NODE_WIDTH
    End If
    
    AddAttribute styleAttributes, GRAPHVIZ_FIXEDSIZE, DESIGNER_NODE_FIXED_SIZE
    AddAttribute styleAttributes, GRAPHVIZ_ORIENTATION, DESIGNER_NODE_ORIENTATION
    
    ' Fill Color attributes
    AddFillColorAttribute styleAttributes, DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_ANGLE
    
    ' Border attributes
    AddAttribute styleAttributes, GRAPHVIZ_COLOR, DESIGNER_BORDER_COLOR
    AddAttribute styleAttributes, GRAPHVIZ_PENWIDTH, DESIGNER_BORDER_PEN_WIDTH
    AddAttribute styleAttributes, GRAPHVIZ_PERIPHERIES, DESIGNER_BORDER_PERIPHERIES
 
    ' Font attributes
    AddFontNameAttribute styleAttributes
    AddAttribute styleAttributes, GRAPHVIZ_FONTSIZE, DESIGNER_FONT_SIZE
    AddAttribute styleAttributes, GRAPHVIZ_FONTCOLOR, DESIGNER_FONT_COLOR
      
    ' Image attributes
    
    cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME).value)
    If cellValue <> vbNullString Then
        styleAttributes = styleAttributes & " image=" & AddQuotes(cellValue)
    End If

    'AddAttribute styleAttributes, "image", DESIGNER_NODE_IMAGE_NAME
    AddAttribute styleAttributes, GRAPHVIZ_IMAGESCALE, DESIGNER_NODE_IMAGE_SCALE
    AddAttribute styleAttributes, GRAPHVIZ_IMAGEPOS, DESIGNER_NODE_IMAGE_POSITION
    
    ' Style attributes
    AddStyleAttribute styleAttributes, DESIGNER_BORDER_STYLE1, DESIGNER_BORDER_STYLE2, DESIGNER_BORDER_STYLE3, _
                      GetGradientType(DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_TYPE)
    
    ' Include the label in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_LABEL, StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT).value)
    End If
    
    ' Include the xlabel in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_XLABEL, StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).value)
    End If
    
    ' Return the finished string of style attributes
    GetNodeStyle = Trim$(styleAttributes)
    
End Function

' ==========================================================================
' FUNCTION: GetFontStyle
'
' PURPOSE:
'   Constructs a composite font name string that includes weight and slant
'   modifiers (Bold/Italic) based on Style Designer toggles.
'
' TECHNICAL WORKFLOW:
'   1. BASE RETRIEVAL: Initializes the string with the primary font family
'      name from the 'DESIGNER_FONT_NAME' range.
'   2. WEIGHT MODIFICATION: Appends a " Bold" suffix if the specified
'      'boldCell' toggle is set to 'TOGGLE_YES'.
'   3. SLANT MODIFICATION: Appends an " Italic" suffix if the specified
'      'italicCell' toggle is set to 'TOGGLE_YES'.
'   4. NORMALIZATION: Returns a trimmed string to ensure proper syntax
'      for the Graphviz 'fontname' attribute.
'
' USAGE:
'   - Helper function used by 'AddFontNameAttribute' to create valid
'     postscript-style font references.
' ==========================================================================
Private Function GetFontStyle(ByVal boldCell As String, ByVal italicCell As String) As String

    Dim fontStyle As String
    fontStyle = StyleDesignerSheet.Range(DESIGNER_FONT_NAME).value
    
    If StyleDesignerSheet.Range(boldCell).value = TOGGLE_YES Then
        fontStyle = fontStyle & " Bold"
    End If
    
    If StyleDesignerSheet.Range(italicCell).value = TOGGLE_YES Then
        fontStyle = fontStyle & " Italic"
    End If
    
    GetFontStyle = Trim$(fontStyle)

End Function

' ==========================================================================
' FUNCTION: GetEdgeStyle
'
' PURPOSE:
'   Synthesizes a complete Graphviz DOT attribute string for an Edge element
'   by harvesting configuration data from the Style Designer.
'
' TECHNICAL WORKFLOW:
'   1. STROKE & COLOR: Maps core aesthetics including color schemes, pen
'      width, and multi-color edge gradients (via 'AddAttributeGroup').
'   2. STRUCTURAL GEOMETRY: Configures edge direction, weight, and
'      logically validates 'radius'—ensuring it is numeric and greater
'      than zero before inclusion.
'   3. ARROWHEAD DYNAMICS: Orchestrates complex arrow configurations for
'      both head and tail, allowing for multi-part arrow shapes.
'   4. LABEL TOPOGRAPHY: Sets global edge labels and specific modifiers
'      like 'decorate', 'labelangle', and 'labeldistance'.
'   5. TYPOGRAPHY: Applies font settings for the main label and secondary
'      head/tail labels.
'   6. PORT & CLIPPING: Defines 'headport', 'tailport', and clipping
'      behaviors to control exactly where edges attach to nodes.
'   7. UI SYNC: Conditionally appends Label, XLabel, HeadLabel, and
'      TailLabel text based on user "Include" toggles.
'
' RETURN:
'   A String containing the serialized DOT attributes for an Edge.
' ==========================================================================
Private Function GetEdgeStyle() As String

    Dim styleAttributes As String
    
    styleAttributes = vbNullString

    GetEdgeStyle = styleAttributes
    
    ' Color Scheme
    AddAttribute styleAttributes, GRAPHVIZ_COLORSCHEME, DESIGNER_COLOR_SCHEME
    
    ' Style attributes
    AddAttribute styleAttributes, GRAPHVIZ_STYLE, DESIGNER_EDGE_STYLE
    AddAttributeGroup styleAttributes, GRAPHVIZ_COLOR, DESIGNER_EDGE_COLOR_1, DESIGNER_EDGE_COLOR_2, DESIGNER_EDGE_COLOR_3, ":"
    AddAttribute styleAttributes, GRAPHVIZ_PENWIDTH, DESIGNER_EDGE_PEN_WIDTH
    AddAttribute styleAttributes, GRAPHVIZ_DIR, DESIGNER_EDGE_DIRECTION
    AddAttribute styleAttributes, GRAPHVIZ_WEIGHT, DESIGNER_EDGE_WEIGHT
    
    ' Radius attribute is managed as a number, not a string
    Dim radius As Long
    Dim rawValue As Variant
    
    rawValue = StyleDesignerSheet.Range(DESIGNER_EDGE_RADIUS).value
    
    If IsNumeric(rawValue) Then
        radius = CLng(rawValue)
    Else
        radius = 0
    End If
    
    If radius > 0 Then
        AddAttribute styleAttributes, GRAPHVIZ_RADIUS, DESIGNER_EDGE_RADIUS
    End If

    ' Label attributes
    AddAttribute styleAttributes, GRAPHVIZ_DECORATE, DESIGNER_EDGE_DECORATE
    AddAttribute styleAttributes, GRAPHVIZ_LABELANGLE, DESIGNER_EDGE_LABEL_ANGLE
    AddAttribute styleAttributes, GRAPHVIZ_LABELFLOAT, DESIGNER_EDGE_LABEL_FLOAT
    AddAttribute styleAttributes, GRAPHVIZ_LABELDISTANCE, DESIGNER_EDGE_LABEL_DISTANCE

    ' Arrow attributes
    AddAttributeGroup styleAttributes, GRAPHVIZ_ARROWHEAD, DESIGNER_EDGE_ARROW_HEAD_1, DESIGNER_EDGE_ARROW_HEAD_2, DESIGNER_EDGE_ARROW_HEAD_3, vbNullString
    AddAttributeGroup styleAttributes, GRAPHVIZ_ARROWTAIL, DESIGNER_EDGE_ARROW_TAIL_1, DESIGNER_EDGE_ARROW_TAIL_2, DESIGNER_EDGE_ARROW_TAIL_3, vbNullString
    AddAttribute styleAttributes, GRAPHVIZ_ARROWSIZE, DESIGNER_EDGE_ARROW_SIZE

    ' Font attributes
    AddFontNameAttribute styleAttributes
    AddAttribute styleAttributes, GRAPHVIZ_FONTSIZE, DESIGNER_FONT_SIZE
    AddAttribute styleAttributes, GRAPHVIZ_FONTCOLOR, DESIGNER_FONT_COLOR
    
    ' Head/Tail label attributes
    AddAttribute styleAttributes, GRAPHVIZ_LABELFONTNAME, DESIGNER_EDGE_LABEL_FONT_NAME
    AddAttribute styleAttributes, GRAPHVIZ_LABELFONTSIZE, DESIGNER_EDGE_LABEL_FONT_SIZE
    AddAttribute styleAttributes, GRAPHVIZ_LABELFONTCOLOR, DESIGNER_EDGE_LABEL_FONT_COLOR
    
    ' Port attributes
    AddAttribute styleAttributes, GRAPHVIZ_HEADPORT, DESIGNER_EDGE_HEAD_PORT
    AddAttribute styleAttributes, GRAPHVIZ_TAILPORT, DESIGNER_EDGE_TAIL_PORT
    
    ' Clip attributes
    AddAttribute styleAttributes, GRAPHVIZ_HEADCLIP, DESIGNER_EDGE_HEAD_CLIP
    AddAttribute styleAttributes, GRAPHVIZ_TAILCLIP, DESIGNER_EDGE_TAIL_CLIP
   
    ' Include the label in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_LABEL, StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT).value)
    End If
    
    ' Include the xlabel in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_XLABEL, StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).value)
    End If
    
    ' Include the taillabel in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_TAIL_LABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_TAILLABEL, StyleDesignerSheet.Range(DESIGNER_TAIL_LABEL_TEXT).value)
    End If
    
    ' Include the xlabel in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_HEAD_LABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_HEADLABEL, StyleDesignerSheet.Range(DESIGNER_HEAD_LABEL_TEXT).value)
    End If
    
    GetEdgeStyle = Trim$(styleAttributes)
    
End Function

' ==========================================================================
' PROCEDURE: AddFontNameAttribute
'
' PURPOSE:
'   Appends a formatted Graphviz 'fontname' attribute to the style string,
'   incorporating weight and slant modifiers.
'
' TECHNICAL WORKFLOW:
'   1. STYLE RESOLUTION: Calls 'GetFontStyle' to retrieve the base font
'      name combined with "Bold" or "Italic" suffixes as needed.
'   2. VALIDATION: Exits immediately if the resulting font name is empty.
'   3. STRING COMPOSITION: Concatenates the 'fontname=' key with the
'      escaped font string (via 'AddQuotesConditionally') onto the existing
'      'styleAttributes' variable.
'
' USAGE:
'   - Internal utility used by 'GetNodeStyle', 'GetEdgeStyle', and
'     'GetClusterStyle' to ensure consistent font attribute syntax.
' ==========================================================================
Private Sub AddFontNameAttribute(ByRef styleAttributes As String)
    Dim fontName As String
    fontName = GetFontStyle(DESIGNER_FONT_BOLD, DESIGNER_FONT_ITALIC)
    If fontName = vbNullString Then Exit Sub
    
    styleAttributes = Join(Array(styleAttributes, " fontname=", AddQuotesConditionally(fontName)), vbNullString)
End Sub

' ==========================================================================
' FUNCTION: GetClusterStyle
'
' PURPOSE:
'   Constructs a Graphviz attribute string specifically for Cluster
'   (subgraph) elements, integrating layout-specific logic for the
'   Osage engine.
'
' TECHNICAL WORKFLOW:
'   1. UI MAPPING: Aggregates standard label justification, location,
'      color schemes, and border properties from the Style Designer.
'   2. FILL & GRADIENTS: Invokes 'AddFillColorAttribute' to handle background
'      coloring, including advanced linear gradients.
'   3. TYPOGRAPHY: Sets font family, size, and color for the cluster label.
'   4. OSAGE ENGINE LOGIC:
'      - Implements 'pack' and 'packmode' attributes if Osage is active.
'      - Complex Array Parsing: If 'packmode' is set to "array", it
'        concatenates modifiers (major, align, justify, sort, split)
'        into a single "array_..." string.
'   5. LABEL SYNC: Conditionally appends the cluster label based on the
'      UI's "Include" checkbox.
'
' RETURN:
'   A String containing the serialized DOT attributes for a Cluster.
' ==========================================================================
Private Function GetClusterStyle() As String

    Dim styleAttributes As String
    
    styleAttributes = vbNullString

    GetClusterStyle = styleAttributes
    
    ' Label attributes
    AddAttribute styleAttributes, GRAPHVIZ_LABELJUST, DESIGNER_LABEL_JUSTIFICATION
    AddAttribute styleAttributes, GRAPHVIZ_LABELLOC, DESIGNER_LABEL_LOCATION
    
    ' Color scheme
    AddAttribute styleAttributes, GRAPHVIZ_COLORSCHEME, DESIGNER_COLOR_SCHEME
    
    ' Border attributes
    AddAttribute styleAttributes, GRAPHVIZ_PENWIDTH, DESIGNER_BORDER_PEN_WIDTH
    AddAttribute styleAttributes, GRAPHVIZ_PENCOLOR, DESIGNER_BORDER_COLOR
   
    ' Fill and Gradient Color attributes
    AddFillColorAttribute styleAttributes, DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_ANGLE

    ' Font attributes
    AddFontNameAttribute styleAttributes
    AddAttribute styleAttributes, GRAPHVIZ_FONTSIZE, DESIGNER_FONT_SIZE
    AddAttribute styleAttributes, GRAPHVIZ_FONTCOLOR, DESIGNER_FONT_COLOR
        
    ' Style attributes
    AddStyleAttribute styleAttributes, DESIGNER_BORDER_STYLE1, DESIGNER_BORDER_STYLE2, DESIGNER_BORDER_STYLE3, _
                      GetGradientType(DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_TYPE)
    
    If SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).value = LAYOUT_OSAGE Then
        ' Pack attribute
        If Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_MARGIN).value) <> vbNullString Then
            AddAttribute styleAttributes, GRAPHVIZ_PACK, DESIGNER_CLUSTER_MARGIN
        End If
        
        ' Packmode attribute
        Dim packmode As String
        packmode = Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_PACKMODE).value)
        
        If LCase$(packmode) = GRAPHVIZ_PACKMODE_ARRAY Then
            Dim major As String
            major = Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_MAJOR).value)
            
            Dim split As String
            split = Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_SPLIT).value)
            
            Dim align As String
            align = Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_ALIGN).value)
            
            Dim justify As String
            justify = Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_JUSTIFY).value)
            
            Dim sort As String
            sort = Trim$(StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_SORT).value)
            If LCase$(sort) = TOGGLE_YES Then
                sort = GRAPHVIZ_PACKMODE_SORT
            Else
                sort = vbNullString
            End If
            
            Dim modifiers As String
            modifiers = major & align & justify & sort & split
            
            If modifiers <> vbNullString Then
                styleAttributes = styleAttributes & " packmode=array_" & modifiers
            Else
                AddAttribute styleAttributes, GRAPHVIZ_PACKMODE, DESIGNER_CLUSTER_PACKMODE
            End If
        Else
            AddAttribute styleAttributes, GRAPHVIZ_PACKMODE, DESIGNER_CLUSTER_PACKMODE
        End If
    End If
    
    ' Include the label in the Format String if checked
    If StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT_INCLUDE).value = True Then
        styleAttributes = styleAttributes & FormatLabel(GRAPHVIZ_LABEL, StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT).value)
    End If
    
    GetClusterStyle = Trim$(styleAttributes)
    
End Function

' ==========================================================================
' PROCEDURE: DisplayStyleDesignerRows
'
' PURPOSE:
'   Controls the visibility of the primary input range within the Style
'   Designer worksheet, allowing for a "clean" interface toggle.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY RESOLUTION: Calculates the dynamic row range using
'      the 'GRAPHVIZ_COLORSCHEME' constant as a top anchor and the
'      'AddCaption' range as a bottom anchor (with buffers).
'   2. BULK MODIFICATION: Iterates through the calculated row indices.
'   3. STATE APPLICATION: Sets the '.Hidden' property to the inverse
'      of the 'isVisible' parameter to show or hide the designer controls.
'
' USAGE:
'   - Used to minimize the designer UI when not in use or during specific
'     worksheet transitions.
' ==========================================================================
Public Sub DisplayStyleDesignerRows(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    Dim row As Long
    
    rowFrom = StyleDesignerSheet.Range(GRAPHVIZ_COLORSCHEME).row - 5
    rowTo = StyleDesignerSheet.Range("AddCaption").row + 3
    
    For row = rowFrom To rowTo
        StyleDesignerSheet.rows.item(row).Hidden = Not isVisible
    Next row
End Sub

' ==========================================================================
' PROCEDURE: StyleDesignerToggleShowSettings
'
' PURPOSE:
'   Acts as the event handler for the "Show/Hide Settings" checkbox (Form Control)
'   on the Style Designer worksheet.
'
' TECHNICAL WORKFLOW:
'   1. SHAPE RESOLUTION: Accesses the specific checkbox shape named
'      "ToggleStyleDesignerSettings" using the default item accessor.
'   2. STATE EVALUATION: Checks the 'Selection.value' of the control.
'   3. BRANCHED EXECUTION:
'      - If checked (xlOn): Invokes 'ShowStyleDesignerSettings' to expand the UI.
'      - If unchecked: Invokes 'HideStyleDesignerSettings' to collapse the UI.
'   4. OBJECT CLEANUP: Releases the shape object reference.
'
' USAGE:
'   - Assigned to the checkbox macro on the StyleDesignerSheet.
'   - Provides a user-friendly way to toggle advanced Graphviz configuration rows.
' ==========================================================================
'@Ignore ProcedureNotUsed
Public Sub StyleDesignerToggleShowSettings()

    Dim s As shape
    Set s = StyleDesignerSheet.Shapes.[_Default]("ToggleStyleDesignerSettings")
 
    s.Select
    
    If Selection.value = xlOn Then
       ShowStyleDesignerSettings
    Else
       HideStyleDesignerSettings
    End If

    Set s = Nothing
End Sub

' ==========================================================================
' PROCEDURE: HideStyleDesignerSettings
'
' PURPOSE:
'   Collapses the Style Designer's input interface to provide a streamlined,
'   non-intrusive view.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to prevent visual flicker
'      during row manipulation.
'   2. UI COLLAPSE: Invokes 'DisplayStyleDesignerRows' with a FALSE
'      parameter to hide the configuration range.
'   3. FOCUS MANAGEMENT: Shifts the active selection to the 'DESIGNER_FORMAT_STRING'
'      range, ensuring the user lands on a relevant, visible cell.
'   4. REFRESH: Re-enables 'ScreenUpdating' to commit the visual changes.
' ==========================================================================
Private Sub HideStyleDesignerSettings()
    Application.ScreenUpdating = False
    DisplayStyleDesignerRows False
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: ShowStyleDesignerSettings
'
' PURPOSE:
'   Expands the Style Designer's input interface, making all Graphviz
'   configuration and attribute rows visible to the user.
'
' TECHNICAL WORKFLOW:
'   1. UI STABILIZATION: Disables 'ScreenUpdating' to ensure a smooth
'      transition while multiple rows are unhidden.
'   2. UI EXPANSION: Invokes 'DisplayStyleDesignerRows' with a TRUE
'      parameter to reveal the configuration range.
'   3. FOCUS MANAGEMENT: Selects the 'DESIGNER_FORMAT_STRING' range to
'      orient the user toward the primary output field.
'   4. REFRESH: Re-enables 'ScreenUpdating' to display the updated
'      worksheet layout.
' ==========================================================================
Private Sub ShowStyleDesignerSettings()
    Application.ScreenUpdating = False
    DisplayStyleDesignerRows True
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
    Application.ScreenUpdating = True
End Sub

' ==========================================================================
' PROCEDURE: SaveToStylesWorksheet
'
' PURPOSE:
'   Persists the current Style Designer configuration into the global
'   Style Gallery, handling both updates to existing styles and the
'   intelligent creation of new style entries.
'
' TECHNICAL WORKFLOW:
'   1. ENVIRONMENT PREPARATION: Ensures the 'Styles' sheet is visible and
'      retrieves worksheet schema (column indices) via 'GetSettingsForStylesWorksheet'.
'   2. IDENTITY RESOLUTION: Captures the 'styleName' and 'styleType'; if the name
'      is blank, it invokes 'CreateStyleName' to generate a unique identifier.
'   3. LOCATION DISCOVERY: Searches for an existing row matching the style
'      name. If not found, it triggers an 'insertRow' logic to find the next
'      available data row.
'   4. DATA COMMIT: Writes the DOT format string and object type to the
'      target row and applies visual defaults via 'SetStyleViewDefaults'.
'   5. CLUSTER HANDLING: If 'DESIGNER_MODE' is 'CLUSTER', it automatically
'      manages the "Open/Close" subgraph row pair to maintain gallery integrity.
'   6. UI SYNCHRONIZATION: Activates the 'Styles' sheet, focuses on the saved
'      row, and triggers 'GenerateStylesPreview' to render the new thumbnail.
'
' USAGE:
'   - Linked to the "Save Style" button in the Designer UI.
'   - Crucial for transitioning from a draft design to a reusable gallery asset.
' ==========================================================================
Public Sub SaveToStylesWorksheet()
    Dim row As Long
    Dim rowFocus As Long
    Dim col As Long
    Dim styleName As String
    Dim styleType As String
    
    Dim insertRow As Boolean
    insertRow = False
    
    ' Unhide the styles sheet if hidden
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLES).value = TOGGLE_HIDE Then
        SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLES).value = TOGGLE_SHOW
    End If
    
    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    ' Determine which type of style to add
    styleType = GetStyleDesignerStyleType()
    
    ' Establish a style name, either user-specified, or generate one
    styleName = Trim$(StyleDesignerSheet.Range("StyleNameText").value)
    If styleName = vbNullString Then
        styleName = CreateStyleName(styles)
    End If
    
    ' Find the row where the style should be saved
    row = GetStyleRowForSave(styleName, styles)
    If row = 0 And styleType = TYPE_SUBGRAPH_OPEN Then
        row = GetStyleRowForSave(styleName & " " & styles.suffixOpen, styles)
    End If
    
    ' Style does not exist, insert a new one
    If row = 0 Then
        insertRow = True
        row = GetStyleRowForInsert(styles)
    End If
    
    ' Store the format from the Style Designer
    StylesSheet.Cells.item(row, styles.formatColumn).value = StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).value
    
    ' Save the row number so we know where to place the focus if the DESIGNER_MODE = CLUSTER
    rowFocus = row
    
    If insertRow Then
        ' Set the format string and the object type
        StylesSheet.Cells.item(row, styles.nameColumn).value = styleName
        StylesSheet.Cells.item(row, styles.formatColumn).value = StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).value
        StylesSheet.Cells.item(row, styles.typeColumn).value = styleType
        
        ' Add default values for the view columns
        SetStyleViewDefaults row, styles
        
        ' If the style is CLUSTER we want to add a row for the subgraph-close, as it improves filtering capabilities
        If StyleDesignerSheet.Range(DESIGNER_MODE).value = KEYWORD_CLUSTER Then
            If EndsWith(styleName, styles.suffixOpen) Then
                styleName = Left(styleName, Len(styleName) - Len(styles.suffixOpen) - 1)
            End If
            StylesSheet.Cells.item(row, styles.nameColumn).value = styleName & " " & styles.suffixOpen
         
            ' Last row information changed if a new style was appended
            styles = GetSettingsForStylesWorksheet()
            
            ' Look for a row that does not have a style name
            row = GetStyleRowForInsert(styles)
    
            ' Set the format string and the object type
            StylesSheet.Cells.item(row, styles.nameColumn).value = styleName & " " & styles.suffixClose
            StylesSheet.Cells.item(row, styles.formatColumn).value = vbNullString
            StylesSheet.Cells.item(row, styles.typeColumn).value = TYPE_SUBGRAPH_CLOSE
            
            ' Add default values for the view columns
            SetStyleViewDefaults row, styles
        End If
    End If
    
    ' Put the focus on the cell where the style name has to be entered
    StylesSheet.Activate
    ActiveSheet.Cells(rowFocus, styles.nameColumn).Select
    
    ' Generate a preview image on the styles worksheet
    GenerateStylesPreview rowFocus
End Sub

' ==========================================================================
' PROCEDURE: SetStyleViewDefaults
'
' PURPOSE:
'   Initializes a new style row by enabling all available view/filter
'   columns with a default "yes" status.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN SCANNING: Iterates through the 'StylesSheet' starting from
'      the 'firstYesNoColumn' defined in the worksheet schema.
'   2. DYNAMIC BOUNDARY: Uses 'GetLastColumn' to determine the end of
'      the header row, ensuring the loop covers all configured views.
'   3. TERMINATION LOGIC: Stops processing if it encounters a null header
'      to prevent writing data into unmanaged worksheet space.
'   4. STATE ASSIGNMENT: Populates every valid view cell in the target
'      'row' with the 'TOGGLE_YES' constant.
'
' USAGE:
'   - Internal helper for 'SaveToStylesWorksheet'.
'   - Ensures new styles are visible in all "Views" by default.
' ==========================================================================
Private Sub SetStyleViewDefaults(ByVal row As Long, ByRef styles As stylesWorksheet)
    ' Loop through the columns which have column headings and put a value of 'yes' in the cell
    Dim moreViews As Boolean
    moreViews = True
    
    Dim col As Long
    For col = styles.firstYesNoColumn To GetLastColumn(StylesSheet.name, styles.headingRow)
        ' Stop when the first null column is encountered
        If StylesSheet.Cells.item(styles.headingRow, col) = vbNullString Then
            moreViews = False
        End If
        
        ' Add a 'yes' value to a view column
        If moreViews Then
            StylesSheet.Cells.item(row, col).value = TOGGLE_YES
        End If
    Next col
End Sub

' ==========================================================================
' FUNCTION: GetStyleRowForSave
'
' PURPOSE:
'   Performs a linear search of the Styles worksheet to locate an existing
'   style entry by its name.
'
' TECHNICAL WORKFLOW:
'   1. ITERATION: Loops through the worksheet rows starting from the schema's
'      'firstRow' through to the 'lastRow'.
'   2. COMPARISON: Checks the value in the 'nameColumn' against the
'      provided 'styleName' parameter.
'   3. EARLY EXIT: Terminates the loop and returns the current row index
'      immediately upon finding a match.
'   4. NULL RESULT: Returns 0 if no matching style name is found in the
'      defined range.
'
' USAGE:
'   - Called by 'SaveToStylesWorksheet' to determine if an operation
'     should be an "Update" (existing row) or an "Insert" (new row).
' ==========================================================================
Private Function GetStyleRowForSave(ByVal styleName As String, ByRef styles As stylesWorksheet) As Long
    Dim styleRow As Long
    styleRow = 0
    
    ' Look for a row which matches the style name
    Dim row As Long
    For row = styles.firstRow To styles.lastRow
        If StylesSheet.Cells.item(row, styles.nameColumn).value = styleName Then
            styleRow = row
            Exit For
        End If
    Next row
    GetStyleRowForSave = styleRow
End Function

' ==========================================================================
' FUNCTION: GetStyleRowForInsert
'
' PURPOSE:
'   Identifies the first available empty slot in the Styles worksheet
'   to append a new style definition.
'
' TECHNICAL WORKFLOW:
'   1. LINEAR SCAN: Iterates through the data range defined by the
'      'styles' worksheet schema (firstRow to lastRow).
'   2. AVAILABILITY CRITERIA: Validates each row against two conditions:
'      - The 'flagColumn' must not contain a comment indicator (FLAG_COMMENT).
'      - The 'nameColumn' must be empty (vbNullString).
'   3. INDEX RETURN: Exits the loop and returns the index of the first
'      compliant row found.
'
' USAGE:
'   - Internal utility for 'SaveToStylesWorksheet' to ensure new entries
'     do not overwrite existing styles or system comments.
' ==========================================================================
Private Function GetStyleRowForInsert(ByRef styles As stylesWorksheet) As Long
    ' Look for a row that does not have a style name
    Dim row As Long
    For row = styles.firstRow To styles.lastRow
        If StylesSheet.Cells.item(row, styles.flagColumn) <> FLAG_COMMENT And _
           StylesSheet.Cells.item(row, styles.nameColumn).value = vbNullString Then
            Exit For
        End If
    Next row
    GetStyleRowForInsert = row
End Function

' ==========================================================================
' FUNCTION: CreateStyleName
'
' PURPOSE:
'   Generates a unique, localized default name for a new style when the
'   user has not provided one in the Style Designer.
'
' TECHNICAL WORKFLOW:
'   1. TYPE IDENTIFICATION: Resolves the active designer mode (Node, Edge,
'      or Cluster) to determine the naming category.
'   2. INDEX CALCULATION: Invokes 'GetStyleCount' to determine how many
'      styles of that specific type already exist, then increments by 1.
'   3. LOCALIZED ASSEMBLY: Combines a localized prefix (retrieved via
'      'GetLabel') with the new index number (e.g., "Node 5").
'   4. BRANCHED LOGIC: Uses a 'Select Case' on 'DESIGNER_MODE' to ensure
'      the correct prefix key is requested for the current context.
'
' USAGE:
'   - Helper function for 'SaveToStylesWorksheet' to prevent null
'     name errors during the save process.
' ==========================================================================
Private Function CreateStyleName(ByRef styles As stylesWorksheet) As String
    Dim styleType As String
    styleType = GetStyleDesignerStyleType()
    
    ' Increment the count to reflect the style we are adding
    Dim objectCount As Long
    objectCount = GetStyleCount(styleType, styles) + 1
    
    ' Create default style name
    Select Case StyleDesignerSheet.Range(DESIGNER_MODE).value
        Case KEYWORD_NODE
            CreateStyleName = GetLabel("SaveStyleNode") & " " & objectCount
        Case KEYWORD_EDGE
            CreateStyleName = GetLabel("SaveStyleEdge") & " " & objectCount
        Case KEYWORD_CLUSTER
            CreateStyleName = GetLabel("SaveStyleCluster") & " " & objectCount
    End Select
End Function

' ==========================================================================
' FUNCTION: GetStyleDesignerStyleType
'
' PURPOSE:
'   Maps the user-friendly "Design Mode" dropdown selection to the internal
'   Graphviz object type constants used for data storage and rendering.
'
' TECHNICAL WORKFLOW:
'   1. UI EVALUATION: Reads the active selection from the 'DESIGNER_MODE'
'      range on the 'StyleDesignerSheet'.
'   2. CONSTANT MAPPING: Translates high-level keywords (Node, Edge, Cluster)
'      into specific internal types:
'      - KEYWORD_NODE    -> TYPE_NODE
'      - KEYWORD_EDGE    -> TYPE_EDGE
'      - KEYWORD_CLUSTER -> TYPE_SUBGRAPH_OPEN
'
' RETURN:
'   A Variant (typically String) containing the internal type constant.
' ==========================================================================
Private Function GetStyleDesignerStyleType()
    ' Map the 'Design Mode' dropdown value to the Object Type
    Select Case StyleDesignerSheet.Range(DESIGNER_MODE).value
        Case KEYWORD_NODE
            GetStyleDesignerStyleType = TYPE_NODE
        Case KEYWORD_EDGE
            GetStyleDesignerStyleType = TYPE_EDGE
        Case KEYWORD_CLUSTER
            GetStyleDesignerStyleType = TYPE_SUBGRAPH_OPEN
    End Select
End Function

' ==========================================================================
' FUNCTION: GetStyleCount
'
' PURPOSE:
'   Calculates the total number of existing style definitions for a specific
'   object type within the Styles worksheet.
'
' TECHNICAL WORKFLOW:
'   1. ITERATION: Loops through the defined style range (firstRow to lastRow)
'      using the provided worksheet schema.
'   2. TYPE MATCHING: Compares the value in the 'typeColumn' of each row
'      against the target 'styleType' (e.g., TYPE_NODE).
'   3. ACCUMULATION: Increments a local counter for every successful match
'      found in the gallery.
'   4. RETURN: Provides the final tally to the calling function.
'
' USAGE:
'   - Primary dependency for 'CreateStyleName' to ensure auto-generated
'     names carry the correct sequential index.
' ==========================================================================
Private Function GetStyleCount(ByVal styleType As String, ByRef styles As stylesWorksheet) As Long
    Dim row As Long
    Dim styleCount As Long
    
    styleCount = 0
    
    For row = styles.firstRow To styles.lastRow
        If StylesSheet.Cells.item(row, styles.typeColumn).value = styleType Then
            styleCount = styleCount + 1
        End If
    Next row

    GetStyleCount = styleCount
End Function


