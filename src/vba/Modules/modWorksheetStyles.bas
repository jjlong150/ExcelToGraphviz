Attribute VB_Name = "modWorksheetStyles"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetStyles
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Styles
'
' ROLE:
'   Worksheet-to-Graphviz bridge for the Style Gallery. Generates live
'   thumbnails for Node, Edge, and Cluster styles, manages worksheet
'   geometry, and enables round-trip editing between saved styles and the
'   interactive Style Designer.
'
' RESPONSIBILITIES:
'   - Visual rendering:
'       • Convert DOT format strings into PNG thumbnails via Graphviz
'       • Handle per-row preview generation and full-gallery batch refresh
'
'   - Worksheet layout:
'       • Auto-size rows to fit rendered images
'       • Compute preview-column placement dynamically
'
'   - Round-trip editing:
'       • Restore saved styles back into the Style Designer
'       • Parse DOT attribute strings into Designer UI fields
'
'   - State management:
'       • Clear previews, purge images, and maintain workbook responsiveness
'       • Display progress indicators during bulk rendering
'
' ARCHITECTURAL NOTES:
'   - Uses the Graphviz class for rendering and console capture
'   - Uses WIA (Windows Image Acquisition) for pixel-to-point scaling on Windows
'   - macOS uses a fixed-height fallback due to sandboxed image metadata access
'   - Designed for high-volume rendering with DoEvents-based responsiveness
'
' USAGE:
'   - Invoked by the Styles Ribbon tab (Preview, Preview All, Restore)
'   - Used to maintain a visual catalog of reusable Graphviz styles
'
' RELATED WIKI PAGES:
'   - Style Gallery Overview
'   - Rendering Pipeline (Graphviz Integration)
'   - Restoring Styles into the Style Designer
' =============================================================================


Option Explicit

' ==========================================================================
' PROCEDURE: GenerateStylesPreviewAll
' @Ignore ProcedureNotUsed, ParameterNotUsed
'
' PURPOSE:
'   Orchestrates a full-scale visual refresh of the entire Style Gallery.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Resolves the physical boundaries of the 'Styles'
'      sheet using 'GetSettingsForStylesWorksheet'.
'   2. PROGRESS TRACKING: Calculates 'styleCount' to drive a localized
'      percentage-based progress indicator in the Excel StatusBar.
'   3. ITERATIVE RENDERING: Loops through every row, delegating the
'      actual image creation to 'GenerateStylesPreview'.
'   4. UI RESPONSIVENESS: Employs 'DoEvents' to prevent the application
'      from appearing frozen during high-volume Graphviz calls.
'
' USAGE:
'   - Triggered by the "Preview All" button on the Styles Ribbon tab.
'   - Vital for reconciling visual thumbnails after a global style edit
'     or a change in the Graphviz engine settings.
' ==========================================================================
Public Sub GenerateStylesPreviewAll()
    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    Dim styleCount As Long
    styleCount = styles.lastRow - styles.firstRow + 1
    
    Dim statusMsg As String
    statusMsg = GetLabel("stylesProgressIndicator")
    ' Loop through the rows, generating preview images from the format strings
    Dim row As Long
    For row = styles.firstRow To styles.lastRow
        Application.StatusBar = statusMsg & " " & format(((row - 1) * 100) / styleCount, "0") & "%"
        GenerateStylesPreview row
        DoEvents
    Next row
End Sub

' ==========================================================================
' PROCEDURE: GenerateStylesPreview
' @Ignore ProcedureNotUsed, ParameterNotUsed
'
' PURPOSE:
'   Translates a specific style definition into a high-fidelity visual preview.
'
' TECHNICAL WORKFLOW:
'   1. VALIDATION: Skips rows marked as comments (#) or rows lacking a style
'      name to ensure gallery integrity.
'   2. DYNAMIC MAPPING: Calculates the 'previewColumn' by finding the end
'      of the row and appending a buffer, ensuring the image doesn't
'      overlap the configuration data.
'   3. DOT TEMPLATING: Constructs a specialized Graphviz source string
'      based on the 'styleType' (Node, Edge, or Subgraph):
'      - NODES: Renders a single node using its name as the ID.
'      - EDGES: Renders a horizontal rank (LR) connection between invisible points.
'      - SUBGRAPHS: Renders a cluster container containing dummy nodes (A->Z).
'   4. SMART LABELING: Checks for existing 'label=' attributes; if missing,
'      it automatically injects the style name as a label to provide context.
'   5. IMAGE PIPELINE: Calls 'PreviewStyleAndAutosize' to execute the
'      render and physically place the image on the worksheet.
' ==========================================================================
Public Sub GenerateStylesPreview(ByRef row As Long)

    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    If StylesSheet.Cells.item(row, styles.flagColumn) = FLAG_COMMENT Then
        Exit Sub
    End If
    
    If StylesSheet.Cells.item(row, styles.nameColumn).value = vbNullString Then
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
    styleName = StylesSheet.Cells.item(row, styles.nameColumn).value
    
    Dim styleType As String
    styleType = StylesSheet.Cells.item(row, styles.typeColumn).value
    
    Dim styleFormat As String
    styleFormat = StylesSheet.Cells.item(row, styles.formatColumn).value
    
    Dim graphvizSource As String
    
    ' Check for the label attribute
    If InStr(1, styleFormat, "label=", vbTextCompare) > 0 Then
        ' Contains the label attribute
        Select Case styleType
            Case TYPE_NODE
                graphvizSource = "digraph preview { bgcolor=transparent imagepath=" & AddQuotes(GetImagePath()) & " " & AddQuotes(styleName) & " [" & styleFormat & "] }"
            Case TYPE_EDGE
                graphvizSource = "digraph preview { bgcolor=transparent imagepath=" & AddQuotes(GetImagePath()) & " layout=dot rankdir=LR tail[shape=point color=invis]; head[shape=point color=invis]; tail->head[" & styleFormat & "] }"
            Case TYPE_SUBGRAPH_OPEN
                graphvizSource = "digraph preview { bgcolor=transparent imagepath=" & AddQuotes(GetImagePath()) & " layout=dot rankdir=LR subgraph cluster_1 {" & styleFormat & " node[style=filled fillcolor=white]; A->Z; } }"
            Case Else
        End Select
    Else
        ' Supply a label
        Select Case styleType
            Case TYPE_NODE
                graphvizSource = "digraph preview { bgcolor=transparent imagepath=" & AddQuotes(GetImagePath()) & " " & AddQuotes(styleName) & " [label=" & AddQuotes(replace(styleName, " ", "\n")) & " " & styleFormat & "] }"
            Case TYPE_EDGE
                graphvizSource = "digraph preview { bgcolor=transparent layout=dot rankdir=LR tail[shape=point color=invis]; head[shape=point color=invis]; tail->head[label=" & AddQuotes(styleName) & " " & styleFormat & "] }"
            Case TYPE_SUBGRAPH_OPEN
                graphvizSource = "digraph preview { bgcolor=transparent layout=dot rankdir=LR subgraph cluster_1 { label=" & AddQuotes(styleName) & " " & styleFormat & " node[style=filled fillcolor=white]; A->Z; } }"
            Case Else
        End Select
    End If
    If graphvizSource <> vbNullString Then
        PreviewStyleAndAutosize styleName, graphvizSource, previewColumn, row
    End If
    
    ' Repaint the screen
    DoEvents
End Sub

' ==========================================================================
' PROCEDURE: ClearStylesPreview
' PURPOSE:
'   Removes all visual thumbnails and resets row formatting on the Styles sheet.
'
' TECHNICAL WORKFLOW:
'   1. IMAGE RECLAMATION: Calls 'DeleteAllPictures' specifically for the
'      'StylesSheet' to clear the "Preview" column of all Graphviz renders.
'   2. LAYOUT RESTORATION:
'      - Resolves the functional range of the sheet via 'GetSettingsForStylesWorksheet'.
'      - Iterates through the data rows and applies '.AutoFit'.
'   3. STATE RESET: Returns the sheet to a standard text-based view,
'      removing the expanded rows typically required to house thumbnails.
'
' USAGE:
'   - Essential for "re-indexing" the gallery or reducing file size before
'     distribution.
' ==========================================================================
Public Sub ClearStylesPreview()
    ' Delete the images
    DeleteAllPictures StylesSheet.name
    
    ' Get the 'styles' sheet layout
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()

    ' Reset the row height
    Dim row As Long
    For row = styles.firstRow To styles.lastRow
        StylesSheet.rows.item(row).AutoFit
    Next row
End Sub

' ==========================================================================
' PROCEDURE: PreviewStyleForCurrentRow
' PURPOSE:
'   Triggers a visual render for the style entry currently selected by the user.
'
' TECHNICAL WORKFLOW:
'   1. FOCUS MANAGEMENT: Ensures 'StylesSheet' is the active context.
'   2. ATOMIC INVOCATION: Passes the 'ActiveCell.row' to 'GenerateStylesPreview'
'      to execute the specific DOT-to-image pipeline for that style.
'   3. UI CLEANUP: Resets the 'StatusBar' to clear any "Rendering..." messages,
'      maintaining a clean user interface.
'
' USAGE:
'   - Linked to the floating button on the Styles worksheet.
'   - Used for rapid iterative testing during style development.
' ==========================================================================
Public Sub PreviewStyleForCurrentRow()
    PreviewStyleForRow ActiveCell.row
End Sub

Public Sub PreviewStyleForRow(row As Long)
    StylesSheet.Activate
    GenerateStylesPreview row
    ClearStatusBar
End Sub

' ==========================================================================
' PROCEDURE: PreviewStyleAndAutosize
' PURPOSE:
'   Executes the Graphviz pipeline for a specific style and embeds the
'   resulting image into the worksheet with automatic row-height adjustment.
'
' TECHNICAL WORKFLOW:
'   1. OBJECT INSTANTIATION: Creates a new 'Graphviz' class instance to
'      encapsulate the rendering parameters (Path, Format, Engine).
'   2. PRE-RENDER CLEANUP: Invokes 'DeleteCellPictures' to ensure the
'      target cell is empty before the new thumbnail is placed.
'   3. ASYNCHRONOUS RENDERING:
'      - Exports the DOT source to a temporary file.
'      - Invokes 'RenderGraph' to call the Graphviz binary (dot.exe).
'      - Redirects console feedback to the 'Console' worksheet for debugging.
'   4. IMAGE INJECTION: Calls 'InsertPicture' to physically embed the
'      PNG into the target cell with an automated Alt-Text description.
'   5. DYNAMIC LAYOUT (#Else / Win):
'      - Uses the 'WIA.ImageFile' (Windows Image Acquisition) library to
'        inspect the physical height of the generated PNG.
'      - Calls 'determineRowHeight' to expand the Excel row to fit the image.
'   6. MAC FALLBACK (#If Mac): Applies a standard height (126 points)
'      due to sandboxing restrictions on image metadata inspection.
'   7. RESOURCE HYGIENE: Systematically deletes temporary DOT and PNG
'      files to prevent disk clutter.
' ==========================================================================
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
    Dim shapeObject As shape
    '@Ignore AssignmentNotUsed
    Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range(targetCell), False, True, _
                                    "Image showing the rendering of style named " & styleName)
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

' ==========================================================================
' FUNCTION: determineRowHeight
' PURPOSE:
'   Calculates the optimal Excel row height required to house a rendered image.
'
' TECHNICAL WORKFLOW:
'   1. UNIT CONVERSION: Applies the standard 96 DPI logic (image pixels * 72 / 96)
'      to translate raw image dimensions into Excel "Points."
'   2. BOUNDARY CLAMPING:
'      - FLOOR: Enforces a 20pt minimum to keep the row selectable.
'      - CEILING: Enforces a 546pt maximum to comply with Excel's internal
'        row height limitations.
'   3. NON-DESTRUCTIVE SCALING: Compares the calculated height against the
'      existing row height; it will only expand the row, never shrink it
'      below its current state.
'
' USAGE:
'   - Called by 'PreviewStyleAndAutosize' during the Windows rendering pipeline.
'   - Ensures a professional, uniform look for the Style Gallery.
' ==========================================================================
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

' ==========================================================================
' PROCEDURE: RestoreStyleDesigner
' PURPOSE:
'   Populates the Style Designer with the attributes of a selected Style Gallery row.
'
' TECHNICAL WORKFLOW:
'   1. DATA EXTRACTION: Retrieves the raw format string, style name, and
'      object type (Node, Edge, Cluster) from the active 'Styles' row.
'   2. WORKSPACE RESET: Purges previous designer data via 'ClearStyleDesignerRanges'
'      and 'ClearStyleDesignerLabels' to ensure a clean slate.
'   3. ATTRIBUTE PARSING: Calls 'ParseAttributeString' to break the DOT code
'      into a key-value Dictionary (e.g., "color" -> "red").
'   4. GUI MAPPING: Iterates through the dictionary, calling 'RestoreStyleDesignerSetting'
'      to physically place values into the correct Designer input cells.
'   5. UI RE-INDEXING: Triggers 'ShowLabelRows' to hide/show input fields
'      relevant to the specific mode (e.g., hiding Head/Tail labels for Nodes).
'   6. SESSION ACTIVATION: Makes the 'StyleDesignerSheet' visible, refreshes
'      the Ribbon, and triggers a 'RenderPreview' to show the restored style.
' ==========================================================================
Public Sub RestoreStyleDesigner()
    ' Turn off screen updates
    OptimizeCode_Begin
        
    Dim row As Long
    row = ActiveCell.row
        
    ' Get the Format String
    Dim formatText As String
    formatText = CStr(StylesSheet.Cells(row, GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)).value)
    
    ' Get the Style Name
    Dim styleName As String
    styleName = GetStyleNameForRestore(row)
    
    ' Set the edit mode for the ribbon to switch to
    Dim mode As String
    mode = StylesSheet.Cells(row, GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)).value
    
    ' Ribbon mode uses different values for its run mode than styles do
    If mode = TYPE_SUBGRAPH_OPEN Then mode = TYPE_CLUSTER

    ' Ensure mode is uppercase
    mode = UCase$(mode)
    
    ' Establish the mode the Style Designer ribbon should switch to
    StyleDesignerSheet.Range(DESIGNER_MODE).value = UCase$(mode)
    
    ' Restore the Style Name
    StyleDesignerSheet.Range(DESIGNER_STYLE_NAME_TEXT).value = styleName

    ' Reset all the Style Designer ribbon settings
    ClearStyleDesignerRanges
    
    ' Clear the color sheme
    StyleDesignerSheet.Range(DESIGNER_COLOR_SCHEME).value = vbNullString

    ' Clear the Label fields on the worksheet
    ClearStyleDesignerLabels
    
    ' Default the style name as the label
    StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT).value = styleName
    
    ' Show only the label rows which are appropriate for the mode (Node,Edge,Cluster)
    ShowLabelRows mode

    ' Parse the format string into a Dictionary
    Dim attributeDictionary As Dictionary
    Set attributeDictionary = ParseAttributeString(formatText)
    
    ' Iterate the dictionary to restore values into the correct settings cells
    Dim key As Variant
    For Each key In attributeDictionary.Keys
        RestoreStyleDesignerSetting mode, CStr(key), CStr(attributeDictionary(key))
    Next key
    Set attributeDictionary = Nothing
    
    ' Turn screen updates on (RenderPreview also uses OptimizeCode routines)
    OptimizeCode_End
    
    ' Regenerate the preview image
    RenderPreview
    
    ' Refresh the ribbon controls
    RefreshRibbon
    
    ' Switch to the Style Designer Worksheet
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).value = TOGGLE_SHOW
    StyleDesignerSheet.visible = True
    StyleDesignerSheet.Activate
End Sub

' ==========================================================================
' FUNCTION: GetStyleNameForRestore
' PURPOSE:
'   Cleans a style name from the Gallery for use in the Style Designer.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA LOOKUP: Retrieves the 'Styles' sheet layout and defined
'      suffixes (e.g., "_OPEN") via 'GetSettingsForStylesWorksheet'.
'   2. DATA EXTRACTION: Trims and captures the raw 'styleName' and 'styleType'.
'   3. SUFFIX STRIPPING:
'      - Specifically checks for 'SUBGRAPH_OPEN' (Cluster) styles.
'      - If the name ends with the global 'Open Suffix', it removes that
'        suffix to return the user to the "root" style name.
'   4. NORMALIZATION: Returns a clean, trimmed string ready for display in
'      the Style Designer's input fields.
' ==========================================================================
Private Function GetStyleNameForRestore(row As Long)
    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    ' Get the Style Name
    Dim styleName As String
    styleName = Trim$(StylesSheet.Cells(row, styles.nameColumn).value)
    
    Dim styleType As String
    styleType = Trim$(StylesSheet.Cells(row, styles.typeColumn).value)
    
    ' If the style is associated with a cluster, trim off the suffix
    If styleType = TYPE_SUBGRAPH_OPEN And EndsWith(styleName, styles.suffixOpen) Then
        styleName = Left(styleName, Len(styleName) - Len(styles.suffixOpen) - 1)
    End If

    GetStyleNameForRestore = Trim$(styleName)
End Function

' ==========================================================================
' PROCEDURE: RestoreStyleDesignerSetting
' PURPOSE:
'   Routes a specific DOT attribute to the correct Style Designer UI control.
'
' TECHNICAL WORKFLOW:
'   1. NORMALIZATION: Trims and lowercases the 'attributeName' to ensure
'      case-insensitive matching against Graphviz constants.
'   2. DISPATCH LOGIC: Uses a large 'Select Case' block to determine the
'      destination of the 'attributeValue':
'      - DIRECT ASSIGNMENT: Writes simple values (like 'width' or 'shape')
'        straight to a specific Named Range.
'      - COMPLEX PROCESSING: Delegates multi-part attributes (like 'color'
'        gradients, 'style' arrays, or 'labels') to specialized 'Apply...'
'        helper subroutines.
'   3. CONTEXT AWARENESS: Passes the 'mode' (Node, Edge, Cluster) to helpers
'      to ensure attributes like 'penwidth' are applied to the correct
'      logical UI group (e.g., Border vs. Line).
'   4. DIAGNOSTICS: Prints unhandled attributes to the Immediate Window
'      during development to identify gaps in the designer's coverage.
' ==========================================================================
Private Sub RestoreStyleDesignerSetting(mode As String, attributeName As String, attributeValue As String)
    Dim result() As String
    Dim cnt As Long
    Dim i As Long

    Select Case LCase$(Trim$(attributeName))
        Case GRAPHVIZ_ARROWHEAD:      ApplyArrowheadSettings attributeValue
        Case GRAPHVIZ_ARROWSIZE:      StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_SIZE).value = attributeValue
        Case GRAPHVIZ_ARROWTAIL:      ApplyArrowtailSettings attributeValue
        Case GRAPHVIZ_COLOR:          ApplyColorSettings attributeValue, mode
        Case GRAPHVIZ_COLORSCHEME:    StyleDesignerSheet.Range(DESIGNER_COLOR_SCHEME).value = attributeValue
        Case GRAPHVIZ_DECORATE:       StyleDesignerSheet.Range(DESIGNER_EDGE_DECORATE).value = attributeValue
        Case GRAPHVIZ_DIR:            StyleDesignerSheet.Range(DESIGNER_EDGE_DIRECTION).value = attributeValue
        Case GRAPHVIZ_DISTORTION:     StyleDesignerSheet.Range(DESIGNER_NODE_DISTORTION).value = attributeValue
        Case GRAPHVIZ_FILLCOLOR:      ApplyFillColorSettings attributeValue
        Case GRAPHVIZ_FIXEDSIZE:      StyleDesignerSheet.Range(DESIGNER_NODE_FIXED_SIZE).value = attributeValue
        Case GRAPHVIZ_FONTNAME:       ApplyFontNameSettings attributeValue
        Case GRAPHVIZ_FONTCOLOR:      StyleDesignerSheet.Range(DESIGNER_FONT_COLOR).value = attributeValue
        Case GRAPHVIZ_FONTSIZE:       StyleDesignerSheet.Range(DESIGNER_FONT_SIZE).value = attributeValue
        Case GRAPHVIZ_GRADIENTANGLE:  StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_ANGLE).value = attributeValue
        Case GRAPHVIZ_HEADCLIP:       StyleDesignerSheet.Range(DESIGNER_EDGE_HEAD_CLIP).value = attributeValue
        Case GRAPHVIZ_HEADLABEL:      ApplyLabelSetting DESIGNER_HEAD_LABEL_TEXT, DESIGNER_HEAD_LABEL_TEXT_INCLUDE, attributeValue
        Case GRAPHVIZ_HEADPORT:       StyleDesignerSheet.Range(DESIGNER_EDGE_HEAD_PORT).value = attributeValue
        Case GRAPHVIZ_HEIGHT:         StyleDesignerSheet.Range(DESIGNER_NODE_HEIGHT).value = attributeValue
        Case GRAPHVIZ_IMAGE:          StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME).value = attributeValue
        Case GRAPHVIZ_IMAGEPOS:       StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_POSITION).value = attributeValue
        Case GRAPHVIZ_IMAGESCALE:     StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_SCALE).value = attributeValue
        Case GRAPHVIZ_LABEL:          ApplyLabelSetting DESIGNER_LABEL_TEXT, DESIGNER_LABEL_TEXT_INCLUDE, attributeValue
        Case GRAPHVIZ_LABELANGLE:     StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_ANGLE).value = attributeValue
        Case GRAPHVIZ_LABELDISTANCE:  StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_DISTANCE).value = attributeValue
        Case GRAPHVIZ_LABELFLOAT:     StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FLOAT).value = attributeValue
        Case GRAPHVIZ_LABELFONTCOLOR: StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_COLOR).value = attributeValue
        Case GRAPHVIZ_LABELFONTNAME:  StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_NAME).value = attributeValue
        Case GRAPHVIZ_LABELFONTSIZE:  StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_SIZE).value = attributeValue
        Case GRAPHVIZ_LABELJUST:      StyleDesignerSheet.Range(DESIGNER_LABEL_JUSTIFICATION).value = attributeValue
        Case GRAPHVIZ_LABELLOC:       StyleDesignerSheet.Range(DESIGNER_LABEL_LOCATION).value = attributeValue
        Case GRAPHVIZ_MARGIN:         StyleDesignerSheet.Range(DESIGNER_CLUSTER_MARGIN).value = attributeValue
        Case GRAPHVIZ_ORIENTATION:    StyleDesignerSheet.Range(DESIGNER_NODE_ORIENTATION).value = attributeValue
        Case GRAPHVIZ_PACK:           StyleDesignerSheet.Range(DESIGNER_CLUSTER_MARGIN).value = attributeValue
        Case GRAPHVIZ_PACKMODE:       ApplyPackmodeSettings attributeValue
        Case GRAPHVIZ_PENCOLOR:       StyleDesignerSheet.Range(DESIGNER_BORDER_COLOR).value = attributeValue
        Case GRAPHVIZ_PENWIDTH:       ApplyPenWidthSetting attributeValue, mode
        Case GRAPHVIZ_PERIPHERIES:    StyleDesignerSheet.Range(DESIGNER_BORDER_PERIPHERIES).value = attributeValue
        Case GRAPHVIZ_REGULAR:        StyleDesignerSheet.Range(DESIGNER_NODE_REGULAR).value = attributeValue
        Case GRAPHVIZ_RADIUS:         StyleDesignerSheet.Range(DESIGNER_EDGE_RADIUS).value = attributeValue
        Case GRAPHVIZ_SHAPE:          StyleDesignerSheet.Range(DESIGNER_NODE_SHAPE).value = attributeValue
        Case GRAPHVIZ_SIDES:          StyleDesignerSheet.Range(DESIGNER_NODE_SIDES).value = attributeValue
        Case GRAPHVIZ_SKEW:           StyleDesignerSheet.Range(DESIGNER_NODE_SKEW).value = attributeValue
        Case GRAPHVIZ_STYLE:          ApplyStyleValue attributeValue, mode
        Case GRAPHVIZ_TAILCLIP:       StyleDesignerSheet.Range(DESIGNER_EDGE_TAIL_CLIP).value = attributeValue
        Case GRAPHVIZ_TAILLABEL:      ApplyLabelSetting DESIGNER_TAIL_LABEL_TEXT, DESIGNER_TAIL_LABEL_TEXT_INCLUDE, attributeValue
        Case GRAPHVIZ_TAILPORT:       StyleDesignerSheet.Range(DESIGNER_EDGE_TAIL_PORT).value = attributeValue
        Case GRAPHVIZ_WEIGHT:         StyleDesignerSheet.Range(DESIGNER_EDGE_WEIGHT).value = attributeValue
        Case GRAPHVIZ_WIDTH:          StyleDesignerSheet.Range(DESIGNER_NODE_WIDTH).value = attributeValue
        Case GRAPHVIZ_XLABEL:         ApplyLabelSetting DESIGNER_XLABEL_TEXT, DESIGNER_XLABEL_TEXT_INCLUDE, attributeValue
        Case Else
            Debug.Print attributeName & " : " & attributeValue & " was not handled"
    End Select
End Sub

' ==========================================================================
' PROCEDURE: ApplyArrowheadSettings
' PURPOSE:
'   Maps a Graphviz 'arrowhead' string back to the Style Designer UI.
'
' TECHNICAL WORKFLOW:
'   1. STRING DECONSTRUCTION: Calls 'ParseGraphvizArrowheads' to break a
'      combined string into its primitive components (e.g., "onormal" becomes
'      "o" and "normal").
'   2. UI INJECTION: Iterates through the resulting array and populates
'      the sequential arrowhead ranges (DESIGNER_EDGE_ARROW_HEAD 1, 2, 3...).
'   3. STATE VALIDATION: Only proceeds if the first component contains data,
'      ensuring the UI isn't cleared by empty attribute values.
'
' USAGE:
'   - Used by 'RestoreStyleDesignerSetting' to "rehydrate" the complex
'     arrowhead selectors from raw DOT source code.
' ==========================================================================
Private Sub ApplyArrowheadSettings(ByVal attributeValue As String)
    Dim result() As String
    Dim i As Long

    result = ParseGraphvizArrowheads(attributeValue)

    If Len(result(0)) > 0 Then
        For i = 0 To UBound(result)
            StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD & (i + 1)).value = result(i)
        Next i
    End If
End Sub

' ==========================================================================
' PROCEDURE: ApplyArrowtailSettings
' PURPOSE:
'   Maps a Graphviz 'arrowtail' string back to the Style Designer UI.
'
' TECHNICAL WORKFLOW:
'   1. STRING DECONSTRUCTION: Leverages 'ParseGraphvizArrowheads' to split
'      the combined tail string into its constituent parts.
'   2. UI INJECTION: Iterates through the result and populates the
'      sequential arrowtail ranges (DESIGNER_EDGE_ARROW_TAIL 1, 2, 3, etc.).
'   3. STATE VALIDATION: Checks the first element of the result array to
'      prevent erroneous writes from empty or invalid strings.
'
' USAGE:
'   - Used by 'RestoreStyleDesignerSetting' to restore complex multi-part
'     tail shapes from existing DOT code.
' ==========================================================================
Private Sub ApplyArrowtailSettings(ByVal attributeValue As String)
    Dim result() As String
    Dim i As Long

    result = ParseGraphvizArrowheads(attributeValue)

    If Len(result(0)) > 0 Then
        For i = 0 To UBound(result)
            StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL & (i + 1)).value = result(i)
        Next i
    End If
End Sub

' ==========================================================================
' PROCEDURE: ApplyColorSettings
' PURPOSE:
'   Directs 'color' attributes to either Edge paths or Node/Cluster borders.
'
' TECHNICAL WORKFLOW:
'   1. MODE DISCRIMINATION: Determines the logical target (Edge vs. Node/Cluster).
'   2. EDGE LOGIC (Multi-color support):
'      - Splits the attribute string by ':' to support Graphviz multi-color
'        parallel lines.
'      - Maps up to 3 individual color values to the sequential Edge color
'        input ranges (DESIGNER_EDGE_COLOR 1, 2, 3).
'   3. NODE/CLUSTER LOGIC (Single color):
'      - Directly assigns the value to the 'Border Color' input range.
'   4. DATA SANITIZATION: Applies 'Trim$' to ensure no leading/trailing
'      whitespace disrupts the color-picker lookup.
' ==========================================================================
Private Sub ApplyColorSettings(ByVal attributeValue As String, ByVal mode As String)
    Select Case LCase$(mode)
        Case TYPE_EDGE
            Dim colors() As String
            Dim color As Variant
            Dim cnt As Long

            colors = split(attributeValue, ":")
            cnt = 0

            For Each color In colors
                cnt = cnt + 1
                If cnt <= 3 Then
                    StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR & cnt).value = Trim$(color)
                End If
            Next color

        Case TYPE_NODE, TYPE_CLUSTER
            StyleDesignerSheet.Range(DESIGNER_BORDER_COLOR).value = Trim$(attributeValue)
    End Select
End Sub

' ==========================================================================
' PROCEDURE: ApplyFillColorSettings
'
' PURPOSE:
'   Deconstructs a complex Graphviz fill attribute string and maps the
'   individual components to the Style Designer's input fields.
'
' TECHNICAL WORKFLOW:
'   1. STRING PARSING: Splits the 'attributeValue' using a colon (:) delimiter
'      to separate primary fill settings from secondary gradient colors.
'   2. GRADIENT ANGLE EXTRACTION: Further parses the primary segment using
'      a semicolon (;) to isolate the fill color from the gradient angle.
'   3. UNIT CONVERSION: Converts the Graphviz angle value into a percentage-
'      based "gradient weight" (Angle * 100) for the Excel UI.
'   4. UI MAPPING: Writes the resulting 'fillColor', 'gradientColor', and
'      'gradientWeight' directly to the 'StyleDesignerSheet' using
'      predefined range constants.
'
' USAGE:
'   - Internal helper called during the "Restore Style" process.
'   - Essential for handling advanced Graphviz styling like linear gradients.
' ==========================================================================
Private Sub ApplyFillColorSettings(ByVal attributeValue As String)
    Dim fillAttribute As String
    Dim fillColor As String
    Dim gradientWeight As String
    Dim gradientColor As String

    fillAttribute = Trim$(attributeValue)
    gradientWeight = ""
    gradientColor = ""

    ' Split on colon
    Dim colonParts() As String
    colonParts = split(fillAttribute, ":")

    ' Parse left side: fillColor and optional angle
    Dim leftParts() As String
    leftParts = split(colonParts(0), ";")
    fillColor = Trim$(leftParts(0))

    If UBound(leftParts) >= 1 Then
        Dim angleVal As Double
        angleVal = val(Trim$(leftParts(1)))
        gradientWeight = CStr(Int(angleVal * 100))
    End If

    ' Parse right side: gradient color
    If UBound(colonParts) >= 1 Then
        gradientColor = Trim$(colonParts(1))
    End If

    ' Apply to sheet
    With StyleDesignerSheet
        .Range(DESIGNER_FILL_COLOR).value = fillColor
        .Range(DESIGNER_GRADIENT_FILL_COLOR).value = gradientColor
        .Range(DESIGNER_GRADIENT_FILL_WEIGHT).value = gradientWeight
    End With
End Sub

' ==========================================================================
' PROCEDURE: ApplyFillColorSettings
' PURPOSE:
'   Decodes Graphviz fill strings (including gradients) into UI settings.
'
' TECHNICAL WORKFLOW:
'   1. GRADIENT DETECTION: Splits the attribute by ':' to separate the
'      primary fill color from the secondary gradient color.
'   2. WEIGHT/ANGLE EXTRACTION: Splits the primary color segment by ';'
'      to find optional Graphviz gradient angle data.
'   3. UNIT CONVERSION: Converts the Graphviz decimal angle/weight into
'      a percentage-based 'gradientWeight' (0-100) for the Excel UI.
'   4. UI SYNCHRONIZATION: Populates three distinct Named Ranges:
'      - 'DESIGNER_FILL_COLOR': The base or start color.
'      - 'DESIGNER_GRADIENT_FILL_COLOR': The target transition color.
'      - 'DESIGNER_GRADIENT_FILL_WEIGHT': The visual balance of the gradient.
'
' USAGE:
'   - Essential for "rehydrating" complex visual styles that use
'     linear or radial gradients from raw DOT source code.
' ==========================================================================
Private Sub ApplyLabelSetting(ByVal textCell As String, ByVal includeFlagCell As String, ByVal attributeValue As String)
    With StyleDesignerSheet
        .Range(textCell).value = attributeValue
        .Range(includeFlagCell).value = True
    End With
End Sub

' ==========================================================================
' PROCEDURE: ApplyPenWidthSetting
' PURPOSE:
'   Routes the 'penwidth' attribute to the correct Style Designer UI field.
'
' TECHNICAL WORKFLOW:
'   1. MODE DISCRIMINATION: Checks the object type (Edge vs. Node/Cluster)
'      to determine the logical meaning of "line thickness" for the user.
'   2. EDGE MAPPING: Assigns the value to the 'DESIGNER_EDGE_PEN_WIDTH' range,
'      representing the weight of the connecting line.
'   3. NODE/CLUSTER MAPPING: Assigns the value to the 'DESIGNER_BORDER_PEN_WIDTH'
'      range, representing the weight of the object's outline.
'
' USAGE:
'   - Ensures that when a style is restored, the "Thickness" dropdown
'     matches the original DOT specification.
' ==========================================================================
Private Sub ApplyPenWidthSetting(ByVal attributeValue As String, ByVal mode As String)
    Select Case LCase$(mode)
        Case TYPE_EDGE
            StyleDesignerSheet.Range(DESIGNER_EDGE_PEN_WIDTH).value = attributeValue

        Case TYPE_NODE, TYPE_CLUSTER
            StyleDesignerSheet.Range(DESIGNER_BORDER_PEN_WIDTH).value = attributeValue
    End Select
End Sub

' ==========================================================================
' PROCEDURE: ApplyFontNameSettings
' PURPOSE:
'   Extracts font family, bold, and italic states from a Graphviz font name.
'
' TECHNICAL WORKFLOW:
'   1. SUFFIX ANALYSIS: Evaluates the trailing characters of the attribute
'      value using a 'Select Case True' pattern to detect style keywords.
'   2. STATE TOGGLING:
'      - "Bold Italic": Sets both Bold and Italic toggles to 'Yes'.
'      - "Bold": Sets Bold to 'Yes' and Italic to 'No'.
'      - "Italic": Sets Bold to 'No' and Italic to 'Yes'.
'   3. STRING CLEANING: Truncates the style suffix from the full string
'      to isolate the 'baseFontName' (e.g., "Times New Roman").
'   4. UI RESTORATION: Populates the font dropdown and the boolean
'      toggles in the Style Designer simultaneously.
' ==========================================================================
Private Sub ApplyFontNameSettings(ByVal attributeValue As String)
    Dim fullFontName As String
    Dim baseFontName As String

    fullFontName = Trim$(attributeValue)

    Select Case True
        Case Right$(fullFontName, 11) = "Bold Italic"
            StyleDesignerSheet.Range(DESIGNER_FONT_BOLD).value = TOGGLE_YES
            StyleDesignerSheet.Range(DESIGNER_FONT_ITALIC).value = TOGGLE_YES
            baseFontName = Left$(fullFontName, Len(fullFontName) - 12)

        Case Right$(fullFontName, 4) = "Bold"
            StyleDesignerSheet.Range(DESIGNER_FONT_BOLD).value = TOGGLE_YES
            StyleDesignerSheet.Range(DESIGNER_FONT_ITALIC).value = TOGGLE_NO
            baseFontName = Left$(fullFontName, Len(fullFontName) - 5)

        Case Right$(fullFontName, 6) = "Italic"
            StyleDesignerSheet.Range(DESIGNER_FONT_BOLD).value = TOGGLE_NO
            StyleDesignerSheet.Range(DESIGNER_FONT_ITALIC).value = TOGGLE_YES
            baseFontName = Left$(fullFontName, Len(fullFontName) - 7)

        Case Else
            StyleDesignerSheet.Range(DESIGNER_FONT_BOLD).value = TOGGLE_NO
            StyleDesignerSheet.Range(DESIGNER_FONT_ITALIC).value = TOGGLE_NO
            baseFontName = fullFontName
    End Select

    StyleDesignerSheet.Range(DESIGNER_FONT_NAME).value = Trim$(baseFontName)
End Sub

' ==========================================================================
' PROCEDURE: ApplyPackmodeSettings
' PURPOSE:
'   Decodes complex Graphviz 'packmode' strings into individual UI toggles.
'
' TECHNICAL WORKFLOW:
'   1. OBJECT PARSING: Calls 'ParseGraphvizPackmode' to break the string
'      into its base Mode, Flags, and Suffix components.
'   2. ALIGNMENT RESOLUTION: Scans the Flags for 't' (Top) or 'b' (Bottom)
'      markers to set the vertical alignment dropdown.
'   3. JUSTIFICATION MAPPING: Identifies 'l' (Left) or 'r' (Right) markers
'      to synchronize the horizontal justification UI.
'   4. ARRAY CONFIGURATION:
'      - Detects the 's' flag to toggle the 'Array Sort' setting.
'      - Checks for 'c' to set the 'Column Major' vs. 'Row Major' orientation.
'   5. SUFFIX CAPTURE: Restores the numeric suffix (often used for row/column
'      counts) to the 'Array Split' input range.
' ==========================================================================
Private Sub ApplyPackmodeSettings(ByVal attributeValue As String)
    Dim packmode As Object
    Set packmode = ParseGraphvizPackmode(attributeValue)

    ' Apply main mode
    StyleDesignerSheet.Range(DESIGNER_CLUSTER_PACKMODE).value = packmode("Mode")

    ' Apply flags
    Dim Flags As String
    Flags = Trim$(packmode("Flags"))

    StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_SORT).value = TOGGLE_NO

    ' Bottom/Top Alignment
    If InStr(1, Flags, GRAPHVIZ_PACKMODE_ALIGN_TOP, vbTextCompare) > 0 Then
        StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_ALIGN).value = GRAPHVIZ_PACKMODE_ALIGN_TOP
    ElseIf InStr(1, Flags, GRAPHVIZ_PACKMODE_ALIGN_BOTTOM, vbTextCompare) > 0 Then
        StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_ALIGN).value = GRAPHVIZ_PACKMODE_ALIGN_BOTTOM
    End If

    ' Left/Right Justification
    If InStr(1, Flags, GRAPHVIZ_PACKMODE_JUSTIFY_LEFT, vbTextCompare) > 0 Then
        StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_JUSTIFY).value = GRAPHVIZ_PACKMODE_JUSTIFY_LEFT
    ElseIf InStr(1, Flags, GRAPHVIZ_PACKMODE_JUSTIFY_RIGHT, vbTextCompare) > 0 Then
        StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_JUSTIFY).value = GRAPHVIZ_PACKMODE_JUSTIFY_RIGHT
    End If

    ' Sort Order
    If InStr(1, Flags, GRAPHVIZ_PACKMODE_SORT, vbTextCompare) > 0 Then
        StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_SORT).value = TOGGLE_YES
    End If

    ' Column/Row Major
    If InStr(1, Flags, GRAPHVIZ_PACKMODE_MAJOR_COLUMN, vbTextCompare) > 0 Then
        StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_MAJOR).value = GRAPHVIZ_PACKMODE_MAJOR_COLUMN
    End If

    ' Apply suffix
    StyleDesignerSheet.Range(DESIGNER_CLUSTER_ARRAY_SPLIT).value = packmode("Suffix")

    Set packmode = Nothing
End Sub

' ==========================================================================
' PROCEDURE: ApplyNodeStyles
' PURPOSE:
'   Distributes 'style' attributes to the correct gradient or border controls.
'
' TECHNICAL WORKFLOW:
'   1. TOKENIZATION: Splits the 'attributeValue' by commas to handle multi-
'      part Graphviz style declarations (e.g., "filled, dashed, rounded").
'   2. GRADIENT DETECTION: Specifically scans for 'radial' or 'filled'
'      keywords to populate the 'Gradient Fill Type' dropdown.
'   3. BORDER SEQUENCING:
'      - Iterates through the remaining style tokens (e.g., 'dotted', 'bold').
'      - Maps up to 3 individual properties to the sequential UI ranges
'        (BorderStyle1, BorderStyle2, BorderStyle3).
'   4. DEFENSIVE BOUNDING: Limits the border assignment to 3 slots to match
'      the Style Designer's physical UI constraints.
' ==========================================================================
Private Sub ApplyNodeStyles(ByVal attributeValue As String)
    Dim styles() As String
    Dim style As Variant
    Dim trimmedStyle As String
    Dim cnt As Long

    styles = split(attributeValue, ",")
    cnt = 0

    For Each style In styles
        trimmedStyle = Trim$(style)

        Select Case True
            Case InStr(1, trimmedStyle, GRAPHVIZ_STYLE_GRADIENT_RADIAL, vbTextCompare) > 0, _
                 InStr(1, trimmedStyle, GRAPHVIZ_STYLE_GRADIENT_FILLED, vbTextCompare) > 0
                StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_TYPE).value = trimmedStyle

            Case Else
                cnt = cnt + 1
                If cnt <= 3 Then ' Defensive: avoid overflow
                    StyleDesignerSheet.Range("BorderStyle" & cnt).value = trimmedStyle
                End If
        End Select
    Next style
End Sub

' ==========================================================================
' PROCEDURE: ApplyStyleValue
' PURPOSE:
'   Routes the 'style' attribute to either complex node logic or simple edge settings.
'
' TECHNICAL WORKFLOW:
'   1. TARGET ANALYSIS: Evaluates the 'mode' to determine if the style
'      applies to a shape or a line.
'   2. CONTAINER LOGIC (Node/Cluster): Delegates to 'ApplyNodeStyles' to
'      handle complex comma-separated values like "filled,dashed,rounded".
'   3. CONNECTOR LOGIC (Edge): Directly assigns the value to the single
'      'DESIGNER_EDGE_STYLE' range (e.g., "dotted" or "bold").
'
' USAGE:
'   - Standardizes the restoration of visual line/fill patterns from raw
'     DOT source code back into the Designer GUI.
' ==========================================================================
Private Sub ApplyStyleValue(ByVal attributeValue As String, ByVal mode As String)
    Select Case LCase$(mode)
        Case TYPE_NODE, TYPE_CLUSTER
            Call ApplyNodeStyles(attributeValue)

        Case TYPE_EDGE
            StyleDesignerSheet.Range(DESIGNER_EDGE_STYLE).value = attributeValue
    End Select
End Sub

' ==========================================================================
' FUNCTION: IsStylesRowActive
' PURPOSE:
'   Determines if a specific row in the Style Gallery is eligible for
'   interactive floating buttons.
'
' TECHNICAL WORKFLOW:
'   1. DEFAULT STATE: Initializes to 'False' to ensure a "deny-by-default"
'      safety posture.
'   2. COMMENT CHECK: Scans the 'Comment' column for the '#' indicator.
'      If found, the row is treated as documentation and ignored by the UI.
'   3. TYPE VERIFICATION: Retrieves the 'Object Type' and checks it against
'      supported functional types:
'      - TYPE_NODE (Standard Nodes)
'      - TYPE_EDGE (Connections)
'      - TYPE_SUBGRAPH_OPEN (Cluster boundaries)
'   4. VALIDATION GRANT: Returns 'True' only if the row is an active,
'      renderable style definition.
'
' USAGE:
'   - The primary 'ValidationFunc' used by the floating button engine on
'     the 'Styles' worksheet.
' ==========================================================================
Public Function IsStylesRowActive(ByVal row As Long) As Boolean
    IsStylesRowActive = False               ' Establish default setting
    
    Dim commentIndicator As String
    commentIndicator = StylesSheet.Cells(row, GetSettingColNum(SETTINGS_STYLES_COL_COMMENT)).value
    If commentIndicator = "#" Then Exit Function ' Commented out, exit
    
    ' We only use floating buttons for styles "node", "edge", and "subgraph-open"
    Dim styleType As String: styleType = StylesSheet.Cells(row, GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)).value
    
    If styleType = TYPE_NODE Or styleType = TYPE_EDGE Or styleType = TYPE_SUBGRAPH_OPEN Then
        IsStylesRowActive = True
    End If
End Function

