Attribute VB_Name = "modCreateCommon"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modCreateCommon
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Shared Logic Layer — Cross-cutting routines for Graphviz and
'            Knowledge Graph generation.
'
' ROLE:
'   Provides the unified foundation for all parsing, normalization, style
'   resolution, label synthesis, filesystem preparation, and environment
'   detection used by both the Graphviz (DOT) pipeline and the JSON Viewer
'   pipeline. Centralizes logic that must behave identically across modules.
'
' RESPONSIBILITIES:
'   o Data-Row Construction:
'       - Map worksheet rows into structured dataRow UDTs.
'       - Normalize identifiers, labels, metadata, and style names.
'
'   o Style & Attribute Resolution:
'       - Load enabled styles into a high-speed cache.
'       - Apply extra-attribute overrides and style-format templates.
'       - Expand placeholder tokens within label templates.
'
'   o Label Classification:
'       - Detect HTML-like labels for Graphviz's HTML label syntax.
'       - Provide helper routines for ports, node IDs, and syntactic cleanup.
'
'   o Filesystem & Environment Support:
'       - Validate output directories and filename prefixes.
'       - Resolve image paths, environment variables, and platform separators.
'
'   o Shared Parsing Logic:
'       - Classify rows into NODE, EDGE, SUBGRAPH, NATIVE, or KEYWORD types.
'       - Provide syntactic helpers used by both DOT and JSON pipelines.
'
' INTERACTIONS:
'   o modCreateGraph:
'       - Consumes normalized dataRow records.
'       - Uses style caches, label overrides, and classification helpers.
'
'   o modCreateJson:
'       - Reuses parsing, normalization, and label-classification logic.
'       - Shares filesystem routines for viewer emission.
'
'   o ViewerAssets / Settings Sheets:
'       - Supplies style definitions, worksheet mappings, and configuration.
'
' CROSS-PLATFORM NOTES:
'   o All routines avoid platform-specific assumptions.
'   o Path separators, environment variables, and encoding behaviors are
'     normalized for Windows and macOS.
'
' ERROR HANDLING:
'   o Fail-fast for critical filesystem prerequisites.
'   o Graceful fallback for style lookups, attribute overrides, and parsing.
'   o Diagnostic emission via EmitMessage for user-visible issues.
'
' RELATED WIKI PAGES:
'   o Data Mapping & Normalization
'   o Styles & Attribute Inheritance
'   o Label Formatting & HTML Detection
'   o File Output & Environment Resolution
'   o Graph Generation Pipeline (DOT)
'   o JSON Viewer Pipeline
' =============================================================================

Option Explicit

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
    GetDataRow.xlabel = GetCell(worksheetName, row, ini.data.xLabelColumn)
    GetDataRow.taillabel = GetCell(worksheetName, row, ini.data.tailLabelColumn)
    GetDataRow.headlabel = GetCell(worksheetName, row, ini.data.headLabelColumn)
    GetDataRow.tooltip = GetCell(worksheetName, row, ini.data.tooltipColumn)
    GetDataRow.relatedItem = GetCell(worksheetName, row, ini.data.isRelatedToItemColumn)
    GetDataRow.styleName = GetCell(worksheetName, row, ini.data.styleNameColumn)
    GetDataRow.extraAttrs = GetCell(worksheetName, row, ini.data.extraAttributesColumn)
    GetDataRow.properties = GetCell(worksheetName, row, ini.data.propertiesColumn)
    GetDataRow.row = row

End Function

' ==========================================================================
' ROUTINE: ApplyLabelOverrides
'
' PURPOSE:
'   Resolves node-attribute override rules and placeholder substitutions for
'   a single Graphviz label field. Produces a synthesized label value that
'   incorporates extra-attribute overrides, style-format templates, and
'   placeholder expansion based on the current data-row context.
'
' FUNCTIONAL WORKFLOW:
'   1. INITIALIZATION:
'        - Begins with the caller-provided fieldValue and copies it into
'          finalValue for subsequent transformation.
'
'   2. EXTRA-ATTRIBUTE OVERRIDES:
'        - When extra attributes are enabled and present, parses the
'          attribute string and checks for a matching fieldName.
'        - If found, replaces both finalValue and fieldValue with the
'          override value. This preserves inheritance semantics for
'          downstream handlers.
'
'   3. STYLE-FORMAT TEMPLATE EXPANSION:
'        - When style-format rules are enabled and a format template exists,
'          parses the template and checks for a matching fieldName.
'        - If the template contains a placeholder-based value, expands it
'          using the current data-row context (ExpandLabelPlaceholders).
'        - Static template values are preserved as-is.
'
' TECHNICAL NOTES:
'   - This routine does not modify the attribute dictionary directly; callers
'     are responsible for inserting or updating finalValue.
'   - Overrides from extra attributes take precedence over style-format
'     templates.
'   - Placeholder expansion is performed only when a matching template entry
'     exists and contains non-empty content.
'   - DeepWiki Context: Implements the override and template-resolution rules
'     described in the "Styles" and "Attribute Inheritance" documentation.
' ==========================================================================
Public Sub ApplyLabelOverrides(ByRef ini As settings, _
                                  ByRef data As dataRow, _
                                  ByVal fieldName As String, _
                                  ByRef fieldValue As String, _
                                  ByRef finalValue As String)
    Dim d As Dictionary

    finalValue = fieldValue
    
    ' --- Extra Attributes Override ---
    If ini.graph.includeExtraAttributes And Len(data.extraAttrs) > 0 Then
        Set d = ParseAttributeString(data.extraAttrs)
        If d.Exists(fieldName) Then
            finalValue = d(fieldName)
            fieldValue = finalValue     ' Intentional override
        End If
    End If
    
    ' --- Style Format Placeholder Expansion ---
    If ini.graph.includeStyleFormat Then
        Dim labelStyle As String
        labelStyle = vbNullString
        
        ' A label in the Style format may have a placeholder that needs expansion
        If Len(data.Format) > 0 Then
            Set d = ParseAttributeString(data.Format)
            If d.Exists(fieldName) Then
                labelStyle = d(fieldName)
            End If
            
            If Len(labelStyle) > 0 Then
                finalValue = ExpandLabelPlaceholders(labelStyle, data)
            End If
        End If
    End If
End Sub

' ==========================================================================
' FUNCTION: ExpandLabelPlaceholders
'
' PURPOSE:
'   Expands placeholder tokens within a label template using values from the
'   current data-row context. Supports all node and edge label variants and
'   produces a fully substituted string for downstream label synthesis.
'
' FUNCTIONAL WORKFLOW:
'   1. EMPTY-TEMPLATE CHECK:
'        - Returns an empty string immediately when the input template is
'          blank, preventing unnecessary processing.
'
'   2. PLACEHOLDER SUBSTITUTION:
'        - Performs case-insensitive replacement of supported tokens:
'             o {label}      -> data.label
'             o {xlabel}     -> data.xlabel
'             o {taillabel}  -> data.taillabel
'             o {headlabel}  -> data.headlabel
'        - Each replacement is global (all occurrences) and preserves any
'          surrounding static text in the template.
'
'   3. OUTPUT:
'        - Returns the fully expanded string containing substituted values
'          and any remaining literal content.
'
' TECHNICAL NOTES:
'   - Placeholder expansion is performed only on the caller-supplied template;
'     no dictionary lookups or inheritance rules occur here.
'   - Tokens not present in the template are ignored; no errors are raised.
'   - Complements ApplyLabelOverrides and ProcessLabel, which determine when
'     placeholder expansion should occur.
'   - DeepWiki Context: Implements the placeholder-expansion rules described
'     in the "Labels" and "Styles" documentation.
' ==========================================================================
Private Function ExpandLabelPlaceholders(ByVal Text As String, ByRef data As dataRow) As String
    If Len(Text) = 0 Then
        ExpandLabelPlaceholders = ""
        Exit Function
    End If

    Dim result As String
    result = Text

    result = replace(result, "{label}", data.label, 1, -1, vbTextCompare)
    result = replace(result, "{xlabel}", data.xlabel, 1, -1, vbTextCompare)
    result = replace(result, "{taillabel}", data.taillabel, 1, -1, vbTextCompare)
    result = replace(result, "{headlabel}", data.headlabel, 1, -1, vbTextCompare)

    ExpandLabelPlaceholders = result
End Function

' ==========================================================================
' ROUTINE: IsLabelHTMLLike
'
' PURPOSE:
'   Performs a fast, lightweight determination of whether a label appears to
'   contain HTML-like content. Used to decide whether Graphviz should receive
'   the label as an HTML string ("<...>") or as plain text. This routine does
'   not validate full HTML correctness; it simply detects structural cues that
'   strongly indicate HTML formatting.
'
' FUNCTIONAL WORKFLOW:
'   1. NORMALIZATION:
'        - Removes newline characters to ensure the label is evaluated as a
'          single line, matching Graphviz's HTML-label expectations.
'
'   2. WRAPPER CHECK:
'        - HTML-like labels must begin with "<" and end with ">".
'        - Early exits are used for performance.
'
'   3. INNER-CONTENT EXTRACTION:
'        - Extracts the substring between the outer "<" and ">" delimiters.
'        - Trims whitespace to avoid false negatives.
'
'   4. HTML-LIKENESS TEST:
'        - Searches for common closing-tag patterns ("</" or "/>").
'        - Presence of either indicates HTML-style markup.
'        - This is intentionally heuristic: invalid HTML will still render
'          visibly incorrect in the diagram, prompting user correction.
'
'   5. OUTPUT:
'        - Returns True when the label is likely HTML-formatted.
'        - Returns False otherwise.
'
' TECHNICAL NOTES:
'   - Designed for speed: minimal string scanning, early exits, and no
'     expensive parsing.
'   - Complements GetLabelType, which selects "html" vs "text" based on this
'     routine.
'   - DeepWiki Context: Implements the label-classification rules described in
'     the "Labels", "Formatting", and "Graphviz Serialization" sections.
' ==========================================================================
Public Function IsLabelHTMLLike(ByVal label As String) As Boolean
    Dim singleLine As String
    singleLine = replace(label, Chr$(10), vbNullString)

    ' Must start with "<" and end with ">"
    If Not StartsWith(singleLine, LESS_THAN) Then Exit Function
    If Not EndsWith(singleLine, GREATER_THAN) Then Exit Function

    ' Extract inner content
    Dim inner As String
    inner = Trim$(GetStringBetweenDelimiters(singleLine, LESS_THAN, GREATER_THAN))

    ' Fast HTML-likeness check: look for closing tags
    If InStr(inner, "</") > 0 Or InStr(inner, "/>") > 0 Then
        IsLabelHTMLLike = True
    End If
End Function

' ==========================================================================
' FUNCTION: GetPort
'
' PURPOSE:
'   Extracts the port component from a Graphviz node reference of the form
'   "node:port". Returns the substring following the first colon, or an
'   empty string when no port is present.
'
' FUNCTIONAL WORKFLOW:
'   1. COLON SEARCH:
'        - Scans the input string for the first colon using binary comparison.
'
'   2. PORT EXTRACTION:
'        - When a colon is found:
'             o Returns all characters following the colon.
'        - When no colon is found:
'             o Returns an empty string, indicating that no port is defined.
'
' TECHNICAL NOTES:
'   - This routine performs no validation of port syntax; callers are
'     responsible for ensuring that extracted ports conform to Graphviz
'     conventions.
'   - Used by edge-processing routines to support tailport/headport
'     assignment when includeEdgePorts is enabled.
'   - DeepWiki Context: Implements the port-extraction rules described in
'     the "Edges" and "Syntax" documentation.
' ==========================================================================
Public Function GetPort(ByVal s As String) As String
    Dim pos As Long
    pos = InStr(1, s, ":", vbBinaryCompare)

    If pos > 0 Then
        GetPort = Mid$(s, pos + 1)
    Else
        GetPort = vbNullString
    End If
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
Public Function RemovePort(ByVal nodeId As String) As String
    
    ' Strip off the port (if specified)
    If InStr(nodeId, ":") > 0 Then
        RemovePort = GetStringTokenAtPosition(nodeId, ":", 1)
    Else
        RemovePort = nodeId
    End If

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
Public Function FileLocationProvided(ByRef output As FileOutput) As Boolean
    FileLocationProvided = True
    
    ' Validate that the output directory exists
    If Not DirectoryExists(output.directory) Then
        EmitMessage replace(GetMessage("msgboxDirDoesNotExist"), "{dir}", output.directory), buttons:=vbCritical
        FileLocationProvided = False
    End If

    ' Get the base value of the file name
    If output.fileNamePrefix = vbNullString Then
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

    ' Get file output settings
    Dim output As FileOutput
    output = GetSettingsForFileOutput()
    
    ' Build up the file name from the user-specified prefix
    Dim fileBase As String
    fileBase = output.fileNamePrefix
    
    ' Include Timestamp if desired
    If output.appendTimeStamp Then
        If InStr(fileBase, "%D") Or InStr(fileBase, "%T") Then
            ' Substitute date for %D
            If InStr(fileBase, "%D") Then
                fileBase = replace(fileBase, "%D", output.date)
            End If
            
            ' Substitute time for %D
            If InStr(fileBase, "%T") Then
                fileBase = replace(fileBase, "%T", output.time)
            End If
        Else
            fileBase = fileBase & " " & output.date & " " & output.time
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
    If output.appendOptions Then
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
Public Function DetermineStyleName(ByRef ini As settings, ByVal row As Long) As String

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
                If UCase$(dataItem) = KEYWORD_NODE Or UCase$(dataItem) = KEYWORD_EDGE Or UCase$(dataItem) = KEYWORD_GRAPH Then
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
' SECTION: DATA MAPPING & SYNTACTIC HELPERS
' ==========================================================================


' ==========================================================================
' ROUTINE: CacheEnabledStyles
'
' PURPOSE:
'   Builds a high-speed, in-memory cache of all styles marked "Yes" for the
'   currently active View column. Eliminates repeated worksheet lookups during
'   the main generation loop by preloading Style objects into a Dictionary
'   keyed by uppercase style name. Skips commented rows, ignores duplicates,
'   and constructs fully populated Style instances via GetStyle.
'
' FUNCTIONAL WORKFLOW:
'   1. DICTIONARY INITIALIZATION:
'        - Allocates a new Dictionary to hold enabled styles.
'        - Keys:   UCase$(styleName)
'        - Values: style objects created via GetStyle.
'
'   2. WORKSHEET SCAN:
'        - Iterates from ini.styles.firstRow to ini.styles.lastRow.
'        - Reads the flag column to determine row status.
'
'   3. COMMENT FILTER:
'        - Rows with FLAG_COMMENT in the flag column are ignored entirely.
'        - Allows the Styles sheet to contain notes, separators, or disabled
'          entries without affecting the cache.
'
'   4. VIEW-COLUMN FILTER:
'        - Only rows with TOGGLE_YES in showStyleColumn are considered enabled
'          for the current View.
'        - Supports multiple Views (columns) without duplicating logic.
'
'   5. STYLE INSTANTIATION:
'        - Retrieves the style name, type, format, and description from the
'          corresponding columns.
'        - Converts the style name to uppercase for consistent dictionary keys.
'        - Skips empty names and ignores duplicates already present in the
'          dictionary.
'        - Constructs a new style object via GetStyle and stores it in the
'          dictionary.
'
'   6. OUTPUT:
'        - Returns the populated Dictionary containing all enabled styles for
'          the active View.
'
' TECHNICAL NOTES:
'   - Dictionary lookups are O(1), dramatically reducing overhead during
'     generation of large graphs.
'   - Centralizes style-loading logic so future changes to the Styles sheet
'     structure require updates in only one place.
'   - DeepWiki Context: Implements the style-caching rules described in the
'     "Styles", "Formatting", and "Generation Pipeline" sections.
' ==========================================================================
Public Function CacheEnabledStyles(ByRef ini As settings, ByVal showStyleColumn As Long) As Dictionary

    ' Dictionary to hold the key and associated values
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    ' Loop through the specified range
    Dim row As Long
    Dim styleName As String
    
    For row = ini.styles.firstRow To ini.styles.lastRow
        If StylesSheet.Cells.item(row, ini.styles.flagColumn).value = FLAG_COMMENT Then
            ' Comment row, ignore it
        ElseIf StylesSheet.Cells.item(row, showStyleColumn).value = TOGGLE_YES Then
            ' Retrieve the style name
            styleName = UCase$(StylesSheet.Cells.item(row, ini.styles.nameColumn).value)

            If styleName <> vbNullString Then    ' a style name is present
                If Not dictionaryObj.Exists(styleName) Then ' ignore duplicate style names
                    Set dictionaryObj.item(styleName) = GetStyle(StylesSheet.Cells.item(row, ini.styles.nameColumn), _
                                                                 StylesSheet.Cells.item(row, ini.styles.typeColumn), _
                                                                 StylesSheet.Cells.item(row, ini.styles.formatColumn), _
                                                                 StylesSheet.Cells.item(row, ini.styles.descriptionColumn))
                End If
            End If
        End If
    Next row

    Set CacheEnabledStyles = dictionaryObj
    
End Function

' ==========================================================================
' ROUTINE: GetStyle
'
' PURPOSE:
'   Constructs and returns a fully populated style object. Centralizes the
'   creation of style records so callers can instantiate a style with a
'   single, declarative call rather than manually setting properties each
'   time. Ensures consistent initialization across all style-related
'   subsystems (formatting, serialization, and UI display).
'
' FUNCTIONAL WORKFLOW:
'   1. OBJECT INSTANTIATION:
'        - Allocates a new style instance.
'
'   2. PROPERTY ASSIGNMENT:
'        - Sets the following fields on the newly created object:
'             o styleName        - logical identifier
'             o styleType        - category (node, edge, cluster, etc.)
'             o styleFormat      - Graphviz format string or template
'             o styleDescription - human-readable description
'
'   3. OUTPUT:
'        - Returns the fully initialized style object to the caller.
'
' TECHNICAL NOTES:
'   - This routine acts as a lightweight factory method for the style class.
'   - Ensures consistent initialization even if the style class later gains
'     additional fields or validation logic.
'   - DeepWiki Context: Supports the style-construction rules described in the
'     "Styles", "Formatting", and "Serialization" sections.
' ==========================================================================
Private Function GetStyle(ByVal styleName As String, ByVal styleType As String, ByVal styleFormat As String, ByVal styleDescription) As style

    Dim value As style
    Set value = New style
        
    value.styleName = styleName
    value.styleType = styleType
    value.styleFormat = styleFormat
    value.styleDescription = styleDescription
    
    Set GetStyle = value

End Function


' ==========================================================================
' SECTION: DATA SOURCE RESOLUTION & VALIDATION
' ==========================================================================

' ==========================================================================
' ROUTINE: GetDataWorksheetName
'
' PURPOSE:
'   Determines the worksheet to use for graph-data operations. Returns the
'   active worksheet name when it is a valid data worksheet; otherwise falls
'   back to the canonical DataSheet. Uses a fast Select Case block to filter
'   out non-data sheets, then validates the worksheet's layout by comparing
'   key column headings against the DataSheet template.
'
' FUNCTIONAL WORKFLOW:
'   1. ACTIVE WORKSHEET IDENTIFICATION:
'        - Retrieves the name of the currently active worksheet.
'
'   2. BLOCKED WORKSHEET FILTER:
'        - Uses a Select Case block to match the active sheet against the
'          complete list of worksheets that cannot contain graph data.
'        - If matched, immediately returns DataSheet.Name.
'        - Eliminates long OR chains and improves readability and performance.
'
'   3. LAYOUT VALIDATION:
'        - For non-blocked sheets, retrieves layout metadata via
'          GetSettingsForDataWorksheet.
'        - Compares the worksheet's key headings (item, label, isRelatedToItem)
'          against the corresponding headings in DataSheet.
'        - If any heading differs, the worksheet is considered incompatible
'          and the routine falls back to DataSheet.Name.
'
'   4. OUTPUT:
'        - Returns either the validated worksheet name or DataSheet.Name when
'          the sheet is blocked or layout-incompatible.
'
' TECHNICAL NOTES:
'   - Select Case provides the fastest equality-matching path in VBA.
'   - Heading comparison ensures structural compatibility without scanning
'     the entire worksheet.
'   - DeepWiki Context: Implements the worksheet-selection rules described in
'     the "Data Worksheets", "Validation", and "Graph Input Pipeline" sections.
' ==========================================================================
Public Function GetDataWorksheetName() As String

    Dim wsName As String
    wsName = ActiveSheet.name

    ' ----------------------------------------------------------------------
    ' BLOCKED WORKSHEETS
    ' If the active sheet is one of the non-data sheets, immediately return
    ' the canonical DataSheet name.
    ' ----------------------------------------------------------------------
    Select Case wsName
        Case DataSheet.name, _
             AboutSheet.name, _
             ChoicesSheet.name, _
             ConsoleSheet.name, _
             DiagnosticsSheet.name, _
             GraphSheet.name, _
             HelpAttributesSheet.name, _
             HelpColorsSheet.name, _
             HelpShapesSheet.name, _
             ListsSheet.name, _
             LocaleDeDeSheet.name, _
             LocaleEnGbSheet.name, _
             LocaleEnUsSheet.name, _
             LocaleFrFrSheet.name, _
             LocaleItItSheet.name, _
             LocalePlPlSheet.name, _
             SettingsSheet.name, _
             SourceSheet.name, _
             SqlSheet.name, _
             StyleDesignerSheet.name, _
             StylesSheet.name, _
             ViewerAssetsSheet.name

            GetDataWorksheetName = DataSheet.name
            Exit Function
    End Select

    ' ----------------------------------------------------------------------
    ' LAYOUT VALIDATION
    ' Ensure the worksheet has the same column headings as the DataSheet.
    ' If any key heading differs, fall back to DataSheet.
    ' ----------------------------------------------------------------------
    Dim data As dataWorksheet
    data = GetSettingsForDataWorksheet(wsName)

    If GetCell(wsName, data.headingRow, data.itemColumn) <> _
       DataSheet.Cells.item(data.headingRow, data.itemColumn).value _
       Or GetCell(wsName, data.headingRow, data.labelColumn) <> _
       DataSheet.Cells.item(data.headingRow, data.labelColumn).value _
       Or GetCell(wsName, data.headingRow, data.isRelatedToItemColumn) <> _
       DataSheet.Cells.item(data.headingRow, data.isRelatedToItemColumn).value Then

        wsName = DataSheet.name
    End If

    GetDataWorksheetName = wsName
End Function


