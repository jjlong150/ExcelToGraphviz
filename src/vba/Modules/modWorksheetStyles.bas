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
    
    Dim graphvizSource As String
    Select Case styleType
        Case TYPE_NODE
            graphvizSource = "digraph preview { bgcolor=transparent imagepath=" & AddQuotes(GetImagePath()) & " " & AddQuotes(styleName) & " [label=" & AddQuotes(replace(styleName, " ", "\n")) & " " & StylesSheet.Cells.item(row, styles.formatColumn).value & "] }"
        Case TYPE_EDGE
            graphvizSource = "digraph preview { bgcolor=transparent layout=dot rankdir=LR tail[shape=point color=invis]; head[shape=point color=invis]; tail->head[label=" & AddQuotes(styleName) & " " & StylesSheet.Cells.item(row, styles.formatColumn).value & "] }"
        Case TYPE_SUBGRAPH_OPEN
            graphvizSource = "digraph preview { bgcolor=transparent layout=dot rankdir=LR subgraph cluster_1 { label=" & AddQuotes(styleName) & " " & StylesSheet.Cells.item(row, styles.formatColumn).value & " node[style=filled fillcolor=white]; A->Z; } }"
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
        StylesSheet.rows.item(row).AutoFit
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
    Dim shapeObject As shape
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

Private Sub ApplyLabelSetting(ByVal textCell As String, ByVal includeFlagCell As String, ByVal attributeValue As String)
    With StyleDesignerSheet
        .Range(textCell).value = attributeValue
        .Range(includeFlagCell).value = True
    End With
End Sub

Private Sub ApplyPenWidthSetting(ByVal attributeValue As String, ByVal mode As String)
    Select Case LCase$(mode)
        Case TYPE_EDGE
            StyleDesignerSheet.Range(DESIGNER_EDGE_PEN_WIDTH).value = attributeValue

        Case TYPE_NODE, TYPE_CLUSTER
            StyleDesignerSheet.Range(DESIGNER_BORDER_PEN_WIDTH).value = attributeValue
    End Select
End Sub

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

Private Sub ApplyStyleValue(ByVal attributeValue As String, ByVal mode As String)
    Select Case LCase$(mode)
        Case TYPE_NODE, TYPE_CLUSTER
            Call ApplyNodeStyles(attributeValue)

        Case TYPE_EDGE
            StyleDesignerSheet.Range(DESIGNER_EDGE_STYLE).value = attributeValue
    End Select
End Sub

