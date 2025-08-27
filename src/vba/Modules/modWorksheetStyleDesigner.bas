Attribute VB_Name = "modWorksheetStyleDesigner"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Style Designer")

Option Explicit

' Uncomment code below if encontering "Runtime Error 49, Bad DLL calling convention"
' Refer to: https://stackoverflow.com/questions/15758834/runtime-error-49-bad-dll-calling-convention
'Private Enum Something
'    member = 1
'End Enum

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

    ' Only node previews show images
    If elementType = KEYWORD_NODE Then
        AddNameValue graphOptions, GRAPHVIZ_IMAGEPATH, AddQuotes(GetImagePath())
    End If
    
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

Private Function FormatLabel(ByVal labelName As String, ByVal labelValue As String) As String
    If IsLabelHTMLLike(labelValue) Then
        FormatLabel = " " & labelName & "=" & labelValue
    Else
        FormatLabel = " " & labelName & "=" & AddQuotes(ScrubText(labelValue))
    End If
End Function

Private Function FormatOptionalLabel(ByVal labelName As String, ByVal labelValue As String) As String
    If Trim$(labelValue) = vbNullString Then
        FormatOptionalLabel = vbNullString
    Else
        FormatOptionalLabel = FormatLabel(labelName, labelValue)
    End If
End Function

Public Function GetPreviewNodeEdge(ByVal nodeStyle As String) As String
    GetPreviewNodeEdge = GetPreviewNodeEdge & "    %T [" & nodeStyle & "];" & vbNewLine
    GetPreviewNodeEdge = GetPreviewNodeEdge & "    %H [" & nodeStyle & "];" & vbNewLine
    GetPreviewNodeEdge = GetPreviewNodeEdge & "    %T->%H"
End Function

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
    Set shapeObject = InsertPicture(graphvizObj.DiagramFilename, ActiveSheet.Range(targetCell), False, True)
    Set shapeObject = Nothing
                    
    ' Delete the temporary files
    DeleteFile graphvizObj.GraphvizFilename
    DeleteFile graphvizObj.DiagramFilename
    
    ' Release the Graphviz object
    Set graphvizObj = Nothing
End Sub

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

Private Function ConvertMillimetersToInches(ByVal mm As String) As String
    Dim inches As Double
    inches = CDbl(mm) / 25.4
    ConvertMillimetersToInches = CStr(format(inches, "#0.0000"))
End Function

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

Private Sub AddFontNameAttribute(ByRef styleAttributes As String)
    Dim fontName As String
    fontName = GetFontStyle(DESIGNER_FONT_BOLD, DESIGNER_FONT_ITALIC)
    If fontName = vbNullString Then Exit Sub
    
    styleAttributes = Join(Array(styleAttributes, " fontname=", AddQuotesConditionally(fontName)), vbNullString)
End Sub

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

Private Sub HideStyleDesignerSettings()
    Application.ScreenUpdating = False
    DisplayStyleDesignerRows False
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
    Application.ScreenUpdating = True
End Sub

Private Sub ShowStyleDesignerSettings()
    Application.ScreenUpdating = False
    DisplayStyleDesignerRows True
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
    Application.ScreenUpdating = True
End Sub

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


