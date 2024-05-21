Attribute VB_Name = "modWorksheetStyleDesigner"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

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
    labels.label = Trim$(StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT).Value)
    Select Case elementType
        Case KEYWORD_NODE
            labels.xLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).Value)
            labels.headLabel = vbNullString
            labels.tailLabel = vbNullString
        Case KEYWORD_EDGE
            labels.xLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).Value)
            labels.headLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_HEAD_LABEL_TEXT).Value)
            labels.tailLabel = Trim$(StyleDesignerSheet.Range(DESIGNER_TAIL_LABEL_TEXT).Value)
        Case KEYWORD_CLUSTER
            labels.xLabel = vbNullString
            labels.headLabel = vbNullString
            labels.tailLabel = vbNullString
        End Select

    If createFormat Then
        ' Clear the Style cell (can't use .ClearContents on merged cells)
        StyleDesignerSheet.Range(formatCellName).Value = vbNullString
        
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
        StyleDesignerSheet.Range(formatCellName).Value = styleAttributes
    Else
        ' The user has composed/edited the format. Use the value in the format cell
        styleAttributes = Trim$(StyleDesignerSheet.Range(formatCellName).Value)
    End If
    
    ' Get the user-specified cell where the preview image should be displayed
    previewCell = Trim$(StyleDesignerSheet.Range(previewCellName).Value)
    If previewCell <> vbNullString Then
        
        ' Find out if the user wants the graph options included in the preview
        If StyleDesignerSheet.Range(DESIGNER_ADD_CAPTION).Value = TOGGLE_YES Then
            addCaption = True
        End If
        
        ' Create the Graphviz statements which can preview the style
        dotSource = GeneratePreviewGraph(elementType, labels, styleAttributes, addCaption)
        
        ' Generate the image, and display it at the location specified
        PreviewStyle dotSource, previewCell
    End If

End Sub

Public Function GeneratePreviewGraph(ByVal elementType As String, _
                                     ByRef labels As LabelSet, _
                                     ByVal styleAttributes As String, _
                                     ByVal addCaption As Boolean) As String

    ' Graph Options section
    Dim direction As String
    
    Dim layout As String
    layout = SettingsSheet.Range(SETTINGS_GRAPHVIZ_ENGINE).Value
    
    Dim graphSplines As String
    graphSplines = SettingsSheet.Range(SETTINGS_SPLINES).Value
    
    Dim splines As String
    If graphSplines <> vbNullString Then
        AddNameValue splines, "splines", graphSplines
    End If
    
    Dim graphOptions As String
    AddNameValue graphOptions, "layout", layout
    graphOptions = graphOptions & " " & SettingsSheet.Range(SETTINGS_GRAPH_OPTIONS).Value & SEMICOLON & vbNewLine
    
    ' Tweak the graph options to give the previews a tiny border
    AddNameValue graphOptions, "pad", "0.03125,0.03125"
    
    ' If the graphing layout is "dot" add in the direction specification
    If layout = LAYOUT_DOT Then
        direction = SettingsSheet.Range(SETTINGS_RANKDIR).Value
        AddNameValue graphOptions, "rankdir", direction
        AddNameValue graphOptions, "ranksep", "1.25"
    End If
    
    AddNameValue graphOptions, "imagepath", GetImagePath()

    ' =====================================================================
    ' Convert the data to graphviz format
    ' =====================================================================
    
    Dim dotSource As String
    dotSource = "digraph g{" & splines & graphOptions & vbNewLine
   
    If addCaption Then
        dotSource = dotSource & " " & GetPreviewCaption(elementType, layout, graphSplines, direction) & vbNewLine
    End If

    If elementType = KEYWORD_NODE Then
        dotSource = dotSource & "  " & AddQuotes("node1") & " [" & FormatLabel("label", labels.label) & FormatOptionalLabel("xlabel", labels.xLabel) & " " & styleAttributes & " ];" & vbNewLine
        
    ElseIf elementType = KEYWORD_EDGE Then
        dotSource = dotSource & GetPreviewNodeEdge(GetPreviewNodeStyle("gray", "gray"))
        dotSource = dotSource & " [" & FormatLabel("label", labels.label) & FormatOptionalLabel("xlabel", labels.xLabel) & FormatOptionalLabel("headlabel", labels.headLabel) & FormatOptionalLabel("taillabel", labels.tailLabel) & " " & styleAttributes & " ];" & vbNewLine
        
    ElseIf elementType = KEYWORD_CLUSTER Then
        dotSource = dotSource & " subgraph cluster_1{ " & vbNewLine
        dotSource = dotSource & styleAttributes & FormatLabel("label", labels.label) & " " & vbNewLine
        dotSource = dotSource & GetPreviewNodeEdge(GetPreviewNodeStyle("black", "black"))
        dotSource = dotSource & CLOSE_BRACE & vbNewLine
    End If
    
    dotSource = dotSource & CLOSE_BRACE

    GeneratePreviewGraph = dotSource
    
End Function

Private Function FormatLabel(ByVal labelName As String, ByVal labelValue As String) As String
    If IsLabelHTMLLike(labelValue) Then
        FormatLabel = " " & labelName & "=" & labelValue
    Else
        FormatLabel = " " & labelName & "=" & AddQuotes(labelValue)
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
    GetPreviewNodeEdge = GetPreviewNodeEdge & "  " & AddQuotes("TAIL") & "[" & nodeStyle & "];" & vbNewLine
    GetPreviewNodeEdge = GetPreviewNodeEdge & "  " & AddQuotes("HEAD") & "[" & nodeStyle & "];" & vbNewLine
    GetPreviewNodeEdge = GetPreviewNodeEdge & "  " & AddQuotes("TAIL") & "->" & AddQuotes("HEAD")
End Function

Public Function GetPreviewCaption(ByVal elementType As String, ByVal layout As String, ByVal graphSplines As String, ByVal direction As String) As String

    Dim label As String
    label = elementType & "\l\lLayout: " & layout & " \lSplines: " & graphSplines & "\l"
    
    If layout = LAYOUT_DOT Then
        label = label & "Direction: " & direction & "\l"
    End If
    
    If elementType = KEYWORD_CLUSTER Then
        If layout = LAYOUT_CIRCO Or layout = LAYOUT_NEATO Or layout = LAYOUT_PATCHWORK Or layout = LAYOUT_SFDP Or layout = LAYOUT_TWOPI Then
            label = label & "\lNOTE: '" & layout & "' layout does not support clusters.\l"
        End If
    End If

    Dim caption As String
    caption = AddQuotes("legend") & "["
    AddNameValue caption, "shape", "plaintext"
    AddNameValue caption, "fontname", "Arial"
    AddNameValue caption, "fontsize", "10"
    AddNameValue caption, "label", label
    GetPreviewCaption = caption & "];"
    
End Function

Public Function GetPreviewNodeStyle(ByVal pencolor As String, ByVal fontColor As String) As String

    Dim styleAttributes As String
    
    AddNameValue styleAttributes, "shape", "polygon"
    AddNameValue styleAttributes, "sides", "8"
    AddNameValue styleAttributes, "color", pencolor
    AddNameValue styleAttributes, "fixedsize", "true"
    AddNameValue styleAttributes, "fontname", "Arial"
    AddNameValue styleAttributes, "fontsize", "10"
    AddNameValue styleAttributes, "fontcolor", fontColor
    AddNameValue styleAttributes, "height", "0.50"
    AddNameValue styleAttributes, "width", "0.50"
    AddNameValue styleAttributes, "style", "filled"
    AddNameValue styleAttributes, "fillcolor", "white"

    GetPreviewNodeStyle = styleAttributes
End Function

Public Sub PreviewStyle(ByVal dotSource As String, ByVal targetCell As String)
    
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    ' The type of image to generate
    Dim outputFormat As String
    outputFormat = "png"
    
    ' =====================================================================
    ' Prepare the output files
    ' =====================================================================
    
    Dim outputDirectory As String
    outputDirectory = GetTempDirectory()
    
    ' File name variables
    Dim graphvizFile As String
    Dim diagramFile As String
    
    graphvizFile = outputDirectory & Application.pathSeparator & "PreviewStyle.gv"
    diagramFile = outputDirectory & Application.pathSeparator & "PreviewStyle." & outputFormat
          
    ' Remove any image from a previous run of the macro
    DeleteAllPictures StyleDesignerSheet.name
      
    ' -----------------------------------------------------------------------------
    ' Obtain the run time parameters from the 'settings' worksheet
    ' -----------------------------------------------------------------------------
    
    ' "Graph Options" section
    Dim graphEngine As String
    graphEngine = GetGraphvizEngine()
   
    ' "Command Line Options" section
    Dim commandLineParameters As String
    commandLineParameters = SettingsSheet.Range(SETTINGS_COMMAND_LINE_PARAMETERS).Value
        
    ' Write the Graphviz data to a file so it can be sent to a rendering engine
#If Mac Then
    WriteTextToFile dotSource, graphvizFile
#Else
    WriteTextToUTF8FileFileWithoutBOM dotSource, graphvizFile
#End If
    
    ' Generate an image using grapviz
    Dim returnCode As Long
    returnCode = CreateGraphDiagram(graphvizFile, diagramFile, outputFormat, _
                                    graphEngine, commandLineParameters, 60000)
    Select Case returnCode
        Case ShellAndWaitResult.success
            ' Display the generated image
            '@Ignore VariableNotUsed
            Dim shapeObject As Shape
            '@Ignore AssignmentNotUsed
            Set shapeObject = InsertPicture(diagramFile, ActiveSheet.Range(targetCell), False, True)
            Set shapeObject = Nothing
       
        Case ShellAndWaitResult.timeout
            MsgBox GetMessage("msgboxShellAndWaitTimeout"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        
        Case Else
            MsgBox GetMessage("msgboxGraphvizCommandFailed"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End Select
              
    ' Delete the temporary files
    DeleteFile graphvizFile
    DeleteFile diagramFile
    
End Sub

Private Sub AddAttribute(ByRef styleAttributes As String, _
                            ByVal attrName As String, _
                            ByVal cellName As String)
    ' Get the cell value
    Dim cellValue As String
    cellValue = Trim$(StyleDesignerSheet.Range(cellName).Value)
    
    If cellValue <> vbNullString Then
        styleAttributes = styleAttributes & " " & attrName & "=" & AddQuotes(cellValue)
    End If

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
    cellValue = Trim$(StyleDesignerSheet.Range(cellName1).Value)
    
    If cellValue <> vbNullString Then
        ' Start building the group attribute
        styleAttributes = styleAttributes & " " & attrName & "=" & Chr$(34) & cellValue
    
        ' Get the second attribute of the group
        cellValue = Trim$(StyleDesignerSheet.Range(cellName2).Value)
        
        ' Add to set of attributes if not blank
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & separator & cellValue
    
            ' Get the third group attribute
            cellValue = Trim$(StyleDesignerSheet.Range(cellName3).Value)
            
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
    
    ' Get first style attribute. If blank, ignore the others
    cellValue = Trim$(StyleDesignerSheet.Range(cellName1).Value)
    
    If cellValue <> vbNullString Then
        ' Start building the style attribute
        styleAttributes = styleAttributes & " style=" & Chr$(34) & cellValue
    
        ' Get the second style attribute
        cellValue = Trim$(StyleDesignerSheet.Range(cellName2).Value)
        
        ' Add to set of styles if not blank
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & COMMA & cellValue
    
            ' Get the third style attribute
            cellValue = Trim$(StyleDesignerSheet.Range(cellName3).Value)
            
            ' Add to set of styles if not blank
            If cellValue <> vbNullString Then
                styleAttributes = styleAttributes & COMMA & cellValue
            End If
        End If
        
        ' If a fill color attribute was specified, a value of "filled" or "radial" must be included
        ' as one of the values in the 'style' attribute.
        If gradientType <> vbNullString Then
            styleAttributes = styleAttributes & COMMA & gradientType
        End If
    
        ' Close the double quotes around the style attributes
        styleAttributes = styleAttributes & Chr$(34)
    
        ' Even though the style attributes are blank, we still need to return a style attribute if a
        ' fill color was specified elsewhere. gradientType will tell us if this is required.
    ElseIf gradientType <> vbNullString Then
        styleAttributes = styleAttributes & " style=" & AddQuotes(gradientType)
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
    fillColor = Trim$(StyleDesignerSheet.Range(cellNameFillColor1).Value)
    If fillColor <> vbNullString Then
        ' Since we have a fill color, check for a gradient color
        gradientColor = Trim$(StyleDesignerSheet.Range(cellNameFillColor2).Value)
        If gradientColor <> vbNullString Then
            Dim gradientWeight As String
            gradientWeight = Trim$(StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_WEIGHT).Value)
            If gradientWeight <> vbNullString Then
                gradientWeight = ";0." & Right$("00" & gradientWeight, 2)
            End If
            fillColor = fillColor & gradientWeight & ":" & gradientColor
            
            ' Gradient angle attribute
            AddAttribute styleAttributes, "gradientangle", cellNameGradientAngle
        End If
        
        ' Complete the attribute statement and return
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
    cellValue = Trim$(StyleDesignerSheet.Range(cellNameGradientType).Value)
    
    If cellValue <> vbNullString Then
        GetGradientType = cellValue
    
        ' Gradient type cell is empty. If a gradient fill color has been specified, return "filled"
    ElseIf Trim$(StyleDesignerSheet.Range(cellNameFillColor2).Value) <> vbNullString Then
        GetGradientType = "filled"
    
        ' Gradient type cell is empty. If a fill color has been specified, return "filled"
    ElseIf Trim$(StyleDesignerSheet.Range(cellNameFillColor1).Value) <> vbNullString Then
        GetGradientType = "filled"
    End If
    
End Function

Private Function ConvertMillimetersToInches(ByVal mm As String) As String
    Dim inches As Double
    inches = CDbl(mm) / 25.4
    ConvertMillimetersToInches = CStr(format(inches, "#0.0000"))
End Function
Private Function GetNodeStyle() As String
    
    Dim styleAttributes As String
    
    ' Label attributes
    AddAttribute styleAttributes, "labelloc", DESIGNER_LABEL_LOCATION
    
    ' Color Scheme
    AddAttribute styleAttributes, "colorscheme", DESIGNER_COLOR_SCHEME
    
    ' Shape attributes
    AddAttribute styleAttributes, "shape", DESIGNER_NODE_SHAPE

    ' If the shape is 'polygon', get the number of polygon sides
    If Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_SHAPE).Value) = "polygon" Then
        AddAttribute styleAttributes, "sides", DESIGNER_NODE_SIDES
        AddAttribute styleAttributes, "skew", DESIGNER_NODE_SKEW
        AddAttribute styleAttributes, "distortion", DESIGNER_NODE_DISTORTION
        AddAttribute styleAttributes, "regular", DESIGNER_NODE_REGULAR
    End If
    
    ' If metric units were specified, they have to be converted to inches as that is what Graphviz expects
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value = TOGGLE_YES Then
        Dim cellValue As String
        cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_HEIGHT).Value)
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & " height=" & AddQuotes(ConvertMillimetersToInches(cellValue))
        End If

        cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_WIDTH).Value)
        If cellValue <> vbNullString Then
            styleAttributes = styleAttributes & " width=" & AddQuotes(ConvertMillimetersToInches(cellValue))
        End If
    Else
        AddAttribute styleAttributes, "height", DESIGNER_NODE_HEIGHT
        AddAttribute styleAttributes, "width", DESIGNER_NODE_WIDTH
    End If
    
    AddAttribute styleAttributes, "fixedsize", DESIGNER_NODE_FIXED_SIZE
    AddAttribute styleAttributes, "orientation", DESIGNER_NODE_ORIENTATION
    
    ' Fill Color attributes
    AddFillColorAttribute styleAttributes, DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_ANGLE
    
    ' Border attributes
    AddAttribute styleAttributes, "color", DESIGNER_BORDER_COLOR
    AddAttribute styleAttributes, "penwidth", DESIGNER_BORDER_PEN_WIDTH
    AddAttribute styleAttributes, "peripheries", DESIGNER_BORDER_PERIPHERIES
 
    ' Font attributes
    AddFontNameAttribute styleAttributes
    AddAttribute styleAttributes, "fontsize", DESIGNER_FONT_SIZE
    AddAttribute styleAttributes, "fontcolor", DESIGNER_FONT_COLOR
      
    ' Image attributes
    AddAttribute styleAttributes, "image", DESIGNER_NODE_IMAGE_NAME
    AddAttribute styleAttributes, "imagescale", DESIGNER_NODE_IMAGE_SCALE
    AddAttribute styleAttributes, "imagepos", DESIGNER_NODE_IMAGE_POSITION
    
    ' Style attributes
    AddStyleAttribute styleAttributes, DESIGNER_BORDER_STYLE1, DESIGNER_BORDER_STYLE2, DESIGNER_BORDER_STYLE3, _
                      GetGradientType(DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_TYPE)
    
    ' Return the finished string of style attributes
    GetNodeStyle = Trim$(styleAttributes)
    
End Function

Private Function GetFontStyle(ByVal boldCell As String, ByVal italicCell As String) As String

    Dim fontStyle As String
    fontStyle = StyleDesignerSheet.Range(DESIGNER_FONT_NAME).Value
    
    If StyleDesignerSheet.Range(boldCell).Value = TOGGLE_YES Then
        fontStyle = fontStyle & " Bold"
    End If
    
    If StyleDesignerSheet.Range(italicCell).Value = TOGGLE_YES Then
        fontStyle = fontStyle & " Italic"
    End If
    
    GetFontStyle = Trim$(fontStyle)

End Function

Private Function GetEdgeStyle() As String

    Dim styleAttributes As String
    
    styleAttributes = vbNullString

    GetEdgeStyle = styleAttributes
    
    ' Color Scheme
    AddAttribute styleAttributes, "colorscheme", DESIGNER_COLOR_SCHEME
    
    ' Style attributes
    AddAttribute styleAttributes, "style", DESIGNER_EDGE_STYLE
    AddAttributeGroup styleAttributes, "color", DESIGNER_EDGE_COLOR_1, DESIGNER_EDGE_COLOR_2, DESIGNER_EDGE_COLOR_3, ":"
    AddAttribute styleAttributes, "penwidth", DESIGNER_EDGE_PEN_WIDTH
    AddAttribute styleAttributes, "dir", DESIGNER_EDGE_DIRECTION
    AddAttribute styleAttributes, "weight", DESIGNER_EDGE_WEIGHT
    
    ' Label attributes
    AddAttribute styleAttributes, "decorate", DESIGNER_EDGE_DECORATE
    AddAttribute styleAttributes, "labelangle", DESIGNER_EDGE_LABEL_ANGLE
    AddAttribute styleAttributes, "labelfloat", DESIGNER_EDGE_LABEL_FLOAT
    AddAttribute styleAttributes, "labeldistance", DESIGNER_EDGE_LABEL_DISTANCE

    ' Arrow attributes
    AddAttributeGroup styleAttributes, "arrowhead", DESIGNER_EDGE_ARROW_HEAD_1, DESIGNER_EDGE_ARROW_HEAD_2, DESIGNER_EDGE_ARROW_HEAD_3, vbNullString
    AddAttributeGroup styleAttributes, "arrowtail", DESIGNER_EDGE_ARROW_TAIL_1, DESIGNER_EDGE_ARROW_TAIL_2, DESIGNER_EDGE_ARROW_TAIL_3, vbNullString
    AddAttribute styleAttributes, "arrowsize", DESIGNER_EDGE_ARROW_SIZE

    ' Font attributes
    AddFontNameAttribute styleAttributes
    AddAttribute styleAttributes, "fontsize", DESIGNER_FONT_SIZE
    AddAttribute styleAttributes, "fontcolor", DESIGNER_FONT_COLOR
    
    ' Head/Tail label attributes
    AddAttribute styleAttributes, "labelfontname", DESIGNER_EDGE_LABEL_FONT_NAME
    AddAttribute styleAttributes, "labelfontsize", DESIGNER_EDGE_LABEL_FONT_SIZE
    AddAttribute styleAttributes, "labelfontcolor", DESIGNER_EDGE_LABEL_FONT_COLOR
    
    ' Port attributes
    AddAttribute styleAttributes, "headport", DESIGNER_EDGE_HEAD_PORT
    AddAttribute styleAttributes, "tailport", DESIGNER_EDGE_TAIL_PORT
    
    ' Clip attributes
    AddAttribute styleAttributes, "headclip", DESIGNER_EDGE_HEAD_CLIP
    AddAttribute styleAttributes, "tailclip", DESIGNER_EDGE_TAIL_CLIP
   
    GetEdgeStyle = Trim$(styleAttributes)
    
End Function

Private Sub AddFontNameAttribute(ByRef styleAttributes As String)
    Dim fontName As String
    fontName = GetFontStyle(DESIGNER_FONT_BOLD, DESIGNER_FONT_ITALIC)
    If fontName <> vbNullString Then
        styleAttributes = styleAttributes & " fontname=" & AddQuotes(fontName)
    End If
End Sub
Private Function GetClusterStyle() As String

    Dim styleAttributes As String
    
    styleAttributes = vbNullString

    GetClusterStyle = styleAttributes
    
    ' Label attributes
    AddAttribute styleAttributes, "labeljust", DESIGNER_LABEL_JUSTIFICATION
    AddAttribute styleAttributes, "labelloc", DESIGNER_LABEL_LOCATION
    
    ' Color scheme
    AddAttribute styleAttributes, "colorscheme", DESIGNER_COLOR_SCHEME
    
    ' Border attributes
    AddAttribute styleAttributes, "penwidth", DESIGNER_BORDER_PEN_WIDTH
    AddAttribute styleAttributes, "pencolor", DESIGNER_BORDER_COLOR
   
    ' Fill and Gradient Color attributes
    AddFillColorAttribute styleAttributes, DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_ANGLE

    ' Font attributes
    AddFontNameAttribute styleAttributes
    AddAttribute styleAttributes, "fontsize", DESIGNER_FONT_SIZE
    AddAttribute styleAttributes, "fontcolor", DESIGNER_FONT_COLOR
        
    ' Style attributes
    AddStyleAttribute styleAttributes, DESIGNER_BORDER_STYLE1, DESIGNER_BORDER_STYLE2, DESIGNER_BORDER_STYLE3, _
                      GetGradientType(DESIGNER_FILL_COLOR, DESIGNER_GRADIENT_FILL_COLOR, DESIGNER_GRADIENT_FILL_TYPE)
    
    GetClusterStyle = Trim$(styleAttributes)
    
End Function

Public Sub DisplayStyleDesignerRows(ByVal isVisible As Boolean)
    Dim rowFrom As Long
    Dim rowTo As Long
    Dim row As Long
    
    rowFrom = StyleDesignerSheet.Range("ColorScheme").row - 5
    rowTo = StyleDesignerSheet.Range("AddCaption").row + 3
    
    For row = rowFrom To rowTo
        StyleDesignerSheet.rows.Item(row).Hidden = Not isVisible
    Next row
End Sub

'@Ignore ProcedureNotUsed
Public Sub StyleDesignerToggleShowSettings()

    Dim s As Shape
    Set s = StyleDesignerSheet.Shapes.[_Default]("Check Box 5")    ' Brittle
 
    s.Select
    
    If Selection.Value = xlOn Then
       ShowStyleDesignerSettings
    Else
       HideStyleDesignerSettings
    End If

    Set s = Nothing
End Sub

Private Sub HideStyleDesignerSettings()
    Application.screenUpdating = False
    DisplayStyleDesignerRows False
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
    Application.screenUpdating = True
End Sub

Private Sub ShowStyleDesignerSettings()
    Application.screenUpdating = False
    DisplayStyleDesignerRows True
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
    Application.screenUpdating = True
End Sub


