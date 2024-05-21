Attribute VB_Name = "modRibbonTabStyleDesigner"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule UseMeaningfulName, UnassignedVariableUsage, ProcedureNotUsed, ParameterNotUsed, ImplicitByRefModifier

Option Explicit

Private fontList As Variant
Private x11Colors As Variant
Private svgColors As Variant
Private brewerColors As Variant
Private brewerColorsAreFresh As Boolean
Private fontImageDir As String
Private fontImageCache As Dictionary
Private colorImageDir As String
Private colorScheme As String
Private colorImageCache As Dictionary
Private colorCount As Long
Private fontCount As Long

' ===========================================================================
' Ribbon callbacks for "Style Designer" ribbon tab
' ===========================================================================

' ===========================================================================
' Callbacks for colorScheme

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub colorScheme_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If Left$(controlId, 4) = "cs_x" Then Exit Sub ' Blank gallery image selected
    
    OptimizeCode_Begin
    
    If index = 0 Then
        colorScheme = vbNullString
    Else
        colorScheme = Mid$(controlId, Len("cs_") + 1)
    End If
    
    ' If color scheme is not X11 or SVG then it is a Brewer color scheme.
    ' Loading the brewerColors array is deferred until the next time the array is
    ' referenced (i.e. lazy load).
    If colorScheme <> COLOR_SCHEME_X11 And colorScheme <> COLOR_SCHEME_SVG Then
        brewerColorsAreFresh = False
    End If
    
    StyleDesignerSheet.Range(DESIGNER_COLOR_SCHEME).Value = colorScheme
    StyleDesignerSheet.Range("FontColor,BorderColor,FillColor,GradientFillColor,GradientFillType,GradientFillAngle,GradientFillWeight,EdgeColor1,EdgeColor2,EdgeColor3,EdgeLabelFontColor").ClearContents
    OptimizeCode_End
    
    InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_GRP_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_ANGLE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_WEIGHT
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub colorScheme_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = GetListIndex(LISTS_COLOR_SCHEMES, DESIGNER_COLOR_SCHEME)
End Sub

' ===========================================================================
' Callbacks for fontColor

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    color_getImage DESIGNER_FONT_COLOR, COLOR_BLACK, returnedVal
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub null_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = vbNullString
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    If StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_COLOR).Value = vbNullString Then
        color_getImage DESIGNER_FONT_COLOR, COLOR_BLACK, returnedVal
    Else
        color_getImage DESIGNER_EDGE_LABEL_FONT_COLOR, COLOR_BLACK, returnedVal
    End If
#End If
End Sub

Private Sub color_getImage(ByVal cellName As String, ByVal defaultColor, ByRef returnedVal As Variant)
    On Error Resume Next
    
    ' Performance aid
    If IsProgressIndicatorNeeded() Then
        returnedVal = vbNullString
        Exit Sub
    End If

    Dim color As String
    color = StyleDesignerSheet.Range(cellName).Value
    
    ' Try to return the color image from cache
    Dim colorCacheKey As String
    
    If color = vbNullString Then
        colorCacheKey = COLOR_SCHEME_X11 & "_" & defaultColor
    Else
        colorCacheKey = colorScheme & "_" & color
    End If
    
    If colorImageCache.Exists(colorCacheKey) Then
        Set returnedVal = colorImageCache(colorCacheKey)
        Exit Sub
    End If
    
    ' Build the path to where the images are kept
    Dim imageFile As String
    imageFile = GetColorImageDir() & Application.pathSeparator & LCase$(colorCacheKey & ".bmp")

    ' If the image already exists we should be able to load it
    Set returnedVal = LoadPicture(imageFile)
    
    ' Add the loaded picture to the image cache
    colorImageCache.Add colorCacheKey, returnedVal
    
    On Error GoTo 0
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontColor_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SaveColor index, DESIGNER_FONT_COLOR
    InvalidateRibbonColorList RIBBON_CTL_FONT_COLOR
    InvalidateRibbonColorList RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontColor_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SaveColor index, DESIGNER_EDGE_LABEL_FONT_COLOR
    InvalidateRibbonColorList RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub color_getItemImage(ByVal control As IRibbonControl, ByVal index As Long, ByRef image As Variant)
        
    If index < 0 Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Dim colorName As String
    Dim interiorColor As Long
    Dim imageFile As String
    Dim scheme As String
    scheme = colorScheme
    
    If index = 0 Then   ' Determine the default color for the attribute
        If control.ID = RIBBON_CTL_FILL_COLOR Then
            interiorColor = 16777215
            colorName = COLOR_WHITE
            scheme = COLOR_SCHEME_X11
        ElseIf control.ID = RIBBON_CTL_GRADIENT_FILL_COLOR Then
            ' The default gradient fill will be white if fill color has not been set
            ' been set, otherwise make the default gradient fill the same as the fill color
            If StyleDesignerSheet.Range(DESIGNER_FILL_COLOR).Value = vbNullString Then
                interiorColor = 16777215
                colorName = COLOR_WHITE
                scheme = COLOR_SCHEME_X11
            Else
                colorName = LCase$(StyleDesignerSheet.Range(DESIGNER_FILL_COLOR).Value)
            End If
        Else
            interiorColor = 0
        End If
    Else
        ' Get the color name based on the current color scheme
        If colorScheme = COLOR_SCHEME_X11 Then
            colorName = x11Colors(index)
        ElseIf colorScheme = COLOR_SCHEME_SVG Then
            colorName = svgColors(index)
        Else
            colorName = brewerColors(index)
        End If
    End If
    
    imageFile = GetColorImageDir() & Application.pathSeparator & LCase$(scheme & "_" & colorName & ".bmp")

    ' See if the image has been previously loaded. If so, return the cached reference
    Dim colorCacheKey As String
    colorCacheKey = scheme & "_" & colorName
    
    If colorImageCache.Exists(colorCacheKey) Then
        Set image = colorImageCache(colorCacheKey)
        Exit Sub
    End If
    
    ' If the image already exists we should be able to load it
    Set image = LoadPicture(imageFile)

    If image Is Nothing Then    ' the image does not exist, create one
        UpdateStatusBar replace(GetMessage("statusbarCreateImage"), "{colorScheme}", scheme) & " " & colorName
       
        ' Get the RGB color for this color scheme index
        If index > 0 Then
            If scheme = COLOR_SCHEME_X11 Or scheme = COLOR_SCHEME_SVG Then
                ' Color list is arranged in a column of cells
                interiorColor = HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & scheme).Cells.Item(index, 1).Interior.color
            Else
                ' Color list is aranged in a row of cells
                interiorColor = HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & scheme).Cells.Item(1, index).Interior.color
            End If
        End If

        ' Generate a thumbnail image
        If color_createThumbnail(interiorColor, imageFile) Then
            ' Load the new image
            Set image = LoadPicture(imageFile)
        End If
    End If

    ' By getting this far we know the image is not in the cache. Add it.
    colorImageCache.Add colorCacheKey, image
    
    On Error GoTo 0
End Sub

Private Function color_createThumbnail(ByVal colorRGB As Long, ByVal imageFile As String) As Boolean
    color_createThumbnail = False
    On Error Resume Next
               
    Dim chartObj As Chart
    With StyleDesignerSheet
        ' Chart attributes are in points. 15 points = 20 pixels
        Set chartObj = .ChartObjects.Add(0, 0, 15, 15).Chart
                      
        ' Set the background fill color of the chart to the fill color
        ' passed to this function, then write the chart out as
        ' an image file.
        With chartObj
            .Parent.Activate
            .ChartArea.format.Fill.visible = msoTrue
            .ChartArea.format.Fill.ForeColor.RGB = colorRGB
            .ChartArea.format.line.visible = msoTrue
            .Export filename:=imageFile
            .Parent.Delete
        End With
    End With
    color_createThumbnail = True

    Set chartObj = Nothing
    On Error GoTo 0
End Function

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub color_getItemCount(ByVal control As IRibbonControl, ByRef Count As Variant)
    ' Set the global colorScheme value
    colorScheme = StyleDesignerSheet.Range(DESIGNER_COLOR_SCHEME).Value
    If colorScheme = vbNullString Then
        colorScheme = COLOR_SCHEME_DEFAULT       ' X11 is Graphviz's default color scheme
    End If
    
    ' Lazy creation of colorImageCache dictionary
    If colorImageCache Is Nothing Then
        Set colorImageCache = New Dictionary
    End If
    
    ' Lazy cache the large color lists in arrays. Is supposed to improve performance over individual cell access
    If colorScheme = COLOR_SCHEME_X11 Then
        If IsEmpty(x11Colors) Then
            x11Colors = Application.WorksheetFunction.Transpose(HelpColorsSheet.Range("CS_X11"))  ' 656 colors
        End If
        Count = (UBound(x11Colors) - LBound(x11Colors) + 2)
    ElseIf colorScheme = COLOR_SCHEME_SVG Then
        If IsEmpty(svgColors) Then
            svgColors = Application.WorksheetFunction.Transpose(HelpColorsSheet.Range("CS_SVG"))  ' 147 colors
        End If
        Count = (UBound(svgColors) - LBound(svgColors) + 2)
    Else
        If Not brewerColorsAreFresh Then
            brewerColors = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.Transpose(HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & colorScheme)))
            brewerColorsAreFresh = True
        End If
        Count = (UBound(brewerColors) - LBound(brewerColors) + 2)
    End If
    
    ' Save count for calculating percent complete
    colorCount = Count
    
    ' Hack to disable loading the hidden dropdowns
    If StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE Or StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER Then
        If control.ID = RIBBON_CTL_EDGE_COLOR1 Or control.ID = RIBBON_CTL_EDGE_COLOR2 Or control.ID = RIBBON_CTL_EDGE_COLOR3 Or control.ID = RIBBON_CTL_EDGE_LABEL_FONT_COLOR Then
            Count = 0
        End If
        
        If control.ID = RIBBON_CTL_GRADIENT_FILL_COLOR And IsEmpty(StyleDesignerSheet.Range(DESIGNER_FILL_COLOR)) Then
            Count = 0
        End If
        
    ElseIf StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_EDGE Then
        If control.ID = RIBBON_CTL_FILL_COLOR Or control.ID = RIBBON_CTL_GRADIENT_FILL_COLOR Or control.ID = RIBBON_CTL_BORDER_COLOR Then
            Count = 0
        ElseIf control.ID = RIBBON_CTL_EDGE_COLOR2 Then
            If IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_1)) Then
                Count = 0
            End If
        ElseIf control.ID = RIBBON_CTL_EDGE_COLOR3 Then
            If IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_2)) Then
                Count = 0
            End If
        End If
    End If
    
    If Count > 0 Then
        Application.StatusBar = LocalizeGetString(control.ID, LOCALE_COL_LABEL_VERBOSE)
        If IsProgressIndicatorNeeded() Then
            ShowProgressIndicator LocalizeGetString(control.ID, LOCALE_COL_LABEL_VERBOSE)
        End If
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub color_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = vbNullString
    Else
        If colorScheme = COLOR_SCHEME_X11 Then
            label = x11Colors(index)
        ElseIf colorScheme = COLOR_SCHEME_SVG Then
            label = svgColors(index)
        Else
            label = brewerColors(index)
        End If
        
        UpdateProgressIndicator ((index * 100) / colorCount)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = GetColorIndex(DESIGNER_FONT_COLOR)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = GetColorIndex(DESIGNER_EDGE_LABEL_FONT_COLOR)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

' ===========================================================================
' Callbacks for borderColor

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    color_getImage DESIGNER_BORDER_COLOR, COLOR_BLACK, returnedVal
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = GetColorIndex(DESIGNER_BORDER_COLOR)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderColor_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SaveColor index, DESIGNER_BORDER_COLOR
    InvalidateRibbonColorList RIBBON_CTL_BORDER_COLOR
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for fontName

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getItemCount(ByVal control As IRibbonControl, ByRef Count As Variant)
    
    ' Cache the list of fonts in an array
    If IsEmpty(fontList) Then
        fontList = getFontList()
    End If
    
    ' Lazy creation of fontImageCache dictionary
    If fontImageCache Is Nothing Then
        Set fontImageCache = New Dictionary
        fontImageCache.CompareMode = TextCompare
    End If
    
    If IsEmpty(fontList) Then
        Count = 0
    Else
        Count = (UBound(fontList) - LBound(fontList) + 1)
        If Count > 1000 Then    ' Microsoft caps dropdown lists at 1000 items
            Count = 1000
        End If
    End If
    
    ' Hack to disable loading the font which will not be displayed
    If StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE Or StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER Then
        If control.ID = RIBBON_CTL_EDGE_LABEL_FONT_NAME Then
            Count = 0
        End If
    End If
    
    ' Save count for updating progress indicator
    fontCount = Count
    
    If Count > 0 Then
        ShowProgressIndicator LocalizeGetString(control.ID, LOCALE_COL_LABEL_VERBOSE)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef returnedVal As Variant)
    If index = 0 Then
        returnedVal = vbNullString
    Else
       UpdateProgressIndicator ((index * 100) / fontCount)
       returnedVal = fontList(index)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_FONT_NAME).Value = getFontName(index)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef listIndex As Variant)
    listIndex = getFontIndex(StyleDesignerSheet.Range(DESIGNER_FONT_NAME).Value)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getItemImage(ByVal control As IRibbonControl, ByVal index As Long, ByRef image As Variant)
        
    If index < 0 Then Exit Sub
    
    On Error Resume Next
    
    ' Get the font name
    Dim fontName As String
    fontName = getFontName(index)
    If fontName = vbNullString Then
        fontName = "defaultFont"
    End If
    
    ' See if the font's image is already in cache
    If fontImageCache.Exists(fontName) Then
        Set image = fontImageCache(fontName)
        Exit Sub
    End If
  
    Dim outputFormat As String
    outputFormat = "jpg"
    
    ' Build the path to where the images are kept
    Dim imageFile As String
    imageFile = GetFontImageDir() & Application.pathSeparator & fontName & "." & outputFormat
    
    ' If the image already exists we should be able to load it
    Set image = LoadPicture(imageFile)

    If image Is Nothing Then    ' the image does not exist, create one
        UpdateStatusBar replace(GetMessage("statusbarCreateFontImage"), "{fontName}", fontName)
        
        If fontName_createItemImage(fontName, imageFile, outputFormat) Then
            Set image = LoadPicture(imageFile)
        End If
    End If

    ' Add the loaded font image to the cache
    fontImageCache.Add fontName, image
    
    On Error GoTo 0
End Sub

Private Function fontName_createItemImage(ByVal fontName As String, ByVal imageFile As String, ByVal outputFormat As String) As Boolean
    fontName_createItemImage = False
    On Error Resume Next
               
    ' Define a simple one node DOT graph which will create a 48x48 pixel image suitable for display in the ribbon
    Dim dotSource As String
    dotSource = "digraph g{ pad=0.01 margin=0.01 a[ shape=square style=filled fillcolor=white fontcolor=black fontsize=38 dpi=96 labelloc=top height=0.46 width=0.46 fixedsize=true penwidth=0"
    If fontName <> "defaultFont" Then
        dotSource = dotSource & " fontname=" & AddQuotes(fontName)
    End If
    dotSource = dotSource & " label=" & AddQuotes("A") & " ]; }"
    
    ' Write the Graphviz data to a file so it can be sent to a rendering engine
    Dim outputDirectory As String
    outputDirectory = GetTempDirectory()
    
    Dim graphvizFile As String
    graphvizFile = outputDirectory & Application.pathSeparator & "PreviewFontName.gv"
    
    ' Write the Graphviz data to a file so it can be sent to a rendering engine
#If Mac Then
    WriteTextToFile dotSource, graphvizFile
#Else
    WriteTextToUTF8FileFileWithoutBOM dotSource, graphvizFile
#End If
    
    ' Generate an image using grapviz
    If CreateGraphDiagram(graphvizFile, imageFile, outputFormat, LAYOUT_DOT, vbNullString, 5000) = ShellAndWaitResult.success Then
        fontName_createItemImage = True
    End If
    
    On Error GoTo 0
End Function

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontName_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_NAME).Value = getFontName(index)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontName_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef listIndex As Variant)
    listIndex = getFontIndex(StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_NAME).Value)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

Private Function getFontName(ByVal index As Long) As String
    If index = 0 Then
        getFontName = vbNullString
    Else
        getFontName = Trim$(CStr(fontList(index)))
    End If
End Function

Private Function getFontIndex(ByVal fontName As String) As Long
    getFontIndex = 0
    
    Dim index As Long
    index = 0
    
    If fontName <> vbNullString Then
        Dim font As Variant
        ' Find the font name
        For Each font In fontList
            If fontName = font Then
                getFontIndex = index
                Exit For
            End If
            index = index + 1
        Next font
    End If
End Function

' ===========================================================================
' Callbacks for fontSize

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    listSize = ListsSheet.Range(LISTS_FONT_SIZES).Count + 1
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = vbNullString
    Else
        label = ListsSheet.Range(LISTS_FONT_SIZES).Cells.Item(index, 1).Value
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        StyleDesignerSheet.Range(DESIGNER_FONT_SIZE).Value = vbNullString
    Else
        StyleDesignerSheet.Range(DESIGNER_FONT_SIZE).Value = ListsSheet.Range(LISTS_FONT_SIZES).Cells.Item(index, 1).Value
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_FONT_SIZES, DESIGNER_FONT_SIZE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontSize_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_FONT_SIZES, DESIGNER_EDGE_LABEL_FONT_SIZE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontSize_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_SIZE).Value = vbNullString
    Else
        StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FONT_SIZE).Value = ListsSheet.Range(LISTS_FONT_SIZES).Cells.Item(index, 1).Value
    End If
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeWeight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeWeight_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_WEIGHT).Value = Mid$(controlId, Len("weight_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeWeight_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_EDGE_WEIGHTS, DESIGNER_EDGE_WEIGHT)
End Sub

' ===========================================================================
' Callbacks for edgeLabelAngle

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelAngle_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_ANGLE).Value = Mid$(controlId, Len("angle_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelAngle_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_EDGE_ANGLES, DESIGNER_EDGE_LABEL_ANGLE)
End Sub

' ===========================================================================
' Callbacks for edgeLabelDistance

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelDistance_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_DISTANCE).Value = Mid$(controlId, Len("distance_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelDistance_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_EDGE_LABEL_DISTANCES, DESIGNER_EDGE_LABEL_DISTANCE)
End Sub

' ===========================================================================
' Callbacks for borderPenWidth

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPenWidth_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_BORDER_PEN_WIDTH).Value = Mid$(controlId, Len("bw_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPenWidth_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_PEN_WIDTH, DESIGNER_BORDER_PEN_WIDTH)
End Sub

' ===========================================================================
' Callbacks for borderPeripheries

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPeripheries_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_BORDER_PERIPHERIES).Value = Mid$(controlId, Len("p_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPeripheries_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_PERIPHERIES, DESIGNER_BORDER_PERIPHERIES)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPeripheries_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE
End Sub

' ===========================================================================
' Callbacks for designModeNode

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeNode_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE
    ShowLabelRows KEYWORD_NODE
    InvalidateDesignMode
    RefreshStyleDesignerRibbon
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeNode_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE
End Sub

Private Sub InvalidateDesignMode()
    InvalidateRibbonControl RIBBON_CTL_DESIGN_MODE_NODE
    InvalidateRibbonControl RIBBON_CTL_DESIGN_MODE_EDGE
    InvalidateRibbonControl RIBBON_CTL_DESIGN_MODE_CLUSTER
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_ANGLE
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_DECORATE
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_DISTANCE
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FLOAT
    InvalidateRibbonControl RIBBON_CTL_LABEL_STYLE_SEPARATOR
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_NAME
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
End Sub

' ===========================================================================
' Callbacks for designModeEdge

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeEdge_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_EDGE
    ShowLabelRows KEYWORD_EDGE
    InvalidateDesignMode
    RefreshStyleDesignerRibbon
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeEdge_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_EDGE
End Sub

' ===========================================================================
' Callbacks for designModeCluster

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeCluster_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER
    ShowLabelRows KEYWORD_CLUSTER
    InvalidateDesignMode
    RefreshStyleDesignerRibbon
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeCluster_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER
End Sub

Private Sub ShowLabelRows(ByVal designerMode As String)
    Application.screenUpdating = False
    
    Dim labelRow As Long
    Dim xlabelRow As Long
    Dim tailLabelRow As Long
    Dim headLabelRow As Long
    
    ' Get the row numbers dynamically instead of using constants in case
    ' additional rows get introduced on the worksheet or rows are rearranged
    labelRow = StyleDesignerSheet.Range("TitleStyleDesignerLabelText").row
    xlabelRow = StyleDesignerSheet.Range("TitleStyleDesignerXlabelText").row
    tailLabelRow = StyleDesignerSheet.Range("TitleStyleDesignerTailLabelText").row
    headLabelRow = StyleDesignerSheet.Range("TitleStyleDesignerHeadLabelText").row
    
    ' Show/hide rows based upon what Graphviz supports for the element
    Select Case designerMode
        Case KEYWORD_NODE
            StyleDesignerSheet.rows.Item(labelRow).Hidden = False
            StyleDesignerSheet.rows.Item(xlabelRow).Hidden = False
            StyleDesignerSheet.rows.Item(tailLabelRow).Hidden = True
            StyleDesignerSheet.rows.Item(headLabelRow).Hidden = True
        Case KEYWORD_EDGE
            StyleDesignerSheet.rows.Item(labelRow).Hidden = False
            StyleDesignerSheet.rows.Item(xlabelRow).Hidden = False
            StyleDesignerSheet.rows.Item(tailLabelRow).Hidden = False
            StyleDesignerSheet.rows.Item(headLabelRow).Hidden = False
        Case KEYWORD_CLUSTER
            StyleDesignerSheet.rows.Item(labelRow).Hidden = False
            StyleDesignerSheet.rows.Item(xlabelRow).Hidden = True
            StyleDesignerSheet.rows.Item(tailLabelRow).Hidden = True
            StyleDesignerSheet.rows.Item(headLabelRow).Hidden = True
    End Select
    
    Application.screenUpdating = True
End Sub

Private Sub RefreshStyleDesignerRibbon()
    InvalidateRibbonControl RIBBON_GRP_LABELS
    InvalidateRibbonControl RIBBON_CTL_LABEL_JUSTIFICATION
    InvalidateRibbonControl RIBBON_GRP_BORDERS
    InvalidateRibbonControl RIBBON_GRP_FILL_COLOR
    InvalidateRibbonControl RIBBON_GRP_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_GRP_NODE_SHAPE
    InvalidateRibbonControl RIBBON_CTL_BORDER_PERIPHERIES
    InvalidateRibbonControl RIBBON_GRP_NODE_DIMENSIONS
    InvalidateRibbonControl RIBBON_GRP_NODE_IMAGE
    InvalidateRibbonControl RIBBON_GRP_EDGE_STYLE
    InvalidateRibbonControl RIBBON_GRP_EDGE_COLORS
    InvalidateRibbonControl RIBBON_GRP_EDGE_HEAD_TAIL
    InvalidateRibbonControl RIBBON_GRP_EDGE_ARROW
    InvalidateRibbonControl RIBBON_GRP_EDGE_ARROW_HEAD
    InvalidateRibbonControl RIBBON_GRP_EDGE_ARROW_TAIL
    InvalidateRibbonControl RIBBON_CTL_ALIGN_BOTTOM
    InvalidateRibbonControl RIBBON_CTL_ALIGN_TOP
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_LEFT
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_RIGHT
End Sub

' ===========================================================================
' Callbacks for fillColor

'@Ignore ProcedureNotUsed
Private Sub fillColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    color_getImage DESIGNER_FILL_COLOR, COLOR_WHITE, returnedVal
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fillColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetColorIndex(DESIGNER_FILL_COLOR)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fillColor_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SaveColor index, DESIGNER_FILL_COLOR
    If StyleDesignerSheet.Range(DESIGNER_FILL_COLOR).Value = vbNullString Then
        StyleDesignerSheet.Range("GradientFillColor,GradientFillType,GradientFillAngle,GradientFillWeight").ClearContents
    End If
    
    InvalidateRibbonColorList RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_WEIGHT
    InvalidateRibbonControl RIBBON_GRP_GRADIENT_FILL_COLOR
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for gradientFillColor

'@Ignore ProcedureNotUsed
Private Sub gradientFillColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)

    ' Performance aid
    If IsProgressIndicatorNeeded() Then
        returnedVal = vbNullString
        Exit Sub
    End If

#If Mac Then
    returnedVal = vbNullString
#Else
    If StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_COLOR).Value = vbNullString Then
        ' Default gradient color to match fillColor
        color_getImage DESIGNER_FILL_COLOR, COLOR_WHITE, returnedVal
    Else
        ' Set it to the color chosen
        color_getImage DESIGNER_GRADIENT_FILL_COLOR, COLOR_WHITE, returnedVal
    End If
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetColorIndex(DESIGNER_GRADIENT_FILL_COLOR)
    HideProgressIndicator
    ClearStatusBar
    DoEvents
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillColor_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SaveColor index, DESIGNER_GRADIENT_FILL_COLOR
    If StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_COLOR).Value = vbNullString Then
        StyleDesignerSheet.Range("GradientFillType,GradientFillWeight,GradientFillAngle").ClearContents
    End If
    InvalidateRibbonColorList RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_WEIGHT
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_ANGLE
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for gradientFillType

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillType_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_TYPE).Value = Mid$(controlId, Len("ft_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillType_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_GRADIENT_TYPE, DESIGNER_GRADIENT_FILL_TYPE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillType_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_COLOR))
End Sub

' ===========================================================================
' Callbacks for gradientFillAngle

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillAngle_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_ANGLE).Value = Mid$(controlId, Len("a_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillAngle_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ANGLE, DESIGNER_GRADIENT_FILL_ANGLE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillAngle_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_COLOR))
End Sub

' ===========================================================================
' Callbacks for GradientFillWeight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillWeight_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_WEIGHT).Value = Mid$(controlId, Len("gw_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillWeight_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_COLOR_WEIGHT, DESIGNER_GRADIENT_FILL_WEIGHT)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillWeight_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_GRADIENT_FILL_COLOR))
End Sub

' ===========================================================================
' Callbacks for labelJustification

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelJustification_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER
End Sub

' ===========================================================================
' Callbacks for headPort

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadPort_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_HEAD_PORT).Value = Mid$(controlId, Len("hp_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadPort_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_PORTS, DESIGNER_EDGE_HEAD_PORT)
End Sub

' ===========================================================================
' Callbacks for tailPort

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeTailPort_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_TAIL_PORT).Value = Mid$(controlId, Len("tp_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeTailPort_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_PORTS, DESIGNER_EDGE_TAIL_PORT)
End Sub

' ===========================================================================
' Callbacks for edgeStyle

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeStyle_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_STYLE).Value = Mid$(controlId, Len("es_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeStyle_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_EDGE_STYLES, DESIGNER_EDGE_STYLE)
End Sub

' ===========================================================================
' Callbacks for nodeShape

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeShape_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_SHAPE).Value = Mid$(controlId, Len("s_") + 1)
    StyleDesignerSheet.Range("NodeSides,NodeOrientation,NodeRegular,NodeSkew,NodeDistortion").ClearContents
    InvalidateRibbonControl RIBBON_CTL_NODE_SIDES
    InvalidateRibbonControl RIBBON_CTL_NODE_REGULAR
    InvalidateRibbonControl RIBBON_CTL_NODE_ROTATION
    InvalidateRibbonControl RIBBON_CTL_POLYGON_SEPARATOR
    InvalidateRibbonControl RIBBON_CTL_NODE_SKEW
    InvalidateRibbonControl RIBBON_CTL_NODE_DISTORTION
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeShape_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_SHAPES, DESIGNER_NODE_SHAPE)
End Sub

' GetVisible callback for polygon shape

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeShape_isPolygon(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_NODE_SHAPE).Value = "polygon"
End Sub

' ===========================================================================
' Callbacks for nodeSides

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSides_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_SIDES).Value = Mid$(controlId, Len("si_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSides_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_SIDES, DESIGNER_NODE_SIDES)
End Sub

' ===========================================================================
' Callbacks for nodeRotation

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeRotation_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_ORIENTATION).Value = Mid$(controlId, Len("r_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeRotation_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ANGLE, DESIGNER_NODE_ORIENTATION)
End Sub

' ===========================================================================
' Callbacks for borderStyle1

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle1_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE1).Value = Mid$(controlId, Len("bs1_") + 1)
    If StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE1).Value = vbNullString Then
        StyleDesignerSheet.Range("BorderStyle2,BorderStyle3").ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_BORDER_STYLE2
    InvalidateRibbonControl RIBBON_CTL_BORDER_STYLE3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle1_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_STYLES, DESIGNER_BORDER_STYLE1)
End Sub

' ===========================================================================
' Callbacks for BorderStyle2

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle2_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE2).Value = Mid$(controlId, Len("bs2_") + 1)
    If StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE2).Value = vbNullString Then
        StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE3).ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_BORDER_STYLE3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle2_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_STYLES, DESIGNER_BORDER_STYLE2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE1))
End Sub

' ===========================================================================
' Callbacks for borderStyle3

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle3_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE3).Value = Mid$(controlId, Len("bs3_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle3_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_STYLES, DESIGNER_BORDER_STYLE3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_BORDER_STYLE2))
End Sub

' ===========================================================================
' Callbacks for nodeHeight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeight_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_HEIGHT).Value = Mid$(controlId, Len("h_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeight_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_SHAPE_SIZES, DESIGNER_NODE_HEIGHT)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeight_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value
        Case TOGGLE_YES
            visible = False
        Case TOGGLE_NO
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for nodeHeightMetric

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeightMetric_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_HEIGHT).Value = Mid$(controlId, Len("mmh_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeightMetric_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value = TOGGLE_YES Then
        Dim cellValue As String
        cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_HEIGHT).Value)
        If cellValue = vbNullString Then
            returnedVal = 0
        Else
            returnedVal = CInt(cellValue) + 1
        End If
    Else
        returnedVal = 0
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeightMetric_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value
        Case TOGGLE_NO
            visible = False
        Case TOGGLE_YES
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for nodeWidth

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidth_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_WIDTH).Value = Mid$(controlId, Len("w_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidth_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_SHAPE_SIZES, DESIGNER_NODE_WIDTH)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidth_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value
        Case TOGGLE_YES
            visible = False
        Case TOGGLE_NO
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for nodeWidthMetric

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidthMetric_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_WIDTH).Value = Mid$(controlId, Len("mmw_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidthMetric_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value = TOGGLE_YES Then
        Dim cellValue As String
        cellValue = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_WIDTH).Value)
        If cellValue = vbNullString Then
            returnedVal = 0
        Else
            returnedVal = CInt(cellValue) + 1
        End If
    Else
        returnedVal = 0
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidthMetric_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value
        Case TOGGLE_NO
            visible = False
        Case TOGGLE_YES
            visible = True
        Case Else
            visible = False
    End Select
End Sub

' ===========================================================================
' Callbacks for nodeFixedSize

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeFixedSize_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_FIXED_SIZE).Value = Mid$(controlId, Len("fs_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeFixedSize_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_FIXED_SIZE, DESIGNER_NODE_FIXED_SIZE)
End Sub

' ===========================================================================
' Callbacks for nodeMetric

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeMetric_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value = TOGGLE_YES
    Else
        StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value = TOGGLE_NO
    End If
    StyleDesignerSheet.Range("NodeHeight,NodeWidth").ClearContents
    InvalidateRibbonControl RIBBON_CTL_NODE_HEIGHT
    InvalidateRibbonControl RIBBON_CTL_NODE_HEIGHT_METRIC
    InvalidateRibbonControl RIBBON_CTL_NODE_WIDTH
    InvalidateRibbonControl RIBBON_CTL_NODE_WIDTH_METRIC
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeMetric_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).Value = vbNullString Then
        returnedVal = False
    Else
        returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_NODE_METRIC)
    End If
End Sub


' ===========================================================================
' Callbacks for edgeColor1

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor1_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    color_getImage DESIGNER_EDGE_COLOR_1, COLOR_BLACK, returnedVal
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor1_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    HideProgressIndicator
    returnedVal = GetColorIndex(DESIGNER_EDGE_COLOR_1)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor1_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    Application.EnableEvents = False
    
    SaveColor index, DESIGNER_EDGE_COLOR_1
    
    If StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_1).Value = vbNullString Then
        StyleDesignerSheet.Range("EdgeColor2,EdgeColor3").ClearContents
    End If
    
    InvalidateRibbonColorList controlId
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    
    Application.EnableEvents = True
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeColor2

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    color_getImage DESIGNER_EDGE_COLOR_2, COLOR_BLACK, returnedVal
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    HideProgressIndicator
    returnedVal = GetColorIndex(DESIGNER_EDGE_COLOR_2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    Application.EnableEvents = False
    
    SaveColor index, DESIGNER_EDGE_COLOR_2
    
    If StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_2).Value = vbNullString Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_3).ClearContents
    End If
    
    InvalidateRibbonColorList controlId
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    Application.EnableEvents = True
    
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_1))
End Sub

' ===========================================================================
' Callbacks for edgeColor3

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = vbNullString
#Else
    color_getImage DESIGNER_EDGE_COLOR_3, COLOR_BLACK, returnedVal
#End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    HideProgressIndicator
    returnedVal = GetColorIndex(DESIGNER_EDGE_COLOR_3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SaveColor index, DESIGNER_EDGE_COLOR_3
    InvalidateRibbonColorList controlId
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_COLOR_2))
End Sub

' ===========================================================================
' Callbacks for Arrow Tail

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub groupArrowHead_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Dim direction As String
    Dim mode As String
    
    direction = StyleDesignerSheet.Range(DESIGNER_EDGE_DIRECTION).Value
    mode = StyleDesignerSheet.Range(DESIGNER_MODE).Value
    
    visible = mode = KEYWORD_EDGE And (direction = vbNullString Or direction = "forward" Or direction = "both")
End Sub

' ===========================================================================
' Callbacks for edgeArrowHead1

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead1_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_HEADS, DESIGNER_EDGE_ARROW_HEAD_1)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead1_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_1).Value = Mid$(controlId, Len("h1_") + 1)
    If StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_1).Value = vbNullString Then
        StyleDesignerSheet.Range("EdgeArrowHead2,EdgeArrowHead3").ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD2
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD3
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeArrowHead2

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead2_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_HEADS, DESIGNER_EDGE_ARROW_HEAD_2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead2_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_2).Value = Mid$(controlId, Len("h2_") + 1)
    If StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_2).Value = vbNullString Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_3).ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_1))
End Sub

' ===========================================================================
' Callbacks for edgeArrowHead3

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead3_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_HEADS, DESIGNER_EDGE_ARROW_HEAD_3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead3_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_3).Value = Mid$(controlId, Len("h3_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_HEAD_2))
End Sub

' ===========================================================================
' Callbacks for Arrow Tail

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub groupArrowTail_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Dim direction As String
    Dim mode As String
    
    direction = StyleDesignerSheet.Range(DESIGNER_EDGE_DIRECTION).Value
    mode = StyleDesignerSheet.Range(DESIGNER_MODE).Value
    
    visible = mode = KEYWORD_EDGE And (direction = "back" Or direction = "both")
End Sub

' ===========================================================================
' Callbacks for edgeArrowTail1

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail1_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_TAILS, DESIGNER_EDGE_ARROW_TAIL_1)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail1_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_1).Value = Mid$(controlId, Len("t1_") + 1)
    If StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_1).Value = vbNullString Then
        StyleDesignerSheet.Range("EdgeArrowTail2,EdgeArrowTail3").ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL2
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL3
    InvalidateRibbonControl RIBBON_CTL_EDGE_DIRECTION
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeArrowTail2

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail2_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_TAILS, DESIGNER_EDGE_ARROW_TAIL_2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail2_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_2).Value = Mid$(controlId, Len("t2_") + 1)
    If StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_2).Value = vbNullString Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_3).ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_1))
End Sub

' ===========================================================================
' Callbacks for edgeArrowTail3

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail3_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_TAILS, DESIGNER_EDGE_ARROW_TAIL_3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail3_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_3).Value = Mid$(controlId, Len("t3_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not IsEmpty(StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_TAIL_2))
End Sub

' ===========================================================================
' Callbacks for edgeDirection

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeDirection_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    Dim direction As String
    direction = Mid$(controlId, Len("ed_") + 1)
    
    StyleDesignerSheet.Range(DESIGNER_EDGE_DIRECTION).Value = direction
    
    If direction = vbNullString Then
        StyleDesignerSheet.Range("EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3").ClearContents
    
    ElseIf direction = "back" Then
        StyleDesignerSheet.Range("EdgeArrowHead1,EdgeArrowHead2,EdgeArrowHead3").ClearContents
    
    '@Ignore EmptyIfBlock
    ElseIf direction = "both" Then
        ' No action to take
    
    ElseIf direction = "forward" Then
        StyleDesignerSheet.Range("EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3").ClearContents
   
    ElseIf direction = "none" Then
        StyleDesignerSheet.Range("EdgeArrowHead1,EdgeArrowHead2,EdgeArrowHead3,EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3,EdgeArrowSize").ClearContents
    End If
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_SIZE
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD1
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD2
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD3
    InvalidateRibbonControl RIBBON_GRP_EDGE_ARROW_HEAD
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL1
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL2
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL3
    InvalidateRibbonControl RIBBON_GRP_EDGE_ARROW_TAIL
    
    InvalidateRibbonControl RIBBON_GRP_EDGE_ARROW
    
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeDirection_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_DIRECTION, DESIGNER_EDGE_DIRECTION)
End Sub

' ===========================================================================
' Callbacks for edgeArrowSize

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowSize_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim direction As String
    direction = StyleDesignerSheet.Range(DESIGNER_EDGE_DIRECTION).Value
    
    returnedVal = Not (direction = "none")
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowSize_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_ARROW_SIZE).Value = Mid$(controlId, Len("as_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowSize_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_ARROW_SIZE, DESIGNER_EDGE_ARROW_SIZE)
End Sub

' ===========================================================================
' Callbacks for edgePenWidth

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgePenWidth_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_EDGE_PEN_WIDTH).Value = Mid$(controlId, Len("ew_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgePenWidth_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_PEN_WIDTH, DESIGNER_EDGE_PEN_WIDTH)
End Sub

' ===========================================================================
' Callbacks for nodeImageName

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageName_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME).Value = Text
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageName_getText(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME))
End Sub

' ===========================================================================
' Callbacks for nodeRegular

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub regular_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_NODE_REGULAR).Value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub regular_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_NODE_REGULAR).Value = TOGGLE_YES
End Sub

' ===========================================================================
' Callbacks for nodeSkew

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSkew_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    StyleDesignerSheet.Range(DESIGNER_NODE_SKEW).Value = Text
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSkew_getText(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_SKEW))
End Sub

' ===========================================================================
' Callbacks for nodeDistortion

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeDistortion_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    StyleDesignerSheet.Range(DESIGNER_NODE_DISTORTION).Value = Text
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeDistortion_getText(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Trim$(StyleDesignerSheet.Range(DESIGNER_NODE_DISTORTION))
End Sub

' ===========================================================================
' Callbacks for nodeImageSelect

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageChoose_onAction(ByVal control As IRibbonControl)
    Dim filePath As String
    Dim filename As String
    Dim directoryName As String
    
#If Mac Then
    filePath = RunAppleScriptTask("chooseImageFile", "Select an image file")
    If filePath = vbNullString Then
        ' No image was chosen
        Exit Sub
    End If
#Else
    Dim fDialog As FileDialog
    Dim choice As Long
    
    ' Set the options for the file picker dialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = False
    fDialog.title = "Select an image file"
    fDialog.InitialFileName = ActiveWorkbook.path
    fDialog.Filters.Clear
    fDialog.Filters.Add "Image files", "*.bmp;*.gif;*.jpg;*.jpeg;*.png"
    fDialog.Filters.Add "All files", "*.*"
    
    'get the number of the button chosen
    choice = fDialog.show
    If choice <> -1 Then                         ' user selected cancel, do not continue any farther
        Exit Sub
    End If

    ' get the path from the file dialog
    filePath = fDialog.SelectedItems.Item(1)
    
    ' Clean up objects
    Set fDialog = Nothing
#End If
    Dim envVarSeparator As String
    envVarSeparator = GetEnvVarSeparator
    
    ' Split the complete file name into directory and filename components. The Graphviz image
    ' attribute only wants the filename specified, and looks to find the file on the image path
    Dim pathComponents() As String
    pathComponents = Split(filePath, Application.pathSeparator)
    filename = pathComponents(UBound(pathComponents))
    directoryName = Left$(filePath, Len(filePath) - Len(filename) - 1)
    
    '@Ignore EmptyIfBlock
    If ImageFoundInEnvVariablePath(directoryName) Then
        ' No need to alter the saved image path
    ElseIf Not ImageFoundInCurrentDir(directoryName) Then ' Image is not in the workbook directory.
        ' All .gv files created by this tool specify the workbook path on the image path.
        ' If the image file is in the current workbook directory then nothing more needs
        ' to be done to make it render properly.
        
        Dim settingsImagePath As String
        settingsImagePath = SettingsSheet.Range(SETTINGS_IMAGE_PATH).Value
        If settingsImagePath = vbNullString Then
            ' If an image path has not been specified in the Settings worksheet the easiest thing
            ' to do is save the image directory there.
            SettingsSheet.Range(SETTINGS_IMAGE_PATH).Value = directoryName
        Else
            ' One of more paths are already specified. We don't want to add duplicates, so run
            ' a test to see if the directory is already within the path concatenation
            Dim pathArray() As String
            pathArray = Split(settingsImagePath, envVarSeparator)
            
            Dim index As Long
            Dim boolOnPath As Boolean
            boolOnPath = False
            
            For index = LBound(pathArray) To UBound(pathArray)
                If pathArray(index) = directoryName Then
                    ' The directory is in the path concatenation, no need to do any more checks
                    boolOnPath = True
                    Exit For
                End If
            Next index
            
            If Not boolOnPath Then
                ' Append the directory to the current ImagePath setting and save it to the Settings worksheet
                settingsImagePath = settingsImagePath & envVarSeparator & directoryName
                SettingsSheet.Range(SETTINGS_IMAGE_PATH).Value = settingsImagePath
            End If
        End If
    End If
    
    ' Display the filename in the ribbon
    StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME).Value = filename
    InvalidateRibbonControl RIBBON_CTL_NODE_IMAGE_NAME
    
    ' Update the Node preview
    RenderPreview
End Sub

Private Function ImageFoundInEnvVariablePath(ByVal directoryName As String) As Boolean
#If Mac Then
    ' Environment variable specifying a directory of images is not supported in the
    ' Mac version at this time.
    ImageFoundInEnvVariablePath = False
#Else
    ImageFoundInEnvVariablePath = UCase$(directoryName) = UCase$(Trim$(Environ$("ExcelToGraphvizImages")))
#End If
End Function

Private Function ImageFoundInCurrentDir(ByVal directoryName As String) As Boolean
    ImageFoundInCurrentDir = UCase$(directoryName) = UCase$(ActiveWorkbook.path)
End Function

' ===========================================================================
' Callbacks for nodeImagePosition

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImagePosition_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_POSITION).Value = Mid$(controlId, Len("imagepos_") + 1)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImagePosition_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_IMAGE_POSITION, DESIGNER_NODE_IMAGE_POSITION)
End Sub

' ===========================================================================
' Callbacks for nodeImageScale

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    listSize = ListsSheet.Range(LISTS_IMAGE_SCALE).Count + 1
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = vbNullString
    Else
        Dim listId As String
        listId = "is_" & ListsSheet.Range(LISTS_IMAGE_SCALE).Cells.Item(index, 1).Value
        label = GetLabel(listId)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_IMAGE_SCALE, DESIGNER_NODE_IMAGE_SCALE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    If index = 0 Then
        StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_SCALE).Value = vbNullString
    Else
        StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_SCALE).Value = ListsSheet.Range(LISTS_IMAGE_SCALE).Cells.Item(index, 1).Value
    End If
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeHeadClip

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadClip_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_HEAD_CLIP).ClearContents
    Else
        StyleDesignerSheet.Range(DESIGNER_EDGE_HEAD_CLIP).Value = TOGGLE_NO
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadClip_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_EDGE_HEAD_CLIP).Value = vbNullString Then
        returnedVal = True
    Else
        returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_EDGE_HEAD_CLIP)
    End If
End Sub

' ===========================================================================
' Callbacks for edgeTailClip

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeTailClip_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_TAIL_CLIP).ClearContents
    Else
        StyleDesignerSheet.Range(DESIGNER_EDGE_TAIL_CLIP).Value = TOGGLE_NO
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeTailClip_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_EDGE_TAIL_CLIP).Value = vbNullString Then
        returnedVal = True
    Else
        returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_EDGE_TAIL_CLIP)
    End If
End Sub

' ===========================================================================
' Callbacks for edgeDecorate

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeDecorate_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_DECORATE).Value = TOGGLE_YES
    Else
        StyleDesignerSheet.Range(DESIGNER_EDGE_DECORATE).ClearContents
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeDecorate_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_EDGE_DECORATE)
End Sub

' ===========================================================================
' Callbacks for edgeLabelFloat

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelFloat_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FLOAT).Value = TOGGLE_YES
    Else
        StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FLOAT).ClearContents
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelFloat_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_EDGE_LABEL_FLOAT)
End Sub

' ===========================================================================
' Callbacks for clearStyleRibbon

Public Sub ClearStyleRibbon()
    OptimizeCode_Begin
    StyleDesignerSheet.Range("ColorScheme,FontName,FontSize,FontColor,BorderColor,BorderColor,BorderPenWidth,BorderPeripheries").ClearContents
    StyleDesignerSheet.Range("FillColor,GradientFillColor,GradientFillType,GradientFillAngle,GradientFillWeight,LabelLocation,LabelJustification,EdgeStyle,EdgeHeadPort,EdgeTailPort,EdgeColor1,EdgeColor2,EdgeColor3").ClearContents
    StyleDesignerSheet.Range("NodeShape,NodeSides,NodeOrientation,NodeRegular,NodeSkew,NodeDistortion,BorderStyle1,BorderStyle2,BorderStyle3").ClearContents
    StyleDesignerSheet.Range("NodeHeight,NodeWidth,NodeFixedSize,EdgeArrowHead1,EdgeArrowHead2,EdgeArrowHead3,EdgeDecorate,EdgeLabelFloat").ClearContents
    StyleDesignerSheet.Range("EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3,EdgeDirection,EdgeArrowSize,EdgeWeight,EdgeLabelAngle,EdgeLabelDistance").ClearContents
    StyleDesignerSheet.Range("EdgePenWidth,NodeImageName,NodeImageScale,NodeImagePosition,EdgeHeadClip,EdgeTailClip,EdgeLabelFontName,EdgeLabelFontSize,EdgeLabelFontColor").ClearContents
    StyleDesignerSheet.Range("FontBold,FontItalic").ClearContents
    OptimizeCode_End
    RenderPreview
  
    RefreshRibbon tag:="*Tab"

    StyleDesignerSheet.Range("ElapsedTime").ClearContents
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub clearStyleRibbon_onAction(ByVal control As IRibbonControl)
    ClearStyleRibbon
End Sub

' ===========================================================================
' Callbacks for saveToStylesWorksheet

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub saveToStylesWorksheet_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not (StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Value = vbNullString)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub saveToStylesWorksheet_onAction(ByVal control As IRibbonControl)
    Dim row As Long
    Dim rowFocus As Long
    Dim col As Long
    Dim styleType As String
    Dim defaultStyleName As String
    
    ' Obtain the layout of the "styles' worksheet
    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    ' Map the 'Design Mode' dropdown value to the Object Type
    Select Case StyleDesignerSheet.Range(DESIGNER_MODE).Value
        Case KEYWORD_NODE
            styleType = TYPE_NODE
        Case KEYWORD_EDGE
            styleType = TYPE_EDGE
        Case KEYWORD_CLUSTER
            styleType = TYPE_SUBGRAPH_OPEN
    End Select
    
    ' Increment the count to reflect the style we are adding
    Dim objectCount As Long
    objectCount = GetStyleCount(styleType, styles) + 1
    
    ' Create default style name
    Select Case StyleDesignerSheet.Range(DESIGNER_MODE).Value
        Case KEYWORD_NODE
            defaultStyleName = "Node style " & objectCount
        Case KEYWORD_EDGE
            defaultStyleName = "Edge style " & objectCount
        Case KEYWORD_CLUSTER
            defaultStyleName = "Cluster style " & objectCount & " " & styles.suffixOpen
    End Select
    
    ' Look for a row that does not have a style name
    For row = styles.firstRow To styles.lastRow
        If StylesSheet.Cells.Item(row, styles.flagColumn) <> FLAG_COMMENT And _
           StylesSheet.Cells.Item(row, styles.nameColumn).Value = vbNullString Then
            Exit For
        End If
    Next row
    
    ' Save the row number so we know where to place the focus if the DESIGNER_MODE = CLUSTER
    rowFocus = row
    
    ' Set the format string and the object type
    
    StylesSheet.Cells.Item(row, styles.nameColumn).Value = defaultStyleName
    StylesSheet.Cells.Item(row, styles.formatColumn).Value = StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Value
    StylesSheet.Cells.Item(row, styles.typeColumn).Value = styleType

    ' Loop through the columns which have column headings and put a value of 'yes' in the cell
    For col = styles.firstYesNoColumn To GetLastColumn(StylesSheet.name, styles.headingRow)
        If StylesSheet.Cells.Item(styles.headingRow, col) <> vbNullString Then
            StylesSheet.Cells.Item(row, col).Value = TOGGLE_YES
        End If
    Next col
    
    ' If the style is CLUSTER we want to add a row for the subgraph-close, as it improves filtering capabilities
    If StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER Then
        styleType = "subgraph-close"
        defaultStyleName = "Cluster style " & objectCount & " " & styles.suffixClose
   
        ' Look for a row that does not have a style name
        For row = rowFocus To styles.lastRow + 1
            If StylesSheet.Cells.Item(row, styles.flagColumn) <> FLAG_COMMENT And _
               StylesSheet.Cells.Item(row, styles.nameColumn).Value = vbNullString Then
                Exit For
            End If
        Next row

        ' Set the format string and the object type
        StylesSheet.Cells.Item(row, styles.nameColumn).Value = defaultStyleName
        StylesSheet.Cells.Item(row, styles.formatColumn).Value = vbNullString
        StylesSheet.Cells.Item(row, styles.typeColumn).Value = styleType

        ' Loop through the columns which have column headings and put a value of 'yes' in the cell
        For col = styles.firstYesNoColumn To GetLastColumn(StylesSheet.name, styles.headingRow)
            If StylesSheet.Cells.Item(styles.headingRow, col) <> vbNullString Then
                StylesSheet.Cells.Item(row, col) = TOGGLE_YES
            End If
        Next col
    End If
    
    ' Put the focus on the cell where the style name has to be entered
    StylesSheet.Activate
    ActiveSheet.Cells(rowFocus, styles.nameColumn).Select
    
End Sub

Private Function GetStyleCount(ByVal styleType As String, ByRef styles As stylesWorksheet) As Long
    Dim row As Long
    Dim styleCount As Long
    
    styleCount = 0
    
    For row = styles.firstRow To styles.lastRow
        If StylesSheet.Cells.Item(row, styles.typeColumn).Value = styleType Then
            styleCount = styleCount + 1
        End If
    Next row

    GetStyleCount = styleCount
End Function

' ===========================================================================
' Callbacks for copyToClipboard

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub copyToClipboard_onAction(ByVal control As IRibbonControl)
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Copy
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub copyToClipboard_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not (StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Value = vbNullString)
End Sub

' ===========================================================================
' Callbacks for alignTop

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignTop_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_LABEL_LOCATION).Value = Toggle(pressed, "top", vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ALIGN_BOTTOM
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignTop_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_LABEL_LOCATION).Value = "top"
End Sub

' ===========================================================================
' Callbacks for alignBottom

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignBottom_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_LABEL_LOCATION).Value = Toggle(pressed, "bottom", vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ALIGN_TOP
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignBottom_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_LABEL_LOCATION).Value = "bottom"
End Sub

' ===========================================================================
' Callbacks for justifyLeft

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyLeft_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_LABEL_JUSTIFICATION).Value = Toggle(pressed, "left", vbNullString)
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_RIGHT
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyLeft_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_LABEL_JUSTIFICATION).Value = "left"
End Sub

' ===========================================================================
' Callbacks for justifyRight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyRight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_LABEL_JUSTIFICATION).Value = Toggle(pressed, "right", vbNullString)
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_LEFT
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyRight_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSheet.Range(DESIGNER_LABEL_JUSTIFICATION).Value = "right"
End Sub

' ===========================================================================
' Callbacks for fontBold

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontBold_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_FONT_BOLD).Value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontBold_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_FONT_BOLD)
End Sub

' ===========================================================================
' Callbacks for fontItalic

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontItalic_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    StyleDesignerSheet.Range(DESIGNER_FONT_ITALIC).Value = Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontItalic_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_FONT_ITALIC)
End Sub

' ===========================================================================
' Group visibility callbacks

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupLabels_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE Or _
              StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupBorders_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE Or _
              StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupFillColor_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE Or _
              StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupGradientFillColor_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = False
    
    If StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE Or _
       StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_CLUSTER Then
        If StyleDesignerSheet.Range(DESIGNER_FILL_COLOR).Value <> vbNullString Then
            visible = True
        End If
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupNodeShape_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupNodeDimensions_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupNodeImage_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_NODE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupEdgeStyle_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_EDGE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupEdgeColors_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_EDGE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupEdgeArrows_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSheet.Range(DESIGNER_MODE).Value = KEYWORD_EDGE
End Sub

' ===========================================================================
' Utility routines

Public Sub RenderPreview()
#If Mac Then
    ' For some reason, my Mac fails when I code it as "#If Not Mac Then"
#Else
    Dim timex As Stopwatch
    Set timex = New Stopwatch
    timex.start
#End If

    StyleDesignerSheet.Activate
    OptimizeCode_Begin
    RenderPreviewFromLists
    OptimizeCode_End
    
#If Mac Then
    ' For some reason, my Mac fails when I code it as "#If Not Mac Then"
#Else
    timex.stop_it
    StyleDesignerSheet.Range("ElapsedTime").Value = timex.Elapsed_sec & " seconds"
#End If
End Sub

Private Function GetColorIndex(ByVal cellName As String) As Long
    
    GetColorIndex = 0
    
    Dim color As String
    color = StyleDesignerSheet.Range(cellName).Value
    
    If color <> vbNullString Then
        
        Dim index As Long
        index = 0
        Dim arrayItem As Variant
       
        If colorScheme = COLOR_SCHEME_X11 Then
            For Each arrayItem In x11Colors
                index = index + 1
                If arrayItem = color Then
                    Exit For
                End If
            Next arrayItem
        ElseIf colorScheme = COLOR_SCHEME_SVG Then
            For Each arrayItem In svgColors
                index = index + 1
                If arrayItem = color Then
                    Exit For
                End If
            Next arrayItem
        Else
            For Each arrayItem In brewerColors
                index = index + 1
                If arrayItem = color Then
                    Exit For
                End If
            Next arrayItem
        End If
        
        GetColorIndex = index
    End If
End Function

Private Sub SaveColor(ByVal index As Long, ByVal cellName As String)
    Dim color As String
    If index = 0 Then
        color = vbNullString
    Else
        If colorScheme = COLOR_SCHEME_X11 Or colorScheme = COLOR_SCHEME_SVG Then
            ' Color list is in cells along a column
            color = HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & colorScheme).Cells.Item(index, 1).Value
        Else
            ' Color list is in cells along a row
            color = HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & colorScheme).Cells.Item(1, index).Value
        End If
    End If
    StyleDesignerSheet.Range(cellName).Value = color

End Sub

Private Function GetListIndex(ByVal listName As String, ByVal cellName As String) As Long
    Dim index As Long
    Dim cellValue As String
    
    GetListIndex = 0
    
    cellValue = StyleDesignerSheet.Range(cellName).Value
    
    If cellValue <> vbNullString Then
        ' Iterating arrays is faster than iterating cells
        Dim listArray As Variant
        listArray = Application.WorksheetFunction.Transpose(ListsSheet.Range(listName))
        
        Dim listItem As Variant
        
        ' Iterate top to bottom
        For Each listItem In listArray
            index = index + 1
            If UCase$(Trim$(listItem)) = UCase$(cellValue) Then
                GetListIndex = index
                Exit For
            End If
        Next listItem
    End If
    
End Function

Public Sub SetStyleDesignerNodeShape(ByVal shapeName As String)

    SettingsSheet.Range(RIBBON_CTL_HELP_SHAPES).Value = TOGGLE_HIDE
    StyleDesignerSheet.Activate
    
    OptimizeCode_Begin
    
    StyleDesignerSheet.Range(DESIGNER_NODE_SHAPE).Value = shapeName
    If shapeName <> "polygon" Then
        StyleDesignerSheet.Range("NodeSides,NodeOrientation,NodeSkew,NodeDistortion").ClearContents
    End If
    
    InvalidateRibbonControl RIBBON_CTL_HELP_DESIGN_SHAPES
    InvalidateRibbonControl RIBBON_CTL_NODE_SHAPE
    InvalidateRibbonControl RIBBON_CTL_NODE_SIDES
    InvalidateRibbonControl RIBBON_CTL_POLYGON_SEPARATOR
    InvalidateRibbonControl RIBBON_CTL_NODE_ROTATION
    InvalidateRibbonControl RIBBON_CTL_NODE_SKEW
    InvalidateRibbonControl RIBBON_CTL_NODE_DISTORTION

    OptimizeCode_End
    
    RenderPreview
End Sub

Public Sub SetStyleDesignerColorScheme(ByVal colorScheme As String)
    
    SettingsSheet.Range(RIBBON_CTL_HELP_COLORS).Value = TOGGLE_HIDE
    StyleDesignerSheet.Activate
    
    OptimizeCode_Begin
    
    StyleDesignerSheet.Range(DESIGNER_COLOR_SCHEME).Value = colorScheme
    StyleDesignerSheet.Range("FontColor,BorderColor,FillColor,GradientFillColor,GradientFillType,GradientFillAngle,EdgeColor1,EdgeColor2,EdgeColor3,EdgeLabelFontColor").ClearContents
    
    InvalidateRibbonControl RIBBON_CTL_HELP_DESIGN_COLORS
    InvalidateRibbonControl RIBBON_CTL_COLOR_SCHEME
    InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_ANGLE
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    
    OptimizeCode_End
    
    RenderPreview

End Sub

Public Sub RenderPreviewFromLists()
    RenderElement DESIGNER_FORMAT_STRING, _
                  DESIGNER_PREVIEW_CELL, _
                  GetCellString(StyleDesignerSheet.name, DESIGNER_MODE), _
                  True
    
    InvalidateRibbonControl RIBBON_CTL_SAVE_TO_STYLES_WORKSHEET
    InvalidateRibbonControl RIBBON_CTL_COPY_TO_CLIPBOARD
End Sub

Public Sub RenderPreviewFromFormatString()
    RenderElement DESIGNER_FORMAT_STRING, _
                  DESIGNER_PREVIEW_CELL, _
                  GetCellString(StyleDesignerSheet.name, DESIGNER_MODE), _
                  False
    
    InvalidateRibbonControl RIBBON_CTL_SAVE_TO_STYLES_WORKSHEET
    InvalidateRibbonControl RIBBON_CTL_COPY_TO_CLIPBOARD
End Sub


Public Sub ribbon_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetLabel(control.ID)
End Sub

Public Sub ribbon_getScreenTip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetScreentip(control.ID)
End Sub

Public Sub ribbon_getSuperTip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSupertip(control.ID)
End Sub

Private Function getFontList() As Variant
    
#If Mac Then
    getFontList = Application.Transpose(ListsSheet.Range(LISTS_FONTS))
#Else
    ' The list of fonts on Windows is availble through a menu control
    Dim tmpFontList As CommandBarControl
    On Error Resume Next
    Set tmpFontList = Application.CommandBars.Item("Formatting").FindControl(ID:=1728)
    On Error GoTo 0
    
    'If Font control is missing, create it on a temporary CommandBar
    If tmpFontList Is Nothing Then
        Dim tmpCommandBar As Variant
        Set tmpCommandBar = Application.CommandBars.Add
        Set tmpFontList = tmpCommandBar.Controls.Add(ID:=1728)
        tmpCommandBar.Delete
    End If
    
    ' Cache the list of fonts in an array
    On Error GoTo ErrorHandler
    Dim i As Long
    '@Ignore VariableNotAssigned
    Dim fontList As Variant
    '@Ignore MemberNotOnInterface
    For i = 1 To tmpFontList.listCount
        ' Office 365 has exploded the number of fonts, blowing past
        ' the 1000 items a dropdown list is limited to. To compensate,
        ' filter out variations of font names to try to bring the list
        ' down to a managable size before truncating the list at 1000
        ' font names.
        '@Ignore MemberNotOnInterface
        If addToFontList(tmpFontList.List(i)) Then
            If IsEmpty(fontList) Then   ' Allocate an array
                ReDim fontList(1)
                '@Ignore MemberNotOnInterface
                fontList(UBound(fontList)) = tmpFontList.List(i)
            Else    ' Grow the array by 1
                ReDim Preserve fontList(0 To UBound(fontList) + 1)
                '@Ignore MemberNotOnInterface
                fontList(UBound(fontList)) = tmpFontList.List(i)
            End If
        End If
    Next i
    
    ' Clean up
    Set tmpFontList = Nothing
    getFontList = fontList
    Exit Function
ErrorHandler:
    MsgBox GetMessage("msgboxNoListOfFonts"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    ReDim fontList(0)
    getFontList = fontList
#End If

End Function
Private Function addToFontList(ByVal fontName As String) As Boolean
    addToFontList = True
    
    ' The Graphviz font mapper does not recogonize these fonts, and maps them to Arial
    If StartsWith(fontName, "Abadi") Then
        addToFontList = False
    ElseIf StartsWith(fontName, "Abel") Then addToFontList = False
    ElseIf StartsWith(fontName, "Abril") Then addToFontList = False
    ElseIf StartsWith(fontName, "Aharoni") Then addToFontList = False
    ElseIf StartsWith(fontName, "Aldhabi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Alef") Then addToFontList = False
    ElseIf StartsWith(fontName, "Aleo") Then addToFontList = False
    ElseIf StartsWith(fontName, "Amatic") Then addToFontList = False
    ElseIf StartsWith(fontName, "Angsana") Then addToFontList = False
    ElseIf StartsWith(fontName, "Anton") Then addToFontList = False
    ElseIf StartsWith(fontName, "Aparajita") Then addToFontList = False
    ElseIf StartsWith(fontName, "Arabic") Then addToFontList = False
    ElseIf StartsWith(fontName, "Aref") Then addToFontList = False
    ElseIf StartsWith(fontName, "Assistant") Then addToFontList = False
    ElseIf StartsWith(fontName, "Athiti") Then addToFontList = False
    ElseIf StartsWith(fontName, "Baguet") Then addToFontList = False
    ElseIf StartsWith(fontName, "Bahnschrift") Then addToFontList = False
    ElseIf StartsWith(fontName, "Barlow") Then addToFontList = False
    ElseIf StartsWith(fontName, "Batang") Then addToFontList = False
    ElseIf StartsWith(fontName, "Bauhaus") Then addToFontList = False
    ElseIf StartsWith(fontName, "Bebas") Then addToFontList = False
    ElseIf StartsWith(fontName, "Bembo") Then addToFontList = False
    ElseIf StartsWith(fontName, "Berlin") Then addToFontList = False
    ElseIf StartsWith(fontName, "Bierstadt") Then addToFontList = False
    ElseIf StartsWith(fontName, "Biome") Then addToFontList = False
    ElseIf StartsWith(fontName, "Bookshelf") Then addToFontList = False
    ElseIf StartsWith(fontName, "Browallia") Then addToFontList = False
    ElseIf StartsWith(fontName, "Brush") Then addToFontList = False
    ElseIf StartsWith(fontName, "Buxton") Then addToFontList = False
    ElseIf StartsWith(fontName, "Cascadia") Then addToFontList = False
    ElseIf StartsWith(fontName, "Caveat") Then addToFontList = False
    ElseIf StartsWith(fontName, "Cavolini") Then addToFontList = False
    ElseIf StartsWith(fontName, "Chamberi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Charmonman") Then addToFontList = False
    ElseIf StartsWith(fontName, "Chonburi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Concert") Then addToFontList = False
    ElseIf StartsWith(fontName, "Congenial") Then addToFontList = False
    ElseIf StartsWith(fontName, "Convection") Then addToFontList = False
    ElseIf StartsWith(fontName, "Cordia") Then addToFontList = False
    ElseIf StartsWith(fontName, "DM") Then addToFontList = False
    ElseIf StartsWith(fontName, "Dante") Then addToFontList = False
    ElseIf StartsWith(fontName, "DaunPenh") Then addToFontList = False
    ElseIf StartsWith(fontName, "David") Then addToFontList = False
    ElseIf StartsWith(fontName, "Daytona") Then addToFontList = False
    ElseIf StartsWith(fontName, "DengXian") Then addToFontList = False
    ElseIf StartsWith(fontName, "Didact") Then addToFontList = False
    ElseIf StartsWith(fontName, "Dillenia") Then addToFontList = False
    ElseIf StartsWith(fontName, "DokChampa") Then addToFontList = False
    ElseIf StartsWith(fontName, "Dosis") Then addToFontList = False
    ElseIf StartsWith(fontName, "Dotum") Then addToFontList = False
    ElseIf StartsWith(fontName, "EB Garamond") Then addToFontList = False
    ElseIf StartsWith(fontName, "Ebrima") Then addToFontList = False
    ElseIf StartsWith(fontName, "Eucrosia") Then addToFontList = False
    ElseIf StartsWith(fontName, "Euphemia") Then addToFontList = False
    ElseIf StartsWith(fontName, "Fahkwang") Then addToFontList = False
    ElseIf StartsWith(fontName, "Fairwater") Then addToFontList = False
    ElseIf StartsWith(fontName, "Fira") Then addToFontList = False
    ElseIf StartsWith(fontName, "Fjalla") Then addToFontList = False
    ElseIf StartsWith(fontName, "Frank") Then addToFontList = False
    ElseIf StartsWith(fontName, "Fredoka") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gabriela") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gaegu") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gautami") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gill") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gisha") Then addToFontList = False
    ElseIf StartsWith(fontName, "Goudy") Then addToFontList = False
    ElseIf StartsWith(fontName, "Grandview") Then addToFontList = False
    ElseIf StartsWith(fontName, "Grotesque") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gulim") Then addToFontList = False
    ElseIf StartsWith(fontName, "Gungsuh") Then addToFontList = False
    ElseIf StartsWith(fontName, "HG") Then addToFontList = False
    ElseIf StartsWith(fontName, "Hadassah") Then addToFontList = False
    ElseIf StartsWith(fontName, "Hammersmith") Then addToFontList = False
    ElseIf StartsWith(fontName, "Heebo") Then addToFontList = False
    ElseIf StartsWith(fontName, "Hind") Then addToFontList = False
    ElseIf StartsWith(fontName, "HoloLens") Then addToFontList = False
    ElseIf StartsWith(fontName, "IBM") Then addToFontList = False
    ElseIf StartsWith(fontName, "Inconsolata") Then addToFontList = False
    ElseIf StartsWith(fontName, "Informal") Then addToFontList = False
    ElseIf StartsWith(fontName, "Iris") Then addToFontList = False
    ElseIf StartsWith(fontName, "Iskoola") Then addToFontList = False
    ElseIf StartsWith(fontName, "Jasmine") Then addToFontList = False
    ElseIf StartsWith(fontName, "Josefin") Then addToFontList = False
    ElseIf StartsWith(fontName, "Jumble") Then addToFontList = False
    ElseIf StartsWith(fontName, "KaiTi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Kalinga") Then addToFontList = False
    ElseIf StartsWith(fontName, "Karla") Then addToFontList = False
    ElseIf StartsWith(fontName, "Kartika") Then addToFontList = False
    ElseIf StartsWith(fontName, "Kigelia") Then addToFontList = False
    ElseIf StartsWith(fontName, "KleeOne") Then addToFontList = False
    ElseIf StartsWith(fontName, "Kodchiang") Then addToFontList = False
    ElseIf StartsWith(fontName, "Kokila") Then addToFontList = False
    ElseIf StartsWith(fontName, "Krub") Then addToFontList = False
    ElseIf StartsWith(fontName, "Lalezar") Then addToFontList = False
    ElseIf StartsWith(fontName, "Latha") Then addToFontList = False
    ElseIf StartsWith(fontName, "Lato") Then addToFontList = False
    ElseIf StartsWith(fontName, "Leelawadee") Then addToFontList = False
    ElseIf StartsWith(fontName, "Levenim") Then addToFontList = False
    ElseIf StartsWith(fontName, "Libre") Then addToFontList = False
    ElseIf StartsWith(fontName, "Ligconsolata") Then addToFontList = False
    ElseIf StartsWith(fontName, "Lily") Then addToFontList = False
    ElseIf StartsWith(fontName, "Livvic") Then addToFontList = False
    ElseIf StartsWith(fontName, "Lobster") Then addToFontList = False
    ElseIf StartsWith(fontName, "Lora") Then addToFontList = False
    ElseIf StartsWith(fontName, "Microsoft") Then addToFontList = False
    ElseIf StartsWith(fontName, "MS") Then addToFontList = False
    ElseIf StartsWith(fontName, "MT") Then addToFontList = False
    ElseIf StartsWith(fontName, "Mangal") Then addToFontList = False
    ElseIf StartsWith(fontName, "Marlett") Then addToFontList = False
    ElseIf StartsWith(fontName, "Meddon") Then addToFontList = False
    ElseIf StartsWith(fontName, "Meiryo") Then addToFontList = False
    ElseIf StartsWith(fontName, "Merriweather") Then addToFontList = False
    ElseIf StartsWith(fontName, "Ming") Then addToFontList = False
    ElseIf StartsWith(fontName, "Miriam") Then addToFontList = False
    ElseIf StartsWith(fontName, "Mitr") Then addToFontList = False
    ElseIf StartsWith(fontName, "Modern") Then addToFontList = False
    ElseIf StartsWith(fontName, "Monotype") Then addToFontList = False
    ElseIf StartsWith(fontName, "Montserrat") Then addToFontList = False
    ElseIf StartsWith(fontName, "MoolBoran") Then addToFontList = False
    ElseIf StartsWith(fontName, "Mystical") Then addToFontList = False
    ElseIf StartsWith(fontName, "Nanum") Then addToFontList = False
    ElseIf StartsWith(fontName, "Narkisim") Then addToFontList = False
    ElseIf StartsWith(fontName, "News") Then addToFontList = False
    ElseIf StartsWith(fontName, "Nina") Then addToFontList = False
    ElseIf StartsWith(fontName, "Nordique") Then addToFontList = False
    ElseIf StartsWith(fontName, "Noto") Then addToFontList = False
    ElseIf StartsWith(fontName, "Nunito") Then addToFontList = False
    ElseIf StartsWith(fontName, "Nyala") Then addToFontList = False
    ElseIf StartsWith(fontName, "OCR") Then addToFontList = False
    ElseIf StartsWith(fontName, "Open Sans") Then addToFontList = False
    ElseIf StartsWith(fontName, "Oranienbaum") Then addToFontList = False
    ElseIf StartsWith(fontName, "Oswald") Then addToFontList = False
    ElseIf StartsWith(fontName, "Oxygen") Then addToFontList = False
    ElseIf StartsWith(fontName, "PT") Then addToFontList = False
    ElseIf StartsWith(fontName, "Pacifico") Then addToFontList = False
    ElseIf StartsWith(fontName, "Palace") Then addToFontList = False
    ElseIf StartsWith(fontName, "Palanquin") Then addToFontList = False
    ElseIf StartsWith(fontName, "Patrick") Then addToFontList = False
    ElseIf StartsWith(fontName, "Petit") Then addToFontList = False
    ElseIf StartsWith(fontName, "Playfair") Then addToFontList = False
    ElseIf StartsWith(fontName, "Plantagenet") Then addToFontList = False
    ElseIf StartsWith(fontName, "PMing") Then addToFontList = False
    ElseIf StartsWith(fontName, "Poiret") Then addToFontList = False
    ElseIf StartsWith(fontName, "Poppins") Then addToFontList = False
    ElseIf StartsWith(fontName, "Posterama") Then addToFontList = False
    ElseIf StartsWith(fontName, "Pridi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Prompt") Then addToFontList = False
    ElseIf StartsWith(fontName, "Quattro") Then addToFontList = False
    ElseIf StartsWith(fontName, "Questrial") Then addToFontList = False
    ElseIf StartsWith(fontName, "QuickType") Then addToFontList = False
    ElseIf StartsWith(fontName, "Quire") Then addToFontList = False
    ElseIf StartsWith(fontName, "Raleway") Then addToFontList = False
    ElseIf StartsWith(fontName, "Rastanty") Then addToFontList = False
    ElseIf StartsWith(fontName, "Reem") Then addToFontList = False
    ElseIf StartsWith(fontName, "Roboto") Then addToFontList = False
    ElseIf StartsWith(fontName, "Rod") Then addToFontList = False
    ElseIf StartsWith(fontName, "STCaiyun") Then addToFontList = False
    ElseIf StartsWith(fontName, "STF") Then addToFontList = False
    ElseIf StartsWith(fontName, "STH") Then addToFontList = False
    ElseIf StartsWith(fontName, "STK") Then addToFontList = False
    ElseIf StartsWith(fontName, "STX") Then addToFontList = False
    ElseIf StartsWith(fontName, "STZ") Then addToFontList = False
    ElseIf StartsWith(fontName, "Sacramento") Then addToFontList = False
    ElseIf StartsWith(fontName, "Sagona") Then addToFontList = False
    ElseIf StartsWith(fontName, "Sans Serif Collection") Then addToFontList = False
    ElseIf StartsWith(fontName, "Sakkal") Then addToFontList = False
    ElseIf StartsWith(fontName, "Seaford") Then addToFontList = False
    ElseIf StartsWith(fontName, "Secular") Then addToFontList = False
    ElseIf StartsWith(fontName, "Segoe") Then addToFontList = False
    ElseIf StartsWith(fontName, "Selawik") Then addToFontList = False
    ElseIf StartsWith(fontName, "Shadows") Then addToFontList = False
    ElseIf StartsWith(fontName, "Shonar") Then addToFontList = False
    ElseIf StartsWith(fontName, "Shruti") Then addToFontList = False
    ElseIf StartsWith(fontName, "SimHei") Then addToFontList = False
    ElseIf StartsWith(fontName, "Simplified") Then addToFontList = False
    ElseIf StartsWith(fontName, "Sitka") Then addToFontList = False
    ElseIf StartsWith(fontName, "Skeena") Then addToFontList = False
    ElseIf StartsWith(fontName, "Statliches") Then addToFontList = False
    ElseIf StartsWith(fontName, "Suez") Then addToFontList = False
    ElseIf StartsWith(fontName, "TH") Then addToFontList = False
    ElseIf StartsWith(fontName, "Tahoma") Then addToFontList = False
    ElseIf StartsWith(fontName, "Tenorite") Then addToFontList = False
    ElseIf StartsWith(fontName, "Titillum") Then addToFontList = False
    ElseIf StartsWith(fontName, "Times New Roman") Then addToFontList = False
    ElseIf StartsWith(fontName, "Trade") Then addToFontList = False
    ElseIf StartsWith(fontName, "Traditional") Then addToFontList = False
    ElseIf StartsWith(fontName, "Trirong") Then addToFontList = False
    ElseIf StartsWith(fontName, "Tunga") Then addToFontList = False
    ElseIf StartsWith(fontName, "UD Digi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Ubuntu") Then addToFontList = False
    ElseIf StartsWith(fontName, "Univers") Then addToFontList = False
    ElseIf StartsWith(fontName, "Urdu") Then addToFontList = False
    ElseIf StartsWith(fontName, "Utsaah") Then addToFontList = False
    ElseIf StartsWith(fontName, "Vani") Then addToFontList = False
    ElseIf StartsWith(fontName, "Varela") Then addToFontList = False
    ElseIf StartsWith(fontName, "Vijaya") Then addToFontList = False
    ElseIf StartsWith(fontName, "Vivaldi") Then addToFontList = False
    ElseIf StartsWith(fontName, "Vrinda") Then addToFontList = False
    ElseIf StartsWith(fontName, "Walbaum") Then addToFontList = False
    ElseIf StartsWith(fontName, "Webdings") Then addToFontList = False
    ElseIf StartsWith(fontName, "Wingdings") Then addToFontList = False
    ElseIf StartsWith(fontName, "Work Sans") Then addToFontList = False
    ElseIf StartsWith(fontName, "Yesteryear") Then addToFontList = False
    ElseIf StartsWith(fontName, "Yu") Then addToFontList = False
        
    ' These are variations of a font
    ElseIf EndsWith(fontName, "Black") Then addToFontList = False
    ElseIf EndsWith(fontName, "Bold ITC") Then addToFontList = False
    ElseIf EndsWith(fontName, "Bold") Then addToFontList = False
    ElseIf EndsWith(fontName, "Compressed") Then addToFontList = False
    ElseIf EndsWith(fontName, "Cond") Then addToFontList = False
    ElseIf EndsWith(fontName, "Conde") Then addToFontList = False
    ElseIf EndsWith(fontName, "Conden") Then addToFontList = False
    ElseIf EndsWith(fontName, "Condensed") Then addToFontList = False
    ElseIf EndsWith(fontName, "Demi ITC") Then addToFontList = False
    ElseIf EndsWith(fontName, "Demi") Then addToFontList = False
    ElseIf EndsWith(fontName, "Expanded") Then addToFontList = False
    ElseIf EndsWith(fontName, "ExtB") Then addToFontList = False
    ElseIf EndsWith(fontName, "Extended") Then addToFontList = False
    ElseIf EndsWith(fontName, "Hand") Then addToFontList = False
    ElseIf EndsWith(fontName, "Heavy") Then addToFontList = False
    ElseIf EndsWith(fontName, "Light ITC") Then addToFontList = False
    ElseIf EndsWith(fontName, "Light") Then addToFontList = False
    ElseIf EndsWith(fontName, "Lt") Then addToFontList = False
    ElseIf EndsWith(fontName, "Medium ITC") Then addToFontList = False
    ElseIf EndsWith(fontName, "Medium") Then addToFontList = False
    ElseIf EndsWith(fontName, "Nova") Then addToFontList = False
    ElseIf EndsWith(fontName, "Pro") Then addToFontList = False
    ElseIf EndsWith(fontName, "Schoolbook") Then addToFontList = False
    ElseIf EndsWith(fontName, "Text") Then addToFontList = False
    ElseIf EndsWith(fontName, "Thin") Then addToFontList = False
    ElseIf EndsWith(fontName, "UI") Then addToFontList = False
    ElseIf EndsWith(fontName, "XBd") Then addToFontList = False
    ElseIf EndsWith(fontName, ".tmp") Then addToFontList = False
    End If
End Function

Public Sub CreateColorImageDir()
#If Mac Then
    colorImageDir = GetTempDirectory()
#Else
    colorImageDir = Environ$("AppData")
#End If

    colorImageDir = colorImageDir & Application.pathSeparator & PRODUCT_TEMPDIR
    CreateDirectory colorImageDir
    
    colorImageDir = colorImageDir & Application.pathSeparator & "colors"
    CreateDirectory colorImageDir
End Sub

Public Sub CreateFontImageDir()
#If Mac Then
    fontImageDir = GetTempDirectory()
#Else
    fontImageDir = Environ$("AppData")
#End If

    fontImageDir = fontImageDir & Application.pathSeparator & PRODUCT_TEMPDIR
    CreateDirectory fontImageDir
    
    fontImageDir = fontImageDir & Application.pathSeparator & "fonts"
    CreateDirectory fontImageDir
End Sub

Public Function GetColorImageDir() As String
    GetColorImageDir = colorImageDir
End Function

Public Function GetFontImageDir() As String
    GetFontImageDir = fontImageDir
End Function

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub designerHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLStyleDesignerTab").Value, NewWindow:=True
End Sub

Private Function IsProgressIndicatorNeeded() As Boolean
    IsProgressIndicatorNeeded = False
    
    If colorScheme = COLOR_SCHEME_X11 Or colorScheme = COLOR_SCHEME_SVG Then
        IsProgressIndicatorNeeded = True
    End If
End Function

Private Sub InvalidateRibbonColorList(ByVal controlName As String)
    If Not IsProgressIndicatorNeeded() Then
        InvalidateRibbonControl controlName
    End If
End Sub


