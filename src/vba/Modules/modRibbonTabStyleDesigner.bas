Attribute VB_Name = "modRibbonTabStyleDesigner"
' Copyright (c) 2015-2025 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule IntegerDataType, AssignmentNotUsed, UseMeaningfulName, UnassignedVariableUsage, ProcedureNotUsed, ParameterNotUsed, ImplicitByRefModifier

Option Explicit

' Cached color names
Private x11Colors As Variant
Private svgColors As Variant
Private brewerColors As Variant

' Trigger load Brewer color scheme if selection changes
Private brewerColorsAreFresh As Boolean

' Where the color gallery images are stored
Private colorImageDir As String

' An in-memory cache of image objects enables reuse across multiple color galleries.
' This reduces overall memory consumption and improves performance by avoiding
' redundant image loading for each color gallery control.
Private colorImageCache As Dictionary

' The list of font names
Private fontList As Variant

' Where the font gallery images are stored
Private fontImageDir As String

' An in-memory cache of image objects enables reuse across multiple font galleries.
' This reduces overall memory consumption and improves performance by avoiding
' redundant image loading for each font gallery control.
Private fontImageCache As Dictionary

' Lists of font name filters used to exclude fonts which Graphviz cannot render.
Private excludedPrefixes As Dictionary
Private excludedSuffixes As Dictionary
Private fontExclusionsInitialized As Boolean

' There are seven color galleries on the Style Designer ribbon tab. X11 is Graphviz’s default
' color scheme and includes 656 colors. For each color, two callbacks are triggered:
' one to retrieve the label (shown on hover), and one to retrieve or generate the image.
' This results in 656 × 2 = 1,312 callbacks per gallery, or 9,184 total callbacks when
' all galleries are loaded with the ribbon. Saving even a few milliseconds per callback
' yields noticeable performance improvements.

' A local variable named "kolorScheme" (note the initial "K") is used to cache the color
' scheme name during gallery refresh. A key performance optimization relies on understanding
' the Ribbon callback sequence: getItemCount() is called once, followed by getItemLabel()
' and getItemImage() for each item specified by getItemCount(). Although the color scheme
' is stored in a worksheet cell, reading from the cell repeatedly is slow-especially when
' multiplied across thousands of callbacks.

' To improve performance, the color scheme is read from the cell once in getItemCount()
' and stored in kolorScheme. This cached value remains valid throughout the subsequent
' getItemLabel() and getItemImage() calls. kolorScheme should only be used within the
' color gallery callbacks. The authoritative source for the color scheme is the cell itself;
' do not assume kolorScheme holds a current or valid value outside the callback sequence.
Private kolorScheme As String

' Represents color information for passing to Subs and Functions
Private Type ColorInfo
    name As String
    scheme As String
    imageFile As String
    RGB As Long
End Type

' Type definition for file path components
Private Type FilePathInfo
    fileName As String
    directory As String
End Type

#If Mac Then
' On Mac, some functions are handled via AppleScript. Additional functionality has been added in successive releases.
' This variable is initialized when the spreadsheet opens and is used to verify whether the correct script is available
' to perform a given action. If a version mismatch is detected, the corresponding capability is typically set to
' "visible=false" to prevent the user from attempting to use it.
Private scriptVersion As Long

Public Sub SetScriptVersion(ByVal version As Long)
    scriptVersion = version
End Sub

#End If

' ===========================================================================
' Ribbon callbacks for "Style Designer" ribbon tab
' ===========================================================================

' ===========================================================================
' Callbacks for colorScheme

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub colorScheme_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    If Left$(itemId, 4) = "cs_x" Then Exit Sub ' Blank gallery image selected
    
    Dim colorScheme As String
    If index = 0 Then
        colorScheme = vbNullString
    Else
        colorScheme = Mid$(itemId, Len("cs_") + 1)
    End If
    
    ' If color scheme is not X11 or SVG then it is a Brewer color scheme.
    ' Loading the brewerColors array is deferred until the next time the array is
    ' referenced (i.e. lazy load).
    If colorScheme <> COLOR_SCHEME_X11 And colorScheme <> COLOR_SCHEME_SVG Then
        brewerColorsAreFresh = False
    End If
    
    OptimizeCode_Begin
    SaveStyleDesignerSetting DESIGNER_COLOR_SCHEME, colorScheme
    StyleDesignerSheet.Range("FontColor,BorderColor,FillColor,GradientFillColor,GradientFillType,GradientFillAngle,GradientFillWeight,EdgeColor1,EdgeColor2,EdgeColor3,EdgeLabelFontColor").ClearContents
    OptimizeCode_End
    
    ColorLoadImageCache
    RefreshControlsColor
    
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub colorScheme_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("cs_", DESIGNER_COLOR_SCHEME)
End Sub

' ===========================================================================
' Callbacks for fontColor

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_FONT_COLOR, COLOR_BLACK, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub null_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = vbNullString
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_LABEL_FONT_COLOR) Then
        ColorGetImage DESIGNER_FONT_COLOR, COLOR_BLACK, returnedVal
    Else
        ColorGetImage DESIGNER_EDGE_LABEL_FONT_COLOR, COLOR_BLACK, returnedVal
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontColor_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    If index = 0 Then
        ClearStyleDesignerSetting DESIGNER_FONT_COLOR
    Else
        SaveColor index, DESIGNER_FONT_COLOR
    End If
    InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontColor_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveColor index, DESIGNER_EDGE_LABEL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub colorPicker_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = ColorPickerGetVisible(control.ID)
End Sub

Private Function ColorPickerGetVisible(controlId As String) As Boolean
    Dim visible As Boolean
    
#If Mac Then
    ' Show the color picker if we have access to the "pickColor" AppleScript function.
    ' It was introduced in the 3rd release of the ExcelToGraphviz.applescript script.
    If scriptVersion >= 3 Then
        ColorPickerGetVisible = True
        Exit Function
    End If
#End If
    
    ' Windows does not need the separator between the gallery controls, and the
    ' buttons which launch the color picker dialog.
    Select Case controlId
        Case "fontColorSeparator":      visible = False
        Case "borderColorSeparator":    visible = False
        Case "fillColorSeparator":      visible = False
        Case "edgeColorSeparator":      visible = False
        Case "edgeLabelColorSeparator": visible = False
        Case Else:                      visible = True
    End Select

#If Mac Then
    ' Invert the value of visible on Mac to do the opposite of the behavior on Windows,
    ' i.e. show the separators.
    visible = Not visible
#End If

    ColorPickerGetVisible = visible
End Function

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub color_getItemImage(ByVal control As IRibbonControl, ByVal index As Long, ByRef image As Variant)
    ' Get the color scheme
    Dim color As ColorInfo
    
    ' Initialize RGB color
    color.RGB = 0
    
    ' Get the color name
    If index = 0 Then   ' Determine the default color for the attribute
        color.scheme = GetColorScheme()
        ColorGetDefaultColorByControlId color, control.ID
    Else
        ' See comment at top of module regarding variable "kolorScheme"
        color.scheme = kolorScheme
        color.name = ColorGetNameByIndex(color, index)
    End If
    
    ' Get/create color image
    ColorGetOrCreateImage color, image
End Sub

Private Sub ColorGetImage(ByVal cellName As String, ByVal defaultColor As String, ByRef image As Variant)
    ' Get the color scheme
    Dim color As ColorInfo
    color.scheme = GetColorScheme()
    
    ' Initialize RGB color
    color.RGB = 0
    
    ' Get the color name
    Dim colorName As String
    colorName = StyleDesignerSetting(cellName)
    If Len(colorName) = 0 Then
        If Len(defaultColor) = 0 Then
            Set image = Nothing
            Exit Sub
        End If
        color.name = defaultColor
        color.scheme = COLOR_SCHEME_X11
    Else
        color.name = colorName
    End If
    
    ' Get/create color image
    ColorGetOrCreateImage color, image
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub color_getItemCount(ByVal control As IRibbonControl, ByRef count As Variant)
    ' Lazy creation of colorImageCache dictionary
    If colorImageCache Is Nothing Then
        Set colorImageCache = New Dictionary
    End If
    
    ' Load the array of color names, and obtain the quantity of colors
    count = LoadColorNameArray()
    
    ' See comment at top of module regarding variable "kolorScheme"
    kolorScheme = GetColorScheme()
    
    ' Hack to disable loading the hidden dropdowns
    If StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE Or StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER Then
        If control.ID = RIBBON_CTL_EDGE_COLOR1 Or control.ID = RIBBON_CTL_EDGE_COLOR2 Or control.ID = RIBBON_CTL_EDGE_COLOR3 Or control.ID = RIBBON_CTL_EDGE_LABEL_FONT_COLOR Then
            count = 0
        End If
        
        If control.ID = RIBBON_CTL_GRADIENT_FILL_COLOR And CellIsEmpty(DESIGNER_FILL_COLOR) Then
            count = 0
        End If
        
    ElseIf StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_EDGE Then
        If control.ID = RIBBON_CTL_FILL_COLOR Or control.ID = RIBBON_CTL_GRADIENT_FILL_COLOR Or control.ID = RIBBON_CTL_BORDER_COLOR Then
            count = 0
        ElseIf control.ID = RIBBON_CTL_EDGE_COLOR2 Then
            If CellIsEmpty(DESIGNER_EDGE_COLOR_1) Then
                count = 0
            End If
        ElseIf control.ID = RIBBON_CTL_EDGE_COLOR3 Then
            If CellIsEmpty(DESIGNER_EDGE_COLOR_2) Then
                count = 0
            End If
        End If
    End If
End Sub

Private Function LoadColorNameArray() As Long
    ' Get the color scheme
    Dim colorScheme As String
    colorScheme = GetColorScheme()

    ' Lazy cache the large color lists in arrays to improve performance over individual cell access
    Select Case colorScheme
        Case COLOR_SCHEME_X11
            If IsEmpty(x11Colors) Then
                x11Colors = Application.WorksheetFunction.Transpose(HelpColorsSheet.Range("CS_X11"))  ' 656 colors
            End If
            LoadColorNameArray = UBound(x11Colors) - LBound(x11Colors) + 2

        Case COLOR_SCHEME_SVG
            If IsEmpty(svgColors) Then
                svgColors = Application.WorksheetFunction.Transpose(HelpColorsSheet.Range("CS_SVG"))  ' 147 colors
            End If
            LoadColorNameArray = UBound(svgColors) - LBound(svgColors) + 2

        Case Else
            If Not brewerColorsAreFresh Then
                brewerColors = Application.WorksheetFunction.Transpose( _
                    Application.WorksheetFunction.Transpose( _
                        HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & colorScheme)))
                brewerColorsAreFresh = True
            End If
            LoadColorNameArray = UBound(brewerColors) - LBound(brewerColors) + 2
    End Select
End Function

Public Sub color_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim cellName As String
    Dim cellValue As String
    
    Select Case control.ID
        Case RIBBON_CTL_FONT_COLOR
            cellName = DESIGNER_FONT_COLOR
            
        Case RIBBON_CTL_BORDER_COLOR
            cellName = DESIGNER_BORDER_COLOR

        Case RIBBON_CTL_FILL_COLOR
            cellName = DESIGNER_FILL_COLOR
            
        Case RIBBON_CTL_GRADIENT_FILL_COLOR
            cellName = DESIGNER_GRADIENT_FILL_COLOR
        
        Case RIBBON_CTL_EDGE_COLOR1
            cellName = DESIGNER_EDGE_COLOR_1
            
        Case RIBBON_CTL_EDGE_COLOR2
            cellName = DESIGNER_EDGE_COLOR_2
            
        Case RIBBON_CTL_EDGE_COLOR3
            cellName = DESIGNER_EDGE_COLOR_3
            
        Case RIBBON_CTL_EDGE_LABEL_FONT_COLOR
            ' Fall back to overall font color if label font color is not set
            If CellIsEmpty(DESIGNER_EDGE_LABEL_FONT_COLOR) Then
                cellName = DESIGNER_FONT_COLOR
            Else
                cellName = DESIGNER_EDGE_LABEL_FONT_COLOR
            End If
    End Select
    
    cellValue = StyleDesignerSetting(cellName)
    If cellValue = vbNullString Then
        label = GetLabel(control.ID)
    Else
        label = cellValue
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub color_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    ' See comment at top of module regarding variable "kolorScheme"
    Dim color As ColorInfo
    color.scheme = kolorScheme
    label = ColorGetNameByIndex(color, index)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_FONT_COLOR)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_EDGE_LABEL_FONT_COLOR)
End Sub

' ===========================================================================
' Callbacks for borderColor

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_BORDER_COLOR, COLOR_BLACK, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_BORDER_COLOR)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderColor_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveColor index, DESIGNER_BORDER_COLOR
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for fontName

Public Sub LoadFontImageCache()
    ' Cache the list of fonts in an array
    If IsEmpty(fontList) Then
        fontList = getFontList()
    End If
    
    ' Lazy creation of fontImageCache dictionary
    If fontImageCache Is Nothing Then
        Set fontImageCache = New Dictionary
        fontImageCache.CompareMode = TextCompare
    End If
    
    ' Build the path to where the images are kept
    CreateFontImageDir
    
    ' Load any files which already exist. If they have not been created, defer creation until
    ' the Style Designer worksheet is accessed by the user.
    AddToFontCache "defaultFont", "defaultFont"
    Dim i As Long
    For i = LBound(fontList) To UBound(fontList)
        AddToFontCache CStr(fontList(i)), CStr(fontList(i))
    Next i
End Sub

Private Sub AddToFontCache(cacheKey As String, fontName As String)
    If Len(cacheKey) = 0 Or Len(fontName) = 0 Then Exit Sub
    If fontImageCache.Exists(cacheKey) Then Exit Sub
    
    Dim imageFile As String
    imageFile = GetFontImageDir() & Application.pathSeparator & fontName & DOT & RIBBON_EXT_FONT

    Dim image As StdPicture
    On Error Resume Next
    Set image = LoadPicture(imageFile)
    On Error GoTo 0

    If Not image Is Nothing Then
        fontImageCache.Add cacheKey, image
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getItemCount(ByVal control As IRibbonControl, ByRef count As Variant)
    
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
        count = 0
    Else
        count = (UBound(fontList) - LBound(fontList) + 1)
        If count > 1000 Then    ' Microsoft caps dropdown lists at 1000 items
            count = 1000
        End If
    End If
    
    ' Hack to disable loading the font which will not be displayed
    If StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE Or StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER Then
        If control.ID = RIBBON_CTL_LABEL_FONT_NAME Then
            count = 0
        End If
    End If
End Sub

Public Sub fontName_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim fontName As String
    fontName = StyleDesignerSetting(DESIGNER_FONT_NAME)
    If Len(fontName) = 0 Then
        returnedVal = GetLabel(control.ID)
    Else
        returnedVal = fontName
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef returnedVal As Variant)
    If index = 0 Then
        returnedVal = "Times-Roman"
    Else
        returnedVal = fontList(index)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveStyleDesignerSetting DESIGNER_FONT_NAME, FontGetNameByIndex(index)
    RenderPreview
    InvalidateRibbonControl RIBBON_CTL_FONT_NAME
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_NAME
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef listIndex As Variant)
    listIndex = FontGetIndexByName(StyleDesignerSetting(DESIGNER_FONT_NAME))
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getImage(ByVal control As IRibbonControl, ByRef image As Variant)
    Dim fontName As String
    fontName = StyleDesignerSetting(DESIGNER_FONT_NAME)
    If Len(fontName) = 0 Then fontName = "defaultFont"
    FontGetOrCreateImage fontName, image
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontName_getItemImage(ByVal control As IRibbonControl, ByVal index As Long, ByRef image As Variant)
    If index < 0 Then Exit Sub
    ' Get the font name
    Dim fontName As String
    fontName = FontGetNameByIndex(index)
    If Len(fontName) = 0 Then
        fontName = "defaultFont"
    End If
    FontGetOrCreateImage fontName, image
End Sub

Private Sub FontGetOrCreateImage(fontName As String, ByRef image As Variant)
    ' See if the font's image is already in cache
    If fontImageCache.Exists(fontName) Then
        Set image = fontImageCache.item(fontName)
        Exit Sub
    End If
  
    ' Build the path to where the images are kept
    Dim imageFile As String
    imageFile = GetFontImageDir() & Application.pathSeparator & fontName & DOT & RIBBON_EXT_FONT
    
    ' If the image already exists we should be able to load it
    On Error Resume Next
    Set image = LoadPicture(imageFile)
    On Error GoTo 0

    If image Is Nothing Then    ' the image does not exist, create one
        Application.StatusBar = replace(GetMessage("statusbarCreateFontImage"), "{fontName}", fontName)
        If fontName_createItemImage(fontName, imageFile, RIBBON_EXT_FONT) Then
            On Error Resume Next
            Set image = LoadPicture(imageFile)
            On Error GoTo 0
        End If
        Application.StatusBar = False
    End If

    ' Add the loaded font image to the cache
    fontImageCache.Add fontName, image
End Sub

Private Function fontName_createItemImage(ByVal fontName As String, ByVal imageFile As String, ByVal imageFormat As Variant) As Boolean
    fontName_createItemImage = False
    
    If Len(fontName) = 0 Or Len(imageFile) = 0 Then Exit Function

    On Error GoTo ErrorHandler
               
    ' Define a simple one node DOT graph which will create a 48x48 pixel image suitable for display in the ribbon
    Dim dotSource As String
    dotSource = "digraph g{ bgcolor=gray pad=0 margin=0 a[ shape=square style=filled fillcolor=white fontcolor=black fontsize=38 dpi=96 height=0.48 width=0.48 fixedsize=true penwidth=0"
    If fontName <> "defaultFont" Then
        dotSource = dotSource & " fontname=" & AddQuotesConditionally(fontName)
    End If
    dotSource = dotSource & " label=" & AddQuotes("A") & " ]; }"
    
    ' Determine settings for console output
    Dim console As consoleOptions
    console = GetSettingsForConsole()
    
    ' Instantiate a new Graphviz object
    Dim graphvizObj As Graphviz
    Set graphvizObj = New Graphviz

    ' Configure Graphviz
    With graphvizObj
        .GraphvizPath = SettingsSheet.Range(SETTINGS_GV_PATH).Value2
        .OutputDirectory = GetTempDirectory()
        .FilenameBase = fontName
        .GraphFormat = imageFormat
        .Verbose = console.graphvizVerbose
        .CaptureMessages = console.logToConsole
    
        ' Override the diagram file to use the path specified by the caller
        .DiagramFilename = imageFile
        
        ' Write the Graphviz data to a file so it can be sent to a rendering engine
        .graphvizSource = dotSource
        .SourceToFile
        
        ' Generate an image using graphviz
        .RenderGraph
    End With
    
    ' Display any console output first
    DisplayTextOnConsoleWorksheet graphvizObj.GraphvizCommand, graphvizObj.GraphvizMessages
    
    fontName_createItemImage = True
    
Cleanup:
    Set graphvizObj = Nothing
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Debug.Print "fontName_createItemImage Error: " & Err.Description
    Resume Cleanup
End Function

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontName_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveStyleDesignerSetting DESIGNER_EDGE_LABEL_FONT_NAME, FontGetNameByIndex(index)
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_NAME
    RenderPreview
End Sub

Public Sub labelFontName_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim fontName As String
    fontName = StyleDesignerSetting(DESIGNER_EDGE_LABEL_FONT_NAME)
    If Len(fontName) = 0 Then ' Revert to designer font name
        fontName = StyleDesignerSetting(DESIGNER_FONT_NAME)
    End If
    
    If Len(fontName) = 0 Then
        returnedVal = GetLabel(control.ID)
    Else
        returnedVal = fontName
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontName_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef listIndex As Variant)
    Dim fontName As String
    fontName = StyleDesignerSetting(DESIGNER_EDGE_LABEL_FONT_NAME)
    If Len(fontName) = 0 Then ' Revert to designer font name
        fontName = StyleDesignerSetting(DESIGNER_FONT_NAME)
        If Len(fontName) = 0 Then fontName = "defaultFont" ' revert to Graphviz default font
    End If
    listIndex = FontGetIndexByName(fontName)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelFontName_getImage(ByVal control As IRibbonControl, ByRef image As Variant)
    Dim fontName As String
    fontName = StyleDesignerSetting(DESIGNER_EDGE_LABEL_FONT_NAME)
    If Len(fontName) = 0 Then ' Revert to designer font name
        fontName = StyleDesignerSetting(DESIGNER_FONT_NAME)
        If Len(fontName) = 0 Then fontName = "defaultFont" ' revert to Graphviz default ont
    End If
    FontGetOrCreateImage fontName, image
End Sub

Private Function FontGetNameByIndex(ByVal index As Long) As String
    If index = 0 Then
        FontGetNameByIndex = vbNullString
    Else
        FontGetNameByIndex = Trim$(CStr(fontList(index)))
    End If
End Function

Private Function FontGetIndexByName(ByVal fontName As String) As Long
    FontGetIndexByName = 0
    If Len(fontName) = 0 Then Exit Function
    
    Dim i As Long
    For i = LBound(fontList) To UBound(fontList)
        If fontName = fontList(i) Then
            FontGetIndexByName = i
            Exit Function
        End If
    Next i
End Function

' ===========================================================================
' Callbacks for fontSize

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    listSize = ListsSheet.Range(LISTS_FONT_SIZES).count + 1
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = "14"
    Else
        label = CStr(ListsSheet.Range(LISTS_FONT_SIZES).Cells.item(index, 1).Value2)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontSize_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    If index = 0 Then
        ClearStyleDesignerSetting DESIGNER_FONT_SIZE
    Else
        SaveStyleDesignerSetting DESIGNER_FONT_SIZE, ListsSheet.Range(LISTS_FONT_SIZES).Cells.item(index, 1).Value2
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
Private Sub labelFontSize_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    If index = 0 Then
        ClearStyleDesignerSetting DESIGNER_EDGE_LABEL_FONT_SIZE
    Else
        SaveStyleDesignerSetting DESIGNER_EDGE_LABEL_FONT_SIZE, ListsSheet.Range(LISTS_FONT_SIZES).Cells.item(index, 1).Value2
    End If
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeWeight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeWeight_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "weight_", DESIGNER_EDGE_WEIGHT
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeWeight_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("weight_", DESIGNER_EDGE_WEIGHT)
End Sub

' ===========================================================================
' Callbacks for edgeLabelAngle

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelAngle_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "angle_", DESIGNER_EDGE_LABEL_ANGLE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeLabelAngle_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("angle_", DESIGNER_EDGE_LABEL_ANGLE)
End Sub

' ===========================================================================
' Callbacks for edgeLabelDistance

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelDistance_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "distance_", DESIGNER_EDGE_LABEL_DISTANCE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeLabelDistance_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("distance_", DESIGNER_EDGE_LABEL_DISTANCE)
End Sub

' ===========================================================================
' Callbacks for borderPenWidth

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPenWidth_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "bw_", DESIGNER_BORDER_PEN_WIDTH
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub borderPenWidth_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("bw_", DESIGNER_BORDER_PEN_WIDTH)
End Sub

' ===========================================================================
' Callbacks for borderPeripheries

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPeripheries_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "p_", DESIGNER_BORDER_PERIPHERIES
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub borderPeripheries_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("p_", DESIGNER_BORDER_PERIPHERIES)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderPeripheries_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE
End Sub

' ===========================================================================
' Callbacks for designModeNode

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeNode_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_MODE, KEYWORD_NODE
    ShowLabelRows KEYWORD_NODE
    RefreshControlsDesignMode
    RefreshStyleDesignerRibbon
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeNode_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE
End Sub

' ===========================================================================
' Callbacks for designModeEdge

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeEdge_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_MODE, KEYWORD_EDGE
    ShowLabelRows KEYWORD_EDGE
    RefreshControlsDesignMode
    RefreshStyleDesignerRibbon
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeEdge_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_EDGE
End Sub

' ===========================================================================
' Callbacks for designModeCluster

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeCluster_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_MODE, KEYWORD_CLUSTER
    ShowLabelRows KEYWORD_CLUSTER
    RefreshControlsDesignMode
    RefreshStyleDesignerRibbon
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designModeCluster_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER
End Sub

Public Sub ShowLabelRows(ByVal designerMode As String)
    Application.ScreenUpdating = False
    
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
            StyleDesignerSheet.rows.item(labelRow).Hidden = False
            StyleDesignerSheet.rows.item(xlabelRow).Hidden = False
            StyleDesignerSheet.rows.item(tailLabelRow).Hidden = True
            StyleDesignerSheet.rows.item(headLabelRow).Hidden = True
            StyleDesignerSheet.CheckBoxes("IncludeLabelCheckBox").visible = True
            StyleDesignerSheet.CheckBoxes("IncludeExternalLabelCheckBox").visible = True
            StyleDesignerSheet.CheckBoxes("IncludeTailLabelCheckBox").visible = False
            StyleDesignerSheet.CheckBoxes("IncludeHeadLabelCheckBox").visible = False
        Case KEYWORD_EDGE
            StyleDesignerSheet.rows.item(labelRow).Hidden = False
            StyleDesignerSheet.rows.item(xlabelRow).Hidden = False
            StyleDesignerSheet.rows.item(tailLabelRow).Hidden = False
            StyleDesignerSheet.rows.item(headLabelRow).Hidden = False
            StyleDesignerSheet.CheckBoxes("IncludeLabelCheckBox").visible = True
            StyleDesignerSheet.CheckBoxes("IncludeExternalLabelCheckBox").visible = True
            StyleDesignerSheet.CheckBoxes("IncludeTailLabelCheckBox").visible = True
            StyleDesignerSheet.CheckBoxes("IncludeHeadLabelCheckBox").visible = True
        Case KEYWORD_CLUSTER
            StyleDesignerSheet.rows.item(labelRow).Hidden = False
            StyleDesignerSheet.rows.item(xlabelRow).Hidden = True
            StyleDesignerSheet.rows.item(tailLabelRow).Hidden = True
            StyleDesignerSheet.rows.item(headLabelRow).Hidden = True
            StyleDesignerSheet.CheckBoxes("IncludeLabelCheckBox").visible = True
            StyleDesignerSheet.CheckBoxes("IncludeExternalLabelCheckBox").visible = False
            StyleDesignerSheet.CheckBoxes("IncludeTailLabelCheckBox").visible = False
            StyleDesignerSheet.CheckBoxes("IncludeHeadLabelCheckBox").visible = False
    End Select
    
    Application.ScreenUpdating = True
End Sub

Public Sub ColorLoadImageCache()
    ' Load the color scheme
    Dim color As ColorInfo
    color.scheme = GetColorScheme()
    
    ' Load the array of color names, and obtain the quantity of colors
    Dim count As Long
    count = LoadColorNameArray()
    
    ' Lazy creation of colorImageCache dictionary
    If colorImageCache Is Nothing Then
        Set colorImageCache = New Dictionary
        colorImageCache.CompareMode = TextCompare
    End If
    
    ' Build the path to where the images are kept
    CreateColorImageDir
    
    ' Load any files which already exist. If they have not been created, defer creation until
    ' the Style Designer worksheet is accessed by the user.
    Dim arrayItem As Variant
    Select Case color.scheme
        Case COLOR_SCHEME_X11
            For Each arrayItem In x11Colors
                color.name = CStr(arrayItem)
                AddToColorCache color
           Next arrayItem
        Case COLOR_SCHEME_SVG
            For Each arrayItem In svgColors
                color.name = CStr(arrayItem)
                AddToColorCache color
            Next arrayItem
        Case Else
            For Each arrayItem In brewerColors
                color.name = CStr(arrayItem)
                AddToColorCache color
            Next arrayItem
    End Select
End Sub

Private Sub AddToColorCache(color As ColorInfo)
    Dim cacheKey As String
    cacheKey = ColorGetCacheKey(color)

    ' See if the image was previously loaded
    If colorImageCache.Exists(cacheKey) Then Exit Sub
    
    ' Not previously loaded, load thumbnail if previously created
    ColorSetImageFile color
    Dim image As StdPicture
    On Error Resume Next
    Set image = LoadPicture(color.imageFile)
    On Error GoTo 0

    If Not image Is Nothing Then
        colorImageCache.Add cacheKey, image
        ' Handle X11 Gray/Grey special case
        If StartsWith(cacheKey, "X11_Gray") Then
            Dim greyKey As String
            greyKey = "X11_Grey" & Right$(cacheKey, Len(cacheKey) - 8)
            colorImageCache.Add greyKey, image
        End If
    End If
End Sub

' ===========================================================================
' Callbacks for fillColor

'@Ignore ProcedureNotUsed
Private Sub fillColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_FILL_COLOR, vbNullString, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fillColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_FILL_COLOR)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fillColor_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveColor index, DESIGNER_FILL_COLOR
    If CellIsEmpty(DESIGNER_FILL_COLOR) Then
        StyleDesignerSheet.Range("GradientFillColor,GradientFillType,GradientFillAngle,GradientFillWeight").ClearContents
    End If
    RefreshControlsFillColor
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for gradientFillColor

'@Ignore ProcedureNotUsed
Private Sub gradientFillColor_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_GRADIENT_FILL_COLOR, vbNullString, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillColor_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_GRADIENT_FILL_COLOR)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillColor_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveColor index, DESIGNER_GRADIENT_FILL_COLOR
    If CellIsEmpty(DESIGNER_GRADIENT_FILL_COLOR) Then
        StyleDesignerSheet.Range("GradientFillType,GradientFillWeight,GradientFillAngle").ClearContents
    End If
    RefreshControlsGradientFill
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillColor_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If CellIsEmpty(DESIGNER_FILL_COLOR) Then
        visible = False
    Else
        visible = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillColorPicker_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    If CellIsEmpty(DESIGNER_FILL_COLOR) Then
        visible = False
    Else
        visible = True
    End If

#If Mac Then
    visible = visible And ColorPickerGetVisible(control.ID)
#End If
End Sub

' ===========================================================================
' Callbacks for gradientFillType

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillType_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "ft_", DESIGNER_GRADIENT_FILL_TYPE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub gradientFillType_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("ft_", DESIGNER_GRADIENT_FILL_TYPE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillType_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = Not CellIsEmpty(DESIGNER_GRADIENT_FILL_COLOR)
End Sub

Private Function CellIsEmpty(cellName As String) As Boolean
    If Len(StyleDesignerSetting(cellName)) = 0 Then
        CellIsEmpty = True
    Else
        CellIsEmpty = False
    End If
End Function
' ===========================================================================
' Callbacks for gradientFillAngle

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillAngle_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "a_", DESIGNER_GRADIENT_FILL_ANGLE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub gradientFillAngle_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("a_", DESIGNER_GRADIENT_FILL_ANGLE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillAngle_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = Not CellIsEmpty(DESIGNER_GRADIENT_FILL_COLOR)
End Sub

' ===========================================================================
' Callbacks for GradientFillWeight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillWeight_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "gw_", DESIGNER_GRADIENT_FILL_WEIGHT
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub gradientFillWeight_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("gw_", DESIGNER_GRADIENT_FILL_WEIGHT)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub gradientFillWeight_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = Not CellIsEmpty(DESIGNER_GRADIENT_FILL_COLOR)
End Sub

' ===========================================================================
' Callbacks for labelJustification

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub labelJustification_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER
End Sub

' ===========================================================================
' Callbacks for headPort

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadPort_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "hp_", DESIGNER_EDGE_HEAD_PORT
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeHeadPort_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("hp_", DESIGNER_EDGE_HEAD_PORT)
End Sub

' ===========================================================================
' Callbacks for tailPort

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeTailPort_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "tp_", DESIGNER_EDGE_TAIL_PORT
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeTailPort_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("tp_", DESIGNER_EDGE_TAIL_PORT)
End Sub

' ===========================================================================
' Callbacks for edgeStyle

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeStyle_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "es_", DESIGNER_EDGE_STYLE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeStyle_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("es_", DESIGNER_EDGE_STYLE)
End Sub

' ===========================================================================
' Callbacks for nodeShape

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeShape_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "s_", DESIGNER_NODE_SHAPE
    StyleDesignerSheet.Range("NodeSides,NodeOrientation,NodeRegular,NodeSkew,NodeDistortion").ClearContents
    RefreshControlsPolygon
    RenderPreview
End Sub

Public Sub nodeShape_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim shape As String
    shape = StyleDesignerSetting(DESIGNER_NODE_SHAPE)
    If Len(shape) = 0 Then
        returnedVal = GetLabel(control.ID)
    Else
        returnedVal = shape
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeShape_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("s_", DESIGNER_NODE_SHAPE)
End Sub

' GetVisible callback for polygon shape

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeShape_isPolygon(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_NODE_SHAPE) = GRAPHVIZ_SHAPE_POLYGON
End Sub

' ===========================================================================
' Callbacks for nodeSides

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSides_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "si_", DESIGNER_NODE_SIDES
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeSides_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("si_", DESIGNER_NODE_SIDES)
End Sub

' ===========================================================================
' Callbacks for nodeRotation

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeRotation_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "r_", DESIGNER_NODE_ORIENTATION
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeRotation_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("r_", DESIGNER_NODE_ORIENTATION)
End Sub

' ===========================================================================
' Callbacks for borderStyle1

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle1_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "bs1_", DESIGNER_BORDER_STYLE1
    If CellIsEmpty(DESIGNER_BORDER_STYLE1) Then
        StyleDesignerSheet.Range("BorderStyle2,BorderStyle3").ClearContents
    End If
    InvalidateRibbonControl RIBBON_CTL_BORDER_STYLE2
    InvalidateRibbonControl RIBBON_CTL_BORDER_STYLE3
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub borderStyle1_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("bs1_", DESIGNER_BORDER_STYLE1)
End Sub

' ===========================================================================
' Callbacks for BorderStyle2

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle2_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "bs2_", DESIGNER_BORDER_STYLE2
    If CellIsEmpty(DESIGNER_BORDER_STYLE2) Then
        ClearStyleDesignerSetting DESIGNER_BORDER_STYLE3
    End If
    InvalidateRibbonControl RIBBON_CTL_BORDER_STYLE3
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub borderStyle2_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("bs2_", DESIGNER_BORDER_STYLE2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_BORDER_STYLE1)
End Sub

' ===========================================================================
' Callbacks for borderStyle3

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle3_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "bs3_", DESIGNER_BORDER_STYLE3
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub borderStyle3_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("bs3_", DESIGNER_BORDER_STYLE3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub borderStyle3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_BORDER_STYLE2)
End Sub

' ===========================================================================
' Callbacks for nodeHeight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeight_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "h_", DESIGNER_NODE_HEIGHT
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeHeight_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("h_", DESIGNER_NODE_HEIGHT)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeight_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = Not IsTrue(DESIGNER_NODE_METRIC)
End Sub

' ===========================================================================
' Callbacks for nodeHeightMetric

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeightMetric_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "mmh_", DESIGNER_NODE_HEIGHT
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeightMetric_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim metricEnabled As Boolean
    metricEnabled = (StyleDesignerSetting(DESIGNER_NODE_METRIC) = TOGGLE_YES)

    If Not metricEnabled Then
        returnedVal = 0
        Exit Sub
    End If

    Dim cellValue As String
    cellValue = StyleDesignerSetting(DESIGNER_NODE_HEIGHT)

    If Len(cellValue) = 0 Then
        returnedVal = 0
    Else
        returnedVal = CInt(cellValue) + 1
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeHeightMetric_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = IsTrue(DESIGNER_NODE_METRIC)
End Sub

' ===========================================================================
' Callbacks for nodeWidth

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidth_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "w_", DESIGNER_NODE_WIDTH
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeWidth_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("w_", DESIGNER_NODE_WIDTH)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidth_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = Not IsTrue(DESIGNER_NODE_METRIC)
End Sub

' ===========================================================================
' Callbacks for nodeWidthMetric

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidthMetric_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "mmw_", DESIGNER_NODE_WIDTH
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidthMetric_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim metricEnabled As Boolean
    metricEnabled = (StyleDesignerSetting(DESIGNER_NODE_METRIC) = TOGGLE_YES)

    If Not metricEnabled Then
        returnedVal = 0
        Exit Sub
    End If

    Dim cellValue As String
    cellValue = StyleDesignerSetting(DESIGNER_NODE_WIDTH)

    If Len(cellValue) = 0 Then
        returnedVal = 0
    Else
        returnedVal = CInt(cellValue) + 1
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeWidthMetric_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = IsTrue(DESIGNER_NODE_METRIC)
End Sub

' ===========================================================================
' Callbacks for nodeFixedSize

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeFixedSize_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "fs_", DESIGNER_NODE_FIXED_SIZE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeFixedSize_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = LCase$("fs_" & StyleDesignerSetting(DESIGNER_NODE_FIXED_SIZE))
End Sub

' ===========================================================================
' Callbacks for edgeColor1

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor1_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_EDGE_COLOR_1, COLOR_BLACK, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor1_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_EDGE_COLOR_1)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor1_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Application.enableEvents = False
    
    SaveColor index, DESIGNER_EDGE_COLOR_1
    
    If CellIsEmpty(DESIGNER_EDGE_COLOR_1) Then
        StyleDesignerSheet.Range("EdgeColor2,EdgeColor3").ClearContents
    End If
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    
    Application.enableEvents = True
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeColor2

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_EDGE_COLOR_2, vbNullString, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_EDGE_COLOR_2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Application.enableEvents = False
    
    SaveColor index, DESIGNER_EDGE_COLOR_2
    
    If CellIsEmpty(DESIGNER_EDGE_COLOR_2) Then
        ClearStyleDesignerSetting DESIGNER_EDGE_COLOR_3
    End If
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
    Application.enableEvents = True
    
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_COLOR_1) Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor2Picker_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_COLOR_1) Then
        returnedVal = False
    Else
        returnedVal = True
    End If
    
#If Mac Then
    returnedVal = returnedVal And ColorPickerGetVisible(control.ID)
#End If
End Sub

' ===========================================================================
' Callbacks for edgeColor3

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_getImage(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    ColorGetImage DESIGNER_EDGE_COLOR_3, vbNullString, returnedVal
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef index As Variant)
    index = ColorGetIndex(DESIGNER_EDGE_COLOR_3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveColor index, DESIGNER_EDGE_COLOR_3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_COLOR_2) Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeColor3Picker_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_COLOR_2) Then
        returnedVal = False
    Else
        returnedVal = True
    End If
    
#If Mac Then
    returnedVal = returnedVal And ColorPickerGetVisible(control.ID)
#End If
End Sub
' ===========================================================================
' Callbacks for Arrow Tail

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub groupArrowHead_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Dim direction As String
    Dim mode As String
    
    direction = StyleDesignerSetting(DESIGNER_EDGE_DIRECTION)
    mode = StyleDesignerSetting(DESIGNER_MODE)
    
    visible = mode = KEYWORD_EDGE And (direction = vbNullString Or direction = "forward" Or direction = "both")
End Sub

' ===========================================================================
' Callbacks for edgeArrowHead1

'@Ignore ParameterNotUsed
Public Sub edgeArrowHead1_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("h1_", DESIGNER_EDGE_ARROW_HEAD_1)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead1_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Dim newValue As String
    newValue = Mid$(itemId, Len("h1_") + 1)

    With StyleDesignerSheet
        .Range(DESIGNER_EDGE_ARROW_HEAD_1).Value2 = newValue

        If Len(newValue) = 0 Then
            .Range("EdgeArrowHead2,EdgeArrowHead3").ClearContents
        End If
    End With

    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD2
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD3
    RenderPreview
End Sub
' ===========================================================================
' Callbacks for edgeArrowHead2

'@Ignore ParameterNotUsed
Public Sub edgeArrowHead2_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("h2_", DESIGNER_EDGE_ARROW_HEAD_2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead2_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Dim newValue As String
    newValue = Mid$(itemId, Len("h2_") + 1)

    With StyleDesignerSheet
        .Range(DESIGNER_EDGE_ARROW_HEAD_2).Value2 = newValue

        If Len(newValue) = 0 Then
            .Range(DESIGNER_EDGE_ARROW_HEAD_3).ClearContents
        End If
    End With

    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_HEAD3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_EDGE_ARROW_HEAD_1)
End Sub

' ===========================================================================
' Callbacks for edgeArrowHead3

'@Ignore ParameterNotUsed
Public Sub edgeArrowHead3_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("h3_", DESIGNER_EDGE_ARROW_HEAD_3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead3_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "h3_", DESIGNER_EDGE_ARROW_HEAD_3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowHead3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_EDGE_ARROW_HEAD_2)
End Sub

' ===========================================================================
' Callbacks for Arrow Tail

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub groupArrowTail_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Dim direction As String
    Dim mode As String

    With StyleDesignerSheet
        direction = Trim$(.Range(DESIGNER_EDGE_DIRECTION).Value2)
        mode = Trim$(.Range(DESIGNER_MODE).Value2)
    End With

    visible = (mode = KEYWORD_EDGE) And (direction = "back" Or direction = "both")
End Sub

' ===========================================================================
' Callbacks for edgeArrowTail1

'@Ignore ParameterNotUsed
Public Sub edgeArrowTail1_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("t1_", DESIGNER_EDGE_ARROW_TAIL_1)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail1_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Dim newValue As String
    newValue = Mid$(itemId, Len("t1_") + 1)

    With StyleDesignerSheet
        .Range(DESIGNER_EDGE_ARROW_TAIL_1).Value2 = newValue

        If Len(newValue) = 0 Then
            .Range("EdgeArrowTail2,EdgeArrowTail3").ClearContents
        End If
    End With

    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL2
    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL3
    InvalidateRibbonControl RIBBON_CTL_EDGE_DIRECTION
    RenderPreview
End Sub
' ===========================================================================
' Callbacks for edgeArrowTail2

'@Ignore ParameterNotUsed
Public Sub edgeArrowTail2_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("t2_", DESIGNER_EDGE_ARROW_TAIL_2)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail2_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Dim newValue As String
    newValue = Mid$(itemId, Len("t2_") + 1)

    With StyleDesignerSheet
        .Range(DESIGNER_EDGE_ARROW_TAIL_2).Value2 = newValue

        If Len(newValue) = 0 Then
            .Range(DESIGNER_EDGE_ARROW_TAIL_3).ClearContents
        End If
    End With

    InvalidateRibbonControl RIBBON_CTL_EDGE_ARROW_TAIL3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail2_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_EDGE_ARROW_TAIL_1)
End Sub

' ===========================================================================
' Callbacks for edgeArrowTail3

'@Ignore ParameterNotUsed
Public Sub edgeArrowTail3_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("t3_", DESIGNER_EDGE_ARROW_TAIL_3)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail3_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "t3_", DESIGNER_EDGE_ARROW_TAIL_3
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowTail3_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_EDGE_ARROW_TAIL_2)
End Sub

' ===========================================================================
' Callbacks for edgeDirection

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeDirection_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    Dim direction As String
    direction = Mid$(itemId, Len("ed_") + 1)

    SaveStyleDesignerSetting DESIGNER_EDGE_DIRECTION, direction

    With StyleDesignerSheet
        Select Case direction
            Case vbNullString
                .Range("EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3").ClearContents

            Case GRAPHVIZ_DIR_BACK
                .Range("EdgeArrowHead1,EdgeArrowHead2,EdgeArrowHead3").ClearContents

            Case GRAPHVIZ_DIR_FORWARD
                .Range("EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3").ClearContents

            Case GRAPHVIZ_DIR_NONE
                .Range("EdgeArrowHead1,EdgeArrowHead2,EdgeArrowHead3," & _
                       "EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3,EdgeArrowSize").ClearContents

            Case GRAPHVIZ_DIR_BOTH
                ' No action needed
        End Select
    End With

    ' Invalidate affected controls
    Dim ctlList As Variant
    ctlList = Array( _
        RIBBON_CTL_EDGE_ARROW_SIZE, _
        RIBBON_CTL_EDGE_ARROW_HEAD1, RIBBON_CTL_EDGE_ARROW_HEAD2, RIBBON_CTL_EDGE_ARROW_HEAD3, RIBBON_GRP_EDGE_ARROW_HEAD, _
        RIBBON_CTL_EDGE_ARROW_TAIL1, RIBBON_CTL_EDGE_ARROW_TAIL2, RIBBON_CTL_EDGE_ARROW_TAIL3, RIBBON_GRP_EDGE_ARROW_TAIL, _
        RIBBON_GRP_EDGE_ARROW _
    )

    Dim i As Long
    For i = LBound(ctlList) To UBound(ctlList)
        InvalidateRibbonControl ctlList(i)
    Next i

    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeDirection_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("ed_", DESIGNER_EDGE_DIRECTION)
End Sub

' ===========================================================================
' Callbacks for edgeArrowSize

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowSize_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Dim direction As String
    direction = StyleDesignerSetting(DESIGNER_EDGE_DIRECTION)

    returnedVal = (direction <> GRAPHVIZ_DIR_NONE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeArrowSize_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "as_", DESIGNER_EDGE_ARROW_SIZE
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgeArrowSize_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("as_", DESIGNER_EDGE_ARROW_SIZE)
End Sub

' ===========================================================================
' Callbacks for edgePenWidth

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgePenWidth_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "ew_", DESIGNER_EDGE_PEN_WIDTH
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub edgePenWidth_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = "ew_" & format(StyleDesignerSheet.Range(DESIGNER_EDGE_PEN_WIDTH).value, "0.0")
End Sub

' ===========================================================================
' Callbacks for nodeImageName

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageName_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SaveStyleDesignerSetting DESIGNER_NODE_IMAGE_NAME, Text
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageName_getText(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_NODE_IMAGE_NAME)
End Sub

' ===========================================================================
' Callbacks for nodeImageRelativePath

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageRelativePath_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_NODE_IMAGE_RELATIVE_PATH, Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageRelativePath_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_NODE_IMAGE_RELATIVE_PATH) = TOGGLE_YES
End Sub


' ===========================================================================
' Callbacks for nodeRegular

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub regular_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_NODE_REGULAR, Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub regular_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_NODE_REGULAR) = TOGGLE_YES
End Sub

' ===========================================================================
' Callbacks for nodeSkew

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSkew_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SaveStyleDesignerSetting DESIGNER_NODE_SKEW, Text
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeSkew_getText(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_NODE_SKEW)
End Sub

' ===========================================================================
' Callbacks for nodeDistortion

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeDistortion_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SaveStyleDesignerSetting DESIGNER_NODE_DISTORTION, Text
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeDistortion_getText(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_NODE_DISTORTION)
End Sub

' ===========================================================================
' Callbacks for nodeImageChoose

' Handles the node image selection action from the ribbon
' @param control The IRibbonControl that triggered the action
Private Sub nodeImageChoose_onAction(ByVal control As IRibbonControl)
    Dim selectedPath As String
    selectedPath = SelectImageFile()
    
    If Len(selectedPath) = 0 Then Exit Sub
    
    Dim fileInfo As FilePathInfo
    fileInfo = SplitFilePath(selectedPath)
    
    UpdateImagePathIfNeeded fileInfo.directory
    
    ' Determine if relative path should be used
    Dim useRelativePath As Boolean
    useRelativePath = StyleDesignerSetting(DESIGNER_NODE_IMAGE_RELATIVE_PATH) = TOGGLE_YES
    
    Dim displayFileName As String
    If useRelativePath Then
        displayFileName = GetRelativePath(fileInfo.fileName, fileInfo.directory, ActiveWorkbook.path)
    Else
        displayFileName = fileInfo.fileName
    End If
    
    UpdateRibbonDisplay displayFileName
    RenderPreview
End Sub

' Selects an image file and returns its path
' @returns String The selected file path or empty string if cancelled
Private Function SelectImageFile() As String
#If Mac Then
    SelectImageFile = RunAppleScriptTask("chooseImageFile", "Select an image file")
#Else
    Dim dialog As FileDialog
    Set dialog = CreateFileDialog()
    
    If dialog.show <> -1 Then
        Set dialog = Nothing
        Exit Function
    End If
    
    SelectImageFile = dialog.SelectedItems.item(1)
    Set dialog = Nothing
#End If
End Function

' Creates and configures a file dialog for Windows
' @returns FileDialog Configured file dialog object
Private Function CreateFileDialog() As FileDialog
    Set CreateFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With CreateFileDialog
        .AllowMultiSelect = False
        .title = "Select an image file"
        .InitialFileName = ActiveWorkbook.path
        .Filters.Clear
        .Filters.Add "Image files", "*.bmp;*.gif;*.jpg;*.jpeg;*.png"
        .Filters.Add "All files", "*.*"
    End With
End Function

' Splits a file path into components
' @param fullPath The complete file path
' @returns FilePathInfo Record containing file and directory components
Private Function SplitFilePath(ByVal fullPath As String) As FilePathInfo
    Dim components() As String
    components = split(fullPath, Application.pathSeparator)
    
    Dim info As FilePathInfo
    info.fileName = components(UBound(components))
    info.directory = Left$(fullPath, Len(fullPath) - Len(info.fileName) - 1)
    
    SplitFilePath = info
End Function

' Calculates relative path from workbook path to image file
' @param fileName The image filename
' @param imageDir The image file directory
' @param workbookPath The workbook directory
' @returns String The relative path to the image
Private Function GetRelativePath(ByVal fileName As String, ByVal imageDir As String, ByVal workbookPath As String) As String
    Dim workbookComponents() As String
    Dim imageComponents() As String
    workbookComponents = split(workbookPath, Application.pathSeparator)
    imageComponents = split(imageDir, Application.pathSeparator)
    
    ' Find common path length
    Dim commonLength As Long
    Dim i As Long
    commonLength = Application.WorksheetFunction.Min(UBound(workbookComponents) + 1, UBound(imageComponents) + 1)
    For i = 0 To commonLength - 1
        If workbookComponents(i) <> imageComponents(i) Then
            Exit For
        End If
        commonLength = i + 1
    Next i
    
    ' Build relative path
    Dim relativePath As String
    relativePath = ""
    
    ' Add parent directory traversals if needed
    For i = commonLength To UBound(workbookComponents)
        relativePath = relativePath & ".." & Application.pathSeparator
    Next i
    
    ' Add remaining image path components
    For i = commonLength To UBound(imageComponents)
        relativePath = relativePath & imageComponents(i) & Application.pathSeparator
    Next i
    
    GetRelativePath = relativePath & fileName
End Function

' Updates the image path in settings if necessary
' @param directory The directory to check/add to the image path
Private Sub UpdateImagePathIfNeeded(ByVal directory As String)
    If IsImagePathValid(directory) Then Exit Sub
    
    Dim currentPath As String
    currentPath = SettingsSheet.Range(SETTINGS_IMAGE_PATH).Value2
    If Len(currentPath) = 0 Then
        SettingsSheet.Range(SETTINGS_IMAGE_PATH).Value2 = directory
    ElseIf Not IsDirectoryInPath(directory, currentPath) Then
        SettingsSheet.Range(SETTINGS_IMAGE_PATH).Value2 = currentPath & GetEnvVarSeparator & directory
    End If
End Sub

' Checks if the directory is valid (in env variable or current workbook path)
' @param directory The directory to check
' @returns Boolean True if the directory is valid
Private Function IsImagePathValid(ByVal directory As String) As Boolean
    IsImagePathValid = ImageFoundInEnvVariablePath(directory) Or ImageFoundInCurrentDir(directory)
End Function

' Checks if directory exists in the path concatenation
' @param directory The directory to check
' @param pathString The current image path string
' @returns Boolean True if directory is in the path
Private Function IsDirectoryInPath(ByVal directory As String, ByVal pathString As String) As Boolean
    Dim paths() As String
    paths = split(pathString, GetEnvVarSeparator)
    
    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        If UCase$(paths(i)) = UCase$(directory) Then
            IsDirectoryInPath = True
            Exit Function
        End If
    Next i
End Function

' Updates the ribbon display with the selected filename
' @param fileName The name of the selected file
Private Sub UpdateRibbonDisplay(ByVal fileName As String)
    SaveStyleDesignerSetting DESIGNER_NODE_IMAGE_NAME, fileName
    InvalidateRibbonControl RIBBON_CTL_NODE_IMAGE_NAME
End Sub

' Checks if directory matches the environment variable path
' @param directory The directory to check
' @returns Boolean True if directory matches the environment variable
Private Function ImageFoundInEnvVariablePath(ByVal directory As String) As Boolean
#If Mac Then
    ImageFoundInEnvVariablePath = False
#Else
    ImageFoundInEnvVariablePath = UCase$(directory) = UCase$(Trim$(Environ$("ExcelToGraphvizImages")))
#End If
End Function

' Checks if directory matches the current workbook path
' @param directory The directory to check
' @returns Boolean True if directory matches the workbook path
Private Function ImageFoundInCurrentDir(ByVal directory As String) As Boolean
    ImageFoundInCurrentDir = UCase$(directory) = UCase$(ActiveWorkbook.path)
End Function

' ===========================================================================
' Callbacks for nodeImage dynamic controls

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImage_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = Not CellIsEmpty(DESIGNER_NODE_IMAGE_NAME)
End Sub

' ===========================================================================
' Callbacks for nodeImagePosition

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImagePosition_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    SaveSelectedItem itemId, "imagepos_", DESIGNER_NODE_IMAGE_POSITION
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub nodeImagePosition_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("imagepos_", DESIGNER_NODE_IMAGE_POSITION)
End Sub

' ===========================================================================
' Callbacks for nodeImageScale

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_getItemCount(ByVal control As IRibbonControl, ByRef listSize As Variant)
    listSize = ListsSheet.Range(LISTS_IMAGE_SCALE).count + 1
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    If index = 0 Then
        label = vbNullString
    Else
        Dim listId As String
        listId = "is_" & ListsSheet.Range(LISTS_IMAGE_SCALE).Cells.item(index, 1).Value2
        label = GetLabel(listId)
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex(LISTS_IMAGE_SCALE, DESIGNER_NODE_IMAGE_SCALE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeImageScale_onAction(ByVal control As IRibbonControl, ByVal itemId As String, ByVal index As Long)
    If index = 0 Then
        ClearStyleDesignerSetting DESIGNER_NODE_IMAGE_SCALE
    Else
        SaveStyleDesignerSetting DESIGNER_NODE_IMAGE_SCALE, ListsSheet.Range(LISTS_IMAGE_SCALE).Cells.item(index, 1).Value2
    End If
    RenderPreview
End Sub

' ===========================================================================
' Callbacks for edgeHeadClip

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadClip_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        ClearStyleDesignerSetting DESIGNER_EDGE_HEAD_CLIP
    Else
        SaveStyleDesignerSetting DESIGNER_EDGE_HEAD_CLIP, TOGGLE_NO
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeHeadClip_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_HEAD_CLIP) Then
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
        ClearStyleDesignerSetting DESIGNER_EDGE_TAIL_CLIP
    Else
        SaveStyleDesignerSetting DESIGNER_EDGE_TAIL_CLIP, TOGGLE_NO
    End If
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeTailClip_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If CellIsEmpty(DESIGNER_EDGE_TAIL_CLIP) Then
        returnedVal = True
    Else
        returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_EDGE_TAIL_CLIP)
    End If
End Sub

' ===========================================================================
' Callbacks for edgeDecorate

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeDecorate_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SetToggleValue StyleDesignerSheet.Range(DESIGNER_EDGE_DECORATE), pressed
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
    SetToggleValue StyleDesignerSheet.Range(DESIGNER_EDGE_LABEL_FLOAT), pressed
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub edgeLabelFloat_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_EDGE_LABEL_FLOAT)
End Sub

' ===========================================================================
' Toggle helpers

Private Sub SetToggleValue(ByVal targetRange As Range, ByVal pressed As Boolean)
    If pressed Then
        targetRange.Value2 = TOGGLE_YES
    Else
        targetRange.ClearContents
    End If
End Sub

Private Function IsTrue(ByVal rangeName As String) As Boolean
    Dim val As Variant
    val = StyleDesignerSetting(rangeName)

    Select Case LCase$(Trim$(CStr(val)))
        Case TOGGLE_YES, TOGGLE_TRUE
            IsTrue = True
        Case Else
            IsTrue = False
    End Select
End Function

' ===========================================================================
' Callbacks for clearStyleRibbon

Public Sub ClearStyleRibbon()
    ClearStyleRibbonFields
    RenderPreview
    RefreshRibbon
    Application.StatusBar = False
End Sub

Public Sub ClearStyleRibbonFields()
    OptimizeCode_Begin
    ClearStyleDesignerRanges
    ClearStyleDesignerLabels
    ClearStyleDesignerStyleName
    CheckRelativePathCheckbox
    OptimizeCode_End
    RenderPreview
    RefreshRibbon
    Application.StatusBar = False
End Sub

Public Sub ClearStyleDesignerRanges()
    StyleDesignerSheet.Range("ColorScheme,FontName,FontSize,FontColor,BorderColor,BorderColor,BorderPenWidth,BorderPeripheries").ClearContents
    StyleDesignerSheet.Range("FillColor,GradientFillColor,GradientFillType,GradientFillAngle,GradientFillWeight,LabelLocation,LabelJustification,EdgeStyle,EdgeHeadPort,EdgeTailPort,EdgeColor1,EdgeColor2,EdgeColor3").ClearContents
    StyleDesignerSheet.Range("NodeShape,NodeSides,NodeOrientation,NodeRegular,NodeSkew,NodeDistortion,BorderStyle1,BorderStyle2,BorderStyle3").ClearContents
    StyleDesignerSheet.Range("NodeHeight,NodeWidth,NodeFixedSize,EdgeArrowHead1,EdgeArrowHead2,EdgeArrowHead3,EdgeDecorate,EdgeLabelFloat").ClearContents
    StyleDesignerSheet.Range("EdgeArrowTail1,EdgeArrowTail2,EdgeArrowTail3,EdgeDirection,EdgeArrowSize,EdgeWeight,EdgeLabelAngle,EdgeLabelDistance").ClearContents
    StyleDesignerSheet.Range("EdgePenWidth,NodeImageName,NodeImageScale,NodeImagePosition,EdgeHeadClip,EdgeTailClip,EdgeLabelFontName,EdgeLabelFontSize,EdgeLabelFontColor").ClearContents
    StyleDesignerSheet.Range("FontBold,FontItalic").ClearContents
    StyleDesignerSheet.Range("ClusterMargin,ClusterPackmode,ClusterArrayMajor,ClusterArrayAlign,ClusterArrayJustify,ClusterArraySplit,ClusterArraySort").ClearContents
End Sub

Public Sub ClearStyleDesignerLabels()
    StyleDesignerSheet.Range(DESIGNER_LABEL_TEXT_INCLUDE).Value2 = False
    SaveStyleDesignerSetting DESIGNER_LABEL_TEXT, replace(LCase$(GetLabel("worksheetStyleDesignerLabelText")), ":", vbNullString)
    
    StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT_INCLUDE).Value2 = False
    StyleDesignerSheet.Range(DESIGNER_XLABEL_TEXT).Value2 = vbNullString        ' Can't use ClearContents on merged cells
    
    StyleDesignerSheet.Range(DESIGNER_TAIL_LABEL_TEXT_INCLUDE).Value2 = False
    StyleDesignerSheet.Range(DESIGNER_TAIL_LABEL_TEXT).Value2 = vbNullString
    
    StyleDesignerSheet.Range(DESIGNER_HEAD_LABEL_TEXT_INCLUDE).Value2 = False
    StyleDesignerSheet.Range(DESIGNER_HEAD_LABEL_TEXT).Value2 = vbNullString
End Sub

Public Sub ClearStyleDesignerStyleName()
    StyleDesignerSheet.Range(DESIGNER_STYLE_NAME_TEXT).Value2 = vbNullString    ' Can't use ClearContents on merged cells
End Sub

Public Sub CheckRelativePathCheckbox()
    SaveStyleDesignerSetting DESIGNER_NODE_IMAGE_RELATIVE_PATH, TOGGLE_YES
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub clearStyleRibbon_onAction(ByVal control As IRibbonControl)
    ClearStyleRibbon
End Sub

' ===========================================================================
' Callbacks for saveToStylesWorksheet

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub saveToStylesWorksheet_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_FORMAT_STRING)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub saveToStylesWorksheet_onAction(ByVal control As IRibbonControl)
    SaveToStylesWorksheet
End Sub

' ===========================================================================
' Callbacks for copyToClipboard

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub copyToClipboard_onAction(ByVal control As IRibbonControl)
    StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Copy
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub copyToClipboard_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = Not CellIsEmpty(DESIGNER_FORMAT_STRING)
End Sub

' ===========================================================================
' Callbacks for alignTop

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignTop_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_LABEL_LOCATION, Toggle(pressed, ALIGN_TOP, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ALIGN_BOTTOM
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignTop_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_LABEL_LOCATION) = ALIGN_TOP
End Sub

' ===========================================================================
' Callbacks for alignBottom

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignBottom_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_LABEL_LOCATION, Toggle(pressed, ALIGN_BOTTOM, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ALIGN_TOP
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub alignBottom_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_LABEL_LOCATION) = ALIGN_BOTTOM
End Sub

' ===========================================================================
' Callbacks for justifyLeft

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyLeft_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_LABEL_JUSTIFICATION, Toggle(pressed, JUSTIFY_LEFT, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_RIGHT
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyLeft_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_LABEL_JUSTIFICATION) = JUSTIFY_LEFT
End Sub

' ===========================================================================
' Callbacks for justifyRight

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyRight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_LABEL_JUSTIFICATION, Toggle(pressed, JUSTIFY_RIGHT, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_LEFT
    RenderPreview
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub justifyRight_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_LABEL_JUSTIFICATION) = JUSTIFY_RIGHT
End Sub

' ===========================================================================
' Callbacks for fontBold

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub fontBold_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_FONT_BOLD, Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
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
    SaveStyleDesignerSetting DESIGNER_FONT_ITALIC, Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
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
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE Or _
              StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupBorders_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE Or _
              StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupFillColor_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE Or _
              StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupGradientFillColor_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = False
    
    If StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE Or _
       StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER Then
        If Not CellIsEmpty(DESIGNER_FILL_COLOR) Then
            visible = True
        End If
    End If
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupNodeShape_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupNodeDimensions_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupNodeImage_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_NODE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupEdgeStyle_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_EDGE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupEdgeColors_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_EDGE
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub designerGroupEdgeArrows_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    visible = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_EDGE
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
    Application.StatusBar = GetRenderInfo() & " | " & timex.Elapsed_sec & " seconds"
#End If
End Sub

Private Sub SaveColor(ByVal index As Long, ByVal cellName As String)
    If index = 0 Then
        ClearStyleDesignerSetting cellName
        Exit Sub
    End If

    Dim colorScheme As String
    colorScheme = GetColorScheme()

    Dim colorRangeName As String
    colorRangeName = COLOR_SCHEME_PREFIX & colorScheme

    Dim color As String
    With HelpColorsSheet.Range(colorRangeName)
        Select Case colorScheme
            Case COLOR_SCHEME_X11, COLOR_SCHEME_SVG
                color = .Cells(index, 1).Value2  ' Vertical list
            Case Else
                color = .Cells(1, index).Value2  ' Horizontal list
        End Select
    End With

    SaveStyleDesignerSetting cellName, color
End Sub

Private Function GetListIndex(ByVal listName As String, ByVal cellName As String) As Long
    GetListIndex = 0
    
    If Len(listName) = 0 Or Len(cellName) = 0 Then Exit Function

    On Error Resume Next
    Dim cellValue As String
    cellValue = LCase$(CStr(StyleDesignerSetting(cellName)))
    If Err.number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    If Len(cellValue) = 0 Then Exit Function
    
    ' Iterating arrays is faster than iterating cells
    Dim listArray As Variant
    listArray = Application.WorksheetFunction.Transpose(ListsSheet.Range(listName))
    If Err.number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    Dim index As Long
    index = 0
    Dim i As Long
    For i = LBound(listArray) To UBound(listArray)
        index = index + 1
        If LCase$(CStr(listArray(i))) = cellValue Then
            GetListIndex = index
            Exit Function
        End If
    Next i
End Function

Public Sub SetStyleDesignerNodeShape(ByVal shapeName As String)

    ' Ensure we are in "node" mode
    SaveStyleDesignerSetting DESIGNER_MODE, KEYWORD_NODE
    
    ' Unhide style designer if hidden
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).Value2 = TOGGLE_HIDE Then
        SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).Value2 = TOGGLE_SHOW
    End If
    
    OptimizeCode_Begin
    
    SaveStyleDesignerSetting DESIGNER_NODE_SHAPE, shapeName
    If shapeName <> GRAPHVIZ_SHAPE_POLYGON Then
        StyleDesignerSheet.Range("NodeSides,NodeOrientation,NodeSkew,NodeDistortion").ClearContents
    End If
    
    RefreshControlsPolygon
    OptimizeCode_End
    RenderPreview
End Sub

Public Sub SetStyleDesignerColorScheme(ByVal colorScheme As String)
    ' Unhide style designer if hidden
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).Value2 = TOGGLE_HIDE Then
        SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).Value2 = TOGGLE_SHOW
    End If
    
    OptimizeCode_Begin
    SaveStyleDesignerSetting DESIGNER_COLOR_SCHEME, colorScheme
    StyleDesignerSheet.Range("FontColor,BorderColor,FillColor,GradientFillColor,GradientFillType,GradientFillAngle,EdgeColor1,EdgeColor2,EdgeColor3,EdgeLabelFontColor").ClearContents
    RefreshControlsColorScheme
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
    Set tmpFontList = Application.CommandBars.item("Formatting").FindControl(ID:=1728)
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
    Set excludedPrefixes = Nothing
    Set excludedSuffixes = Nothing
    Set tmpFontList = Nothing
    getFontList = fontList
    Exit Function
ErrorHandler:
    MsgBox GetMessage("msgboxNoListOfFonts"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    ReDim fontList(0)
    getFontList = fontList
#End If

End Function

' Initialize font exclusion dictionaries
Private Sub InitializeFontExclusions()
    If fontExclusionsInitialized Then Exit Sub
    
    Set excludedPrefixes = New Dictionary
    excludedPrefixes.CompareMode = TextCompare ' Case-insensitive
    With excludedPrefixes
        .Add "Abadi", True
        .Add "Abel", True
        .Add "Abril", True
        .Add "ADLaM", True
        .Add "Agency FB", True
        .Add "Aharoni", True
        .Add "Alasassy", True
        .Add "Aldhabi", True
        .Add "Alef", True
        .Add "Aleo", True
        .Add "Algerian", True
        .Add "Amatic", True
        .Add "Angsana", True
        .Add "Anton", True
        .Add "Aparajita", True
        .Add "Aptos", True
        .Add "Arabic", True
        .Add "Aref", True
        .Add "Arial Narrow", True
        .Add "Assistant", True
        .Add "Athiti", True
        .Add "Baguet", True
        .Add "Bahnschrift", True
        .Add "Barlow", True
        .Add "Batang", True
        .Add "Bauhaus", True
        .Add "Bebas", True
        .Add "Bembo", True
        .Add "Berlin", True
        .Add "Bierstadt", True
        .Add "Biome", True
        .Add "Bookshelf", True
        .Add "Boucherie", True
        .Add "Browallia", True
        .Add "Brush", True
        .Add "Buxton", True
        .Add "Cambria", True
        .Add "Cascadia", True
        .Add "Caveat", True
        .Add "Cavolini", True
        .Add "Chamberi", True
        .Add "Charmonman", True
        .Add "Chiller", True
        .Add "Chonburi", True
        .Add "Concert", True
        .Add "Congenial", True
        .Add "Convection", True
        .Add "Cordia", True
        .Add "DM", True
        .Add "Dante", True
        .Add "DaunPenh", True
        .Add "David", True
        .Add "Daytona", True
        .Add "DengXian", True
        .Add "Didact", True
        .Add "Dillenia", True
        .Add "DokChampa", True
        .Add "Dosis", True
        .Add "Dotum", True
        .Add "Dubai", True
        .Add "EB Garamond", True
        .Add "Ebrima", True
        .Add "Edwardian Script", True
        .Add "Engravers", True
        .Add "Eucrosia", True
        .Add "Euphemia", True
        .Add "Fahkwang", True
        .Add "FangSong", True
        .Add "Fairwater", True
        .Add "Fira", True
        .Add "Fjalla", True
        .Add "Forte", True
        .Add "Frank", True
        .Add "Fredoka", True
        .Add "FreesiaUPC", True
        .Add "Gabriela", True
        .Add "Gabriola", True
        .Add "Gaegu", True
        .Add "Gautami", True
        .Add "Gill", True
        .Add "Gisha", True
        .Add "Goudy", True
        .Add "Grandview", True
        .Add "Grotesque", True
        .Add "Gulim", True
        .Add "Gungsuh", True
        .Add "HG", True
        .Add "Hadassah", True
        .Add "Hammersmith", True
        .Add "Harlow", True
        .Add "Heebo", True
        .Add "Hind", True
        .Add "HoloLens", True
        .Add "IBM", True
        .Add "Inconsolata", True
        .Add "Impact", True
        .Add "Informal", True
        .Add "Iris", True
        .Add "Iskoola", True
        .Add "Jasmine", True
        .Add "Josefin", True
        .Add "Jumble", True
        .Add "KaiTi", True
        .Add "Kalinga", True
        .Add "Karla", True
        .Add "Kartika", True
        .Add "Kermit", True
        .Add "Kigelia", True
        .Add "KleeOne", True
        .Add "Kodchiang", True
        .Add "Kokila", True
        .Add "Kristen", True
        .Add "Krub", True
        .Add "Lalezar", True
        .Add "Latha", True
        .Add "Lato", True
        .Add "Leelawadee", True
        .Add "Levenim", True
        .Add "Libre", True
        .Add "Ligconsolata", True
        .Add "Lily", True
        .Add "Livvic", True
        .Add "Lobster", True
        .Add "Lora", True
        .Add "Lucida", True
        .Add "Magneto", True
        .Add "Microsoft", True
        .Add "MS", True
        .Add "MT", True
        .Add "Mangal", True
        .Add "Marlett", True
        .Add "Meddon", True
        .Add "Meiryo", True
        .Add "Merriweather", True
        .Add "Ming", True
        .Add "Miriam", True
        .Add "Mitr", True
        .Add "Modern", True
        .Add "Monotype", True
        .Add "Montserrat", True
        .Add "MoolBoran", True
        .Add "Mr Gabe", True
        .Add "Mystical", True
        .Add "Nanum", True
        .Add "Narkisim", True
        .Add "News", True
        .Add "Niagara", True
        .Add "Nina", True
        .Add "Nordique", True
        .Add "Noto", True
        .Add "Nunito", True
        .Add "Nyala", True
        .Add "OCR", True
        .Add "Open Sans", True
        .Add "Oranienbaum", True
        .Add "Oswald", True
        .Add "Oxygen", True
        .Add "PT", True
        .Add "Pacifico", True
        .Add "Palace", True
        .Add "Palanquin", True
        .Add "Patrick", True
        .Add "Petit", True
        .Add "Playbill", True
        .Add "Playfair", True
        .Add "Plantagenet", True
        .Add "PMing", True
        .Add "Poiret", True
        .Add "Poppins", True
        .Add "Posterama", True
        .Add "Pridi", True
        .Add "Prompt", True
        .Add "Quattro", True
        .Add "Questrial", True
        .Add "QuickType", True
        .Add "Quire", True
        .Add "Raavi", True
        .Add "Ravie", True
        .Add "Rage", True
        .Add "Raleway", True
        .Add "Rastanty", True
        .Add "Reem", True
        .Add "Roboto", True
        .Add "Rod", True
        .Add "STCaiyun", True
        .Add "STF", True
        .Add "STH", True
        .Add "STK", True
        .Add "STX", True
        .Add "STZ", True
        .Add "Sacramento", True
        .Add "Sagona", True
        .Add "Sans Serif Collection", True
        .Add "Sakkal", True
        .Add "Seaford", True
        .Add "Secular", True
        .Add "Segoe", True
        .Add "Selawik", True
        .Add "Shadows", True
        .Add "Shonar", True
        .Add "Shruti", True
        .Add "SimHei", True
        .Add "Simplified", True
        .Add "Sitka", True
        .Add "Skeena", True
        .Add "Statliches", True
        .Add "Suez", True
        .Add "Symbol", True
        .Add "TH", True
        .Add "Tahoma", True
        .Add "Tenorite", True
        .Add "Titillum", True
        .Add "Times New Roman", True
        .Add "Trade", True
        .Add "Traditional", True
        .Add "Trirong", True
        .Add "Tunga", True
        .Add "UD Digi", True
        .Add "Ubuntu", True
        .Add "Univers", True
        .Add "Urdu", True
        .Add "Utsaah", True
        .Add "Vani", True
        .Add "Varela", True
        .Add "Vijaya", True
        .Add "Vivaldi", True
        .Add "Vrinda", True
        .Add "Walbaum", True
        .Add "Wandohope", True
        .Add "Webdings", True
        .Add "Wingdings", True
        .Add "Wide Latin", True
        .Add "Work Sans", True
        .Add "Yesteryear", True
        .Add "Yu", True
    End With
    
    Set excludedSuffixes = New Dictionary
    excludedSuffixes.CompareMode = TextCompare
    With excludedSuffixes
        .Add "Black", True
        .Add "Bold ITC", True
        .Add "Bold", True
        .Add "Compressed", True
        .Add "Cond", True
        .Add "Conde", True
        .Add "Conden", True
        .Add "Condensed", True
        .Add "Demi ITC", True
        .Add "Demi", True
        .Add "Expanded", True
        .Add "ExtB", True
        .Add "Extended", True
        .Add "Hand", True
        .Add "Heavy", True
        .Add "Light ITC", True
        .Add "Light", True
        .Add "Lt", True
        .Add "Medium ITC", True
        .Add "Medium", True
        .Add "Nova", True
        .Add "Pro", True
        .Add "Schoolbook", True
        .Add "Text", True
        .Add "Thin", True
        .Add "UI", True
        .Add "XBd", True
        .Add ".tmp", True
    End With
    
    fontExclusionsInitialized = True
End Sub

' Optimized addToFontList
Private Function addToFontList(ByVal fontName As String) As Boolean
    If Len(fontName) = 0 Then
        addToFontList = False
        Exit Function
    End If
    
    InitializeFontExclusions
    
    Dim lowerFontName As String
    lowerFontName = LCase$(fontName)
    
    ' Check prefixes
    Dim prefix As Variant
    For Each prefix In excludedPrefixes.Keys
        If Len(lowerFontName) >= Len(prefix) Then
            If Left$(lowerFontName, Len(prefix)) = LCase$(prefix) Then
                addToFontList = False
                Exit Function
            End If
        End If
    Next prefix
    
    ' Check suffixes
    Dim suffix As Variant
    For Each suffix In excludedSuffixes.Keys
        If Len(lowerFontName) >= Len(suffix) Then
            If Right$(lowerFontName, Len(suffix)) = LCase$(suffix) Then
                addToFontList = False
                Exit Function
            End If
        End If
    Next suffix
    
    addToFontList = True
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
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLStyleDesignerTab").Value2, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for Pack / Packmode

Public Sub designerGroupPack_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Dim isClusterMode As Boolean
    Dim isOsageEngine As Boolean

    With StyleDesignerSheet
        isClusterMode = (.Range(DESIGNER_MODE).Value2 = KEYWORD_CLUSTER)
    End With

    With SettingsSheet
        isOsageEngine = (.Range(SETTINGS_GRAPHVIZ_ENGINE).Value2 = LAYOUT_OSAGE)
    End With

    If Not (isClusterMode And isOsageEngine) Then
        visible = False
        Exit Sub
    End If

    Select Case control.ID
        Case RIBBON_CTL_CLUSTER_MARGIN
            visible = Not GetCellBoolean(StyleDesignerSheet.name, DESIGNER_NODE_METRIC)

        Case RIBBON_CTL_CLUSTER_MARGIN_MM
            visible = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_NODE_METRIC)

        Case RIBBON_CTL_CLUSTER_PACKMODE, RIBBON_GRP_PACK
            visible = True

        Case Else
            visible = False
    End Select
End Sub

Public Sub clusterMargin_onAction(ByVal control As IRibbonControl, ID As String, ByVal index As Integer)
    Dim prefix As String
    Dim marginValue As String

    Select Case control.ID
        Case RIBBON_CTL_CLUSTER_MARGIN
            prefix = "margin_"
        Case RIBBON_CTL_CLUSTER_MARGIN_MM
            prefix = "mmmargin_"
        Case Else
            Exit Sub ' Unexpected control, do nothing
    End Select

    If Len(ID) > Len(prefix) Then
        marginValue = Mid$(ID, Len(prefix) + 1)
        SaveStyleDesignerSetting DESIGNER_CLUSTER_MARGIN, marginValue
        RenderPreview
    End If
End Sub

Public Sub clusterMargin_getSelectedItemIndex(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetListIndex("Margin", DESIGNER_CLUSTER_MARGIN)
End Sub

'@Ignore ParameterNotUsed
Public Sub clusterMargin_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    Dim prefix As String

    Select Case control.ID
        Case RIBBON_CTL_CLUSTER_MARGIN
            prefix = "margin_"
        Case RIBBON_CTL_CLUSTER_MARGIN_MM
            prefix = "mmmargin_"
        Case Else
            itemId = vbNullString
            Exit Sub
    End Select

    itemId = prefix & StyleDesignerSetting(DESIGNER_CLUSTER_MARGIN)
End Sub

Public Sub clusterPackmode_onAction(ByVal control As IRibbonControl, ID As String, ByVal index As Integer)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_PACKMODE, Mid$(ID, Len("packmode_") + 1)
    RefreshControlsPackmode
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub clusterPackmode_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("packmode_", DESIGNER_CLUSTER_PACKMODE)
End Sub

Public Sub arraySplit_onAction(ByVal control As IRibbonControl, ID As String, ByVal index As Integer)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_SPLIT, Mid$(ID, Len("arraySplit_") + 1)
    RenderPreview
End Sub

'@Ignore ParameterNotUsed
Public Sub arraySplit_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = GetSelectedItemID("arraySplit_", DESIGNER_CLUSTER_ARRAY_SPLIT)
End Sub

Public Sub arrayAlignTop_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_ALIGN, Toggle(pressed, GRAPHVIZ_PACKMODE_ALIGN_TOP, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ARRAY_ALIGN_BOTTOM
    RenderPreview
End Sub

Public Sub arrayAlignTop_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_CLUSTER_ARRAY_ALIGN) = GRAPHVIZ_PACKMODE_ALIGN_TOP
End Sub

Public Sub array_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_MODE) = KEYWORD_CLUSTER And _
        StyleDesignerSetting(DESIGNER_CLUSTER_PACKMODE) = GRAPHVIZ_PACKMODE_ARRAY
End Sub

Public Sub arrayAlignBottom_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_ALIGN, Toggle(pressed, GRAPHVIZ_PACKMODE_ALIGN_BOTTOM, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ARRAY_ALIGN_TOP
    RenderPreview
End Sub

Public Sub arrayAlignBottom_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_CLUSTER_ARRAY_ALIGN) = GRAPHVIZ_PACKMODE_ALIGN_BOTTOM
End Sub

Public Sub arrayJustifyLeft_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_JUSTIFY, Toggle(pressed, GRAPHVIZ_PACKMODE_JUSTIFY_LEFT, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ARRAY_JUSTIFY_RIGHT
    RenderPreview
End Sub

Public Sub arrayJustifyLeft_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_CLUSTER_ARRAY_JUSTIFY) = GRAPHVIZ_PACKMODE_JUSTIFY_LEFT
End Sub

Public Sub arrayJustifyRight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_JUSTIFY, Toggle(pressed, GRAPHVIZ_PACKMODE_JUSTIFY_RIGHT, vbNullString)
    InvalidateRibbonControl RIBBON_CTL_ARRAY_JUSTIFY_LEFT
    RenderPreview
End Sub

Public Sub arrayJustifyRight_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = StyleDesignerSetting(DESIGNER_CLUSTER_ARRAY_JUSTIFY) = GRAPHVIZ_PACKMODE_JUSTIFY_RIGHT
End Sub

Public Sub arraySort_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_SORT, Toggle(pressed, TOGGLE_YES, TOGGLE_NO)
    RenderPreview
End Sub

Public Sub arraySort_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = IsTrue(DESIGNER_CLUSTER_ARRAY_SORT)
End Sub

Public Sub arrayMajor_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        SaveStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_MAJOR, GRAPHVIZ_PACKMODE_MAJOR_COLUMN
    Else
        ClearStyleDesignerSetting DESIGNER_CLUSTER_ARRAY_MAJOR
    End If
    RenderPreview
End Sub

Public Sub arrayMajor_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Select Case LCase$(StyleDesignerSetting(DESIGNER_CLUSTER_ARRAY_MAJOR))
        Case GRAPHVIZ_PACKMODE_MAJOR_COLUMN
            returnedVal = True
        Case Else
            returnedVal = False
    End Select
End Sub

Public Sub colorScheme_getLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Dim colorScheme As String
    colorScheme = StyleDesignerSetting(DESIGNER_COLOR_SCHEME)
    If colorScheme = vbNullString Then
        label = GetLabel(control.ID)
    Else
        label = colorScheme
    End If
End Sub

' Caches image and handles X11 Gray/Grey special case
Private Sub ColorCacheImage(color As ColorInfo, image As Variant)
    Dim cacheKey As String
    cacheKey = ColorGetCacheKey(color)
    
    On Error Resume Next
    colorImageCache.Add cacheKey, image
    
    ' Handle X11 Gray/Grey special case
    If StartsWith(cacheKey, "X11_Gray") Then
        Dim greyKey As String
        greyKey = "X11_Grey" & Right$(cacheKey, Len(cacheKey) - 8)
        colorImageCache.Add greyKey, image
    End If
    On Error GoTo 0
End Sub

Private Function ColorGetCacheKey(color As ColorInfo)
    ColorGetCacheKey = Trim$(color.scheme) & "_" & Trim$(color.name)
End Function

Private Sub ColorSetImageFile(ByRef color As ColorInfo)
    Dim colorCacheKey As String
    colorCacheKey = ColorGetCacheKey(color)
    color.imageFile = GetColorImageDir() & Application.pathSeparator & LCase$(colorCacheKey) & DOT & RIBBON_EXT_COLOR
    color.imageFile = replace(color.imageFile, "#", vbNullString)
End Sub

Private Function ColorGetRGBByIndex(color As ColorInfo, index As Long) As Long
    ' Get the RGB color for this color scheme index
    If color.scheme = COLOR_SCHEME_X11 Or color.scheme = COLOR_SCHEME_SVG Then
        ' Color list is arranged in a column of cells
        ColorGetRGBByIndex = HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & color.scheme).Cells.item(index, 1).Interior.color
    Else
        ' Color list is aranged in a row of cells
        ColorGetRGBByIndex = HelpColorsSheet.Range(COLOR_SCHEME_PREFIX & color.scheme).Cells.item(1, index).Interior.color
    End If
End Function

Private Function ColorGetNameByIndex(color As ColorInfo, index As Long) As String
    ' Get the color name based on the current color scheme
    If index = 0 Then Exit Function

    Select Case color.scheme
        Case COLOR_SCHEME_X11
            If index <= UBound(x11Colors) Then
                ColorGetNameByIndex = x11Colors(index)
            End If
        Case COLOR_SCHEME_SVG
            If index <= UBound(svgColors) Then
                ColorGetNameByIndex = svgColors(index)
            End If
        Case Else
            If index <= UBound(brewerColors) Then
                ColorGetNameByIndex = brewerColors(index)
            End If
    End Select
End Function

Private Function ColorGetIndexByName(color As ColorInfo) As Long
    If color.name = vbNullString Then
        ColorGetIndexByName = 0
        Exit Function
    End If

    Dim colorArray As Variant
    Select Case color.scheme
        Case COLOR_SCHEME_X11
            If Not IsArray(x11Colors) Or IsEmpty(x11Colors) Then LoadColorNameArray
            colorArray = x11Colors
        Case COLOR_SCHEME_SVG
            If Not IsArray(svgColors) Or IsEmpty(svgColors) Then LoadColorNameArray
            colorArray = svgColors
        Case Else
            If Not IsArray(brewerColors) Or IsEmpty(brewerColors) Then LoadColorNameArray
            colorArray = brewerColors
    End Select

    If Not IsArray(colorArray) Or IsEmpty(colorArray) Then
        ColorGetIndexByName = 0
        Exit Function
    End If

    Dim i As Long
    For i = LBound(colorArray) To UBound(colorArray)
        If StrComp(color.name, CStr(colorArray(i)), vbTextCompare) = 0 Then
            ColorGetIndexByName = i + 1
            Exit Function
        End If
    Next i

    ColorGetIndexByName = 0
End Function

Private Function ColorGetIndex(ByVal cellName As String) As Long
    ColorGetIndex = 0
    
    ' Validate input
    If Len(cellName) = 0 Then Exit Function
    
    Dim color As String
    color = LCase$(StyleDesignerSetting(cellName))
    If Len(color) = 0 Then Exit Function
    
    If Left$(color, 1) = "#" Then Exit Function
     
    Dim index As Long
    index = 0
    Dim arrayItem As Variant

    ' This looks like inefficient, repetitive code, but is intentionally
    ' coded in-line. Consolidation requires copying an array which noticably
    ' degrades performance for X11 and SVG schemes.
    Select Case GetColorScheme()
        Case COLOR_SCHEME_X11
            For Each arrayItem In x11Colors
                index = index + 1
                If StrComp(arrayItem, color, vbTextCompare) = 0 Then
                    Exit For
                End If
            Next arrayItem
        Case COLOR_SCHEME_SVG
            For Each arrayItem In svgColors
                index = index + 1
                If StrComp(arrayItem, color, vbTextCompare) = 0 Then
                    Exit For
                End If
            Next arrayItem
        Case Else
            For Each arrayItem In brewerColors
                index = index + 1
                If StrComp(arrayItem, color, vbTextCompare) = 0 Then
                    Exit For
                End If
            Next arrayItem
    End Select
 
    ColorGetIndex = index
End Function

Private Function ColorCreateThumbnail(color As ColorInfo, Optional ByVal sizePoints As Single = 15) As Boolean
    ColorCreateThumbnail = False
    
    If color.RGB < 0 Or color.RGB > &HFFFFFF Then Exit Function
    
    If Len(color.imageFile) = 0 Then
        ColorSetImageFile color
    End If
    
    On Error GoTo ErrorHandler
    
    Dim borderRGB As Long
    Select Case LCase$(color.name)
        Case "transparent", "invis", "none" ' Put red border around "invisible" colors
            borderRGB = RGB(255, 0, 0)
        Case Else
            borderRGB = RGB(200, 200, 200)  ' Put light gray border around all the rest
    End Select

    ' Chart attributes are in points. e.g. 15 points = 20 pixels
    Dim chartObj As ChartObject
    Set chartObj = StyleDesignerSheet.ChartObjects.Add(0, 0, sizePoints, sizePoints)
    
    ' Set the background fill color of the chart to the fill color
    ' passed to this function, then write the chart out as
    ' an image file.
    With chartObj.Chart
        .ChartArea.format.Fill.visible = msoTrue
        .ChartArea.format.Fill.ForeColor.RGB = color.RGB
        
        ' Light gray border for visibility
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.color = borderRGB
        .ChartArea.Border.Weight = xlThin
    
        ' Optional: Uncomment to remove plot area border if not needed
        '.PlotArea.Border.LineStyle = xlNone
    
        .HasLegend = False
        
        DoEvents    ' Give chart time to render
        
        .Export fileName:=color.imageFile
        
        If Len(Dir(color.imageFile)) = 0 Then
            Kill color.imageFile
            ColorCreateThumbnail = False
            MsgBox "DEBUG: ColorCreateThumbnail() - 0 byte file detected and deleted - " & color.imageFile
        Else
            ColorCreateThumbnail = True
        End If
    End With
    
    chartObj.Delete
    
Cleanup:
    Set chartObj = Nothing
    Exit Function

ErrorHandler:
    If Not chartObj Is Nothing Then chartObj.Delete
    Resume Cleanup
End Function

Private Sub ColorGetOrCreateImage(ByRef color As ColorInfo, ByRef image As Variant)
    Set image = Nothing
    
    If Len(color.name) = 0 Then Exit Sub
    If Len(color.scheme) = 0 Then Exit Sub
    
    ' Handle colors passed as hex values (e.g. #FF00CC)
    If Left(color.name, 1) = "#" Then color.scheme = "rgb"
    
    ' Build the cache key
    Dim colorCacheKey As String
    colorCacheKey = ColorGetCacheKey(color)
       
    ' Initialize cache if needed
    If colorImageCache Is Nothing Then
        Set colorImageCache = New Dictionary
    End If
    
    ' Try to return the color image from cache
    On Error Resume Next
    If colorImageCache.Exists(colorCacheKey) Then
        Set image = colorImageCache.item(colorCacheKey)
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Build the path to where the images are kept
    ColorSetImageFile color

    ' If the image already exists we should be able to load it
    On Error Resume Next
    Set image = LoadPicture(color.imageFile)
    On Error GoTo 0
    
    ' If LoadPicture did not fail silently, cache the image
    If Not image Is Nothing Then
        ColorCacheImage color, image
        Exit Sub
    End If
    
    ' ================================================================
    ' If we got this far, the image was not in cache, and file was
    ' not found. Generate a thumbnail.
     
    Application.StatusBar = replace(GetMessage("statusbarCreateImage"), "{colorScheme}", color.scheme) & " " & color.name
    
    If Left(color.name, 1) = "#" Then
        color.RGB = ColorHexToRGBLong(color.name)
        color.scheme = "rgb"
    Else
        ' Get the RGB color for this color scheme index
        Dim index As Long
        index = ColorGetIndexByName(color)
        
        If index = 0 Then   ' Color not found, default to black as this should not occur
            color.RGB = 0
        Else                ' Difference between ranges and gallery
            index = index - 1
            color.RGB = ColorGetRGBByIndex(color, index)
        End If
    End If
    
    ' Generate a thumbnail image and load it into memory
    If ColorCreateThumbnail(color) Then
        On Error Resume Next
        Set image = LoadPicture(color.imageFile)
        On Error GoTo 0
    End If

    ' If LoadPicture did not fail silently, cache the image
    If Not image Is Nothing Then
        ColorCacheImage color, image
    End If
    
    Application.StatusBar = False
End Sub

Private Sub ColorGetDefaultColorByControlId(color As ColorInfo, controlId As String)
    Select Case controlId
        Case RIBBON_CTL_FILL_COLOR
            ColorSetDefaultsWhite color

        Case RIBBON_CTL_GRADIENT_FILL_COLOR
            ' Default to white if fill color is not set; otherwise use the fill color
            Dim fillColor As String
            fillColor = StyleDesignerSetting(DESIGNER_FILL_COLOR)

            If Len(fillColor) = 0 Then
                ColorSetDefaultsWhite color
            Else
                color.name = LCase$(fillColor)
            End If
            
        Case RIBBON_CTL_EDGE_COLOR2
            color.name = GetRGBColorInCell(RIBBON_CTL_EDGE_COLOR1)

        Case RIBBON_CTL_EDGE_COLOR3
            color.name = GetRGBColorInCell(RIBBON_CTL_EDGE_COLOR2)

        Case Else
            ColorSetDefaultsBlack color
    End Select
End Sub

Private Sub ColorSetDefaultsWhite(ByRef color As ColorInfo)
    color.RGB = 16777215
    color.name = COLOR_WHITE
    color.scheme = COLOR_SCHEME_DEFAULT
End Sub

Private Sub ColorSetDefaultsBlack(ByRef color As ColorInfo)
    color.RGB = 0
    color.name = COLOR_BLACK
    color.scheme = COLOR_SCHEME_DEFAULT
End Sub

Private Function GetColorScheme() As String
    Dim colorScheme As String
    colorScheme = StyleDesignerSetting(DESIGNER_COLOR_SCHEME)
    If Len(colorScheme) = 0 Then colorScheme = COLOR_SCHEME_DEFAULT
    GetColorScheme = colorScheme
End Function

Public Function ColorHexToRGBLong(hexColor As String) As Long
    Dim r As Long, g As Long, b As Long

    If Len(hexColor) = 7 And Left(hexColor, 1) = "#" Then
        On Error Resume Next
        r = CLng("&H" & Mid(hexColor, 2, 2))
        g = CLng("&H" & Mid(hexColor, 4, 2))
        b = CLng("&H" & Mid(hexColor, 6, 2))
        On Error GoTo 0
        ColorHexToRGBLong = RGB(r, g, b)
    Else
        ColorHexToRGBLong = RGB(255, 255, 255) ' fallback to white
    End If
End Function

Sub DemoHexToRGB()
    Dim hexColor As String
    Dim rgbValue As Long

    hexColor = "#B7DEE8"
    rgbValue = ColorHexToRGBLong(hexColor)

    ' Apply to cell background
    Debug.Print "Hex Color " & hexColor & " = RGB " & rgbValue
End Sub

Private Sub RefreshControlsColor()
    InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
    
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
    
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_GRP_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_ANGLE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_WEIGHT
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_DUMMY_BUTTON1
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_DUMMY_BUTTON2
    
    InvalidateRibbonControl RIBBON_CTL_COLOR_SCHEME
End Sub

Private Sub RefreshControlsColorPicker(ByVal control As IRibbonControl)
    Select Case control.ID
        Case RIBBON_CTL_FONT_COLOR_PICKER
            InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
            
        Case RIBBON_CTL_BORDER_COLOR_PICKER
            InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR

        Case RIBBON_CTL_FILL_COLOR_PICKER
            InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
            InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
            InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
            
        Case RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
            InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
            InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
        
        Case RIBBON_CTL_EDGE_COLOR1_PICKER
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
            
        Case RIBBON_CTL_EDGE_COLOR2_PICKER
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
            
        Case RIBBON_CTL_EDGE_COLOR3_PICKER
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
            InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
            
        Case RIBBON_CTL_EDGE_LABEL_FONT_COLOR_PICKER
            InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    End Select
End Sub

Private Sub RefreshControlsColorScheme()
    InvalidateRibbonControl RIBBON_CTL_COLOR_SCHEME
    InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
    
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_ANGLE
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
End Sub

Private Sub RefreshControlsDesignMode()
    InvalidateRibbonControl RIBBON_CTL_DESIGN_MODE_NODE
    InvalidateRibbonControl RIBBON_CTL_DESIGN_MODE_EDGE
    InvalidateRibbonControl RIBBON_CTL_DESIGN_MODE_CLUSTER
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_ANGLE
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_DECORATE
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_DISTANCE
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FLOAT
    InvalidateRibbonControl RIBBON_CTL_LABEL_STYLE_SEPARATOR
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR1_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR2_PICKER
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3
    InvalidateRibbonControl RIBBON_CTL_EDGE_COLOR3_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_LABEL_FONT_NAME
    
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_BORDER_COLOR
    
    InvalidateRibbonControl RIBBON_CTL_ARRAY_ALIGN_TOP
    InvalidateRibbonControl RIBBON_CTL_ARRAY_ALIGN_BOTTOM
    InvalidateRibbonControl RIBBON_CTL_ARRAY_JUSTIFY_LEFT
    InvalidateRibbonControl RIBBON_CTL_ARRAY_JUSTIFY_RIGHT
    InvalidateRibbonControl RIBBON_CTL_ARRAY_MAJOR
    InvalidateRibbonControl RIBBON_CTL_ARRAY_SPLIT
    InvalidateRibbonControl RIBBON_CTL_ARRAY_SORT
    InvalidateRibbonControl RIBBON_CTL_PACK_SEPARATOR
    InvalidateRibbonControl RIBBON_CTL_ARRAY_SEPARATOR
    InvalidateRibbonControl RIBBON_GRP_PACK
    InvalidateRibbonControl RIBBON_CTL_CLUSTER_PACK
    InvalidateRibbonControl RIBBON_CTL_CLUSTER_MARGIN
    InvalidateRibbonControl RIBBON_CTL_CLUSTER_PACKMODE
    
    InvalidateRibbonControl RIBBON_CTL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_COLOR
    InvalidateRibbonControl RIBBON_CTL_FONT_NAME
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_NAME
End Sub

Private Sub RefreshControlsPackmode()
    InvalidateRibbonControl RIBBON_CTL_ARRAY_ALIGN_TOP
    InvalidateRibbonControl RIBBON_CTL_ARRAY_ALIGN_BOTTOM
    InvalidateRibbonControl RIBBON_CTL_ARRAY_JUSTIFY_LEFT
    InvalidateRibbonControl RIBBON_CTL_ARRAY_JUSTIFY_RIGHT
    InvalidateRibbonControl RIBBON_CTL_ARRAY_MAJOR
    InvalidateRibbonControl RIBBON_CTL_ARRAY_SPLIT
    InvalidateRibbonControl RIBBON_CTL_ARRAY_SORT
    InvalidateRibbonControl RIBBON_CTL_PACK_SEPARATOR
    InvalidateRibbonControl RIBBON_CTL_ARRAY_SEPARATOR
End Sub

Private Sub RefreshControlsPolygon()
    InvalidateRibbonControl RIBBON_CTL_NODE_SHAPE
    InvalidateRibbonControl RIBBON_CTL_NODE_SIDES
    InvalidateRibbonControl RIBBON_CTL_NODE_REGULAR
    InvalidateRibbonControl RIBBON_CTL_POLYGON_SEPARATOR
    InvalidateRibbonControl RIBBON_CTL_NODE_ROTATION
    InvalidateRibbonControl RIBBON_CTL_NODE_SKEW
    InvalidateRibbonControl RIBBON_CTL_NODE_DISTORTION
End Sub

Private Sub RefreshControlsFillColor()
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_WEIGHT
    InvalidateRibbonControl RIBBON_GRP_GRADIENT_FILL_COLOR
End Sub

Private Sub RefreshControlsGradientFill()
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
    
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_WEIGHT
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_TYPE
    InvalidateRibbonControl RIBBON_CTL_GRADIENT_FILL_ANGLE
End Sub

Public Sub RefreshStyleDesignerRibbon()
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
    InvalidateRibbonControl RIBBON_GRP_PACK
    InvalidateRibbonControl RIBBON_CTL_ALIGN_BOTTOM
    InvalidateRibbonControl RIBBON_CTL_ALIGN_TOP
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_LEFT
    InvalidateRibbonControl RIBBON_CTL_JUSTIFY_RIGHT
    InvalidateRibbonControl RIBBON_CTL_FONT_NAME
    InvalidateRibbonControl RIBBON_CTL_EDGE_LABEL_FONT_NAME
End Sub

Private Sub colorPicker_onAction(ByVal control As IRibbonControl)
    Dim picked As Boolean
    picked = False
    
    ' Bring up the RGB color chooser dialog. Fill controls have
    ' white default color, all others default to black.
    Select Case control.ID
        Case RIBBON_CTL_FONT_COLOR_PICKER
            picked = PickColor(DESIGNER_FONT_COLOR, COLOR_BLACK_RGB)
            
        Case RIBBON_CTL_BORDER_COLOR_PICKER
            picked = PickColor(DESIGNER_BORDER_COLOR, COLOR_BLACK_RGB)

        Case RIBBON_CTL_FILL_COLOR_PICKER
            picked = PickColor(DESIGNER_FILL_COLOR, COLOR_WHITE_RGB)
            
        Case RIBBON_CTL_GRADIENT_FILL_COLOR_PICKER
            picked = PickColor(DESIGNER_GRADIENT_FILL_COLOR, COLOR_WHITE_RGB)
        
        Case RIBBON_CTL_EDGE_COLOR1_PICKER
            picked = PickColor(DESIGNER_EDGE_COLOR_1, COLOR_BLACK_RGB)
            
        Case RIBBON_CTL_EDGE_COLOR2_PICKER
            picked = PickColor(DESIGNER_EDGE_COLOR_2, COLOR_BLACK_RGB)
            
        Case RIBBON_CTL_EDGE_COLOR3_PICKER
            picked = PickColor(DESIGNER_EDGE_COLOR_3, COLOR_BLACK_RGB)
            
        Case RIBBON_CTL_EDGE_LABEL_FONT_COLOR_PICKER
            picked = PickColor(DESIGNER_EDGE_LABEL_FONT_COLOR, COLOR_BLACK_RGB)
    End Select
    
    ' If a rgb color was chosen, generate a new preview and refresh the controls
    If picked Then
        RenderPreview
        RefreshControlsColorPicker control
    End If
End Sub

Private Function PickColor(cellName As String, defaultColorRGB) As Boolean
    PickColor = False
    
    ' Get the current color in Hex format
    Dim currentColorHex As String
    currentColorHex = GetRGBColorInCell(cellName)
    If Len(currentColorHex) = 0 Then
        currentColorHex = defaultColorRGB
    End If

    ' Show the color picker dialog
    Dim selectedColor As Long
    Dim rgbColorPickerToUse As String
    rgbColorPickerToUse = StyleDesignerSetting("DefaultColorPicker")
    
    ' The color picker to display is configurable on Windows.
    ' Optional parameter "pickerType" is ignored on Mac
    If rgbColorPickerToUse = "cpWindowsAPI" Then
        selectedColor = ShowColorChooser(currentColorHex, cpWindowsAPI)
    Else
        selectedColor = ShowColorChooser(currentColorHex)
    End If
    
    If selectedColor = -1 Then Exit Function ' User chose to cancel
    
    Dim chosenColorHex As String
    chosenColorHex = RGBToHex(selectedColor)
    If chosenColorHex <> currentColorHex Then
        SaveStyleDesignerSetting cellName, chosenColorHex
        PickColor = True
    End If
End Function

Private Function GetRGBColorInCell(cellName As String) As String
    GetRGBColorInCell = vbNullString
    
    ' Get and inspect the value in the cell
    Dim cellValue As String
    cellValue = StyleDesignerSetting(cellName)
    
    ' Exit if the cell is empty
    If Len(cellValue) = 0 Then Exit Function
        
    ' Is color in hex, or a value associated with a color scheme?
    If Left(cellValue, 1) = "#" Then
        ' Color is in hex
        GetRGBColorInCell = cellValue
        Exit Function
    End If
    
    ' Color scheme value. Put the color name in context
    Dim color As ColorInfo
    color.scheme = GetColorScheme()
    color.name = cellValue
    
    ' Match the name with the index in the color range
    Dim index As Long
    index = ColorGetIndexByName(color) - 1
    
    ' Get the RGB value associated with the index
    color.RGB = ColorGetRGBByIndex(color, index)
    
    ' Convert the RGB value to the hex string the color chooser accepts
    GetRGBColorInCell = RGBToHex(color.RGB)
End Function

' ===================================================================================
' Helper functions to make the code simpler to read.
' Performance impact should be minimal compared to writing the statements in-line.

' Returns the value in the named cell
Private Function StyleDesignerSetting(ByVal cellName As String) As String
    StyleDesignerSetting = Trim$(StyleDesignerSheet.Range(cellName).Value2)
End Function

' Saves the value in the named cell
Private Sub SaveStyleDesignerSetting(ByVal cellName As String, ByVal cellValue As String)
    StyleDesignerSheet.Range(cellName).Value2 = cellValue
End Sub

' Clears the contents of the named cell
Private Sub ClearStyleDesignerSetting(ByVal cellName As String)
    StyleDesignerSheet.Range(cellName).ClearContents
End Sub

' Creates a list ID by appending a cell value with a prefix
Private Function GetSelectedItemID(ByVal controlPrefix As String, ByVal cellName As String) As String
    GetSelectedItemID = controlPrefix & CStr(StyleDesignerSheet.Range(cellName).Value2)
End Function

' Strips the prefix and saves the value into the named cell
Private Sub SaveSelectedItem(ByVal itemId As String, ByVal itemIdPrefix As String, ByVal cellName As String)
    SaveStyleDesignerSetting cellName, Mid$(itemId, Len(itemIdPrefix) + 1)
End Sub


