Attribute VB_Name = "modLocalize"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Localize")
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed

Option Explicit

Private localeIds As Dictionary
Private localeWorksheet As String
Private masterIds As Dictionary
Private verbose As Boolean

Public Sub SetVerbose(ByVal useVerboseLabels As Boolean)
    verbose = useVerboseLabels
End Sub

Public Function GetVerbose() As Boolean
    GetVerbose = verbose
End Function

Public Sub InitializeLocalization()
    ' Default localization. Used if specified localization is missing keys
    Set masterIds = LocalizeCacheKeys(WORKSHEET_LOCALE_MASTER)
    
    ' Get the language worksheet name specified in the settings
    Dim localeWorksheet As String
    localeWorksheet = SettingsSheet.Range(SETTINGS_LANGUAGE).Value
    
    ' Handle the case where language worksheet name has not been specified, or specifies
    ' a worksheet which is not present in the spreadsheet (possibly through imported values).
    If (localeWorksheet = vbNullString) Or (Not WorksheetExists(localeWorksheet)) Then
        localeWorksheet = WORKSHEET_LOCALE_MASTER
    End If
    
    ' Cache the keys of the specified language
    Set localeIds = LocalizeCacheKeys(localeWorksheet)
    
    ' Determine if compact or verbose labels are to be used
    InitVerbose
End Sub

Public Sub Localize()
    ' Pause screen updates
    OptimizeCode_Begin
    
    ' Load the index to the language translations
    Set localeIds = LocalizeCacheKeys(SettingsSheet.Range(SETTINGS_LANGUAGE).Value)
    
    ' Localize column headings and/or titles on each worksheet
    LocalizeWorksheetAbout
    LocalizeWorksheetData
    LocalizeWorksheetGraph
    LocalizeWorksheetHelpAttributes
    LocalizeWorksheetHelpColors
    LocalizeWorksheetHelpShapes
    LocalizeWorksheetStyles     ' Styles must precede 'settings' due to INDIRECT formula in 'settings' worksheet
    LocalizeWorksheetSettings
    LocalizeWorksheetSource
    LocalizeWorksheetSql
    LocalizeWorksheetSvg
    LocalizeWorksheetStyleDesigner
    
    ' Resume screen updates
    OptimizeCode_End
End Sub

Private Sub LocalizeWorksheetAbout()
    AboutSheet.name = GetLabel("worksheetAboutName")
    
    AboutSheet.Range("AboutLicenseName").Value = GetSupertip("AboutLicenseName")
    AboutSheet.Range("AboutLicenseCopyright").Value = GetSupertip("AboutLicenseCopyright")
    
    AboutSheet.Range("AboutLicenseParagraph01").Value = GetSupertip("AboutLicenseParagraph01") & vbNewLine
    AboutSheet.Range("AboutLicenseParagraph01").rows.AutoFit
    
    AboutSheet.Range("AboutLicenseParagraph02").Value = GetSupertip("AboutLicenseParagraph02") & vbNewLine
    AboutSheet.Range("AboutLicenseParagraph02").rows.AutoFit
    
    AboutSheet.Range("AboutLicenseParagraph03").Value = GetSupertip("AboutLicenseParagraph03") & vbNewLine
    AboutSheet.Range("AboutLicenseParagraph03").rows.AutoFit
    
    AboutSheet.Range("AboutSpecialThanks01").Value = GetSupertip("AboutSpecialThanks")
    AboutSheet.Range("AboutSpecialThanks02").Value = GetSupertip("AboutSpecialThanks")
    AboutSheet.Range("AboutForProvidingPolski").Value = GetSupertip("AboutProvidingPolskiTranslations")
End Sub

Public Sub LocalizeWorksheetData()
    DataSheet.name = GetLabel("worksheetDataName")
    
    Dim dataWs As dataWorksheet
    dataWs = GetSettingsForDataWorksheet(SettingsSheet.name)
   
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.itemColumn).Value = GetLabel(RIBBON_CTL_SHOW_ITEM)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.tailLabelColumn).Value = GetLabel(RIBBON_CTL_SHOW_TAIL_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.labelColumn).Value = GetLabel(RIBBON_CTL_SHOW_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.xLabelColumn).Value = GetLabel(RIBBON_CTL_SHOW_OUTSIDE_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.headLabelColumn).Value = GetLabel(RIBBON_CTL_SHOW_HEAD_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.tooltipColumn).Value = GetLabel(RIBBON_CTL_SHOW_TOOLTIP)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.isRelatedToItemColumn).Value = GetLabel(RIBBON_CTL_SHOW_IS_RELATED_TO_ITEM)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.styleNameColumn).Value = GetLabel(RIBBON_CTL_SHOW_STYLE)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.extraAttributesColumn).Value = GetLabel(RIBBON_CTL_SHOW_EXTRA_STYLE_ATTRIBUTES)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.errorMessageColumn).Value = GetLabel(RIBBON_CTL_SHOW_MESSAGES)
End Sub

Public Sub LocalizeWorksheetSource()
    SourceSheet.name = GetLabel("worksheetSourceName")
    
    Dim sourceWs As sourceWorksheet
    sourceWs = GetSettingsForSourceWorksheet
    
    SourceSheet.Cells.Item(sourceWs.headingRow, sourceWs.lineNumberColumn).Value = GetLabel("worksheetSourceLine")
    SourceSheet.Cells.Item(sourceWs.headingRow, sourceWs.sourceColumn).Value = GetLabel("worksheetSourceGraphvizSource")
End Sub

Private Sub LocalizeWorksheetGraph()
    GraphSheet.name = GetLabel("worksheetGraphName")
End Sub

Private Sub LocalizeWorksheetHelpAttributes()
    HelpAttributesSheet.name = GetLabel("worksheetHelpAttributesName")
End Sub

Private Sub LocalizeWorksheetHelpColors()
    HelpColorsSheet.name = GetLabel("worksheetHelpColorsName")
    HelpColorsSheet.Range("TitleX11ColorScheme").Value = GetLabel("worksheetHelpColorsX11ColorScheme")
    HelpColorsSheet.Range("TitleX11ColorName").Value = GetLabel("worksheetHelpColorsX11ColorName")
    HelpColorsSheet.Range("TitleSVGColorScheme").Value = GetLabel("worksheetHelpColorsSVGColorScheme")
    HelpColorsSheet.Range("TitleSVGColorName").Value = GetLabel("worksheetHelpColorsSVGColorName")
    HelpColorsSheet.Range("TitleBrewerColorSchemes").Value = GetLabel("worksheetHelpColorsBrewerColorSchemes")
    HelpColorsSheet.Range("TitleBrewerColorScheme").Value = GetLabel("worksheetHelpColorsBrewerColorScheme")
    HelpColorsSheet.Range("TitleBrewerColorName").Value = GetLabel("worksheetHelpColorsBrewerColorName")
End Sub

Private Sub LocalizeWorksheetHelpShapes()
    HelpShapesSheet.name = GetLabel("worksheetHelpShapesName")
    HelpShapesSheet.Range("TitleHelpShapesInstruction").Value = GetLabel("worksheetHelpShapesInstructions")
    HelpShapesSheet.Range("TitleHelpShapesRequires238").Value = GetLabel("worksheetHelpShapesRequires238")
End Sub

Private Sub LocalizeWorksheetSettings()
    SettingsSheet.name = GetLabel("worksheetSettingsName")
    SettingsSheet.Range("StylesWorksheetName").Value = StylesSheet.name ' Used by INDIRECT formulas
    SettingsSheet.Range("TitleSettingsImagePath").Value = GetLabel("worksheetSettingsImagePath")
    SettingsSheet.Range("TitleSettingsAdditionalOpt").Value = GetLabel("worksheetSettingsAdditionalOptions")
    SettingsSheet.Range("TitleSettingsCancel").Value = GetLabel("worksheetSettingsCancelGraphing")
    SettingsSheet.Range("TitleSettingsSeconds").Value = GetLabel("worksheetSettingsSeconds")
    SettingsSheet.Range("TitleSettingsPicture").Value = GetLabel("worksheetSettingsPictureName")
    SettingsSheet.Range("TitleSettingsPathToDot").Value = GetLabel("worksheetSettingsPathToDot")
    SettingsSheet.Range("TitleSettingsCmdParms").Value = GetLabel("worksheetSettingsAdditionalCmdParameters")
End Sub

Public Sub LocalizeWorksheetSql()
    SqlSheet.name = GetLabel("worksheetSqlName")
    
    Dim sql As sqlWorksheet
    sql = GetSettingsForSqlWorksheet
    
    SqlSheet.Cells.Item(sql.headingRow, sql.sqlStatementColumn).Value = GetLabel("worksheetSqlSqlSelectStatement")
    SqlSheet.Cells.Item(sql.headingRow, sql.excelFileColumn).Value = GetLabel("worksheetSqlExcelDataFile")
    SqlSheet.Cells.Item(sql.headingRow, sql.statusColumn).Value = GetLabel("worksheetSqlStatus")
End Sub

Public Sub LocalizeWorksheetSvg()
    SvgSheet.name = GetLabel("worksheetSvgName")
    
    Dim svg As svgWorksheet
    svg = GetSettingsForSvgWorksheet
    
    SvgSheet.Cells.Item(svg.headingRow, svg.findColumn).Value = GetLabel("worksheetSvgFind")
    SvgSheet.Cells.Item(svg.headingRow, svg.replaceColumn).Value = GetLabel("worksheetSvgReplace")
End Sub

Private Sub LocalizeWorksheetStyleDesigner()
    StyleDesignerSheet.name = GetLabel("worksheetStyleDesignerName")
    StyleDesignerSheet.Range("TitleStyleDesignerLabelText").Value = GetLabel("worksheetStyleDesignerLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerXlabelText").Value = GetLabel("worksheetStyleDesignerXLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerFormatString").Value = GetLabel("worksheetStyleDesignerFormatString")
    StyleDesignerSheet.Range("TitleStyleDesignerTailLabelText").Value = GetLabel("worksheetStyleDesignerTailLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerHeadLabelText").Value = GetLabel("worksheetStyleDesignerHeadLabelText")
End Sub

Public Sub LocalizeWorksheetStyles()
    StylesSheet.name = GetLabel("worksheetStylesName")
    StylesSheet.Range("TitleStylesStyleName").Value = GetLabel("worksheetStylesStyleName")
    StylesSheet.Range("TitleStylesFormat").Value = GetLabel("worksheetStylesFormat")
    StylesSheet.Range("TitleStylesStyleType").Value = GetLabel("worksheetStylesStyleType")
End Sub

Public Function LocalizeCacheKeys(ByVal worksheetName As String) As Dictionary
    Dim keysToRow As Dictionary
    Set keysToRow = New Dictionary
    
    localeWorksheet = worksheetName
    
    Dim controlId As String
    
    Dim row As Long
    
    ' Find last row with data
    Dim lastRow As Long
    With ActiveWorkbook.Worksheets.[_Default](worksheetName).UsedRange
        lastRow = .Cells(.Cells.Count).row
    End With

    For row = 2 To lastRow
        controlId = Trim$(ActiveWorkbook.Sheets.[_Default](worksheetName).Cells(row, 1).Value)
        If controlId <> vbNullString Then
            If Not keysToRow.Exists(controlId) Then
                keysToRow.Add controlId, row
            End If
        End If
    Next
    
    Set LocalizeCacheKeys = keysToRow
End Function

Public Function LocalizeGetString(ByVal controlId As String, ByVal col As Long) As String
    On Error GoTo ErrorHandler
    Dim row As Long
    
    ' Lazy initialization. Plus, if an error is thrown, VBA wipes out global
    ' variables. This test is to ensure the localization information gets reloaded.
    If localeIds Is Nothing Then
        InitializeLocalization
    End If
    
    ' Look for the specified key in the set of keys of the selected language
    If localeIds.Exists(controlId) Then
        row = localeIds.Item(controlId)
        LocalizeGetString = ActiveWorkbook.Sheets.[_Default](localeWorksheet).Cells(row, col).Value
    ElseIf masterIds.Exists(controlId) Then
        ' Key not found in selected language, look in master set of keys
        row = masterIds.Item(controlId)
        LocalizeGetString = ActiveWorkbook.Sheets.[_Default](WORKSHEET_LOCALE_MASTER).Cells(row, col).Value
    Else
        ' not found in master set either, return the controlId as the value
        LocalizeGetString = controlId
    End If
    Exit Function
ErrorHandler:
    LocalizeGetString = controlId
End Function

Public Function GetMessage(ByVal controlId As String) As String
    GetMessage = LocalizeGetString(controlId, LOCALE_COL_SCREENTIP)
End Function

Public Function GetScreentip(ByVal controlId As String) As String
    GetScreentip = LocalizeGetString(controlId, LOCALE_COL_SCREENTIP)
End Function

Public Function GetSupertip(ByVal controlId As String) As String
    GetSupertip = LocalizeGetString(controlId, LOCALE_COL_SUPERTIP)
End Function

Public Function GetLabel(ByVal controlId As String) As String
    If GetVerbose() Then
        GetLabel = LocalizeGetString(controlId, LOCALE_COL_LABEL_VERBOSE)
    Else
        GetLabel = LocalizeGetString(controlId, LOCALE_COL_LABEL_COMPACT)
    End If
End Function

Private Sub InitVerbose()
#If Mac Then
    SetVerbose (True)
#Else
    SetVerbose (False)
#End If
End Sub
