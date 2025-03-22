Attribute VB_Name = "modLocalize"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Locale")
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed

Option Explicit

Private localeIds As Dictionary
Private localeWorksheet As String
Private masterIds As Dictionary
Private Verbose As Boolean

Public Sub SetVerbose(ByVal useVerboseLabels As Boolean)
    Verbose = useVerboseLabels
End Sub

Public Function GetVerbose() As Boolean
    GetVerbose = Verbose
End Function

Public Sub InitializeLocalization()
    ' Default localization. Used if specified localization is missing keys
    Set masterIds = LocalizeCacheKeys(WORKSHEET_LOCALE_MASTER)
    
    ' Get the language worksheet name specified in the settings
    Dim localeWorksheet As String
    localeWorksheet = SettingsSheet.Range(SETTINGS_LANGUAGE).value
    
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
    Set localeIds = LocalizeCacheKeys(SettingsSheet.Range(SETTINGS_LANGUAGE).value)
    
    ' Localize column headings and/or titles on each worksheet
    LocalizeWorksheetAbout
    LocalizeWorksheetConsole
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
    
    AboutSheet.Range("AboutLicenseName").value = GetSupertip("AboutLicenseName")
    AboutSheet.Range("AboutLicenseCopyright").value = GetSupertip("AboutLicenseCopyright")
    
    AboutSheet.Range("AboutLicenseParagraph01").value = GetSupertip("AboutLicenseParagraph01") & vbNewLine
    AboutSheet.Range("AboutLicenseParagraph01").rows.AutoFit
    
    AboutSheet.Range("AboutLicenseParagraph02").value = GetSupertip("AboutLicenseParagraph02") & vbNewLine
    AboutSheet.Range("AboutLicenseParagraph02").rows.AutoFit
    
    AboutSheet.Range("AboutLicenseParagraph03").value = GetSupertip("AboutLicenseParagraph03") & vbNewLine
    AboutSheet.Range("AboutLicenseParagraph03").rows.AutoFit
    
    AboutSheet.Range("AboutSpecialThanks01").value = GetSupertip("AboutSpecialThanks")
    AboutSheet.Range("AboutSpecialThanks02").value = GetSupertip("AboutSpecialThanks")
    AboutSheet.Range("AboutForProvidingPolski").value = GetSupertip("AboutProvidingPolskiTranslations")
End Sub

Public Sub LocalizeWorksheetConsole()
    ConsoleSheet.name = GetLabel("worksheetConsoleName")
End Sub

Public Sub LocalizeWorksheetData()
    DataSheet.name = GetLabel("worksheetDataName")
    
    Dim dataWs As dataWorksheet
    dataWs = GetSettingsForDataWorksheet(SettingsSheet.name)
   
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.itemColumn).value = GetLabel(RIBBON_CTL_SHOW_ITEM)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.tailLabelColumn).value = GetLabel(RIBBON_CTL_SHOW_TAIL_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.labelColumn).value = GetLabel(RIBBON_CTL_SHOW_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.xLabelColumn).value = GetLabel(RIBBON_CTL_SHOW_OUTSIDE_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.headLabelColumn).value = GetLabel(RIBBON_CTL_SHOW_HEAD_LABEL)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.tooltipColumn).value = GetLabel(RIBBON_CTL_SHOW_TOOLTIP)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.isRelatedToItemColumn).value = GetLabel(RIBBON_CTL_SHOW_IS_RELATED_TO_ITEM)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.styleNameColumn).value = GetLabel(RIBBON_CTL_SHOW_STYLE)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.extraAttributesColumn).value = GetLabel(RIBBON_CTL_SHOW_EXTRA_STYLE_ATTRIBUTES)
    DataSheet.Cells.Item(dataWs.headingRow, dataWs.errorMessageColumn).value = GetLabel(RIBBON_CTL_SHOW_MESSAGES)
End Sub

Public Sub LocalizeWorksheetSource()
    SourceSheet.name = GetLabel("worksheetSourceName")
    
    Dim sourceWs As sourceWorksheet
    sourceWs = GetSettingsForSourceWorksheet
    
    SourceSheet.Cells.Item(sourceWs.headingRow, sourceWs.lineNumberColumn).value = GetLabel("worksheetSourceLine")
    SourceSheet.Cells.Item(sourceWs.headingRow, sourceWs.sourceColumn).value = GetLabel("worksheetSourceGraphvizSource")
    
    DotSourceForm.CopyButton.caption = GetLabel("sourceFormCopy")
    DotSourceForm.wordWrapToggle.caption = GetLabel("sourceFormWrapText")
End Sub

Private Sub LocalizeWorksheetGraph()
    GraphSheet.name = GetLabel("worksheetGraphName")
End Sub

Private Sub LocalizeWorksheetHelpAttributes()
    HelpAttributesSheet.name = GetLabel("worksheetHelpAttributesName")
End Sub

Private Sub LocalizeWorksheetHelpColors()
    HelpColorsSheet.name = GetLabel("worksheetHelpColorsName")
    HelpColorsSheet.Range("TitleX11ColorScheme").value = GetLabel("worksheetHelpColorsX11ColorScheme")
    HelpColorsSheet.Range("TitleX11ColorName").value = GetLabel("worksheetHelpColorsX11ColorName")
    HelpColorsSheet.Range("TitleSVGColorScheme").value = GetLabel("worksheetHelpColorsSVGColorScheme")
    HelpColorsSheet.Range("TitleSVGColorName").value = GetLabel("worksheetHelpColorsSVGColorName")
    HelpColorsSheet.Range("TitleBrewerColorSchemes").value = GetLabel("worksheetHelpColorsBrewerColorSchemes")
    HelpColorsSheet.Range("TitleBrewerColorScheme").value = GetLabel("worksheetHelpColorsBrewerColorScheme")
    HelpColorsSheet.Range("TitleBrewerColorName").value = GetLabel("worksheetHelpColorsBrewerColorName")
End Sub

Private Sub LocalizeWorksheetHelpShapes()
    HelpShapesSheet.name = GetLabel("worksheetHelpShapesName")
    HelpShapesSheet.Range("TitleHelpShapesInstruction").value = GetLabel("worksheetHelpShapesInstructions")
    HelpShapesSheet.Range("TitleHelpShapesRequires238").value = GetLabel("worksheetHelpShapesRequires238")
End Sub

Private Sub LocalizeWorksheetSettings()
    SettingsSheet.name = GetLabel("worksheetSettingsName")
    SettingsSheet.Range("StylesWorksheetName").value = StylesSheet.name ' Used by INDIRECT formulas
    SettingsSheet.Range("TitleSettingsImagePath").value = GetLabel("worksheetSettingsImagePath")
    SettingsSheet.Range("TitleSettingsAdditionalOpt").value = GetLabel("worksheetSettingsAdditionalOptions")
    SettingsSheet.Range("TitleSettingsPicture").value = GetLabel("worksheetSettingsPictureName")
    SettingsSheet.Range("TitleSettingsPathToDot").value = GetLabel("worksheetSettingsPathToDot")
    SettingsSheet.Range("TitleSettingsCmdParms").value = GetLabel("worksheetSettingsAdditionalCmdParameters")
End Sub

Public Sub LocalizeWorksheetSql()
    SqlSheet.name = GetLabel("worksheetSqlName")
    
    Dim sql As sqlWorksheet
    sql = GetSettingsForSqlWorksheet
    
    SqlSheet.Cells.Item(sql.headingRow, sql.sqlStatementColumn).value = GetLabel("worksheetSqlSqlSelectStatement")
    SqlSheet.Cells.Item(sql.headingRow, sql.excelFileColumn).value = GetLabel("worksheetSqlExcelDataFile")
    SqlSheet.Cells.Item(sql.headingRow, sql.statusColumn).value = GetLabel("worksheetSqlStatus")
End Sub

Public Sub LocalizeWorksheetSvg()
    SvgSheet.name = GetLabel("worksheetSvgName")
    
    Dim svg As svgWorksheet
    svg = GetSettingsForSvgWorksheet
    
    SvgSheet.Cells.Item(svg.headingRow, svg.findColumn).value = GetLabel("worksheetSvgFind")
    SvgSheet.Cells.Item(svg.headingRow, svg.replaceColumn).value = GetLabel("worksheetSvgReplace")
End Sub

Private Sub LocalizeWorksheetStyleDesigner()
    StyleDesignerSheet.name = GetLabel("worksheetStyleDesignerName")
    StyleDesignerSheet.Range("TitleStyleDesignerLabelText").value = GetLabel("worksheetStyleDesignerLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerXlabelText").value = GetLabel("worksheetStyleDesignerXLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerFormatString").value = GetLabel("worksheetStyleDesignerFormatString")
    StyleDesignerSheet.Range("TitleStyleDesignerTailLabelText").value = GetLabel("worksheetStyleDesignerTailLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerHeadLabelText").value = GetLabel("worksheetStyleDesignerHeadLabelText")
End Sub

Public Sub LocalizeWorksheetStyles()
    StylesSheet.name = GetLabel("worksheetStylesName")
    StylesSheet.Range("TitleStylesStyleName").value = GetLabel("worksheetStylesStyleName")
    StylesSheet.Range("TitleStylesFormat").value = GetLabel("worksheetStylesFormat")
    StylesSheet.Range("TitleStylesStyleType").value = GetLabel("worksheetStylesStyleType")
End Sub

Public Function LocalizeCacheKeys(ByVal worksheetName As String) As Dictionary
    Dim keysToRow As Dictionary
    Set keysToRow = New Dictionary
    
    localeWorksheet = worksheetName
    
    Dim controlId As String
    
    Dim row As Long
    
    ' Find last row with data
    Dim lastRow As Long
    With ActiveWorkbook.worksheets.[_Default](worksheetName).UsedRange
        lastRow = .Cells(.Cells.count).row
    End With

    For row = 2 To lastRow
        controlId = Trim$(ActiveWorkbook.Sheets.[_Default](worksheetName).Cells(row, 1).value)
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
        LocalizeGetString = ActiveWorkbook.Sheets.[_Default](localeWorksheet).Cells(row, col).value
    ElseIf masterIds.Exists(controlId) Then
        ' Key not found in selected language, look in master set of keys
        row = masterIds.Item(controlId)
        LocalizeGetString = ActiveWorkbook.Sheets.[_Default](WORKSHEET_LOCALE_MASTER).Cells(row, col).value
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
