Attribute VB_Name = "modLocalize"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modLocalize
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Locale / Localization Subsystem
'
' ROLE:
'   Central internationalization (i18n) engine. Loads, caches, and resolves
'   localized UI strings; orchestrates full-workbook translation, worksheet
'   renaming, and Ribbon/Form label synchronization.
'
' RESPONSIBILITIES:
'   - Locale caching:
'       • Build O(1) lookup dictionaries for active and master locales
'       • Maintain fallback chain: Active -> Master -> Key
'   - Global translation pipeline:
'       • Rehydrate caches on language change
'       • Localize all functional worksheets in dependency-safe order
'       • Update Ribbon captions, UserForms, and Named Ranges
'   - Worksheet-specific localization:
'       • About, Console, Data, Graph, Source, SQL, SVG
'       • Styles and Style Designer (INDIRECT-sensitive)
'       • Help: Attributes, Colors, Shapes
'   - Diagnostic modes:
'       • Verbose mode for key-based auditing
'       • Safe fallback to prevent blank UI labels
'
' ARCHITECTURAL NOTES:
'   - Uses Scripting.Dictionary for high-performance lookup.
'   - Master locale (English) guarantees resilience against partial
'     translations or missing keys.
'   - Localization order is carefully sequenced to avoid INDIRECT breakage
'     and formula invalidation in Settings and Styles.
'   - Integrates with modRibbon, DotSourceForm, and all worksheet controllers.
'
' USAGE:
'   - Triggered when the user changes the Language setting.
'   - Ensures a fully localized experience across worksheets, forms, and UI.
'
' RELATED WIKI PAGES:
'   - Localization Architecture
'   - Locale Worksheet Specification
'   - Ribbon & Worksheet Synchronization
' =============================================================================

Option Explicit

' ==========================================================================
' SECTION: STATE MANAGEMENT
' ==========================================================================

' Internal registries for mapping keys to translated strings
Private localeIds As Dictionary
Private localeWorksheet As String
Private masterIds As Dictionary

' Toggle for diagnostic/development labeling
Private Verbose As Boolean

''
' VERBOSE ACCESSORS: Manages the 'Verbose' state.
' When True, the system can display underlying Keys instead of translated
' values, aiding in debugging and string identification.
'
Public Sub SetVerbose(ByVal useVerboseLabels As Boolean)
    Verbose = useVerboseLabels
End Sub

Public Function GetVerbose() As Boolean
    GetVerbose = Verbose
End Function

' ==========================================================================
' PROCEDURE: InitializeLocalization
' PURPOSE:
'   Initializes the high-speed cache for multi-language UI and messages.
'
' TECHNICAL WORKFLOW:
'   1. MASTER FALLBACK: Loads the 'MASTER' (English) locale into a
'      Dictionary to serve as a safety net for missing translations.
'   2. LOCALE RESOLUTION:
'      - Retrieves the user's preferred language worksheet from Settings.
'      - Validates the existence of the worksheet to prevent runtime errors.
'      - Defaults back to Master if the specified sheet is missing or invalid.
'   3. ACTIVE CACHING: Hydrates the 'localeIds' Dictionary with the
'      selected language's key-value pairs for O(1) lookup performance.
'   4. STATE SYNC: Calls 'InitVerbose' to set the diagnostic labeling mode.
' ==========================================================================
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

' ==========================================================================
' PROCEDURE: Localize
' PURPOSE:
'   Triggers a global refresh of all worksheet-based UI elements.
'
' TECHNICAL WORKFLOW:
'   1. PERFORMANCE LOCKDOWN: Invokes 'OptimizeCode_Begin' to disable screen
'      updating and events, ensuring the bulk update is near-instantaneous.
'   2. CACHE REFRESH: Re-loads the translation dictionary based on the
'      current selection in the 'Settings' worksheet.
'   3. SEQUENTIAL TRANSLATION: Iterates through every functional worksheet
'      in the project, calling specialized 'LocalizeWorksheet...' routines.
'   4. DEPENDENCY MANAGEMENT: Specifically executes 'LocalizeWorksheetStyles'
'      before 'Settings' to prevent #REF! errors in INDIRECT-dependent formulas.
'   5. STATE RESTORATION: Calls 'OptimizeCode_End' to return Excel to its
'      standard interactive state.
'
' USAGE:
'   - Triggered when the user changes the "Language" dropdown in Settings.
'   - Essential for maintaining a cohesive experience across localized tabs.
' ==========================================================================
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

' ==========================================================================
' PROCEDURE: LocalizeWorksheetAbout
' PURPOSE:
'   Translates the informational and legal content of the 'About' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the physical worksheet tab name using the
'      'worksheetAboutName' locale key.
'   2. TEXT INJECTION: Fetches translated licensing paragraphs and credits
'      using 'GetSupertip' (leveraging the high-capacity text storage).
'   3. DYNAMIC FORMATTING:
'      - Appends 'vbNewLine' to paragraphs to maintain visual separation.
'      - Triggers '.rows.AutoFit' to ensure varying translation lengths
'        remain fully visible without manual resizing.
'   4. ACKNOWLEDGMENTS: Updates specific credit lines (e.g., for Polish
'      translations) to maintain accurate contributor attribution across locales.
' ==========================================================================
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

' ==========================================================================
' PROCEDURE: LocalizeWorksheetConsole
' PURPOSE:
'   Updates the visual identity of the Console worksheet.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Retrieves the translated string for the 'Console'
'      identity via 'GetLabel' and applies it to the worksheet's 'Name' property.
'
' USAGE:
'   - Called as part of the global 'Localize' batch process.
'   - Ensures the diagnostic interface remains recognizable in any locale.
' ==========================================================================
Public Sub LocalizeWorksheetConsole()
    ConsoleSheet.name = GetLabel("worksheetConsoleName")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetData
' PURPOSE:
'   Translates the 'Data' worksheet tab and its critical column headers.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'Data' worksheet's tab name via 'GetLabel'.
'   2. LAYOUT DISCOVERY: Calls 'GetSettingsForDataWorksheet' to identify
'      the dynamic column indices and heading row location.
'   3. RIBBON SYNC: Maps column headers (e.g., Item, Label, Style) to the
'      same locale keys used by the Ribbon UI.
'   4. DATA BINDING: Directly updates 'DataSheet.Cells' for every logical
'      column (Item, Tooltip, Relationships, etc.) in the schema.
'
' USAGE:
'   - Ensures the 'Data' tab and the 'Relationship Visualizer' Ribbon tab
'     use identical terminology to prevent user confusion.
' ==========================================================================
Public Sub LocalizeWorksheetData()
    DataSheet.name = GetLabel("worksheetDataName")
    
    Dim dataWs As dataWorksheet
    dataWs = GetSettingsForDataWorksheet(SettingsSheet.name)
   
    DataSheet.Cells.item(dataWs.headingRow, dataWs.itemColumn).value = GetLabel(RIBBON_CTL_SHOW_ITEM)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.tailLabelColumn).value = GetLabel(RIBBON_CTL_SHOW_TAIL_LABEL)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.labelColumn).value = GetLabel(RIBBON_CTL_SHOW_LABEL)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.xLabelColumn).value = GetLabel(RIBBON_CTL_SHOW_OUTSIDE_LABEL)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.headLabelColumn).value = GetLabel(RIBBON_CTL_SHOW_HEAD_LABEL)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.tooltipColumn).value = GetLabel(RIBBON_CTL_SHOW_TOOLTIP)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.isRelatedToItemColumn).value = GetLabel(RIBBON_CTL_SHOW_IS_RELATED_TO_ITEM)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.styleNameColumn).value = GetLabel(RIBBON_CTL_SHOW_STYLE)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.extraAttributesColumn).value = GetLabel(RIBBON_CTL_SHOW_EXTRA_STYLE_ATTRIBUTES)
    DataSheet.Cells.item(dataWs.headingRow, dataWs.errorMessageColumn).value = GetLabel(RIBBON_CTL_SHOW_MESSAGES)
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetSource
' PURPOSE:
'   Translates the 'Source' worksheet and the 'DotSourceForm' dialog.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'Source' worksheet tab name using 'GetLabel'.
'   2. HEADER MAPPING:
'      - Resolves dynamic column locations via 'GetSettingsForSourceWorksheet'.
'      - Updates the 'Line Number' and 'Graphviz Source' headers.
'   3. FORM LOCALIZATION:
'      - Directly updates the 'Copy' button caption on the 'DotSourceForm'.
'      - Translates the 'Word Wrap' toggle button text.
'
' USAGE:
'   - Ensures that developers and power users reviewing the raw DOT code
'     have a localized experience within the viewer form.
' ==========================================================================
Public Sub LocalizeWorksheetSource()
    SourceSheet.name = GetLabel("worksheetSourceName")
    
    Dim sourceWs As sourceWorksheet
    sourceWs = GetSettingsForSourceWorksheet
    
    SourceSheet.Cells.item(sourceWs.headingRow, sourceWs.lineNumberColumn).value = GetLabel("worksheetSourceLine")
    SourceSheet.Cells.item(sourceWs.headingRow, sourceWs.sourceColumn).value = GetLabel("worksheetSourceGraphvizSource")
    
    DotSourceForm.CopyButton.caption = GetLabel("sourceFormCopy")
    DotSourceForm.wordWrapToggle.caption = GetLabel("sourceFormWrapText")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetGraph
' PURPOSE:
'   Updates the visual identity of the primary 'Graph' output worksheet.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Retrieves the translated string for the 'Graph' tab
'      identity via 'GetLabel' and applies it to the sheet's 'Name' property.
'
' USAGE:
'   - Called during the global 'Localize' batch process.
'   - Ensures the main destination for rendered images is consistently
'     named across all supported languages.
' ==========================================================================
Private Sub LocalizeWorksheetGraph()
    GraphSheet.name = GetLabel("worksheetGraphName")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetHelpAttributes
' PURPOSE:
'   Updates the tab name for the Graphviz Attributes help reference.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Fetches the localized name via 'GetLabel' and applies
'      it to the 'HelpAttributesSheet' object.
'
' USAGE:
'   - Part of the documentation-sync process during a locale swap.
'   - Helps users navigate to the correct reference sheet for advanced styling.
' ==========================================================================
Private Sub LocalizeWorksheetHelpAttributes()
    HelpAttributesSheet.name = GetLabel("worksheetHelpAttributesName")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetHelpColors
' PURPOSE:
'   Translates the 'Help Colors' worksheet tab and its internal section headers.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'HelpColorsSheet' tab name via 'GetLabel'.
'   2. SECTION TITLING: Maps specific Named Ranges (e.g., "TitleX11ColorScheme")
'      to their corresponding localized strings.
'   3. PALETTE DOCUMENTATION: Ensures that technical headers for X11, SVG,
'      and the sophisticated 'Brewer' color schemes are translated while
'      maintaining the integrity of the underlying color data.
'
' USAGE:
'   - Synchronizes the built-in color reference guides with the active language.
'   - Supports the "Styles & Color" sections of the help documentation.
' ==========================================================================
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

' ==========================================================================
' PROCEDURE: LocalizeWorksheetHelpShapes
' PURPOSE:
'   Translates the 'Help Shapes' worksheet tab and its internal instructions.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'HelpShapesSheet' tab name via 'GetLabel'.
'   2. INSTRUCTIONAL MAPPING: Updates the "TitleHelpShapesInstruction"
'      range with localized usage guidance.
'   3. VERSION WARNING: Translates the "Requires238" warning, which informs
'      the user that certain advanced shapes require Graphviz version 2.38
'      or higher to render correctly.
'
' USAGE:
'   - Ensures that technical rendering constraints are communicated clearly
'     to users across all supported locales.
' ==========================================================================
Private Sub LocalizeWorksheetHelpShapes()
    HelpShapesSheet.name = GetLabel("worksheetHelpShapesName")
    HelpShapesSheet.Range("TitleHelpShapesInstruction").value = GetLabel("worksheetHelpShapesInstructions")
    HelpShapesSheet.Range("TitleHelpShapesRequires238").value = GetLabel("worksheetHelpShapesRequires238")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetSettings
' PURPOSE:
'   Translates the 'Settings' worksheet tab and its primary input headers.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'Settings' tab name via 'GetLabel'.
'   2. FORMULA SYNCHRONIZATION: Updates the 'StylesWorksheetName' range
'      with the current localized name of the Styles sheet. This is critical
'      for 'INDIRECT' formulas that break if a worksheet is renamed.
'   3. CONFIGURATION TITLING: Maps specific Named Ranges to their
'      localized equivalents for:
'      - Image and Graphviz application paths.
'      - Command-line parameters and additional options.
'      - Default file naming conventions.
'
' USAGE:
'   - Ensures the administrative backbone of the tool remains accessible
'     and functionally sound across all locales.
' ==========================================================================
Private Sub LocalizeWorksheetSettings()
    SettingsSheet.name = GetLabel("worksheetSettingsName")
    SettingsSheet.Range("StylesWorksheetName").value = StylesSheet.name ' Used by INDIRECT formulas
    SettingsSheet.Range("TitleSettingsImagePath").value = GetLabel("worksheetSettingsImagePath")
    SettingsSheet.Range("TitleSettingsAdditionalOpt").value = GetLabel("worksheetSettingsAdditionalOptions")
    SettingsSheet.Range("TitleSettingsPicture").value = GetLabel("worksheetSettingsPictureName")
    SettingsSheet.Range("TitleSettingsPathToDot").value = GetLabel("worksheetSettingsPathToDot")
    SettingsSheet.Range("TitleSettingsCmdParms").value = GetLabel("worksheetSettingsAdditionalCmdParameters")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetSql
' PURPOSE:
'   Translates the 'SQL' worksheet tab and its critical operational headers.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'SQL' worksheet's tab name via 'GetLabel'.
'   2. LAYOUT DISCOVERY: Resolves the current column indices for the SQL
'      environment using 'GetSettingsForSqlWorksheet'.
'   3. HEADER MAPPING: Translates the core interaction columns:
'      - SQL Statement: Where the user writes their queries.
'      - Excel Data File: The source path for external lookups.
'      - Status: The diagnostic feedback area for query results.
'
' USAGE:
'   - Ensures that the "IDE" for writing SQL queries is fully localized
'     for power users in all supported regions.
' ==========================================================================
Public Sub LocalizeWorksheetSql()
    SqlSheet.name = GetLabel("worksheetSqlName")
    
    Dim sql As sqlWorksheet
    sql = GetSettingsForSqlWorksheet
    
    SqlSheet.Cells.item(sql.headingRow, sql.sqlStatementColumn).value = GetLabel("worksheetSqlSqlSelectStatement")
    SqlSheet.Cells.item(sql.headingRow, sql.excelFileColumn).value = GetLabel("worksheetSqlExcelDataFile")
    SqlSheet.Cells.item(sql.headingRow, sql.statusColumn).value = GetLabel("worksheetSqlStatus")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetSvg
' PURPOSE:
'   Translates the 'SVG' worksheet tab and its transformation headers.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'SVG' worksheet's tab name via 'GetLabel'.
'   2. LAYOUT DISCOVERY: Resolves the dynamic column positions using the
'      'GetSettingsForSvgWorksheet' helper.
'   3. HEADER MAPPING: Translates the core interaction columns:
'      - Find: The target string or attribute to locate in the SVG XML.
'      - Replace: The new value or script to inject.
'
' USAGE:
'   - Synchronizes the SVG interactivity engine's UI with the active locale.
'   - Supports the "Interactive SVG" section of the documentation.
' ==========================================================================
Public Sub LocalizeWorksheetSvg()
    SvgSheet.name = GetLabel("worksheetSvgName")
    
    Dim svg As svgWorksheet
    svg = GetSettingsForSvgWorksheet
    
    SvgSheet.Cells.item(svg.headingRow, svg.findColumn).value = GetLabel("worksheetSvgFind")
    SvgSheet.Cells.item(svg.headingRow, svg.replaceColumn).value = GetLabel("worksheetSvgReplace")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetStyleDesigner
' PURPOSE:
'   Translates the 'Style Designer' worksheet and its interactive controls.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the worksheet's tab name via 'GetLabel'.
'   2. INPUT LABEL MAPPING: Synchronizes specific Named Ranges with their
'      localized labels for:
'      - Node Labels and XLabels (Outside labels).
'      - Edge-specific labels (Tail and Head labels).
'      - Format strings and Style Name definitions.
'   3. BUTTON LOCALIZATION: Updates the 'Save' button's physical caption
'      on the sheet using the 'worksheetStyleDesignerSaveButtonText' key.
'
' USAGE:
'   - Ensures the graphical interface for building styles remains intuitive
'     for designers across all supported languages.
' ==========================================================================
Private Sub LocalizeWorksheetStyleDesigner()
    StyleDesignerSheet.name = GetLabel("worksheetStyleDesignerName")
    StyleDesignerSheet.Range("TitleStyleDesignerLabelText").value = GetLabel("worksheetStyleDesignerLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerXlabelText").value = GetLabel("worksheetStyleDesignerXLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerFormatString").value = GetLabel("worksheetStyleDesignerFormatString")
    StyleDesignerSheet.Range("TitleStyleDesignerTailLabelText").value = GetLabel("worksheetStyleDesignerTailLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerHeadLabelText").value = GetLabel("worksheetStyleDesignerHeadLabelText")
    StyleDesignerSheet.Range("TitleStyleDesignerStyleNameText").value = GetLabel("worksheetStyleDesignerStyleNameText")
    StyleDesignerSheet.buttons("StyleDesignerSaveButton").caption = GetLabel("worksheetStyleDesignerSaveButtonText")
End Sub

' ==========================================================================
' PROCEDURE: LocalizeWorksheetStyles
' PURPOSE:
'   Translates the 'Styles' gallery worksheet tab and its header labels.
'
' TECHNICAL WORKFLOW:
'   1. TAB RENAMING: Updates the 'Styles' worksheet tab name via 'GetLabel'.
'   2. HEADER SYNCHRONIZATION: Maps the primary structural headers:
'      - Style Name: The unique identifier for the style.
'      - Format: The DOT attribute string applied to the object.
'      - Style Type: The classification (Node, Edge, Cluster, etc.).
'
' USAGE:
'   - Essential for ensuring the style library remains navigable in any locale.
'   - Must be executed early in the 'Localize' batch because 'Settings'
'     formulas often rely on this sheet's localized name via INDIRECT.
' ==========================================================================
Public Sub LocalizeWorksheetStyles()
    StylesSheet.name = GetLabel("worksheetStylesName")
    StylesSheet.Range("TitleStylesStyleName").value = GetLabel("worksheetStylesStyleName")
    StylesSheet.Range("TitleStylesFormat").value = GetLabel("worksheetStylesFormat")
    StylesSheet.Range("TitleStylesStyleType").value = GetLabel("worksheetStylesStyleType")
End Sub

' ==========================================================================
' FUNCTION: LocalizeCacheKeys
' PURPOSE:
'   Creates an in-memory index of a locale worksheet for O(1) performance.
'
' TECHNICAL WORKFLOW:
'   1. TARGETING: Identifies the specified 'worksheetName' and determines
'      the active 'UsedRange' to identify the data boundary.
'   2. SCANNING: Iterates from row 2 (skipping headers) to the last populated
'      row in the first column.
'   3. REGISTRATION:
'      - Captures the unique 'controlId' (the Key).
'      - Stores the row index in a 'Scripting.Dictionary'.
'   4. COLLISION PREVENTION: Checks 'Exists' before adding to prevent
'      duplicate key errors in the dictionary.
'
' USAGE:
'   - Called during 'InitializeLocalization' for the Master and Active locales.
'   - Essential for decoupling the UI text from physical cell addresses.
' ==========================================================================
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

' ==========================================================================
' FUNCTION: LocalizeGetString
' PURPOSE:
'   Retrieves a specific string attribute (Label, Supertip, etc.) for a
'   given Key across all available locales.
'
' TECHNICAL WORKFLOW:
'   1. LAZY INITIALIZATION: Checks if the locale cache has been cleared
'      (e.g., due to a VBA reset) and re-initializes if necessary.
'   2. PREFERRED LOOKUP: Searches the 'localeIds' dictionary for the
'      specified 'controlId' in the user's selected language.
'   3. MASTER FALLBACK: If not found, it automatically searches the
'      'masterIds' (English) dictionary to find the value.
'   4. IDENTITY FALLBACK: If the key is missing from all locale sheets,
'      it returns the 'controlId' itself as a diagnostic placeholder.
'   5. ERROR TRAPPING: Returns the raw ID in the event of an unexpected
'      controlId, preventing a UI crash.
'
' USAGE:
'   - The internal worker for 'GetLabel', 'GetMessage', 'GetSupertip', etc.
' ==========================================================================
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
        row = localeIds.item(controlId)
        LocalizeGetString = ActiveWorkbook.Sheets.[_Default](localeWorksheet).Cells(row, col).value
    ElseIf masterIds.Exists(controlId) Then
        ' Key not found in selected language, look in master set of keys
        row = masterIds.item(controlId)
        LocalizeGetString = ActiveWorkbook.Sheets.[_Default](WORKSHEET_LOCALE_MASTER).Cells(row, col).value
    Else
        ' not found in master set either, return the controlId as the value
        LocalizeGetString = controlId
    End If
    Exit Function
ErrorHandler:
    LocalizeGetString = controlId
End Function

' ==========================================================================
' SECTION: SEMANTIC TEXT ACCESSORS
' PURPOSE:
'   High-level wrappers for retrieving specific localized attributes.
'
' TECHNICAL WORKFLOW:
'   1. COLUMN MAPPING: Maps functional requests to specific worksheet columns:
'      - GetMessage/GetScreentip: Targets 'LOCALE_COL_SCREENTIP' (Short help).
'      - GetSupertip: Targets 'LOCALE_COL_SUPERTIP' (Extended descriptions).
'   2. VERBOSE TOGGLE: 'GetLabel' evaluates the global 'Verbose' state:
'      - COMPACT: Returns the standard, user-friendly label.
'      - VERBOSE: Returns an expanded identifier, aiding in string auditing
'        during translation or development.
'   3. ABSTRACTION: Consolidates all lookups through 'LocalizeGetString' to
'      ensure the fallback and lazy-loading logic is applied consistently.
' ==========================================================================

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

' ==========================================================================
' PROCEDURE: InitVerbose
' PURPOSE:
'   Sets the initial 'Verbose' labeling state based on the host platform.
'
' TECHNICAL WORKFLOW:
'   1. PLATFORM DETECTION: Uses conditional compilation to differentiate
'      between macOS and Windows environments.
'   2. MAC CONFIGURATION (#If Mac): Sets Verbose to 'True'. This helps
'      the user by providing longer control labels as macOS ribbons do not
'      provide group information..
'   3. WINDOWS CONFIGURATION (#Else): Sets Verbose to 'False', providing
'      the standard, clean "Compact" user interface by default.
'
' USAGE:
'   - Called during 'InitializeLocalization' at workbook startup.
'   - Determines if 'GetLabel' returns user-friendly text or technical keys.
' ==========================================================================

Private Sub InitVerbose()
#If Mac Then
    SetVerbose (True)
#Else
    SetVerbose (False)
#End If
End Sub
