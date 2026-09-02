Attribute VB_Name = "modLocaleAudit"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modLocaleAudit
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Locale / Maintenance Utilities
'
' ROLE:
'   Translation integrity and synchronization engine. Ensures all non-English
'   locale worksheets remain structurally aligned with the master locale
'   (en-US), preserving key order, detecting drift, and inserting missing
'   translation rows.
'
' RESPONSIBILITIES:
'   - Key auditing:
'       o Compare locale worksheets against the master key list
'       o Detect out-of-sync rows and report discrepancies
'   - Synchronization:
'       o Rebuild locale sheets to match master ordering
'       o Insert missing keys with default English text
'       o Flag untranslated rows for downstream filtering
'   - Structural preservation:
'       o Maintain row-to-row alignment across all locale sheets
'       o Ensure control IDs, compact labels, verbose labels, screentips,
'         and supertips remain in consistent column positions
'
' ARCHITECTURAL NOTES:
'   - Uses Scripting.Dictionary for fast key-to-row mapping.
'   - Operates directly on locale worksheets; no dependency on the runtime
'     localization cache.
'   - Designed for translator workflows and release-engineering validation.
'   - Ensures that the i18n engine (modLocalize) always receives clean,
'     predictable locale sheets.
'
' USAGE:
'   - Run in VBA editor before shipping a new release to ensure all locales
'     are aligned.
'   - Used by translators to identify missing or outdated strings.
'   - Supports automated QA workflows for localization completeness.
'
' RELATED WIKI PAGES:
'   - Locale Worksheet Specification
'   - Translation Workflow & Release Checklist
'   - Localization Integrity Tools
' =============================================================================

Option Explicit

Private Enum LocaleColumns
    controlId = 1
    compactControlLabel = 2
    verboseControlLabel = 3
    screentipHeading = 4
    supertipText = 5
    controlType = 6
    sortOrder = 7
    messages = 8
End Enum

' ==========================================================================
' FUNCTION: SyncTranslations
'
' PURPOSE:
'   Executes full-locale synchronization by applying SyncLocaleToMaster across
'   all supported translation worksheets. Ensures each locale remains aligned
'   with the master (en-US) in key order, structural fields, and translation
'   metadata.
'
' FUNCTIONAL WORKFLOW:
'   1. MASTER-DRIVEN SYNCHRONIZATION:
'        - Invokes SyncLocaleToMaster for each locale worksheet.
'        - Rebuilds row ordering, inserts missing keys, and propagates
'          fallback values from the master locale.
'
'   2. TRANSLATION FLAGGING:
'        - Marks newly added or updated rows with "Requires translation"
'          according to the rules enforced by SyncLocaleToMaster.
'
'   3. LOCALE COVERAGE:
'        - Processes all configured locales (de-DE, en-GB, fr-FR, it-IT, pl-PL)
'          using a consistent master reference (locale_en-US).
'
' TECHNICAL NOTES:
'   - Serves as the orchestration entry point for full-locale maintenance.
'   - Ensures downstream routines (e.g., CompareTranslationKeys, MergeLabelField)
'     operate on structurally aligned worksheets.
'   - DeepWiki Context: Supports the multi-locale workflow described in
'     "Translation Workflow & Release Checklist" and "Locale Worksheet
'     Specification".
' ==========================================================================
Public Sub SyncTranslations()
    SyncLocaleToMaster "locale_en-US", "locale_de-DE"
    SyncLocaleToMaster "locale_en-US", "locale_en-GB"
    SyncLocaleToMaster "locale_en-US", "locale_fr-FR"
    SyncLocaleToMaster "locale_en-US", "locale_it-IT"
    SyncLocaleToMaster "locale_en-US", "locale_pl-PL"
End Sub

' ==========================================================================
' FUNCTION: AuditTranslations
'
' PURPOSE:
'   Performs a structural audit of all locale worksheets by comparing each
'   translation sheet against the master locale (en-US). Identifies key
'   misalignments, ordering drift, and empty-cell inconsistencies that must
'   be corrected before release.
'
' FUNCTIONAL WORKFLOW:
'   1. MASTER-TO-LOCALE COMPARISON:
'        - Invokes CompareTranslationKeys for each supported locale.
'        - Validates control-ID alignment and row-order consistency.
'
'   2. EMPTY-CELL VERIFICATION:
'        - Ensures blank fields in the master locale are mirrored in each
'          translation sheet.
'        - Detects missing fallback values and highlights structural gaps.
'
'   3. LOCALE COVERAGE:
'        - Audits all configured locales (de-DE, en-GB, fr-FR, it-IT, pl-PL)
'          using a consistent master reference (locale_en-US).
'
' TECHNICAL NOTES:
'   - Serves as the diagnostic counterpart to SyncTranslations, focusing on
'     verification rather than synchronization.
'   - Early detection of misaligned keys prevents downstream rendering and
'     serialization errors in the localization pipeline.
'   - DeepWiki Context: Supports the validation rules described in
'     "Locale Worksheet Specification" and "Translation Workflow & Release
'     Checklist".
' ==========================================================================
Public Sub AuditTranslations()
    CompareTranslationKeys "locale_en-US", "locale_de-DE"
    CompareTranslationKeys "locale_en-US", "locale_en-GB"
    CompareTranslationKeys "locale_en-US", "locale_fr-FR"
    CompareTranslationKeys "locale_en-US", "locale_it-IT"
    CompareTranslationKeys "locale_en-US", "locale_pl-PL"
End Sub

' ==========================================================================
' FUNCTION: CompareTranslationKeys
'
' PURPOSE:
'   Validates structural alignment between the master locale worksheet and a
'   secondary locale worksheet. Ensures key order, empty-cell semantics, and
'   fallback value propagation remain consistent across all translation sheets.
'
' FUNCTIONAL WORKFLOW:
'   1. KEY ORDER VERIFICATION:
'        - Compares control-ID values row-by-row between masterSheet and
'          sheetToTest.
'        - Emits diagnostic messages when a mismatch is detected and exits
'          early to prevent cascading errors.
'
'   2. EMPTY-CELL NORMALIZATION (MASTER -> TEST):
'        - For each row and relevant column, clears cells in sheetToTest
'          when the corresponding masterSheet cell is empty.
'        - Preserves structural parity so downstream routines can rely on
'          consistent blank-field semantics.
'
'   3. FALLBACK VALUE PROPAGATION (MASTER -> TEST):
'        - For empty cells in sheetToTest, copies non-empty values from
'          masterSheet.
'        - Marks rows requiring translation by writing "Requires translation"
'          into the Messages column for all non-type fields.
'
' TECHNICAL NOTES:
'   - Operates strictly within the column boundaries defined by LocaleColumns,
'     ensuring compatibility with the locale-sheet specification.
'   - Early exit on key mismatch prevents partial synchronization of
'     misaligned sheets, preserving data integrity.
'   - DeepWiki Context: Supports the locale-alignment rules described in
'     "Locale Worksheet Specification" and "Translation Workflow & Release
'     Checklist".
' ==========================================================================
Public Sub CompareTranslationKeys(ByVal masterSheet As String, ByVal sheetToTest As String)
    Dim lastRow As Long
    Dim row As Long
    Dim col As Long
    
    With LocaleEnUsSheet.UsedRange
        lastRow = .Cells.item(.Cells.Count).row
    End With

    For row = 2 To lastRow
        If GetCell(masterSheet, row, LocaleColumns.controlId) <> GetCell(sheetToTest, row, LocaleColumns.controlId) Then
            Debug.Print sheetToTest & " is out of sync at row " & row
            Debug.Print GetCell(masterSheet, row, 1) & " != " & GetCell(sheetToTest, row, LocaleColumns.controlId)
           Exit Sub
        End If
    Next
    
    ' Rows are in sync, now ensure empty cells in the mastter are in sync with the sheet to test
    For row = 2 To lastRow
        For col = LocaleColumns.compactControlLabel To LocaleColumns.controlType
            If GetCellLen(masterSheet, row, col) = 0 Then ClearCell sheetToTest, row, col
        Next
    Next

    ' Rows are in sync, now ensure empty cells in sheetToTest are filled in from the master
    For row = 2 To lastRow
        For col = LocaleColumns.compactControlLabel To LocaleColumns.controlType
            If GetCellLen(sheetToTest, row, col) = 0 Then       ' Cell is empty in sheetToTest
                If GetCellLen(masterSheet, row, col) > 0 Then   ' Cell is not empty in masterSheet
                    SetCell sheetToTest, row, col, GetCell(masterSheet, row, col)
                        If col <> LocaleColumns.controlType Then
                            SetCell sheetToTest, row, LocaleColumns.messages, "Requires translation"
                        End If
                End If
            End If
        Next
    Next

    Debug.Print sheetToTest & " keys are in sync with " & masterSheet
End Sub

' ==========================================================================
' FUNCTION: SyncLocaleToMaster
'
' PURPOSE:
'   Synchronizes a locale worksheet with the master locale (en-US). Ensures
'   key alignment, row ordering, fallback value propagation, and translation
'   flagging for all missing or newly introduced control-ID entries.
'
' FUNCTIONAL WORKFLOW:
'   1. MASTER KEY INDEXING:
'        - Scans the masterSheet and builds a Dictionary mapping each
'          control-ID to its corresponding row number.
'        - Provides fast lookup for row-order reconstruction and fallback
'          value retrieval.
'
'   2. ROW-ORDER ASSIGNMENT:
'        - Iterates through sheetToSync and writes the master row number
'          into the Sort Order column.
'        - Establishes a sortable ordering so the sheet can be realigned
'          to match the master locale's structure.
'
'   3. MISSING-KEY DETECTION:
'        - Removes encountered keys from the Dictionary.
'        - Remaining keys represent entries present in the masterSheet but
'          absent from sheetToSync.
'
'   4. MISSING-ROW INSERTION:
'        - For each remaining key, writes a new row containing all master
'          values (labels, screentips, supertips, control type).
'        - Marks the Messages column with "Requires translation" for all
'          non-type fields.
'        - Appends rows sequentially to preserve stable ordering.
'
' TECHNICAL NOTES:
'   - Relies on LocaleColumns for column boundaries, ensuring compatibility
'     with the locale-worksheet specification.
'   - Dictionary-based indexing provides O(1) lookup for row-order mapping
'     and fallback value retrieval.
'   - DeepWiki Context: Supports the synchronization rules described in
'     "Locale Worksheet Specification" and "Translation Workflow & Release
'     Checklist".
' ==========================================================================
Public Sub SyncLocaleToMaster(ByVal masterSheet As String, ByVal sheetToSync As String)

    ' Determine the lasr row of the master locale
    Dim lastRow As Long
    With LocaleEnUsSheet.UsedRange
        lastRow = .Cells.item(.Cells.Count).row
    End With

    ' Loop through all the rows in the master, saving the key and row
    ' number in a dictionary
    Dim masterDictionary As Dictionary
    Set masterDictionary = New Dictionary
    
    Dim row As Long
    For row = 2 To lastRow
        masterDictionary.Add GetCell(masterSheet, row, LocaleColumns.controlId), row
    Next
    
    ' Find last row with data in the worksheet to be synced with the master
    With ActiveWorkbook.worksheets.[_Default](sheetToSync).UsedRange
        lastRow = .Cells(.Cells.Count).row
    End With
    
    ' Loop through the sheet to sync with master, and fetch the row number
    ' in the master sheet that corresponds to the key. This information
    ' lets us sort the modified sheet so that the rows are in the same
    ' order as the master worksheet.
    Dim key As String
    Dim value As Long
    
    For row = 2 To lastRow
        key = GetCell(sheetToSync, row, LocaleColumns.controlId)
        value = masterDictionary.item(key)
        SetCell sheetToSync, row, LocaleColumns.sortOrder, value
    Next
    
    ' Loop through the sheet to sync again, removing the keys from the master
    ' dictionary. Any keys left in the dictionary are missing from the sheet
    ' to sync.
    For row = 2 To lastRow
        key = GetCell(sheetToSync, row, LocaleColumns.controlId)
        If masterDictionary.Exists(key) Then
            masterDictionary.Remove key
        End If
    Next
    
    ' Add column headings to the columns we will write to
    SetCell sheetToSync, 1, LocaleColumns.sortOrder, "Sort Order"
    SetCell sheetToSync, 1, LocaleColumns.messages, "Messages"
    
    ' Iterate through the remaining keys in the master dictionary. Fetch the default
    ' text from the master, and write it to the corresponding column in the
    ' sheet to sync. Include the row number, and a value which can be filtered on
    ' which identifies the row as needing to be translated.
    Dim dictKey As Variant
    For Each dictKey In masterDictionary.keys()
        SetCell sheetToSync, row, LocaleColumns.controlId, dictKey
        SetCell sheetToSync, row, LocaleColumns.compactControlLabel, GetCell(masterSheet, masterDictionary.item(dictKey), LocaleColumns.compactControlLabel)
        SetCell sheetToSync, row, LocaleColumns.verboseControlLabel, GetCell(masterSheet, masterDictionary.item(dictKey), LocaleColumns.verboseControlLabel)
        SetCell sheetToSync, row, LocaleColumns.screentipHeading, GetCell(masterSheet, masterDictionary.item(dictKey), LocaleColumns.screentipHeading)
        SetCell sheetToSync, row, LocaleColumns.supertipText, GetCell(masterSheet, masterDictionary.item(dictKey), LocaleColumns.supertipText)
        SetCell sheetToSync, row, LocaleColumns.controlType, GetCell(masterSheet, masterDictionary.item(dictKey), LocaleColumns.controlType)
        SetCell sheetToSync, row, LocaleColumns.sortOrder, masterDictionary.item(dictKey)
        SetCell sheetToSync, row, LocaleColumns.messages, "Requires translation"
        row = row + 1
    Next
    
    Debug.Print sheetToSync & " has been synced to " & masterSheet

End Sub

