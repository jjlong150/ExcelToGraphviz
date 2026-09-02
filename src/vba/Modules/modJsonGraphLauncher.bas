Attribute VB_Name = "modJsonGraphLauncher"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modJsonGraphLauncher
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Visualization / Browser Launch Pipeline
'
' ROLE:
'   Provides the complete launch mechanism for the JSON-based Knowledge Graph
'   viewer. Emits viewer.html and data.js into a writable working folder,
'   applies localization, manages shared-tab vs new-tab behavior, and opens
'   the viewer in the system browser. Serves as the bridge between the JSON
'   synthesis pipeline (modCreateJson) and the external visualization layer.
'
' RESPONSIBILITIES:
'   o Viewer Launch & Session Management:
'       - Emit viewer.html and data.js into a verified writable folder.
'       - Manage shared-tab behavior using module-state flags.
'       - Reopen previously launched viewers when requested.
'
'   o HTML Template Reconstruction:
'       - Load viewer.html from the hidden ViewerAssets worksheet.
'       - Apply localized UI labels via ApplyJsonViewerLabels.
'       - Support multiple viewer variants (shared vs new tab).
'
'   o JSON Payload Emission:
'       - Write data.js containing the escaped JSON payload.
'       - Use EscapeNonAscii to ensure byte-consistent output across platforms.
'
'   o Filesystem & Cross-Platform Handling:
'       - Resolve writable folders using GetHtmlTempDir and fallback logic.
'       - Verify actual write access via CanWriteTo (critical for macOS sandbox).
'       - Normalize path separators and avoid platform-specific assumptions.
'
'   o Utility Routines:
'       - Provide localized label substitution.
'       - Provide viewer relaunch and session-state queries.
'
' MODULE STATE:
'   mViewerOpened As Boolean
'       - Tracks whether a shared-tab JSON viewer has been opened during the
'         current session. Prevents repeated browser launches when the viewer
'         is already polling data.js for updates.
'
'   mLastHtmlPath As String
'       - Stores the full path to the most recently emitted viewer.html file.
'         Used by ReopenGraphViewer and SharedBrowserWasOpened to provide
'         session continuity and manual relaunch support.
'
'   JsonBrowserTab (Enum)
'       - Defines viewer-launch modes:
'           o SharedTab = 1  — Reuse a single browser tab that auto-refreshes
'                              via data.js polling.
'           o NewTab    = 2  — Always open a fresh browser tab.
'       - Maps directly to ViewerAssets column selection for template loading.
'
' INTERACTIONS:
'   o modCreateJson:
'       - Supplies the JSON payload for viewer emission.
'
'   o ViewerAssets Worksheet:
'       - Stores the HTML viewer template and alternate variants.
'
'   o modCreateCommon:
'       - Provides WriteTextFile, EscapeNonAscii, and other shared helpers.
'
'   o Utility Modules:
'       - modUtilityFileSystem: temp directories, path resolution.
'       - modUtilityString: label substitution and affix stripping.
'
' CROSS-PLATFORM NOTES:
'   o Windows:
'       - Standard temp-directory behavior; direct write access.
'
'   o macOS:
'       - Sandbox may allow enumeration but block writes; CanWriteTo ensures
'         reliable detection.
'       - FollowHyperlink works identically for viewer launch.
'
' ERROR HANDLING:
'   o WriteTextFile emits diagnostic messages on failure.
'   o Viewer launch gracefully aborts when writable folder cannot be resolved.
'   o Shared-tab logic avoids repeated browser launches during a session.
'
' RELATED WIKI PAGES:
'   o Visualization Pipeline
'   o Browser Rendering Model
'   o ViewerAssets & Localization
'   o Temp Folder Pipeline
'   o JSON Output & Serialization
' =============================================================================

Option Explicit

Private mViewerOpened As Boolean
Private mLastHtmlPath As String

Private Enum JsonBrowserTab
    SharedTab = 1   ' Use the html that polls, looking for changes
    NewTab = 2      ' Open a fresh browser tab every time
End Enum

' ==========================================================================
' ROUTINE: ShowKnowledgeGraph
'
' PURPOSE:
'   Renders a JSON-based knowledge graph in the user's default system browser.
'   Writes a static HTML viewer template and a dynamic data.js payload into a
'   writable temporary folder, then opens viewer.html via FollowHyperlink.
'   Designed for maximum portability: no ActiveX, no UserForms, no registry
'   dependencies, and no trust-center requirements. Works identically on
'   Windows and macOS.
'
' FUNCTIONAL WORKFLOW:
'   1. WORKING FOLDER RESOLUTION:
'        - Obtains a writable temporary directory via GetHtmlTempDir.
'        - Constructs full paths for:
'             o viewer.html  (static UI/logic)
'             o data.js      (dynamic JSON payload)
'
'   2. VIEWER TEMPLATE SELECTION:
'        - Reads the JsonViewer setting from SettingsSheet.
'        - Maps "new" or "shared" to the correct ViewerAssets column using
'          JsonBrowserTab enumeration.
'
'   3. HTML TEMPLATE LOADING:
'        - Loads viewer.html content from the hidden ViewerAssets sheet via
'          LoadViewerHtmlFromSheet.
'        - Applies label substitutions via ApplyJsonViewerLabels.
'        - Writes the final HTML to viewer.html.
'
'   4. JSON PAYLOAD EMISSION:
'        - Writes data.js containing:
'             var GRAPH_JSON = {escaped JSON};
'        - EscapeNonAscii ensures safe embedding of Unicode characters.
'
'   5. SHARED-TAB BEHAVIOR:
'        - When tabStyle = "shared":
'             o Opens viewer.html only once per session.
'             o viewer.html polls data.js (see pollDataJs in the HTML) and
'               re-renders automatically when the file changes.
'             o Subsequent calls update data.js without opening new tabs.
'
'   6. NEW-TAB BEHAVIOR:
'        - When tabStyle = "new":
'             o Always opens a fresh browser tab via FollowHyperlink.
'
'   7. OUTPUT:
'        - Stores the last viewer path in mLastHtmlPath.
'        - Browser displays the knowledge graph using the latest JSON payload.
'
' TECHNICAL NOTES:
'   - ViewerAssets is a hidden worksheet containing the full viewer.html
'     template in cell A1 (or alternate columns). This keeps the workbook
'     self-contained with no external companion files.
'   - viewer.html is static; data.js is the only file rewritten per call.
'   - DeepWiki Context: Implements the JSON-viewer pipeline described in the
'     "Visualization", "Browser Rendering", and "Serialization" sections.
' ==========================================================================
Public Sub ShowKnowledgeGraph(jsonText As String)
    Dim folderPath As String
    Dim htmlPath As String
    Dim dataPath As String

    folderPath = GetHtmlTempDir()
    htmlPath = folderPath & Application.pathSeparator & "viewer.html"
    dataPath = folderPath & Application.pathSeparator & "data.js"

    Dim tabStyle As String
    tabStyle = SettingsSheet.Range(SETTINGS_JSON_VIEWER).Value2
    
    Dim htmlColumn As Long
    
    Select Case LCase$(tabStyle)
        Case "new"
            htmlColumn = JsonBrowserTab.NewTab
    
        Case "shared"
            htmlColumn = JsonBrowserTab.SharedTab
    
        Case Else
            htmlColumn = JsonBrowserTab.NewTab   ' safe fallback
    End Select

    ' The HTML content is stored in a hidden worksheet. Two versions are present.
    If Not WriteTextFile(htmlPath, ApplyJsonViewerLabels(LoadViewerHtmlFromSheet(htmlColumn))) Then Exit Sub
    If Not WriteTextFile(dataPath, "var GRAPH_JSON = " & EscapeNonAscii(jsonText) & ";") Then Exit Sub
    
    mLastHtmlPath = htmlPath

    If tabStyle = "shared" Then
        ' Only open a new tab the first time this session. viewer.html polls
        ' data.js on its own (see pollDataJs in the HTML) and re-renders when it
        ' changes, so an already-open tab picks up this update within a couple
        ' of seconds without VBA needing to touch it again.
        If Not mViewerOpened Then
            ActiveWorkbook.FollowHyperlink Address:=htmlPath
            mViewerOpened = True
        End If
    Else
        ' Set NewWindow to True to request a new instance/tab
        ActiveWorkbook.FollowHyperlink Address:=htmlPath, NewWindow:=True
    End If
    
End Sub

' ==========================================================================
' ROUTINE: ApplyJsonViewerLabels
'
' PURPOSE:
'   Applies localized UI labels to the JSON viewer HTML template. Replaces
'   placeholder tokens (e.g., "JsonViewerTitle", "JsonViewerSearchInput") with
'   their corresponding localized strings retrieved via GetLabel. Produces a
'   fully substituted viewer.html ready for emission into the working folder.
'
' FUNCTIONAL WORKFLOW:
'   1. PLACEHOLDER KEY LIST:
'        - Defines a fixed array of all supported JSON viewer label tokens:
'             o Titles
'             o Button captions
'             o Status messages
'             o Mode indicators
'             o Tree/raw view labels
'        - Ensures consistent substitution across all viewer variants.
'
'   2. TOKEN SUBSTITUTION:
'        - Iterates through each key in the array.
'        - Performs a single, case-insensitive replacement:
'             htmlViewer = Replace(htmlViewer, key, GetLabel(key), 1, 1, vbTextCompare)
'        - Guarantees one replacement per key, preventing accidental multiple
'          substitutions inside the template.
'
'   3. OUTPUT:
'        - Returns the fully substituted HTML string.
'        - Caller writes the result to viewer.html before launching the browser.
'
' TECHNICAL NOTES:
'   - GetLabel resolves locale-specific text based on the active language sheet.
'   - Replacement is intentionally limited to the first occurrence of each key
'     to preserve template structure and avoid unintended cascading changes.
'   - DeepWiki Context: Implements the JSON-viewer localization rules described
'     in the "Localization", "ViewerAssets", and "Browser Rendering" sections.
' ==========================================================================
Public Function ApplyJsonViewerLabels(htmlViewer As String) As String
    Dim keys As Variant
    Dim k As Variant

    keys = Array( _
        "JsonViewerTitle", _
        "JsonViewerSearchInput", _
        "JsonViewerSearchButton", _
        "JsonViewerPrevButton", _
        "JsonViewerNextButton", _
        "JsonViewerClearButton", _
        "JsonViewerCopyRawButton", _
        "JsonViewerCopyPrettyButton", _
        "JsonViewerSaveRawButton", _
        "JsonViewerSavePrettyButton", _
        "JsonViewerRawButton", _
        "JsonViewerTreeButton", _
        "JsonViewerReloadButton", _
        "JsonViewerReloadDataTitle", _
        "JsonViewerCouldParseJson", _
        "JsonViewerShowingRawText", _
        "JsonViewerCharactersLabel", _
        "JsonViewerTokensLabel", _
        "JsonViewerLinesLabel", _
        "JsonViewerViewLabel", _
        "JsonViewerNodesLabel", _
        "JsonViewerEdgesLabel", _
        "JsonViewerTreeNotAvailable" _
    )

    For Each k In keys
        htmlViewer = replace(htmlViewer, k, HtmlEncode(GetLabel(k)), 1, 1, vbTextCompare)
    Next k

    ApplyJsonViewerLabels = htmlViewer
End Function

' ==========================================================================
' ROUTINE: ReopenGraphViewer
'
' PURPOSE:
'   Reopens the JSON graph viewer in the system browser when the user has
'   previously closed the tab. ShowKnowledgeGraph intentionally avoids
'   reopening an already-opened viewer during the same session, since VBA
'   cannot detect whether the browser tab was closed. This routine provides a
'   manual recovery path by relaunching the last viewer.html location.
'
' FUNCTIONAL WORKFLOW:
'   1. LAST-VIEWER CHECK:
'        - Verifies that mLastHtmlPath contains a previously emitted viewer
'          path.
'        - If empty, informs the user that no viewer has been shown yet and
'          exits cleanly.
'
'   2. BROWSER RELAUNCH:
'        - Uses FollowHyperlink to open the stored viewer.html path in the
'          default system browser.
'        - Works identically on Windows and macOS.
'
'   3. SESSION STATE UPDATE:
'        - Sets mViewerOpened = True to ensure consistent shared-tab behavior
'          for subsequent calls to ShowKnowledgeGraph.
'
' TECHNICAL NOTES:
'   - This routine does not regenerate viewer.html or data.js; it simply
'     reopens the last viewer instance.
'   - mLastHtmlPath is maintained by ShowKnowledgeGraph each time a viewer is
'     emitted.
'   - DeepWiki Context: Implements the viewer-recovery rules described in the
'     "Visualization", "Browser Rendering", and "Session State" sections.
' ==========================================================================
Public Sub ReopenGraphViewer()
    If Len(mLastHtmlPath) = 0 Then
        MsgBox "Nothing has been shown yet - call ShowGraph first.", vbInformation
        Exit Sub
    End If
    ActiveWorkbook.FollowHyperlink mLastHtmlPath
    mViewerOpened = True
End Sub

' ==========================================================================
' ROUTINE: SharedBrowserWasOpened
'
' PURPOSE:
'   Reports whether a shared JSON-viewer browser tab has been opened during
'   the current session. Used by callers that need to know whether the viewer
'   is already active and capable of auto-refreshing via its internal
'   data.js polling loop.
'
' FUNCTIONAL WORKFLOW:
'   1. SESSION-STATE CHECK:
'        - Evaluates mLastHtmlPath, which is set by ShowKnowledgeGraph each
'          time viewer.html is emitted.
'        - When mLastHtmlPath = "", no viewer has been launched yet.
'        - When mLastHtmlPath contains a valid path, at least one viewer tab
'          has been opened this session.
'
'   2. BOOLEAN RETURN:
'        - Returns True when a viewer has been opened.
'        - Returns False otherwise.
'
' TECHNICAL NOTES:
'   - This routine does not verify whether the browser tab is *currently*
'     open; VBA cannot detect tab closure. It only reports whether a viewer
'     has been launched at least once.
'   - Used by ShowKnowledgeGraph to decide whether to reuse a shared tab or
'     open a new one.
'   - DeepWiki Context: Supports the viewer-session rules described in the
'     "Browser Rendering", "Session State", and "Shared Tab Behavior" sections.
' ==========================================================================
Public Function SharedBrowserWasOpened()
    If Len(mLastHtmlPath) = 0 Then
        SharedBrowserWasOpened = False
    Else
        SharedBrowserWasOpened = True
    End If
End Function

' ==========================================================================
' ROUTINE: GetWritableFolder
'
' PURPOSE:
'   Determines a reliable, writable folder for emitting viewer.html and
'   data.js. Prefers the operating system's temporary directory, but verifies
'   actual write access (important on macOS, where sandboxing can silently
'   block writes). Falls back to the workbook's own folder when necessary,
'   then creates a dedicated subfolder named "JsonGraphViewer".
'
' FUNCTIONAL WORKFLOW:
'   1. CANDIDATE FOLDER SELECTION:
'        - Attempts to use GetHtmlTempDir as the preferred working location.
'        - If GetHtmlTempDir returns an empty string, defaults to
'          ThisWorkbook.Path, which Excel is guaranteed to have access to.
'
'   2. PATH NORMALIZATION:
'        - Removes any trailing path separator.
'        - Appends "\JsonGraphViewer" (or "/" on macOS) to create a dedicated
'          working subfolder.
'
'   3. FOLDER CREATION:
'        - Attempts MkDir on the candidate folder.
'        - Errors are ignored intentionally; the folder may already exist.
'
'   4. WRITE-ACCESS VERIFICATION:
'        - Calls CanWriteTo(candidate) to confirm the folder is genuinely
'          writable.
'        - If write access fails (common in macOS sandboxed temp locations),
'          falls back to ThisWorkbook.Path.
'
'   5. OUTPUT:
'        - Returns the final writable folder path.
'
' TECHNICAL NOTES:
'   - macOS sandboxing can allow folder enumeration but block file creation;
'     CanWriteTo provides a definitive check.
'   - The JsonGraphViewer subfolder isolates viewer artifacts from other
'     workbook files and avoids cluttering the root directory.
'   - DeepWiki Context: Implements the viewer-filesystem rules described in
'     the "Browser Rendering", "Temp Folder Pipeline", and "Cross-Platform
'     Compatibility" sections.
' ==========================================================================
Private Function GetWritableFolder() As String
    Dim candidate As String
    Dim sep As String
    sep = Application.pathSeparator

    candidate = GetHtmlTempDir()
    If Len(candidate) = 0 Then candidate = ThisWorkbook.path

    If Right$(candidate, 1) = sep Then candidate = Left$(candidate, Len(candidate) - 1)
    candidate = candidate & sep & "JsonGraphViewer"

    On Error Resume Next
    MkDir candidate
    On Error GoTo 0

    If Not CanWriteTo(candidate) Then
        candidate = ThisWorkbook.path
    End If

    GetWritableFolder = candidate
End Function

' ==========================================================================
' ROUTINE: CanWriteTo
'
' PURPOSE:
'   Determines whether a folder is genuinely writable by attempting to create
'   and delete a temporary file inside it. Provides a definitive, cross-platform
'   check that catches macOS sandbox restrictions, silent permission denials,
'   and other cases where a directory appears accessible but cannot accept
'   new files.
'
' FUNCTIONAL WORKFLOW:
'   1. TEMPORARY FILE PATH:
'        - Constructs a test file named "~writetest.tmp" inside folderPath
'          using the correct OS path separator.
'
'   2. WRITE PROBE:
'        - Opens the file for Output using FreeFile.
'        - Writes a single line ("test") to ensure both creation and write
'          operations succeed.
'        - Closes the file handle.
'        - All operations are wrapped in On Error Resume Next to avoid
'          interrupting the caller.
'
'   3. RESULT DETERMINATION:
'        - If Err.Number = 0 after the write attempt, the folder is writable.
'        - If any error occurred, the folder is not writable.
'
'   4. CLEANUP:
'        - When writable, deletes the temporary file via Kill.
'        - Restores normal error handling afterward.
'
'   5. OUTPUT:
'        - Returns True when the folder supports file creation and deletion.
'        - Returns False otherwise.
'
' TECHNICAL NOTES:
'   - This routine tests actual write capability, not just directory existence
'     or enumeration rights.
'   - Essential for macOS, where sandboxing may allow folder visibility but
'     block file creation.
'   - DeepWiki Context: Supports the filesystem-validation rules described in
'     the "Temp Folder Pipeline", "Cross-Platform Compatibility", and
'     "Browser Rendering" sections.
' ==========================================================================
Private Function CanWriteTo(folderPath As String) As Boolean
    Dim testFile As String
    Dim fnum As Integer
    testFile = folderPath & Application.pathSeparator & "~writetest.tmp"

    On Error Resume Next
    Err.Clear
    fnum = FreeFile
    Open testFile For Output As #fnum
    Print #fnum, "test"
    Close #fnum
    CanWriteTo = (Err.number = 0)
    If CanWriteTo Then Kill testFile
    On Error GoTo 0
End Function

' ==========================================================================
' ROUTINE: WriteTextFile
'
' PURPOSE:
'   Writes the specified text content to a file at the given path. Provides a
'   safe, error-tolerant wrapper around VBA's file I/O primitives, returning
'   True on success and False on failure. Emits a diagnostic message when an
'   error occurs, including the error number, description, and target path.
'
' FUNCTIONAL WORKFLOW:
'   1. INITIALIZATION:
'        - Sets the default result to False (failure).
'        - Prepares a free file handle via FreeFile.
'
'   2. FILE WRITE OPERATION:
'        - Opens the file for Output.
'        - Writes the supplied content using Print #.
'        - On success, sets the function result to True.
'
'   3. CLEAN EXIT:
'        - Ensures the file handle is closed if it was successfully allocated.
'        - Suppresses errors during cleanup to avoid cascading failures.
'
'   4. ERROR HANDLING:
'        - Captures any write-related errors.
'        - Emits a formatted diagnostic message via EmitMessage, including:
'             o Err.Number
'             o Err.Description
'             o The target file path
'        - Returns False by falling through to the CleanExit block.
'
'   5. OUTPUT:
'        - Returns True when the file was written successfully.
'        - Returns False when any error occurred.
'
' TECHNICAL NOTES:
'   - Uses Print # rather than Write # to preserve literal text formatting.
'   - Caller is responsible for ensuring the target folder is writable; see
'     CanWriteTo and GetWritableFolder for validation.
'   - DeepWiki Context: Implements the file-emission rules described in the
'     "Temp Folder Pipeline", "Browser Rendering", and "Cross-Platform
'     Compatibility" sections.
' ==========================================================================
Public Function WriteTextFile(path As String, content As String) As Boolean
    Dim fnum As Integer

    WriteTextFile = False          ' default: failure
    On Error GoTo ErrHandler

    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, content

    WriteTextFile = True           ' success

CleanExit:
    On Error Resume Next
    If fnum > 0 Then Close #fnum
    Exit Function

ErrHandler:
    Dim fullMessage As String
    fullMessage = GetMessage("errormsgWriteTextFileError")
    fullMessage = replace(fullMessage, "{Err.number}", CStr(Err.number), 1, 1, vbTextCompare)
    fullMessage = replace(fullMessage, "{Err.description}", Err.Description, 1, 1, vbTextCompare)
    fullMessage = replace(fullMessage, "{filename}", path, 1, 1, vbTextCompare)
    
    EmitMessageSilent fullMessage, esError
    Resume CleanExit               ' keep result = False
End Function

' ==========================================================================
' ROUTINE: LoadViewerHtmlFromSheet
'
' PURPOSE:
'   Reconstitutes the viewer.html template from the hidden ViewerAssets
'   worksheet so the workbook remains a single, self-contained distributable
'   file. Reads sequential cells in the specified column and concatenates
'   their contents into a complete HTML document. Supports multi-cell storage
'   for future versions of the template that may exceed Excel's ~32,000-
'   character limit for a single cell.
'
' FUNCTIONAL WORKFLOW:
'   1. WORKSHEET RESOLUTION:
'        - Retrieves the hidden worksheet named "ViewerAssets".
'        - Selects the column specified by the caller (col), allowing multiple
'          viewer variants (e.g., new-tab vs shared-tab) to coexist.
'
'   2. TEMPLATE RECONSTRUCTION:
'        - Begins reading at row 2 (row 1 reserved for metadata or future use).
'        - Continues reading downward until encountering an empty cell.
'        - Concatenates each non-empty cell's text into a single HTML string.
'
'   3. MULTI-CELL SUPPORT:
'        - Allows the viewer.html template to span A2, A3, A4, … if a future
'          version ever exceeds Excel's per-cell character limit.
'        - Current versions typically fit entirely in A2.
'
'   4. OUTPUT:
'        - Returns the fully reconstructed HTML viewer template as a string.
'        - Caller applies localization (ApplyJsonViewerLabels) and writes the
'          result to viewer.html before launching the browser.
'
' TECHNICAL NOTES:
'   - ViewerAssets is intentionally stored inside the workbook to avoid
'     external companion files that can be lost when the workbook is shared.
'   - The routine performs no HTML validation; it simply reassembles the
'     stored template verbatim.
'   - DeepWiki Context: Implements the viewer-asset rules described in the
'     "ViewerAssets", "Browser Rendering", and "Distribution Model" sections.
' ==========================================================================
Private Function LoadViewerHtmlFromSheet(col As Long) As String
    Dim ws As Worksheet
    Dim i As Long
    Dim s As String
    Set ws = ThisWorkbook.worksheets("ViewerAssets")
    i = 2
    Do While Len(ws.Cells(i, col).value) > 0
        s = s & ws.Cells(i, col).value
        i = i + 1
    Loop
    LoadViewerHtmlFromSheet = s
End Function

' ==========================================================================
' ROUTINE: EscapeNonAscii
'
' PURPOSE:
'   Converts a string into pure ASCII by escaping all characters above code
'   point 126 as JavaScript/JSON Unicode literals (\uXXXX). Ensures that the
'   emitted data.js file is byte-for-byte identical on Windows and macOS,
'   avoiding discrepancies caused by platform-specific default text encodings
'   when writing files via VBA's Print #.
'
' FUNCTIONAL WORKFLOW:
'   1. CHARACTER SCAN:
'        - Iterates through each character in the input string.
'        - Retrieves the Unicode code point using AscW.
'        - Normalizes negative values (surrogate handling) by adding 65536.
'
'   2. ASCII GATE:
'        - Characters with code <= 126 are appended verbatim.
'        - Characters with code > 126 are escaped as:
'             \uXXXX
'          where XXXX is a zero-padded, four-digit hexadecimal value.
'
'   3. STRING ASSEMBLY:
'        - Builds the escaped string incrementally.
'        - Produces a final ASCII-only representation safe for embedding
'          directly inside JavaScript string literals.
'
'   4. OUTPUT:
'        - Returns the fully escaped ASCII string.
'        - Caller writes the result into data.js:
'             var GRAPH_JSON = {escaped JSON};
'
' TECHNICAL NOTES:
'   - Required because VBA's Print # emits platform-dependent encodings:
'        o Windows: typically ANSI/ACP
'        o macOS: UTF-8 or UTF-16 depending on sandbox context
'     Escaping ensures deterministic output regardless of OS.
'
'   - EscapeNonAscii does not modify JSON structure; it only transforms
'     character encoding.
'
'   - DeepWiki Context: Implements the JSON-encoding rules described in the
'     "Serialization", "Browser Rendering", and "Cross-Platform Compatibility"
'     sections.
' ==========================================================================
Private Function EscapeNonAscii(s As String) As String
    Dim i As Long, code As Long
    Dim sb As String
    sb = ""
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536
        If code > 126 Then
            sb = sb & "\u" & Right$("0000" & Hex$(code), 4)
        Else
            sb = sb & Mid$(s, i, 1)
        End If
    Next i
    EscapeNonAscii = sb
End Function

' ==========================================================================
' FUNCTION: HtmlEncode
'
' PURPOSE:
'   Encodes a string for safe inclusion in Graphviz HTML-like labels or other
'   markup contexts. Converts reserved characters to their HTML entities and
'   emits numeric character references for all non-ASCII code points.
'
' FUNCTIONAL WORKFLOW:
'   1. CHARACTER EXTRACTION:
'        - Iterates through each character using Mid$ and AscW.
'        - Normalizes negative AscW values for code points above U+7FFF.
'
'   2. ENTITY ENCODING:
'        - Replaces '&', '<', and '>' with their standard HTML entities.
'        - Emits "&#NNNN;" numeric references for characters above ASCII 127.
'        - Preserves plain ASCII characters as-is.
'
' TECHNICAL NOTES:
'   - Produces output compatible with Graphviz HTML-like label syntax and
'     general HTML/XML contexts.
'   - Numeric references ensure cross-platform consistency when rendering
'     Unicode characters in environments with limited encoding support.
'   - DeepWiki Context: Supports the encoding rules described in "HTML-Like
'     Labels" and "Unicode Handling in Label Rendering".
' ==========================================================================
Public Function HtmlEncode(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, sb As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536   ' AscW returns negative above U+7FFF
        Select Case ch
            Case "&": sb = sb & "&amp;"
            Case "<": sb = sb & "&lt;"
            Case ">": sb = sb & "&gt;"
            Case Else
                If code > 127 Then
                    sb = sb & "&#" & code & ";"
                Else
                    sb = sb & ch
                End If
        End Select
    Next i
    HtmlEncode = sb
End Function
