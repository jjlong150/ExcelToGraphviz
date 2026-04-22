Attribute VB_Name = "modUtilityDateTime"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityDateTime
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Date & Time
'
' ROLE:
'   Minimal date-time formatting helpers for generating standardized timestamps
'   used across logging, diagnostics, file naming, and status reporting.
'
' RESPONSIBILITIES:
'   - Provide ISO-like date formatting (yyyy-mm-dd).
'   - Provide time formatting suitable for filenames (hh.mm.ss).
'   - Provide combined date-time strings for lightweight timestamping.
'
' ARCHITECTURAL NOTES:
'   - Uses VBA's locale-aware Format function.
'   - Produces stable, sortable output for logs and filenames.
'   - No external dependencies; safe for both Windows and macOS.
'
' USAGE:
'   - Used by diagnostic logging, console output, and file-naming utilities.
'   - Suitable for lightweight timestamp generation where full locale
'     formatting is not required.
'
' RELATED WIKI PAGES:
'   - Diagnostics & Logging Conventions
'   - File Naming & Timestamping Guidelines
' =============================================================================

Option Explicit

Public Function GetDateTime() As String
    GetDateTime = format(date, "yyyy-mm-dd") & " " & format(time, "hh.mm.ss")
End Function

Public Function GetTime() As String
    GetTime = format(time, "hh.mm.ss")
End Function

Public Function GetDate() As String
    GetDate = format(date, "yyyy-mm-dd")
End Function

