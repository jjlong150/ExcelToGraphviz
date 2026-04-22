Attribute VB_Name = "modUtilityClipboard"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityClipboard
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Clipboard
'
' ROLE:
'   Windows-only clipboard subsystem providing safe, late-bound access to the
'   Win32 API for copying text to the system clipboard. Supports Ribbon-level
'   "Copy to Clipboard" actions across SQL, SVG, Source, and Styles workflows.
'
' RESPONSIBILITIES:
'   - Expose a stable, late-bound API wrapper for GlobalAlloc, GlobalLock,
'     SetClipboardData, and related Win32 functions.
'   - Provide ClipBoard_SetData for safe text transfer to the Windows clipboard.
'   - Provide Clipboard_Clear for clearing clipboard contents with defensive
'     error handling.
'   - Abstract away 32-bit vs 64-bit pointer differences (VBA7 vs legacy VBA).
'
' ARCHITECTURAL NOTES:
'   - Windows-only subsystem; excluded on macOS via conditional compilation.
'   - Ribbon controls automatically hide clipboard buttons on macOS.
'   - Defensive unlock/close logic prevents memory leaks and clipboard locks.
'   - Integrated with SQL, SVG, Source, and Styles tabs for copy operations.
'
' VERSION NOTES:
'   - v8.0.0: Removed reliance on Internet Explorer ActiveX.
'
' USAGE:
'   - Invoked by Ribbon "Copy" buttons across multiple tabs.
'   - Used by editor pop-ups and worksheet-driven copy actions.
'
' RELATED WIKI PAGES:
'   - Clipboard Operations (Windows)
'   - Ribbon Copy Actions (SQL, SVG, Source)
' =============================================================================

Option Explicit

#If Not Mac Then

' Source: How To Copy Text To Clipboard Using Excel VBA
' https://www.spreadsheet1.com/how-to-copy-strings-to-clipboard-using-excel-vba.html

'Handle 64-bit and 32-bit Office
#If VBA7 Then
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
#Else
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Public Declare Function CloseClipboard Lib "User32" () As Long
    Public Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "User32" () As Long
    Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Public Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND As Long = &H42
Public Const CF_TEXT As Long = 1

Public Function ClipBoard_SetData(ByRef MyString As String) As Boolean
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx
'FIX: https://community.ifs.com/framework-experience-infrastructure-cloud-integration-dev-tools-50/simple-migration-tool-xlsm-error-could-not-unlock-memory-location-copy-aborted-3842

#If Mac Then

#Else
#If VBA7 Then
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr
#Else
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
#End If


    ' Assume the copy will be successful
    ClipBoard_SetData = True

    'Allocate moveable global memory
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

    'Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    'Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    'Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        Debug.Print "ClipBoard_SetData - Could not unlock memory location. Copy aborted."
        ClipBoard_SetData = False
        GoTo CouldNotCloseClipboardExit
    End If

    'Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        Debug.Print "ClipBoard_SetData - Could not open the Clipboard. Copy aborted."
        ClipBoard_SetData = False
        Exit Function
    End If

    'Clear the Clipboard.
    Dim x As Long
    x = EmptyClipboard()

    'Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

CouldNotCloseClipboardExit:
    If CloseClipboard() = 0 Then
        Debug.Print "ClipBoard_SetData - Could not close Clipboard."
        ClipBoard_SetData = False
    End If
    
#End If
End Function

Public Sub Clipboard_Clear()

    On Error GoTo ErrorHandler_
    
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
    Exit Sub
ErrorHandler_:
    EmitMessage "Error: " & Err.Description, buttons:=vbCritical
End Sub


#End If

