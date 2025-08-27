VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2025 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Workbook")

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
#If Not Mac Then
    ' Release any connections which may have been cached in the connection pool.
    CleanupConnectionPool
#End If
End Sub

Private Sub Workbook_Open()
    Application.Calculation = xlCalculationAutomatic
    
    ' Establish the temporary directory once so we don't make frequent system level calls to retrieve
    ' environment variable values which slow things down on Windows PCs.
    SetTempDirectory
    
    ' Cache the username (retrieving it causes it to be cached).
    If GetUsername() = vbNullString Then
        MsgBox GetMessage("msgboxUnableToGetUsername"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If

    ' Check that Graphviz is installed
#If Not Mac Then
    Dim gvPath As String
    gvPath = Trim$(SettingsSheet.Range(SETTINGS_GV_PATH))
    If gvPath = vbNullString Then
        ' Locate the graphing program on the path
        Dim engine As String
        engine = FindFileOnPath("dot.exe")
        If engine = vbNullString Then
            AlertGraphvizNotFound "dot"
        End If
    Else
        ' Fix path if it does not end with a path separator
        If Not EndsWith(gvPath, Application.pathSeparator) Then
            gvPath = gvPath & Application.pathSeparator
            SettingsSheet.Range(SETTINGS_GV_PATH).value = gvPath
        End If
        
        ' See if the full path to dot.exe exists
        gvPath = gvPath & "dot.exe"
        If Not FileExists(gvPath) Then
            MsgBox replace(GetMessage("msgboxInvalidDotPath"), "{gvPath}", gvPath), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        End If
    End If
#End If

#If Mac Then
    ' Establish the version of "ExcelToGraphviz.applescript" which this spreadsheet requires.
    Dim currentVersion As Long
    currentVersion = 3
    
    ' Check that AppleScriptTask scripts are present, and retrieve the version number
    Dim scriptVersion As String
    scriptVersion = RunAppleScriptTask("getVersion", vbNullString)
    
    If scriptVersion = vbNullString Then
        MsgBox GetMessage("msgboxScriptsNotFound"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        SetScriptVersion 0
    Else
        ' Save the version number for future reference. AppleScriptTask has slow performance, so
        ' we want to minimize how often we call it.
        Dim yourVersion As Long
        yourVersion = CInt(scriptVersion)
        SetScriptVersion yourVersion
        
        ' Give a warning if an old version of "ExcelToGraphviz.applescript" is installed.
        If yourVersion < currentVersion Then
            Dim msg As String
            msg = GetMessage("msgboxVersionMismatch")
            msg = replace(msg, "{currentVersion}", CStr(currentVersion))
            msg = replace(msg, "{yourVersion}", CStr(yourVersion))
            MsgBox msg, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        End If
    End If

    ' Check that you are running an version of Excel with VBA support
    Dim requiredMajor As Long: requiredMajor = 16
    Dim requiredMinor As Long: requiredMinor = 40
    
    Dim actualVersion As String
    Dim actualMajor As Long, actualMinor As Long
    Dim versionParts As Variant
    
    actualVersion = Application.version
    versionParts = split(actualVersion, ".")
    
    If UBound(versionParts) >= 1 Then
        actualMajor = CLng(versionParts(0))
        actualMinor = CLng(versionParts(1))
        
        If (actualMajor < requiredMajor) Or _
           (actualMajor = requiredMajor And actualMinor < requiredMinor) Then
            msg = GetMessage("msgboxIncompatibleExcelVersion")
            msg = replace(msg, "{requiredExcelVersion}", requiredMajor & "." & requiredMinor)
            msg = replace(msg, "{yourExcelVersion}", actualVersion)
            MsgBox msg, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        End If
    End If
#End If

    ' Create directories where images of colors and fonts can be cached
    ' This directory is created when the spreadsheet is launched so we don't
    ' waste cycles attempting to create it for every ribbon image of a color or font.
    CreateColorImageDir
    CreateFontImageDir
    
    ' Sync the worksheet visibility with the saved values (in case someone hides
    ' worksheets manually).
    ShowOrHideWorksheets
End Sub

