VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Workbook")

Option Explicit

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

    ' Check that AppleScriptTask scripts are present
#If Mac Then
    Dim msg As String
    Dim applescriptResult As String
    Dim currentVersion As String
    Dim yourVersion As String
    currentVersion = "1.0"
    
    yourVersion = RunAppleScriptTask("getVersion", vbNullString)
    
    If yourVersion = vbNullString Then
        MsgBox GetMessage("msgboxScriptsNotFound"), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    ElseIf yourVersion < currentVersion Then
        msg = GetMessage("msgboxVersionMismatch")
        msg = replace(msg, "{currentVersion}", currentVersion)
        msg = replace(msg, "{yourVersion}", yourVersion)
        MsgBox msg, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If

    Dim requiredExcelVersion As String
    requiredExcelVersion = "16.40"
    
    If Application.version < requiredExcelVersion Then
        msg = GetMessage("msgboxIncompatibleExcelVersion")
        msg = replace(msg, "{requiredExcelVersion}", requiredExcelVersion)
        msg = replace(msg, "{yourExcelVersion}", Application.version)
        MsgBox msg, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
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

