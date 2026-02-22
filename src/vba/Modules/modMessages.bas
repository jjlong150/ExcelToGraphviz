Attribute VB_Name = "modMessages"
Option Explicit

' ============================================================
' Error Reporting Module
' Centralized pipeline for console, status bar, and message box
' Includes:
'   - Optional title
'   - Optional buttons
'   - Silent mode
'   - Severity enum (future-proofing)
'   - Settings gatekeeper
'   - Test harness
' ============================================================

' -----------------------------
' Severity Levels (extensible)
' -----------------------------
Public Enum ErrorSeverity
    esInfo = 0
    esWarning = 1
    esError = 2
    esCritical = 3
End Enum

' -----------------------------
' Public API
' -----------------------------
Public Sub EmitMessage( _
    errorMessage As String, _
    Optional title As String = vbNullString, _
    Optional severity As ErrorSeverity = esError, _
    Optional buttons As VbMsgBoxStyle = vbOKOnly _
)
    ' Core pipeline
    SendToConsole errorMessage, severity
    SendToStatusBar errorMessage, severity
    SendToMessageBox errorMessage, title, severity, buttons
End Sub

' Silent version (no message box)
Public Sub EmitMessageSilent( _
    errorMessage As String, _
    Optional severity As ErrorSeverity = esError _
)
    SendToConsole errorMessage, severity
    SendToStatusBar errorMessage, severity
End Sub

' -----------------------------
' Settings Gatekeeper
' -----------------------------
Private Function IsEnabled(settingName As String) As Boolean
    On Error Resume Next
    IsEnabled = (LCase$(SettingsSheet.Range(settingName).value) <> TOGGLE_NO)
End Function

' -----------------------------
' Output Channels
' -----------------------------
Private Sub SendToConsole(errorMessage As String, severity As ErrorSeverity)
    If Not IsEnabled(SETTINGS_ERROR_TO_CONSOLE) Then Exit Sub

    On Error Resume Next
    LogToConsoleWorksheet FormatConsoleMessage(errorMessage, severity)
End Sub

Private Sub SendToStatusBar(errorMessage As String, severity As ErrorSeverity)
    If Not IsEnabled(SETTINGS_ERROR_TO_STATUS_BAR) Then Exit Sub

    On Error Resume Next
    Application.StatusBar = FormatStatusBarMessage(errorMessage, severity)
End Sub

Private Sub SendToMessageBox( _
    errorMessage As String, _
    Optional title As String = vbNullString, _
    Optional severity As ErrorSeverity = esError, _
    Optional buttons As VbMsgBoxStyle = vbOKOnly _
)
    If Not IsEnabled(SETTINGS_ERROR_TO_MESSAGE_BOX) Then Exit Sub

    On Error Resume Next
    
    ' Format a title for the message box
    Dim msgBoxTitle As String
    If Trim$(title) = vbNullString Then
        msgBoxTitle = GetLabel("msgboxProductTitle")
    Else
        msgBoxTitle = title
    End If
    
    ' Pop-up the message
    MsgBox FormatMessageBoxText(errorMessage, severity), buttons, msgBoxTitle
End Sub

' -----------------------------
' Formatting Helpers
' -----------------------------
Private Function FormatConsoleMessage(msg As String, sev As ErrorSeverity) As String
    ' Include the Locale-aware date and time
    Dim timestamp As String
    timestamp = format$(Now, "")

    ' Compose the final message
    FormatConsoleMessage = "[" & timestamp & "] " & _
                           SeverityLabel(sev) & _
                           NormalizeMessage(msg)
End Function

Private Function FormatStatusBarMessage(msg As String, sev As ErrorSeverity) As String
    FormatStatusBarMessage = NormalizeMessage(msg)
End Function

Private Function FormatMessageBoxText(msg As String, sev As ErrorSeverity) As String
    FormatMessageBoxText = msg
End Function

Private Function SeverityLabel(sev As ErrorSeverity) As String
    Select Case sev
        Case esInfo:     SeverityLabel = "Info"
        Case esWarning:  SeverityLabel = "Warning"
        Case esError:    SeverityLabel = "Error"
        Case esCritical: SeverityLabel = "Critical"
        Case Else:       SeverityLabel = "Unknown"
    End Select
    SeverityLabel = "[" & SeverityLabel & "]: "
End Function

Public Function NormalizeMessage(message As String) As String
    Dim cleaned As String

    cleaned = message

    ' Replace all newline variants with a single separator
    cleaned = replace(cleaned, vbCrLf, " ")
    cleaned = replace(cleaned, vbCr, " ")
    cleaned = replace(cleaned, vbLf, " ")

    ' Optional: collapse tabs as well
    cleaned = replace(cleaned, vbTab, " ")

    ' Optional: collapse multiple separators into one
    Do While InStr(cleaned, "  ") > 0
        cleaned = replace(cleaned, "  ", " ")
    Loop

    NormalizeMessage = cleaned
End Function

' -----------------------------
' Test Harness
' -----------------------------
Public Sub TestEmitMessage()
    ' Clear the status bar
    Application.StatusBar = False
    
    ' Show what the function toggles are set to
    Debug.Print "Error to Console     = " & SettingsSheet.Range(SETTINGS_ERROR_TO_CONSOLE).value
    Debug.Print "Error to Status Bar  = " & SettingsSheet.Range(SETTINGS_ERROR_TO_STATUS_BAR).value
    Debug.Print "Error to Message Box = " & SettingsSheet.Range(SETTINGS_ERROR_TO_MESSAGE_BOX).value
    Debug.Print String(40, "-")

    ' --- Basic cases ---
    EmitMessage "Hello world!"
    EmitMessage "Hello world!", "Go Blue!"

    ' --- Severity cases ---
    EmitMessage "Something is wrong", "Warning", esWarning
    EmitMessageSilent "Silent background error"

    ' --- Buttons cases ---
    EmitMessage "Choose wisely", "Decision Required", esInfo, vbYesNo
    EmitMessage "Proceed with caution", "Critical Step", esCritical, vbRetryCancel
    EmitMessage "Operation complete", "Success", esInfo, vbOKOnly
End Sub
