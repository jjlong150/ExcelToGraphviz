Attribute VB_Name = "modUtilityADODBDiagnosticLogger"
' Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved

Option Explicit

Private loggingEnabled As Boolean
Private Const LOG_FILE_NAME As String = "Relationship Visualizer ADO Log.txt"

' ===========================
' Public entry point for logging
' ===========================

' Enable or disable logging at runtime
Public Sub SetLoggingEnabled(ByVal enabled As Boolean)
    loggingEnabled = enabled
End Sub

' Query current logging state
Public Function IsLoggingEnabled() As Boolean
    IsLoggingEnabled = loggingEnabled
End Function

Public Sub LogDiagnostic(ByVal message As String, _
                            Optional ByVal errorNumber As Long = 0, _
                            Optional ByVal errorCategory As String = vbNullString, _
                            Optional ByVal attempt As Long = 0, _
                            Optional ByVal sql As String = vbNullString, _
                            Optional ByVal includeFingerprint As Boolean = False)
    If Not loggingEnabled Then Exit Sub
    On Error Resume Next
    
    Dim fso As Object
    Dim ts As Object
    Dim logPath As String
    
    logPath = GetLogFilePath()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logPath, 8, True) ' ForAppending
    
    ts.WriteLine String(80, "-")
    ts.WriteLine vbCrLf & format$(Now, "yyyy-mm-dd hh:nn:ss") & "   : " & message
    
    If attempt > 0 Then
        ts.WriteLine "  Attempt Number      : " & attempt
    End If
    
    If errorNumber <> 0 Then
        ts.WriteLine "  Error Number        : " & errorNumber
    End If
    
    If Trim(errorCategory) <> vbNullString Then
        ts.WriteLine "  Error Category      : " & errorCategory
    End If
    
    If sql <> vbNullString Then
        ts.WriteLine "  SQL Statement       : " & vbCrLf
        ts.WriteLine sql & vbCrLf
    End If
    
    If includeFingerprint Then
        ts.WriteLine GetEnvironmentFingerprint()
    End If
    
    ts.Close
End Sub

' ===========================
' Environment Fingerprint
' ===========================

' Builds a structured environment fingerprint
Private Function GetEnvironmentFingerprint() As String
    Dim s As String
    's = vbCrLf & "Environment" & vbCrLf
    
    ' User + Machine
    s = s & "  User Name           : " & Environ$("USERNAME") & vbCrLf
    s = s & "  Machine Name        : " & Environ$("COMPUTERNAME") & vbCrLf
    
    ' Hardware and OS
    s = s & "  OS Version          : " & Environ$("OS") & vbCrLf
    s = s & "  Processor Count     : " & Environ$("NUMBER_OF_PROCESSORS") & vbCrLf
    s = s & "  Processor Arch.     : " & Environ$("PROCESSOR_ARCHITECTURE") & vbCrLf
    
    ' Excel + VBA
    s = s & "  Application Name    : " & Application.name & vbCrLf
    s = s & "  Applicaton Version  : " & Application.version & vbCrLf
    s = s & "  Applicaton OS       : " & Application.OperatingSystem & vbCrLf
    
    ' Is MS Office 32 or 64 bit?
    s = s & "  32/64 bit           : " & GetOfficeBitness() & vbCrLf
    
#If VBA7 Then
    s = s & "  VBA Version         : VBA7" & vbCrLf
#Else
    s = s & "  VBA7                : VBA6 or earlier" & vbCrLf
#End If

    ' Security + Trust
    s = s & "  Macro Security      : " & GetMacroSecurityMode() & vbCrLf
    s = s & "  Trusted Location    : " & IsWorkbookInTrustedLocation(ThisWorkbook) & vbCrLf
    
    ' OS
    s = s & "  Locale              : " & Application.International(xlCountrySetting) & vbCrLf
    s = s & "  Time Zone           : " & format$(Now, "zzz") & vbCrLf
    
    ' Workbook context
    On Error Resume Next
    s = s & "  Workbook Path       : " & ThisWorkbook.FullName & vbCrLf
    s = s & "  AutoRecover Enabled : " & Application.AutoRecover.enabled & vbCrLf
    s = s & "  On OneDrive         : " & IsWorkbookOnOneDrive(ThisWorkbook) & vbCrLf
    
    ' Provider info
    s = s & "  Provider            : " & DetectBestExcelProvider() & vbCrLf
    
    GetEnvironmentFingerprint = s
End Function

' Detects if workbook is on OneDrive/SharePoint
Private Function IsWorkbookOnOneDrive(wb As Workbook) As Boolean
    On Error Resume Next
    Dim p As String
    p = LCase$(wb.path)
    
    If InStr(p, "onedrive") > 0 Then
        IsWorkbookOnOneDrive = True
    ElseIf InStr(p, "sharepoint") > 0 Then
        IsWorkbookOnOneDrive = True
    Else
        IsWorkbookOnOneDrive = False
    End If
End Function

' Determines log file location
Private Function GetLogFilePath() As String
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.path
    If Len(basePath) = 0 Then basePath = CurDir$
    GetLogFilePath = basePath & Application.pathSeparator & LOG_FILE_NAME
End Function

Private Function GetOfficeBitness() As String
#If Win64 Then
    GetOfficeBitness = "64-bit"
#Else
    GetOfficeBitness = "32-bit"
#End If
End Function

Private Function GetExcelBuildNumber() As String
    On Error Resume Next
    GetExcelBuildNumber = Application.Build
End Function

Private Function GetExcelFullVersion() As String
    On Error Resume Next
    GetExcelFullVersion = Application.version & " (Build " & Application.Build & ")"
End Function

Private Function IsWorkbookInTrustedLocation(wb As Workbook) As Boolean
    On Error Resume Next
    
    Dim trust As Object
    Dim loc As Object
    Dim p As String
    
    p = LCase$(wb.path)
    
    Set trust = Application.FileDialog(msoFileDialogOpen).Application _
                    .CommandBars("File").Controls("Options") ' dummy to force load
    
    ' Loop trusted locations
    For Each loc In Application.TrustCenter.TrustedLocations
        If Len(loc.path) > 0 Then
            If InStr(LCase$(p), LCase$(loc.path)) = 1 Then
                IsWorkbookInTrustedLocation = True
                Exit Function
            End If
        End If
    Next loc
    
    IsWorkbookInTrustedLocation = False
End Function

Private Function GetMacroSecurityMode() As String
    On Error Resume Next
    
    Dim sec As Long
    sec = Application.AutomationSecurity
    
    Select Case sec
        Case 1: GetMacroSecurityMode = "Low"
        Case 2: GetMacroSecurityMode = "Medium"
        Case 3: GetMacroSecurityMode = "High"
        Case 4: GetMacroSecurityMode = "Very High"
        Case Else: GetMacroSecurityMode = "Unknown"
    End Select
End Function

