Attribute VB_Name = "modAppleScript"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Apple")
'@IgnoreModule EmptyModule

Option Explicit

#If Mac Then

Private Const APPLE_SCRIPT_FILE = "ExcelToGraphviz.applescript"

' If an AppleScript command returns a non-zero value, an error is thrown. This code is to ensure
' any use of AppleScriptTask is wrapped with error handling, and all AppleScript tasks have been
' written within a single script file.

Public Function RunAppleScriptTask(ByVal scriptHandler As String, ByVal scriptParameterString As String)

On Error GoTo taskError
    
    RunAppleScriptTask = AppleScriptTask(APPLE_SCRIPT_FILE, scriptHandler, scriptParameterString)
    Exit Function
    
taskError:
    RunAppleScriptTask = vbNullString
End Function


#End If

