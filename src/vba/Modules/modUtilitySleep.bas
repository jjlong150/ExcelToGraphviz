Attribute VB_Name = "modUtilitySleep"
' Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved

Option Explicit

' Simple sleep using Timer + DoEvents (no API).
Public Sub SleepMilliseconds(ByVal ms As Long)
#If Win32 Or Win64 Then
    Dim t As Single
    t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
#End If
End Sub


