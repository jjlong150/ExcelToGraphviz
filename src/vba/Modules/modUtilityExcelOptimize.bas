Attribute VB_Name = "modUtilityExcelOptimize"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

' Public routines to turn code optimizations on and off
Public Sub OptimizeCode_Begin()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Public Sub OptimizeCode_End()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


