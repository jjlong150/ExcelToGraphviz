Attribute VB_Name = "modUtilityDateTime"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Date Time")
'@IgnoreModule ProcedureNotUsed

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

