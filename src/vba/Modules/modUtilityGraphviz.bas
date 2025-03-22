Attribute VB_Name = "modUtilityGraphviz"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Data")
'@IgnoreModule ProcedureNotUsed

Option Explicit

Public Sub AlertGraphvizNotFound(ByVal graphEngine As String)
#If Mac Then
    'TODO Port
#Else
    MsgBox replace(GetMessage("msgboxGraphvizNotFound"), "{graphEngine}", graphEngine), vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
#End If
End Sub


