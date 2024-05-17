VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpColorsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Help Colors")

Option Explicit

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    On Error GoTo ErrorHandler
    SetStyleDesignerColorScheme (Target.TextToDisplay)
    Exit Sub

ErrorHandler:
    Err.Clear
    Resume Next
End Sub


