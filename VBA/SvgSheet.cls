VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.SVG")

Option Explicit

Private Sub Worksheet_Activate()
    SyncHelpToggleButtons
    ActivateTab (RIBBON_TAB_SVG)
End Sub

