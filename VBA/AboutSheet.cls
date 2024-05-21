VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AboutSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.About")

Option Explicit

Private Sub Worksheet_Activate()
    SyncHelpToggleButtons
    ActivateTab (RIBBON_TAB_GRAPHVIZ)
End Sub

