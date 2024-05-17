VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Source")
'@IgnoreModule ParameterNotUsed

Option Explicit

Private Sub Worksheet_Activate()
    SyncHelpToggleButtons
    ActivateTab (RIBBON_TAB_SOURCE)
    ActiveSheet.Range("A1").Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    OptimizeCode_Begin
    UpdateSourceWorksheetLineNumbers
    OptimizeCode_End
    
End Sub

