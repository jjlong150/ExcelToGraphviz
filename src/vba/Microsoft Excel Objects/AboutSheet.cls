VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AboutSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.About")

Option Explicit

Private Sub Worksheet_Activate()
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabAbout"
End Sub

