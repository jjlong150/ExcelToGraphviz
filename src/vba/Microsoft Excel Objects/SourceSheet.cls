VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Source")
'@IgnoreModule ParameterNotUsed

Option Explicit

Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSource"
#Else
    ActivateTabSource
#End If
    ActiveSheet.Range("A1").Select
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    OptimizeCode_Begin
    UpdateSourceWorksheetLineNumbers
    OptimizeCode_End
End Sub

