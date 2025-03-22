VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved
'@Folder("Relationship Visualizer.Sheets.SVG")
'@IgnoreModule ParameterNotUsed


Option Explicit

Private Sub Worksheet_Activate()
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSvg"
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub

Private Sub Worksheet_Deactivate()
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub
