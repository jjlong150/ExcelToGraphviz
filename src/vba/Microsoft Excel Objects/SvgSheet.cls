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
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSvg"
#Else
    ActivateTabSvg
#End If
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub

Private Sub Worksheet_Deactivate()
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If Not TypeOf target Is Range Then Exit Sub

    ' Remove any existing Edit button
    On Error Resume Next
    Me.buttons("SVGEditButton").Delete
    Me.Shapes("SVGEditButton").Delete
    On Error GoTo 0

    If target.Cells.count = 1 Then
        If Not Intersect(target, Me.columns(svgLayoutColumn.replaceColumn)) Is Nothing Then
            If Not Me.ProtectContents Then
                On Error Resume Next
                ' Add clickable pencil icon with subtle button look
                Dim shp As shape
                Set shp = Me.Shapes.AddShape(msoShapeRoundedRectangle, _
                            target.left + target.Width + 5, _
                            target.top + 1, 20, 20)
                With shp
                    .name = "SVGEditButton"
                    .Fill.ForeColor.RGB = RGB(245, 245, 245)      ' very light gray background
                    .Fill.Transparency = 0
                    .Line.Weight = 0.75
                    .Line.ForeColor.RGB = RGB(180, 180, 180)      ' soft border
                    .TextFrame2.TextRange.Characters.Text = ChrW(&H270E)  ' pencil
                    .TextFrame2.TextRange.Font.name = "Segoe UI Symbol"
                    .TextFrame2.TextRange.Font.Size = 16
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(60, 60, 60)  ' dark gray pencil
                    .TextFrame2.VerticalAnchor = msoAnchorMiddle
                    .TextFrame2.HorizontalAnchor = msoAnchorCenter
                    .Placement = xlMoveAndSize
                    .OnAction = "ShowSVGEditForm"
                End With
                On Error GoTo 0
            End If
        End If
    End If

    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
    RefreshRibbon
End Sub
