VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleDesignerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Style Designer")

Option Explicit

Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabStyleDesigner"
#Else
    ActivateTabStyleDesigner
#End If
    RenderPreview
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Whoa

    ' Detect changes in the NodeFormatString text box
    If Not Intersect(Target, StyleDesignerSheet.Range("LabelText,XlabelText,TailLabelText,HeadLabelText")) Is Nothing Then
        Application.EnableEvents = False
        RenderPreviewFromLists
        
    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING)) Is Nothing Then
        Application.EnableEvents = False
        If StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).value = vbNullString Then
            ClearStyleRibbon
        Else
            RenderPreviewFromFormatString
            StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
        End If
        
    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_ADD_CAPTION)) Is Nothing Then
        Application.EnableEvents = False
        RenderPreview

    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME)) Is Nothing Then
        Application.EnableEvents = False
        RefreshControlsImageScale
        RefreshControlsImagePosition
    End If
    
LetsContinue:
    Application.EnableEvents = True
    Exit Sub
Whoa:
    EmitMessage "StyleDesignerSheet " & Err.number & " " & Err.Description
    
    Resume LetsContinue

End Sub

