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

Private Sub Worksheet_Change(ByVal target As Range)

    On Error GoTo Whoa

    ' Detect changes in the NodeFormatString text box
    If Not Intersect(target, StyleDesignerSheet.Range("LabelText,XlabelText,TailLabelText,HeadLabelText")) Is Nothing Then
        Application.enableEvents = False
        RenderPreviewFromLists
        
    ElseIf Not Intersect(target, StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING)) Is Nothing Then
        Application.enableEvents = False
        If StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).value = vbNullString Then
            ClearStyleRibbon
        Else
            RenderPreviewFromFormatString
            StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
        End If
        
    ElseIf Not Intersect(target, StyleDesignerSheet.Range(DESIGNER_ADD_CAPTION)) Is Nothing Then
        Application.enableEvents = False
        RenderPreview

    ElseIf Not Intersect(target, StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME)) Is Nothing Then
        Application.enableEvents = False
        RefreshControlsImageScale
        RefreshControlsImagePosition
    End If
    
LetsContinue:
    Application.enableEvents = True
    Exit Sub
Whoa:
    EmitMessage "StyleDesignerSheet " & Err.number & " " & Err.Description
    
    Resume LetsContinue

End Sub

