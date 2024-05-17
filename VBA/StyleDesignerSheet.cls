VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleDesignerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Style Designer")

Option Explicit

Private Sub Worksheet_Activate()
    SyncHelpToggleButtons
    ActivateTab (RIBBON_TAB_STYLE_DESIGNER)
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
        If StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Value = vbNullString Then
            ClearStyleRibbon
        Else
            RenderPreviewFromFormatString
            StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
        End If
        
    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_ADD_CAPTION)) Is Nothing Then
        Application.EnableEvents = False
        RenderPreview

    End If
    
LetsContinue:
    Application.EnableEvents = True
    Exit Sub
Whoa:
    MsgBox "StyleDesignerSheet " & Err.number & " " & Err.Description, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    
    Resume LetsContinue

End Sub

