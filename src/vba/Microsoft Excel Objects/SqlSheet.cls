VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@IgnoreModule ParameterNotUsed
'@Folder("Relationship Visualizer.Sheets.SQL")

Option Explicit

Private Sub Worksheet_Activate()
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSql"
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
End Sub

Private Sub Worksheet_Deactivate()
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not TypeOf Target Is Range Then Exit Sub

    ' Remove any existing Edit button
    On Error Resume Next
    Me.buttons("SQLEditButton").Delete
    On Error GoTo 0

    If Target.Cells.count = 1 Then
        Dim sqlCol As Long
        sqlCol = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)

        If Not Intersect(Target, Me.columns(sqlCol)) Is Nothing Then
            If Not Me.ProtectContents Then
                On Error Resume Next
                Dim btn As button
                Set btn = Me.buttons.Add(Left:=Target.Left + Target.Width + 5, _
                                         Top:=Target.Top + 1, _
                                         Width:=20, _
                                         height:=20)
                If Not btn Is Nothing Then
                    With btn
                        .name = "SQLEditButton"
                        .caption = "..."
                        .OnAction = "ShowSQLEditForm"
                    End With
                End If
                On Error GoTo 0
            End If
        End If
    End If

    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
    RefreshRibbon
End Sub

