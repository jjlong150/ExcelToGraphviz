Attribute VB_Name = "modUtilityExcelCellButtons"
' Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved

Option Explicit

Public Type ButtonConfig
    ColumnNumber    As Long
    ButtonName      As String       ' must be unique!
    ActionMacro     As String
    IconText        As String
    VOffset         As Long         ' vertical offset in pixels (positive = down)
    HOffset         As Long         ' horizontal offset in pixels (positive = right)
    ValidationFunc  As String       ' name of function that returns Boolean: "MyValidation(row)"
End Type

' Creates one floating button according to the passed config
'
' .IconText values for Action Buttons
'
' ChrW(&H270E)     pencil angled - eraser upper left, tip lower right
' ChrW(&H270F)     pencil horizontal - eraser left, tip right
' ChrW(&H2710)     pencil angled - eraser lower left, tip upper right
' ChrW(&H21BA)     counterclockwise circular arrow (refresh/reload)
' ChrW(&H21BB)     clockwise circular arrow pointing up   (refresh/reload)
' ChrW(&H27F3)     clockwise circular arrow pointing down (loading/refresh)
' ChrW(&H2699)     gear / settings
' ChrW(&H2714)     heavy check mark
' ChrW(&H2716)     X heavy multiplication X / delete
' ChrW(&H2794)     => thick right arrow
' ChrW(&H27A1)     -> narrow right arrow
' ChrW(&H2B07)     down arrow
' ChrW(&H2191)     narrow upwards arrow (simple & classic)
' ChrW(&H2B06)     thick upwards white arrow (bold, good visibility)
' ChrW(&H21D1)     upwards double arrow (sort ascending / go up)
' ChrW(&H25B2)     /\ black up-pointing triangle (solid, prominent)
' ChrW(&H25B3)     /\ white up-pointing triangle (outlined)
' ChrW(&H25B6)     |> black right-pointing triangle (classic "play/run")
' ChrW(&H2303)     ^   caret / up arrowhead (small, technical)
' ChrW(&H23F8)     || pause symbol

Public Sub CreateOneFloatingButton( _
    ByVal ws As Worksheet, _
    ByVal cell As Range, _
    ByRef cfg As ButtonConfig)

    If Len(cfg.ButtonName) = 0 Or Len(cfg.IconText) = 0 Then Exit Sub

    Dim shp As shape
    Dim leftPos As Double, topPos As Double

    On Error GoTo CreateErr

    leftPos = cell.left + cell.Width + cfg.HOffset
    topPos = cell.top + 1 + cfg.VOffset

    ' Remove any existing button with this name
    On Error Resume Next
    ws.Shapes(cfg.ButtonName).Delete
    On Error GoTo CreateErr

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 leftPos, topPos, 20, 20)

    With shp
        .name = cfg.ButtonName
        .Fill.ForeColor.RGB = RGB(245, 245, 245)
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .Line.ForeColor.RGB = RGB(180, 180, 180)
        .TextFrame2.TextRange.Characters.Text = cfg.IconText
        .TextFrame2.TextRange.Font.name = "Segoe UI Symbol"
        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(60, 60, 60)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .Placement = xlMoveAndSize
        .OnAction = cfg.ActionMacro
    End With

    Exit Sub

CreateErr:
    Debug.Print "CreateFloatingButton failed for " & cfg.ButtonName & ": " & Err.Description
    Err.Clear
End Sub

' Processes the supplied ButtonConfig array and creates floating buttons
' where the selected column matches and any validation passes.
Public Sub UpdateFloatingButtonsOnSheet( _
    ByVal ws As Worksheet, _
    ByVal Target As Range, _
    ByRef configs() As ButtonConfig)

    ' Early exits — common safety checks
    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge > 1 Or Target.Cells.count = 0 Then Exit Sub
    If ws.ProtectContents Then Exit Sub

    Dim rowNum As Long
    rowNum = Target.row

    Dim cfg As ButtonConfig
    Dim i As Long

    For i = LBound(configs) To UBound(configs)
        cfg = configs(i)

        ' Only consider this config if the selected cell is in the target column
        If Not Intersect(Target, ws.columns(cfg.ColumnNumber)) Is Nothing Then

            ' Per-button validation (optional)
            Dim showIt As Boolean
            showIt = True

            If Len(cfg.ValidationFunc) > 0 Then
                On Error Resume Next
                showIt = Application.Run(cfg.ValidationFunc, rowNum)
                If Err.number <> 0 Then
                    Debug.Print "Validation function failed: " & cfg.ValidationFunc & _
                                " - Error: " & Err.Description
                    showIt = False
                    Err.Clear
                End If
                On Error GoTo 0
            End If

            ' Create the button if all conditions are met
            If showIt Then
                CreateOneFloatingButton ws, Target, cfg
            End If
        End If
    Next i

End Sub

' Returns an array of button names extracted from a ButtonConfig array
Public Function GetButtonNamesFromConfigs(ByRef configs() As ButtonConfig) As Variant
    If UBound(configs) < LBound(configs) Then
        GetButtonNamesFromConfigs = Array()   ' empty array if no configs
        Exit Function
    End If
    
    Dim names() As String
    ReDim names(0 To UBound(configs))
    
    Dim i As Long
    For i = LBound(configs) To UBound(configs)
        names(i) = configs(i).ButtonName
    Next i
    
    GetButtonNamesFromConfigs = names
End Function

' Removes all floating buttons whose name appears in the supplied array
Public Sub RemoveFloatingButtons( _
    ByVal ws As Worksheet, _
    ByVal buttonNames As Variant)

    If IsEmpty(buttonNames) Then Exit Sub

    Dim shp As shape
    Dim nm As Variant

    On Error Resume Next
    For Each shp In ws.Shapes
        For Each nm In buttonNames
            If shp.name = nm Then
                shp.Delete
                Exit For
            End If
        Next nm
    Next shp
    On Error GoTo 0
End Sub


