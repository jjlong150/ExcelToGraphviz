Attribute VB_Name = "modUtilityExcelPicture"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcelPicture
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER: Utility / Excel Interop
'
' ROLE:
'   Insert, locate, and remove raster and vector images on worksheets. Provides
'   a consistent abstraction over Shapes.AddPicture and worksheet-based picture
'   enumeration, including SVG-aware deletion routines.
'
' RESPONSIBILITIES:
'   - Picture insertion:
'       • InsertPicture: add raster or vector images at a target cell's position
'         using original dimensions, optional linking, and optional AltText
'   - Picture deletion:
'       • DeletePictures: remove pictures whose bounding cells intersect a range
'       • DeleteCellPictures: remove images anchored to a specific cell
'       • DeleteAllPictures: remove all msoPicture and msoGraphic shapes
'
' ARCHITECTURAL NOTES:
'   - Uses Shapes.AddPicture for full compatibility with PNG, JPG, GIF, BMP,
'     and SVG (msoGraphic) formats.
'   - Placement = xlMove ensures images track row/column movement without
'     resizing.
'   - Deletion routines rely on TopLeftCell/BottomRightCell for precise
'     cell-anchored targeting.
'   - v8.0.01 revisions (per changelog): SVG deletion added; Atom CPU-safe
'     enumeration logic preserved.
'
' VERSION NOTES:
'   - v8.0.01: Expanded deletion logic to include SVG (msoGraphic) and maintain
'              compatibility with low-power Atom CPUs.
'
' USAGE:
'   - Used by SVG export, Data sheet previews, Source/Styles workflows, and
'     any feature requiring image placement or cleanup.
'
' RELATED WIKI PAGES:
'   - Image Handling & SVG Support
'   - Worksheet Rendering Pipeline
'   - Shape Enumeration & Deletion Patterns
' =============================================================================

Option Explicit

Public Function InsertPicture(ByVal fname As String, ByVal Where As Range, _
                              Optional ByVal LinkToFile As Boolean = False, _
                              Optional ByVal SaveWithDocument As Boolean = True, _
                              Optional ByVal AltText As String = vbNullString) As shape
   
    'Inserts the picture file FName as link or permanently into Where
    Dim shapeObject As shape
    With Where
        'Insert in original size
        Set shapeObject = Where.Parent.Shapes.AddPicture( _
                          fname, _
                          LinkToFile, _
                          SaveWithDocument, _
                          .Left, _
                          .top, _
                          -1, _
                          -1)
        shapeObject.Placement = xlMove           ' ( xlFreeFloating | xlMove | xlMoveAndSize )
        
        ' Set alternative text if provided
        If Len(AltText) > 0 Then
            shapeObject.AlternativeText = AltText
        End If
    End With
   
    Set InsertPicture = shapeObject
    Set shapeObject = Nothing
End Function

'@Ignore ProcedureNotUsed
Public Sub DeletePictures(ByVal targetSheet As String, ByVal targetCells As String)
    ' Removes any pictures located within the specified range of cells

    Dim shapeObject As String
    Dim pictureImage As Picture
    Dim targetWorksheet As Worksheet
    Dim targetRange As Range

    Set targetWorksheet = ActiveWorkbook.Sheets.[_Default](targetSheet)
    Set targetRange = targetWorksheet.Range(targetCells)

    For Each pictureImage In targetWorksheet.Pictures
        With pictureImage
            shapeObject = .TopLeftCell.Address & ":" & .BottomRightCell.Address
        End With
        If Not Intersect(targetRange, targetWorksheet.Range(shapeObject)) Is Nothing Then
            pictureImage.Delete
        End If
    Next
       
    Set targetWorksheet = Nothing
    Set targetRange = Nothing
    
End Sub

'@Ignore ProcedureNotUsed
Public Sub DeleteCellPictures(ByVal targetSheet As String, ByVal targetCell As String)
    ' Removes raster and vector images located within the specified cell
    ' Revised in v8.0.01 to include SVG deletion and maintain Atom CPU compatibility

    Dim shp As shape
    Dim targetWorksheet As Worksheet

    Set targetWorksheet = ActiveWorkbook.Sheets(targetSheet)

    For Each shp In targetWorksheet.Shapes
        Select Case shp.Type
            Case msoPicture, msoGraphic
                If Not shp.TopLeftCell Is Nothing Then
                    If shp.TopLeftCell.Address = targetCell Then
                        On Error Resume Next
                        shp.Delete
                        On Error GoTo 0
                    End If
                End If
        End Select
    Next

    Set targetWorksheet = Nothing
End Sub

Public Sub DeleteAllPictures(ByVal targetSheet As String)
    ' Removes all raster and vector images (msoPicture and msoGraphic) from a worksheet
    ' Revised in v8.0.01 to include SVG deletion and maintain Atom CPU compatibility

    Dim shp As shape
    Dim targetWorksheet As Worksheet

    Set targetWorksheet = ActiveWorkbook.Sheets(targetSheet)

    For Each shp In targetWorksheet.Shapes
        Select Case shp.Type
            Case msoPicture, msoGraphic
                On Error Resume Next
                shp.Delete
                On Error GoTo 0
        End Select
    Next

    Set targetWorksheet = Nothing
End Sub

