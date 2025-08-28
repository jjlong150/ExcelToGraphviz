Attribute VB_Name = "modUtilityExcelPicture"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved
'@Folder("Utility.Excel")

Option Explicit

Public Function InsertPicture(ByVal FName As String, ByVal Where As Range, _
                              Optional ByVal LinkToFile As Boolean = False, _
                              Optional ByVal SaveWithDocument As Boolean = True) As shape
   
    'Inserts the picture file FName as link or permanently into Where
    Dim shapeObject As shape
    With Where
        'Insert in original size
        Set shapeObject = Where.Parent.Shapes.AddPicture( _
                          FName, _
                          LinkToFile, _
                          SaveWithDocument, _
                          .Left, _
                          .Top, _
                          -1, _
                          -1)
        shapeObject.Placement = xlMove           ' ( xlFreeFloating | xlMove | xlMoveAndSize )
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

