Attribute VB_Name = "modUtilityExcelPicture"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved
'@Folder("Utility.Excel")

Option Explicit

Public Function InsertPicture(ByVal FName As String, ByVal Where As Range, _
                              Optional ByVal LinkToFile As Boolean = False, _
                              Optional ByVal SaveWithDocument As Boolean = True) As Shape
   
    'Inserts the picture file FName as link or permanently into Where
    Dim shapeObject As Shape
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
    ' Removes any pictures located within the specified range of cells

    Dim shapeObject As String
    Dim pictureImage As Picture
    Dim targetWorksheet As Worksheet
    'Dim targetRange As Range

    Set targetWorksheet = ActiveWorkbook.Sheets.[_Default](targetSheet)
    'Set targetRange = targetWorksheet.Range(targetCell)

    For Each pictureImage In targetWorksheet.Pictures
        With pictureImage
            shapeObject = .TopLeftCell.Address
        End With
        If shapeObject = targetCell Then
            pictureImage.Delete
        End If
    Next
       
    Set targetWorksheet = Nothing
    'Set targetRange = Nothing
    
End Sub

Public Sub DeleteAllPictures(ByVal targetSheet As String)
    ' Removes all pictures within a specified worksheet

    Dim pictureImage As Picture
    Dim targetWorksheet As Worksheet

    Set targetWorksheet = ActiveWorkbook.Sheets.[_Default](targetSheet)
    
    For Each pictureImage In targetWorksheet.Pictures
        With pictureImage
            pictureImage.Delete
        End With
    Next
    
    Set targetWorksheet = Nothing
End Sub


