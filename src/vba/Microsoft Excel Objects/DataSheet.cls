VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Data")

Option Explicit

Private Sub Worksheet_Activate()
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabGraphviz"
    Application.StatusBar = False
End Sub

'@Ignore ParameterNotUsed
Private Sub Worksheet_Change(ByVal Target As Range)
    AutoDraw
End Sub

Private Sub Worksheet_Deactivate()
    Application.StatusBar = False
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Exit if multiple cells selected?
    If Target.Cells.CountLarge > 1 Then Exit Sub
    
    ' Exit if the cell is outside the Style Name column
    If Target.Column <> GetSettingColNum(SETTINGS_DATA_COL_STYLE) Then Exit Sub
    
    ' Clear the named range which is associated with the Style Name dropdown list
    ChoicesSheet.Range(RANGE_STYLE_NAMES).ClearContents
    
    ' Exit if heading row
    If Target.row = 1 Then Exit Sub
    
    ' Examine the row and determine how it is being used
    Dim rowType As String
    rowType = getRowType(ActiveSheet.name, Target.row)
    
    ' Use a dictionary to load the unique style names for this row type
    Dim styleNames As Dictionary
    Set styleNames = getMatchingStyles(rowType)
    
    If styleNames.count > 0 Then
        ' For rows which can be styled
        If rowType <> TYPE_BLANK_ROW And rowType <> TYPE_GRAPH Then
            ' Paste the dictionary keys into the named range associated with the
            ' Style Name dropdown list
            ChoicesSheet.Cells.Item(1, 1).Resize(styleNames.count, 1).Value2 = Application.WorksheetFunction.Transpose(styleNames.Keys)
            
            ' Adjust the size of the named range to match the number of values
            Dim namedRange As name
            Set namedRange = ActiveWorkbook.Names.[_Default](RANGE_STYLE_NAMES)
            
            With namedRange
                '@Ignore ImplicitDefaultMemberAccess
                .RefersTo = .RefersToRange.Resize(styleNames.count, 1)
            End With
            
            ' Sort the values within the named range
            Dim columnRange As String
            columnRange = "A1:A" & styleNames.count
            
            ChoicesSheet.sort.SortFields.Clear
            ChoicesSheet.sort.SortFields.Add key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ChoicesSheet.sort
                .SetRange Range(columnRange)
                .Header = xlNo
                .MatchCase = False
                .orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    End If
    
    ' Don't leak memory
    Set namedRange = Nothing
    Set styleNames = Nothing
End Sub

Private Function getRowType(ByVal worksheetName As String, ByVal row As Long) As String

    Dim rowType As String
    '@Ignore AssignmentNotUsed
    rowType = TYPE_BLANK_ROW
    
    Dim dataItem As String
    dataItem = UCase$(GetCell(worksheetName, row, GetSettingColNum(SETTINGS_DATA_COL_ITEM)))

    If dataItem <> vbNullString Then
        If EndsWith(dataItem, OPEN_BRACE) Then
            rowType = TYPE_SUBGRAPH_OPEN
        
        ElseIf dataItem = CLOSE_BRACE Then
            rowType = TYPE_SUBGRAPH_CLOSE
        
        ElseIf dataItem = GREATER_THAN Then
            rowType = TYPE_NATIVE
        
        Else
            Dim dataIsRelatedtoItem As String
            dataIsRelatedtoItem = GetCell(worksheetName, row, GetSettingColNum(SETTINGS_DATA_COL_IS_RELATED_TO))
            
            If dataIsRelatedtoItem = vbNullString Then
                If dataItem = KEYWORD_NODE Then
                    rowType = TYPE_NODE
                ElseIf dataItem = KEYWORD_EDGE Then
                    rowType = TYPE_EDGE
                ElseIf dataItem = KEYWORD_GRAPH Then
                    rowType = TYPE_GRAPH
                Else
                    rowType = TYPE_NODE
                End If
            Else
                rowType = TYPE_EDGE
            End If
        End If
    End If

    getRowType = rowType
    
End Function

Private Function getMatchingStyles(ByVal rowType As String) As Dictionary

    Dim styles As stylesWorksheet
    styles = GetSettingsForStylesWorksheet()
    
    ' Dictionary to hold the key and associated values
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
    
    ' Loop through the specified range
    Dim row As Long
    Dim styleName As String
    
    For row = styles.firstRow To styles.lastRow
        '@Ignore EmptyIfBlock
        If StylesSheet.Cells.Item(row, styles.flagColumn).value = FLAG_COMMENT Then
            ' Comment row, ignore it
        ElseIf StylesSheet.Cells.Item(row, styles.typeColumn).value = rowType Then
            ' Retrieve the style name
            styleName = StylesSheet.Cells.Item(row, styles.nameColumn).value

            If styleName <> vbNullString Then    ' a style name is present
                ' Add style name to the dictionary if not already present
                If Not dictionaryObj.Exists(styleName) Then ' ignore duplicate style names
                    dictionaryObj.Add styleName, "1"    ' value of name/value pair is not used
                End If
            End If
        End If
    Next row
    
    If dictionaryObj.Exists(TYPE_NODE) Then dictionaryObj.Remove (TYPE_NODE)
    If dictionaryObj.Exists(TYPE_EDGE) Then dictionaryObj.Remove (TYPE_EDGE)
    If dictionaryObj.Exists(TYPE_SUBGRAPH_OPEN) Then dictionaryObj.Remove (TYPE_SUBGRAPH_OPEN)
    If dictionaryObj.Exists(TYPE_SUBGRAPH_CLOSE) Then dictionaryObj.Remove (TYPE_SUBGRAPH_CLOSE)
    If dictionaryObj.Exists(TYPE_KEYWORD) Then dictionaryObj.Remove (TYPE_KEYWORD)
    If dictionaryObj.Exists(TYPE_NATIVE) Then dictionaryObj.Remove (TYPE_NATIVE)

    Set getMatchingStyles = dictionaryObj
    
End Function


