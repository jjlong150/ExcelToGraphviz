Attribute VB_Name = "modWorksheetSQLClusters"
' Copyright (c) 2026 Jeffrey J. Long. All rights reserved
' New module for multi-level clustering support in Relationship Visualizer.

'@Folder("Relationship Visualizer.Sheets.SQL")

Option Explicit

' Detect if multi-level clustering is present (prefers over old "CLUSTER")
Public Function DetectMultiLevel(ByVal rs As Object, prefix As String) As Boolean
    DetectMultiLevel = HasField(rs, prefix & "1")
End Function

' Main entry point: Process recordset with multi-level clusters
Public Sub ProcessMultiLevelRecordset( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByRef recordCnt As Long)

    ' Exit early if invalid or empty
    If rs Is Nothing Then Exit Sub
    If rs.State <> ObjectStateEnum.adStateOpen Then Exit Sub
    If rs.EOF And rs.BOF Then Exit Sub

    ' Build simple sort string (no expressions)
    Dim maxLevels As Long
    maxLevels = DetectMaxLevels(rs, ctx.fields.Cluster, ctx.fields.clusterLevelLimit)
    
    Dim sortStr As String
    sortStr = ""
    
    Dim i As Long
    For i = 1 To maxLevels
        Dim fieldName As String
        fieldName = "[" & ctx.fields.Cluster & i & "]"
        
        If Len(sortStr) > 0 Then sortStr = sortStr & ", "
        sortStr = sortStr & fieldName & " ASC"
    Next i

    ' Attempt sort with error handling
    On Error GoTo SortFailed
    rs.sort = sortStr
    On Error GoTo 0

    ' Ensure we're at the beginning after sort
    rs.MoveFirst

    ' Proceed with clustering logic
    Dim currentClusters() As String
    ReDim currentClusters(1 To maxLevels)
    
    ' Per-level cluster counters
    Dim clusterCounters() As Long
    ReDim clusterCounters(1 To maxLevels)
    For i = 1 To maxLevels
        clusterCounters(i) = 0
    Next i
    
    Dim openLevels As Long: openLevels = 0
    Dim currentOpenLevel As Long:  currentOpenLevel = 0   ' 0 = no cluster open yet (flat nodes)
    Dim absoluteClusterCount As Long: absoluteClusterCount = 0
    Dim relativeClusterCount As Long: relativeClusterCount = 0   ' fallback if no clusters ever opened
    
    Do While Not rs.EOF
        Dim thisClusters() As String
        ReDim thisClusters(1 To maxLevels)
        For i = 1 To maxLevels
            thisClusters(i) = SafeFieldValue(rs, ctx.fields.Cluster & i)
        Next i

        Dim changeLevel As Long
        changeLevel = FindChangeLevel(currentClusters, thisClusters, maxLevels)

        ' Close inner levels
        For i = openLevels To changeLevel Step -1
            EmitClusterClose ctx, rs, row, i, absoluteClusterCount, clusterCounters(i)
            openLevels = openLevels - 1
        Next i

        ' After closing, update currentOpenLevel to the new deepest open level
        currentOpenLevel = openLevels
        
        ' Open new clusters
        For i = changeLevel To maxLevels
            If thisClusters(i) <> "" Then
                ' Increment counter for **this level only**
                clusterCounters(i) = clusterCounters(i) + 1
                
                ' Remember this as the most recently opened relative count
                relativeClusterCount = clusterCounters(i)
                
                ' Increment absolute cluster count
                absoluteClusterCount = absoluteClusterCount + 1
                
                EmitClusterOpen ctx, rs, row, i, absoluteClusterCount, clusterCounters(i)
                
                openLevels = openLevels + 1
                currentOpenLevel = i
                currentClusters(i) = thisClusters(i)
                
                ' Reset deeper levels' counters
                Dim deeper As Long
                For deeper = i + 1 To maxLevels
                    currentClusters(deeper) = ""
                    clusterCounters(deeper) = 0
                Next deeper
            End If
        Next i

        ' Increment the rs record count
        recordCnt = recordCnt + 1
        
        ' Emit row
        EmitRows ctx, rs, row, maxLevels, recordCnt, currentOpenLevel, absoluteClusterCount, relativeClusterCount

        rs.MoveNext
    Loop

    ' Close remaining
    rs.MoveFirst
    For i = 1 To openLevels
        EmitClusterClose ctx, rs, row, i, absoluteClusterCount, clusterCounters(i)
    Next i

    Exit Sub

SortFailed:
    ' Continue without sorting (data will still be processed, just not ordered)
    On Error GoTo 0
    rs.MoveFirst
    Resume Next

End Sub

Private Function DetectMaxLevels(ByVal rs As Object, prefix As String, levelLimit As Long) As Long
    Dim i As Long
    i = 1
    While i <= levelLimit And HasField(rs, prefix & i)
        i = i + 1
    Wend
    DetectMaxLevels = i - 1
End Function

Private Function FindChangeLevel(ByRef current() As String, ByRef thisOne() As String, ByVal maxLevels As Long) As Long
    Dim i As Long
    For i = 1 To maxLevels
        If current(i) <> thisOne(i) Then
            FindChangeLevel = i
            Exit Function
        End If
    Next i
    FindChangeLevel = maxLevels + 1  ' No change
End Function

Private Sub EmitClusterOpen( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long)

    With DataSheet
        ' Mandatory opening brace
        .Cells(row, ctx.dataLayout.itemColumn).value = OPEN_BRACE

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterLabel, ctx.dataLayout.labelColumn

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterStyleName, ctx.dataLayout.styleNameColumn, _
            suffixFromSetting:=SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_OPEN).value

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterAttributes, ctx.dataLayout.extraAttributesColumn

        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterTooltip, ctx.dataLayout.tooltipColumn
    End With
    
    row = row + 1
End Sub

Private Sub EmitClusterClose( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long)

    With DataSheet
        ' Mandatory: closing brace
        .Cells(row, ctx.dataLayout.itemColumn).value = CLOSE_BRACE

        ' Only style name is emitted on close (for stylesheet toggle support)
        ProcessClusterProperty ctx, rs, row, levelNumber, _
            absoluteClusterCount, relativeClusterCount, _
            ctx.fields.clusterStyleName, ctx.dataLayout.styleNameColumn, _
            suffixFromSetting:=SafeStr(SettingsSheet.Range(SETTINGS_STYLES_SUFFIX_CLOSE).value)
    End With

    row = row + 1
End Sub

Private Sub ProcessClusterProperty( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef wsRow As Long, _
    ByVal levelNumber As Long, _
    ByVal absCount As Long, _
    ByVal relCount As Long, _
    ByVal templateField As String, _
    ByVal targetColumn As Long, _
    Optional ByVal suffixFromSetting As String = vbNullString)

    Dim clusterPrefix As String
    clusterPrefix = ctx.fields.Cluster & levelNumber

    Dim fieldName As String
    fieldName = replace(templateField, ctx.fields.Cluster, clusterPrefix, , , vbTextCompare)

    If Not HasField(rs, fieldName) Then Exit Sub

    Dim value As String
    value = SafeFieldValue(rs, fieldName)

    ' Apply suffix if provided (different for open vs close)
    If Len(suffixFromSetting) > 0 Then
        value = value & suffixFromSetting
    End If

    ' Apply placeholders (same logic for all properties)
    value = replace(value, ctx.fields.clusterPlaceholder, SafeStr(absCount), , , vbTextCompare)
    value = replace(value, ctx.fields.subclusterPlaceholder, SafeStr(relCount), , , vbTextCompare)
    value = replace(value, ctx.fields.clusterLevelPlaceholder, SafeStr(levelNumber), , , vbTextCompare)

     ' Write result
    DataSheet.Cells(wsRow, targetColumn).value = value
End Sub

Private Sub EmitRows( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal maxLevels As Long, _
    ByVal recordCount As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long)
    
    Dim i As Long

    ' Safety: prevent infinite loop
    If ctx.loop.stepBy = 0 Then Exit Sub

    ' Safety: prevent direction mismatch infinite loop
    If ctx.loop.stepBy > 0 Then
        If ctx.loop.startAt > ctx.loop.stopAt Then Exit Sub
    Else
        If ctx.loop.startAt < ctx.loop.stopAt Then Exit Sub
    End If

    For i = ctx.loop.startAt To ctx.loop.stopAt Step ctx.loop.stepBy
        ctx.loop.count = ctx.loop.count + 1
        If ctx.loop.count > ctx.loop.max Then Exit For
        
        EmitOneRow ctx, rs, row, maxLevels, recordCount, levelNumber, absoluteClusterCount, relativeClusterCount, i
        row = row + 1
    Next i

End Sub

Private Sub EmitOneRow( _
    ByRef ctx As sqlContext, _
    ByVal rs As Object, _
    ByRef row As Long, _
    ByVal maxLevels As Long, _
    ByVal recordCount As Long, _
    ByVal levelNumber As Long, _
    ByVal absoluteClusterCount As Long, _
    ByVal relativeClusterCount As Long, _
    ByVal enumStep As Long)
    
    ' Process all rs fields and write to 'data' worksheet
    With DataSheet
        Dim fld As Object
        Dim v As String
        Dim targetCol As Long
        
        For Each fld In rs.fields
            If Not IsClusterRelatedField(fld.name, maxLevels) Then

                ' Common transformation: null ? "", placeholder replacement
                v = SafeStr(fld.value)
                
                If Len(v) > 0 Then
                    v = replace(v, ctx.fields.recordsetPlaceholder, CStr(recordCount), , , vbTextCompare)
                    v = replace(v, ctx.fields.clusterPlaceholder, CStr(absoluteClusterCount), , , vbTextCompare)
                    v = replace(v, ctx.fields.subclusterPlaceholder, CStr(relativeClusterCount), , , vbTextCompare)
                    v = replace(v, ctx.fields.clusterLevelPlaceholder, CStr(levelNumber), , , vbTextCompare)

                    If ctx.loop.Enabled Then
                        v = replace(v, ctx.fields.enumeratePlaceholder, SafeStr(enumStep), , , vbTextCompare)
                    End If
                End If
                
                Select Case LCase$(fld.name)
                    Case ctx.headings.flag
                        .Cells(row, ctx.dataLayout.flagColumn).value = v
                    
                    Case ctx.headings.item
                        .Cells(row, ctx.dataLayout.itemColumn).value = v
                    
                    Case ctx.headings.label, ctx.headings.xLabel
                        targetCol = IIf(LCase$(fld.name) = ctx.headings.label, _
                                        ctx.dataLayout.labelColumn, _
                                        ctx.dataLayout.xLabelColumn)
                                        
                        ' Apply multiline splitting only when requested & meaningful
                        Dim splitLength As Long
                        splitLength = GetSplitLength(rs, ctx.fields.splitLength)
                        If splitLength > 0 Then
                            Dim lineEnding  As String
                            lineEnding = GetLineEnding(rs, ctx.fields.lineEnding, NEWLINE)
                            v = SplitMultilineText(v, splitLength, lineEnding)
                        End If
                        
                        .Cells(row, targetCol).value = v
                    
                    Case ctx.headings.tailLabel
                        .Cells(row, ctx.dataLayout.tailLabelColumn).value = v
                    
                    Case ctx.headings.headLabel
                        .Cells(row, ctx.dataLayout.headLabelColumn).value = v
                    
                    Case ctx.headings.Tooltip
                        .Cells(row, ctx.dataLayout.tooltipColumn).value = v
                    
                    Case ctx.headings.isRelatedToItem
                        .Cells(row, ctx.dataLayout.isRelatedToItemColumn).value = v
                    
                    Case ctx.headings.styleName
                        .Cells(row, ctx.dataLayout.styleNameColumn).value = v
                    
                    Case ctx.headings.extraAttributes
                        .Cells(row, ctx.dataLayout.extraAttributesColumn).value = v
                    
                    Case ctx.headings.errorMessage
                        .Cells(row, ctx.dataLayout.errorMessageColumn).value = v
                    
                    ' Case Else: ignore unknown columns (intentional, general-purpose)
                End Select
            End If
        Next fld
    End With
End Sub

Private Function IsClusterRelatedField(ByVal fieldName As String, ByVal maxLevels As Long) As Boolean
    Dim i As Long
    Dim lcName As String
    lcName = LCase$(fieldName)
    
    Dim prefix As String
    For i = 1 To maxLevels
        prefix = LCase$("CLUSTER" & i)
        If lcName = prefix Or _
           lcName = prefix & " label" Or _
           lcName = prefix & " style name" Or _
           lcName = prefix & " attributes" Or _
           lcName = prefix & " tooltip" Then
            IsClusterRelatedField = True
            Exit Function
        End If
    Next i
End Function

