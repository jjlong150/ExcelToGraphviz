Attribute VB_Name = "modUtilityADODBRecordSets"
' Copyright (c) 2015-2025 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

'*******************************************************************************
' Subroutine:  MergeRecordsets
' Description: Merges two ADODB recordsets (rsFirst and rsSecond) into a single
'              recordset (rsMergedResults) using late binding. The merged
'              recordset is passed by reference. This subroutine handles
'              recordsets with different field lists.
'
' Parameters:
'   rsFirst  - The first ADODB recordset to be merged.
'   rsSecond - The second ADODB recordset to be merged.
'   rsMergedResults - The resulting merged ADODB recordset, passed by reference.
'
' Process:
'   1. Initialize the merged recordset (rsMergedResults) with fields from both
'      rsFirst and rsSecond, ensuring no duplicate fields.
'   2. Copy records from rsFirst to rsMergedResults.
'   3. Copy records from rsSecond to rsMergedResults.
'   4. Commit the records to the merged recordset.
'
' Note:
'   - The subroutine uses late binding for ADODB objects.
'   - Fields that exist in one recordset but not the other are handled gracefully.
'
'*******************************************************************************

Public Sub MergeRecordsets(ByVal rsFirst As Object, _
                           ByVal rsSecond As Object, _
                           ByRef rsMergedResults As Object)

    On Error GoTo MergeError

    Dim field As Object
    Dim mergedField As Object
    Dim fieldExists As Boolean
    Dim fieldSize As Long

    '---------------------------------------------------------------------------
    ' Defensive guards: ensure both recordsets are valid and open
    '---------------------------------------------------------------------------
    If rsFirst Is Nothing Or rsFirst.State <> adStateOpen Then Exit Sub
    If rsSecond Is Nothing Or rsSecond.State <> adStateOpen Then Exit Sub

    '---------------------------------------------------------------------------
    ' Ensure both recordsets are positioned at BOF
    '---------------------------------------------------------------------------
    On Error Resume Next
    rsFirst.MoveFirst
    rsSecond.MoveFirst
    On Error GoTo MergeError

    '---------------------------------------------------------------------------
    ' If caller passed an existing recordset, close it safely
    '---------------------------------------------------------------------------
    If Not rsMergedResults Is Nothing Then
        If rsMergedResults.State = adStateOpen Then rsMergedResults.Close
    End If

    '---------------------------------------------------------------------------
    ' Create a new recordset to contain the merged records
    '---------------------------------------------------------------------------
    Set rsMergedResults = CreateObject("ADODB.Recordset")
    rsMergedResults.CursorLocation = CursorLocationEnum.adUseClient
    rsMergedResults.CursorType = CursorTypeEnum.adOpenStatic
    rsMergedResults.LockType = LockTypeEnum.adLockBatchOptimistic

    '---------------------------------------------------------------------------
    ' Create fields in merged recordset for rsFirst
    '---------------------------------------------------------------------------
    For Each field In rsFirst.fields
        fieldSize = field.DefinedSize
        If fieldSize < 1 Then fieldSize = 255   ' Prevent provider errors on -1 sizes
        rsMergedResults.fields.Append field.name, field.Type, fieldSize, field.attributes
    Next field

    '---------------------------------------------------------------------------
    ' Create fields in merged recordset for rsSecond if they don't already exist
    '---------------------------------------------------------------------------
    For Each field In rsSecond.fields
        fieldExists = False
        For Each mergedField In rsMergedResults.fields
            If mergedField.name = field.name Then
                fieldExists = True
                Exit For
            End If
        Next mergedField

        If Not fieldExists Then
            fieldSize = field.DefinedSize
            If fieldSize < 1 Then fieldSize = 255
            rsMergedResults.fields.Append field.name, field.Type, fieldSize, field.attributes
        End If
    Next field

    '---------------------------------------------------------------------------
    ' Open the merged recordset
    '---------------------------------------------------------------------------
    rsMergedResults.Open

    '---------------------------------------------------------------------------
    ' Copy records from rsFirst to rsMergedResults
    '---------------------------------------------------------------------------
    rsFirst.MoveFirst
    Do Until rsFirst.EOF
        rsMergedResults.AddNew

        For Each mergedField In rsMergedResults.fields
            On Error Resume Next
            rsMergedResults.fields(mergedField.name).value = rsFirst.fields(mergedField.name).value
            On Error GoTo MergeError
        Next mergedField

        rsMergedResults.Update
        rsFirst.MoveNext
    Loop

    '---------------------------------------------------------------------------
    ' Copy records from rsSecond to rsMergedResults
    '---------------------------------------------------------------------------
    rsSecond.MoveFirst
    Do Until rsSecond.EOF
        rsMergedResults.AddNew

        For Each mergedField In rsMergedResults.fields
            On Error Resume Next
            rsMergedResults.fields(mergedField.name).value = rsSecond.fields(mergedField.name).value
            On Error GoTo MergeError
        Next mergedField

        rsMergedResults.Update
        rsSecond.MoveNext
    Loop

    '---------------------------------------------------------------------------
    ' Commit the records to the merged recordset
    '---------------------------------------------------------------------------
    rsMergedResults.UpdateBatch

    Exit Sub

'---------------------------------------------------------------------------
' Local error handler — keeps merge failures isolated
'---------------------------------------------------------------------------
MergeError:
    LogDiagnostic _
        "MergeRecordsets(): " & Err.Description, _
        errorNumber:=Err.number, _
        errorCategory:="ADO / Merge"

    On Error Resume Next
    If Not rsMergedResults Is Nothing Then
        If rsMergedResults.State = adStateOpen Then rsMergedResults.Close
    End If
    Set rsMergedResults = Nothing
End Sub

