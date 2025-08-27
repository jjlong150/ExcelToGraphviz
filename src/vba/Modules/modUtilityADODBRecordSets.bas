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

Public Sub MergeRecordsets(ByVal rsFirst As Object, ByVal rsSecond As Object, ByRef rsMergedResults As Object)
    Dim field As Object
    Dim mergedField As Object
    Dim fieldExists As Boolean
    
    ' Create a new recordset to contain the merged records
    Set rsMergedResults = CreateObject("ADODB.Recordset")

    ' Initialize the merged recordset
    rsMergedResults.CursorLocation = CursorLocationEnum.adUseClient
    rsMergedResults.CursorType = CursorTypeEnum.adOpenStatic
    rsMergedResults.LockType = LockTypeEnum.adLockBatchOptimistic
    
    ' Create fields in merged recordset for rsFirst
    For Each field In rsFirst.fields
        rsMergedResults.fields.Append field.name, field.Type, field.DefinedSize, field.attributes
    Next field
    
    ' Create fields in merged recordset for rsSecond if they don't already exist
    For Each field In rsSecond.fields
        fieldExists = False
        For Each mergedField In rsMergedResults.fields
            If mergedField.name = field.name Then
                fieldExists = True
                Exit For
            End If
        Next mergedField
        If Not fieldExists Then
            rsMergedResults.fields.Append field.name, field.Type, field.DefinedSize, field.attributes
        End If
    Next field
    
    ' Open the merged recordset
    rsMergedResults.Open
    
    ' Copy records from rsFirst to rsMergedResults
    Do Until rsFirst.EOF
        rsMergedResults.AddNew
        For Each field In rsMergedResults.fields
            On Error Resume Next ' Skip errors if field is not found
            rsMergedResults.fields(field.name).value = rsFirst.fields(field.name).value
            On Error GoTo 0
        Next field
        rsFirst.MoveNext
    Loop
    
    ' Copy records from rsSecond to rsMergedResults
    rsSecond.MoveFirst ' Ensure the recordset is at the beginning
    Do Until rsSecond.EOF
        rsMergedResults.AddNew
        For Each field In rsMergedResults.fields
            On Error Resume Next ' Skip errors if field is not found
            rsMergedResults.fields(field.name).value = rsSecond.fields(field.name).value
            On Error GoTo 0
        Next field
        rsSecond.MoveNext
    Loop
    
    ' Commit the records to the merged recordset
    rsMergedResults.UpdateBatch
End Sub


