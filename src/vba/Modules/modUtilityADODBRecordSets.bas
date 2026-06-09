Attribute VB_Name = "modUtilityADODBRecordSets"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityADODBRecordSets
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / ADO SQL / Recordset Merging
'
' ROLE:
'   Virtual-Join engine for merging heterogeneous ADODB.Recordset objects.
'   Provides schema synthesis, field-normalization, and client-side static
'   recordset construction to support SQL features that cannot be expressed
'   through native ADO JOIN semantics.
'
' RESPONSIBILITIES:
'   - Recordset merging:
'       o Combine two open ADODB.Recordset objects into a unified static set
'       o Resolve mismatched or missing fields across sources
'       o Normalize field sizes (e.g., DefinedSize = -1) for provider safety
'   - Schema synthesis:
'       o Construct merged field lists dynamically
'       o Ensure deterministic field ordering and type compatibility
'   - Data population:
'       o Two-pass append strategy for performance and correctness
'       o Batch-update semantics for large merges
'   - Error isolation:
'       o Defensive handling of merge-time failures
'       o Forward errors to the diagnostic logger without interrupting SQL flow
'
' ARCHITECTURAL NOTES:
'   - Late-bound ADO for cross-version compatibility.
'   - Windows-only subsystem (ADO unavailable on macOS).
'   - Consumed by recursive SQL, iterative SQL, enumeration SQL, placeholder SQL,
'     and multi-stage SQL pipelines.
'   - Integrates with modUtilityADODBConstants and modUtilityADODBDiagnosticLogger.
'
' USAGE:
'   - Invoked by SQL engine modules requiring heterogeneous recordset merging.
'   - Supports advanced SQL workflows where ADO cannot natively JOIN sources.
'
' RELATED WIKI PAGES:
'   - Virtual Joins & Recordset Merging
'   - SQL Engine Architecture
'   - Diagnostics & ADO Error Handling
' =============================================================================

Option Explicit

''
' Subroutine:  MergeRecordsets
' Description: Merges two ADODB recordsets (rsFirst and rsSecond) into a single
'              recordset (rsMergedResults) using late binding. The merged
'              recordset is passed by reference. This subroutine handles
'              recordsets with different field lists.
'
' THE VIRTUAL JOIN ENGINE: Merges two open recordsets into a new target.
' 1. DEFENSIVE GUARDS: Ensures both sources are open and positioned at BOF.
' 2. SCHEMA ASSEMBLY:
'    - Iterates through rsFirst to define the initial field list.
'    - Iterates through rsSecond to append unique fields not found in the first.
'    - Normalizes 'fieldSize' to prevent provider errors on variable-length fields.
' 3. DATA POPULATION:
'    - Performs a double-pass (rsFirst then rsSecond) to populate the new structure.
'    - Uses 'UpdateBatch' for high-performance record commitment.
' 4. ERROR ISOLATION: Captures failures (e.g., type mismatches) and logs them
'    to the ADODBDiagnosticLogger without crashing the main pipeline.
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
' @param rsFirst [Object]: The primary source recordset to be merged.
' @param rsSecond [Object]: The secondary source recordset to be merged.
' @param rsMergedResults [ByRef Object]: The resulting merged recordset.
'
Public Sub MergeRecordsets(ByVal rsFirst As Object, _
                           ByVal rsSecond As Object, _
                           ByRef rsMergedResults As Object)

    On Error GoTo MergeError

    Dim field As Object
    Dim mergedField As Object
    Dim FieldExists As Boolean
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
        FieldExists = False
        For Each mergedField In rsMergedResults.fields
            If mergedField.name = field.name Then
                FieldExists = True
                Exit For
            End If
        Next mergedField

        If Not FieldExists Then
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
' Local error handler - keeps merge failures isolated
'---------------------------------------------------------------------------
MergeError:
    LogDiagnostic _
        "MergeRecordsets(): " & err.Description, _
        errorNumber:=err.number, _
        errorCategory:="ADO / Merge"

    On Error Resume Next
    If Not rsMergedResults Is Nothing Then
        If rsMergedResults.State = adStateOpen Then rsMergedResults.Close
    End If
    Set rsMergedResults = Nothing
End Sub

