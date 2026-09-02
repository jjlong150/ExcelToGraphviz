VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================
' Class: DotScope
'===========================
' Tracks, for one DOT scope (cluster) depth: which cluster this frame
' belongs to, and the currently-active node[]/edge[]/graph[] label-family
' defaults at that depth. One instance per stack frame, created via
' Clone() from the enclosing frame - never a direct Set from the parent -
' so mutating a child scope can never reach back and corrupt the parent.

Public graphLabel As String
Public graphTooltip As String
Public graphRow As Long
Public clusterName As String
Public nodeLabel As String
Public nodeXLabel As String
Public nodeTooltip As String
Public edgeLabel As String
Public edgeXLabel As String
Public edgeTailLabel As String
Public edgeHeadLabel As String
Public edgeTooltip As String

' Returns an independent copy of this frame's current state, safe to
' mutate without affecting the original. NOTE: this copies clusterName
' too - the caller MUST overwrite it with the new cluster's own id right
' after cloning, or the child frame will silently keep reporting the
' parent's cluster name.
Public Function Clone() As DotScope
    Dim copy As New DotScope
    
    copy.graphLabel = graphLabel
    copy.graphTooltip = graphTooltip
    copy.graphRow = graphRow
    
    copy.clusterName = clusterName
    
    copy.nodeLabel = nodeLabel
    copy.nodeXLabel = nodeXLabel
    copy.nodeTooltip = nodeTooltip
    
    copy.edgeLabel = edgeLabel
    copy.edgeXLabel = edgeXLabel
    copy.edgeTailLabel = edgeTailLabel
    copy.edgeHeadLabel = edgeHeadLabel
    copy.edgeTooltip = edgeTooltip
    
    Set Clone = copy
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = (Len(graphLabel) = 0) And _
              (Len(graphTooltip) = 0) And _
              (Len(nodeLabel) = 0) And _
              (Len(nodeXLabel) = 0) And _
              (Len(nodeTooltip) = 0) And _
              (Len(edgeLabel) = 0) And _
              (Len(edgeXLabel) = 0) And _
              (Len(edgeTailLabel) = 0) And _
              (Len(edgeHeadLabel) = 0) And _
              (Len(edgeTooltip) = 0) And _
              graphRow = 0
End Function

Public Function ToDebugString() As String
    ToDebugString = "clusterName=" & clusterName & _
        " | graphLabel=" & graphLabel & _
        " | graphTooltip=" & graphTooltip & _
        " | nodeLabel=" & nodeLabel & _
        " | nodeXLabel=" & nodeXLabel & _
        " | nodeTooltip=" & nodeTooltip & _
        " | edgeLabel=" & edgeLabel & _
        " | edgeXLabel=" & edgeXLabel & _
        " | edgeTailLabel=" & edgeTailLabel & _
        " | edgeHeadLabel=" & edgeHeadLabel & _
        " | edgeTooltip=" & edgeTooltip & _
        " | graphRow=" & graphRow
End Function

