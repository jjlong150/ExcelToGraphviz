VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' CLASS:     Style
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved
'
' DESCRIPTION:
'   A simple data container (DTO) used to store the mapping between an
'   Excel style name and its Graphviz implementation.
'
'   This class is instantiated during the 'CacheEnabledStyles' phase to
'   allow for O(1) high-speed dictionary lookups during the main
'   DOT generation loop.
' ==========================================================================

''
' The category of the style.
' Expected values: TYPE_NODE ("node"), TYPE_EDGE ("edge"),
' TYPE_SUBGRAPH_OPEN ("{"), or TYPE_KEYWORD.
'
'@Ignore VariableNotUsed, MoveFieldCloserToUsage, EncapsulatePublicField
Public styleType As String

''
' The raw Graphviz attribute string (e.g., "shape=box color=blue").
' This string is injected directly into the DOT source during the
' 'ProcessNode' or 'ProcessEdge' phase.
'
'@Ignore VariableNotUsed, MoveFieldCloserToUsage, EncapsulatePublicField
Public styleFormat As String



