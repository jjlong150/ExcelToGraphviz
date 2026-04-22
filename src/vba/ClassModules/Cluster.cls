VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cluster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' CLASS:     Cluster
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved
'
' DESCRIPTION:
'   A data container used to represent Graphviz 'clusters' or
'   'subgraphs'. It stores the unique identity and visual attributes
'   of a group.
'
'   The inclusion of a 'subclusters' Dictionary allows this class to act
'   as a node in a tree structure, enabling the engine to model deeply
'   nested parent-child relationships found in complex SQL resultsets.
'
' PROPERTIES:
'   - id: Unique Graphviz identifier (e.g., cluster_01).
'   - label: The display title shown within the bounding box.
'   - Tooltip: Interactive metadata (SVG only).
'   - styleName: Reference to the Style Gallery definition.
'   - attributes: Raw DOT attribute string overrides.
'   - subclusters: A Dictionary collection of child Cluster objects.
' ==========================================================================

Public id As String
Public label As String
Public Tooltip As String
Public styleName As String
Public attributes As String
Public subclusters As Dictionary

