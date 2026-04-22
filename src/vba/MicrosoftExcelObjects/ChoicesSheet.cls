VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChoicesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: ChoicesSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Internal
'
' ROLE:
'   Hidden worksheet used as a high-performance staging area for dynamically
'   filtered and sorted lists that feed Data Validation dropdowns across the
'   Data, Styles, and Designer worksheets.
'
' DESCRIPTION:
'   This worksheet module supports the styling and data-entry subsystems by
'   maintaining temporary, calculation-ready lists. It acts as an isolated
'   scratchpad where the engine can assemble context-specific style names,
'   filtered subsets, and other dropdown sources without polluting the
'   visible worksheets. Although the sheet itself contains no UI, its module
'   participates in the real-time regeneration of Named Ranges used by
'   validation lists throughout the workbook.
'
' KEY RESPONSIBILITIES:
'   - DATA ISOLATION:
'       • Hosts transient lists used for dropdown population
'       • Keeps Data and Styles sheets free of intermediate calculations
'
'   - DROPDOWN SYNCHRONIZATION:
'       • Maintains RANGE_STYLE_NAMES and other Named Ranges
'       • Supports dynamic resizing and sorting triggered by selection changes
'
'   - PERFORMANCE OPTIMIZATION:
'       • Provides a fast, hidden workspace for recalculating filtered lists
'       • Ensures Data-sheet dropdowns remain instantaneous and context-aware
'
' USAGE:
'   - Never interacted with directly by users
'   - Serves as the backend list generator for Data Validation dropdowns
'
' RELATED WIKI PAGES:
'   - Data Validation Architecture
'   - Style Designer Dropdown Pipeline
'   - Worksheet Isolation & Performance Strategy
' =============================================================================

Option Explicit

