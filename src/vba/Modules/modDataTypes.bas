Attribute VB_Name = "modDataTypes"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Bootstrap")

Option Explicit

Public Type stylesWorksheet
    headingRow As Long                           ' "styles" worksheet heading row
    firstRow As Long                             ' First row of style definitions to use
    lastRow As Long                              ' Last row of style definitions to use. If = 0, use all rows

    flagColumn As Long                           ' Column number where comment indicator ('#') is located
    nameColumn As Long                           ' Column number where Style name is located
    formatColumn As Long                         ' Column number where style attributes such as font associated with the style is located
    typeColumn As Long                           ' Column number where Object Type (NODE/EDGE/NATIVE etc) is located
    firstYesNoColumn As Long                     ' Column number where Yes/No switches begin
    selectedViewColumn As Long                   ' Column number where Yes/No switch to include the Style during rendering is kept
    
    suffixOpen As String                         ' Value to append to subgraph-open style names when created by the Style Designer
    suffixClose As String                        ' Value to append to subgraph-close style names when created by the Style Designer
End Type

Public Type dataWorksheet
    worksheetName As String                      ' Worksheet containing the data
    
    headingRow As Long                           ' "data" worksheet heading row
    firstRow As Long                             ' First row of data to use
    lastRow As Long                              ' Last row of data to use. If = 0, use all rows

    flagColumn As Long                           ' Column number where comment indicator ('#') is located
    itemColumn As Long                           ' Column number where Item ID is located
    labelColumn As Long                          ' Column number where Label is located
    xLabelColumn As Long                         ' Column number where External Label is located
    tailLabelColumn As Long                      ' Column number where Edge Tail Label is located
    headLabelColumn As Long                      ' Column number where Edge Head Label is located
    tooltipColumn As Long                        ' Column number where Tooltip is located
    isRelatedToItemColumn As Long                ' Column number where related Item ID is located
    styleNameColumn As Long                      ' Column number where Style name is located
    extraAttributesColumn As Long                ' Column number where line item style attributes are located
    errorMessageColumn As Long                   ' Column number to write error messages to
    graphDisplayColumn As Long                   ' Column number where graph can be displayed in the data worksheet
    graphDisplayColumnAsAlpha As String          ' Column letter where graph can be displayed in the data worksheet
End Type

Public Type DataWorksheetHeadings
    flag As String                               ' Comment indicator ('#') column heading
    Item As String                               ' Item ID column heading
    label As String                              ' Label column heading
    xLabel As String                             ' External Label column heading
    tailLabel As String                          ' Edge Tail Label column heading
    headLabel As String                          ' Edge Head Label column heading
    tooltip As String                            ' Tooltip column heading
    isRelatedToItem As String                    ' related Item ID column heading
    styleName As String                          ' Style name column heading
    extraAttributes As String                    ' line item style attributes column heading
    errorMessage As String                       ' error messages column heading
End Type

Public Type sqlWorksheet
    headingRow As Long                           ' "source" worksheet heading row
    firstRow As Long                             ' First row of sql data
    lastRow As Long                              ' Last row of sql data

    flagColumn As Long                           ' Column number where comment indicator ('#') is located
    sqlStatementColumn As Long                   ' Column number where SQL statement is located
    excelFileColumn As Long                      ' Column number where full path to Excel data file is located
    statusColumn As Long                         ' Column number where status messages are located
End Type

Public Type svgWorksheet
    headingRow As Long                           ' "source" worksheet heading row
    firstRow As Long                             ' First row of sql data
    lastRow As Long                              ' Last row of sql data

    flagColumn As Long                           ' Column number where comment indicator ('#') is located
    findColumn As Long                           ' Column number where find string is located
    replaceColumn As Long                        ' Column number where replace string is located
End Type

Public Type sourceWorksheet
    headingRow As Long                           ' "sql" worksheet heading row
    firstRow As Long                             ' First row of sql data

    lineNumberColumn As Long                     ' Column number where line number is located
    sourceColumn As Long                         ' Column number where Graphviz source code is located
    indent As Long                               ' Number of spaces in a tab indent
End Type

Public Type FileOutput
    directory As String                          ' Where the diagram should be written to
    fileNamePrefix As String                     ' The base portion of the file name
    appendTimeStamp As Boolean                   ' Switch which controls if date and time is to be appended to file name
    appendOptions As Boolean                     ' Switch which controls if the engine and spline settings are appended to the file name
    date As String                               ' Date when the code was run
    time As String                               ' Time when the code was run
End Type

Public Type graphOptions
    addStrict As Boolean                         ' Designates if the 'strict' keyword should be applied to the parent graph
    blankEdgeLabels As Boolean                   ' How to handle blank edge labels. = TRUE -> use Graphviz default behavior
    blankNodeLabels As Boolean                   ' How to handle blank node labels. = TRUE -> use Graphviz default behavior
    center As Boolean
    clusterrank As String
    command As String                            ' Derived from graphType
    compound As Boolean
    concentrate As Boolean
    debug As Boolean                             ' Turn debug tracing on/off
    edgeOperator As String                       ' Derived from graphType
    engine As String                             ' The Graphviz executable which will draw the graph
    fileDisposition As String                    ' What to do with the .gv file after graphing (keep/delete)
    forceLabels As Boolean
    graphType As String                          ' Specifies if graph is directed or undirected
    imagePath As String                          ' Directory paths where images referenced in styles can be found
    imageTypeFile As String                      ' Type of image to create when "Graph to File" is pressed
    imageTypeWorksheet As String                 ' Type of image to create when "Graph to Worksheet" is pressed
    imageWorksheet As String                     ' Worksheet to display the graph in when "Graph to Worksheet" is pressed
    includeGraphImagePath As Boolean             ' On/off switch for graph "imagepath" attribute
    includeEdgeHeadLabels As Boolean             ' On/off switch for edge head labels
    includeEdgeLabels As Boolean                 ' On/off switch for edge labels
    includeEdgePorts As Boolean                  ' On/off switch for edge ports
    includeEdgeTailLabels As Boolean             ' On/off switch for edge tail labels
    includeEdgeXLabels As Boolean                ' On/off switch for edge xlabels
    includeExtraAttributes As Boolean            ' On/off switch to include the "data" worksheet "Extra Attributes" column with the style
    includeNodeLabels As Boolean                 ' On/off switch for node labels
    includeNodeXLabels As Boolean                ' On/off switch for node xlabels
    includeOrphanEdges As Boolean                ' Switch which allows you to drop relationships without defined nodes
    includeOrphanNodes As Boolean                ' Switch which allows you to drop nodes without relationships
    includeStyleFormat As Boolean                ' On/off switch to include the "styles" worksheet Format information associated with the style
    includeTooltip As Boolean                    ' If file format is SVG we include tooltips, otherwise they are excluded
    layout As String
    layoutDim As String
    layoutDimen As String
    mode As String
    model As String
    newrank As Boolean
    options As String                            ' Additional (optional) Graphviz graph options the user may want
    ordering As String
    orientation As Boolean
    outputOrder As String
    overlap As String
    pictureName As String                        ' Name of the graph image when inserted into a worksheet
    postProcessSVG As Boolean
    rankdir As String                            ' For dot engine, controls manner in which shapes are laid out. LR/RL/TB/BT
    scaleImage As Long
    smoothing As String
    splines As String                            ' The type of splines to draw
    strictKeyword As String                      ' Derived from addStrict
    transparentBackground As Boolean             ' Designates if the background color of the graph should be transparent
End Type

Public Type consoleOptions
    logToConsole As Boolean
    appendConsole As Boolean
    graphvizVerbose As Boolean
End Type

' Command Line Options section
Public Type CommandLine
    parameters As String                         ' Custom parameters to pass to the graphing engine
    GraphvizPath As String                       ' Path to dot.exe when user can't modify system path
End Type

Public Type ExchangeOptionsRow
    number As Boolean
    height As Boolean
    visible As Boolean
End Type

Public Type ExchangeOptionsWorksheet
    include As Boolean
    row As ExchangeOptionsRow
    action As String
End Type

' Working variables for 'Exchage' ribbon
Public Type ExchangeOptions
    data As ExchangeOptionsWorksheet
    styles As ExchangeOptionsWorksheet
    sql As ExchangeOptionsWorksheet
    svg As ExchangeOptionsWorksheet
    includeSettings As Boolean
    includeLayouts As Boolean
    includeMetadata As Boolean
End Type

' Working variables for the run-time options on the "settings" worksheet
Public Type settings
    graph As graphOptions                        ' Runtime graph options
    styles As stylesWorksheet                    ' "styles" worksheet settings
    data As dataWorksheet                        ' "data" Worksheet settings
    source As sourceWorksheet                    ' "source" Worksheet settings
    sql As sqlWorksheet                          ' "sql" Worksheet settings
    svg As svgWorksheet                          ' "svg" Worksheet settings
    output As FileOutput                         ' File output settings
    CommandLine As CommandLine                   ' Extra settings for the command line
    console As consoleOptions                    ' console toggle switches
End Type

' Working variables for row data on the "data" worksheet, and values derived from the "styles" worksheet
Public Type dataRow
    comment As String
    Item As String
    relatedItem As String
    label As String
    xLabel As String
    tailLabel As String
    headLabel As String
    tooltip As String
    styleName As String
    extraAttrs As String
    errorMessage As String
    styleType As String                          ' Not a column on "data" worksheet. Value is derived from the style associated with the style name.
    format As String                             ' Not a column on "data" worksheet. Value is derived from the format associated with the style name.
End Type

' Working variables for row data on the "stylesheet" worksheet
Public Type StylesRow
    comment As String
    styleName As String
    format As String
    styleType As String
    show As String
End Type

' Working variables for row data on the "sql" worksheet
Public Type sqlRow
    comment As String
    sqlStatement As String
    excelFile As String
    status As String
End Type

' Used to interpret SQL Field values
Public Type sqlFieldName
    Cluster As String
    clusterLabel As String
    clusterStyleName As String
    clusterAttributes As String
    clusterTooltip As String
    clusterPlaceholder As String
    subcluster As String
    subclusterLabel As String
    subclusterStyleName As String
    subclusterAttributes As String
    subclusterTooltip As String
    subclusterPlaceholder As String
    recordsetPlaceholder As String
    filterColumn As String
    filterValue As String
    splitLength As String
    lineEnding As String
End Type

' Working variables for row data on the "svg" worksheet
Public Type svgRow
    comment As String
    find As String
    replace As String
End Type

' For passing labels to Style Designer functions
Public Type LabelSet
    label As String
    xLabel As String
    headLabel As String
    tailLabel As String
End Type

