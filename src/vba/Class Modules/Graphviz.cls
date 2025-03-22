VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graphviz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved
'@Folder("Utility.Graphviz")
'@IgnoreModule ModuleWithoutFolder

Option Explicit

Private m_graphvizSource As String
Private m_graphvizFilename As String
Private m_diagramFilename As String
Private m_graphFormat As String
Private m_graphLayout As String
Private m_commandLineParameters As String
Private m_verbose As Boolean
Private m_captureMessages As Boolean
Private m_graphvizMessages As String
Private m_graphvizCommand As String
Private m_graphvizPath As String
Private m_filenameBase As String
Private m_outputDirectory As String

' Where files should be written to
Public Property Get OutputDirectory() As String
    OutputDirectory = m_outputDirectory
End Property

Public Property Let OutputDirectory(ByVal OutputDirectory As String)
    m_outputDirectory = OutputDirectory
    InitializeGraphvizFilename
    InitializeDiagramFilename
End Property

' Base portion of the file name (i.e. no extension)
Public Property Get FilenameBase() As String
    FilenameBase = m_filenameBase
End Property

Public Property Let FilenameBase(ByVal FilenameBase As String)
    m_filenameBase = FilenameBase
    InitializeGraphvizFilename
    InitializeDiagramFilename
End Property

' Graphviz dot source code
Public Property Get graphvizSource() As String
    graphvizSource = m_graphvizSource
End Property

Public Property Let graphvizSource(ByVal graphvizSource As String)
    m_graphvizSource = graphvizSource
End Property

' Name of file containing Graphviz dot source
Public Property Get GraphvizFilename() As String
    GraphvizFilename = m_graphvizFilename
End Property

Public Property Let GraphvizFilename(ByVal GraphvizFilename As String)
    m_graphvizFilename = GraphvizFilename
End Property

' Name of file for the diagram rendering
Public Property Get DiagramFilename() As String
    DiagramFilename = m_diagramFilename
End Property

Public Property Let DiagramFilename(ByVal DiagramFilename As String)
    m_diagramFilename = DiagramFilename
End Property

' Graph format
Public Property Get GraphFormat() As String
    GraphFormat = m_graphFormat
End Property

Public Property Let GraphFormat(ByVal GraphFormat As String)
    m_graphFormat = GraphFormat
    InitializeGraphvizFilename
    InitializeDiagramFilename
End Property

' Graph Layout
Public Property Get GraphLayout() As String
    GraphLayout = m_graphLayout
End Property

Public Property Let GraphLayout(ByVal GraphLayout As String)
    m_graphLayout = GraphLayout
End Property

' Graphviz Path
Public Property Get GraphvizPath() As String
    GraphvizPath = m_graphvizPath
End Property

Public Property Let GraphvizPath(ByVal GraphvizPath As String)
    m_graphvizPath = GraphvizPath
End Property

' Command Line Parameters
Public Property Get CommandLineParameters() As String
    CommandLineParameters = m_commandLineParameters
End Property

Public Property Let CommandLineParameters(ByVal CommandLineParameters As String)
    m_commandLineParameters = CommandLineParameters
End Property

' Verbose
Public Property Get Verbose() As Boolean
    Verbose = m_verbose
End Property

Public Property Let Verbose(ByVal Verbose As Boolean)
    m_verbose = Verbose
End Property

' Capture Graphviz Messages
Public Property Get CaptureMessages() As Boolean
    CaptureMessages = m_captureMessages
End Property

Public Property Let CaptureMessages(ByVal CaptureMessages As Boolean)
    m_captureMessages = CaptureMessages
End Property

' Graphviz messages
Public Property Get GraphvizMessages() As String
    GraphvizMessages = m_graphvizMessages
End Property

Public Property Let GraphvizMessages(ByVal GraphvizMessages As String)
    m_graphvizMessages = GraphvizMessages
End Property

' Graphviz Command
Public Property Get GraphvizCommand() As String
    GraphvizCommand = m_graphvizCommand
End Property

Public Property Let GraphvizCommand(ByVal GraphvizCommand As String)
    m_graphvizCommand = GraphvizCommand
End Property

Public Sub RenderGraph()
    ' Trap any errors
    On Error GoTo EndRenderGraph:
    
    m_graphvizMessages = vbNullString
    
    Dim verboseFlag As String
    If m_verbose Then
        verboseFlag = " -v"
    Else
        verboseFlag = vbNullString
    End If
    
     ' Generate the diagram which corresponds to the Graphviz file
#If Mac Then
    GraphvizCommand = Chr$(34) & m_graphvizFilename & Chr$(34) & " -T" & m_graphFormat & " -o " & Chr$(34) & m_diagramFilename & Chr$(34) & verboseFlag & " " & m_commandLineParameters
    
    ' Execute the command
    Dim dotOutput As String
    dotOutput = RunAppleScriptTask("runDot", GraphvizCommand)

    If m_captureMessages Then
        m_graphvizMessages = dotOutput
    End If
#Else
    GraphvizCommand = m_graphvizPath & m_graphLayout & " " & Chr$(34) & m_graphvizFilename & Chr$(34) & " -T" & m_graphFormat & " -o " & Chr$(34) & m_diagramFilename & Chr$(34) & verboseFlag & " " & m_commandLineParameters
    
    Dim stdOut As String
    Dim stdErr As String
    ExecuteAndCapture GraphvizCommand, stdOut, stdErr
   
    If m_captureMessages Then
        m_graphvizMessages = stdErr & stdOut
    End If
#End If
    
EndRenderGraph:
    On Error GoTo 0
    
End Sub

Public Sub SourceToFile()
    ' Trap any file system I/O-related errors
    On Error GoTo WriteFailed:

#If Mac Then
    WriteTextToFile m_graphvizSource, m_graphvizFilename
#Else
    WriteTextToUTF8FileFileWithoutBOM m_graphvizSource, m_graphvizFilename
#End If

WriteFailed:
    If Err.number > 0 Then
        MsgBox Err.number & " " & Err.Description, vbOKOnly, "Graphviz.SourceToFile()"
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Private Sub InitializeGraphvizFilename()
    If m_outputDirectory = vbNullString Then
        m_graphvizFilename = m_filenameBase & GRAPHVIZ_EXTENSION
    Else
        m_graphvizFilename = m_outputDirectory & Application.pathSeparator & m_filenameBase & GRAPHVIZ_EXTENSION
    End If
End Sub

Private Sub InitializeDiagramFilename()
    If m_outputDirectory = vbNullString Then
        m_diagramFilename = m_filenameBase & "." & m_graphFormat
    Else
        m_diagramFilename = m_outputDirectory & Application.pathSeparator & m_filenameBase & "." & m_graphFormat
    End If
End Sub

Private Sub Class_Initialize()
    m_verbose = False           ' Run Graphviz in quiet mode
    m_captureMessages = False   ' Do not capture Graphviz std err output
    m_graphFormat = "png"
    m_graphLayout = "dot"
    m_filenameBase = "RelationshipVisualizer"
End Sub
