Attribute VB_Name = "modExecuteAndCapture"
'@Folder("Open Source")
'@IgnoreModule UseMeaningfulName, HungarianNotation, VariableNotAssigned, IntegerDataType, DefaultMemberRequired, UnassignedVariableUsage

Option Explicit

' Written by Christos Samaras
' https://myengineeringworld.net/2020/01/call-csharp-console-app-vba.html
' Copyright 2024, Christos Samaras
' MIT License

' Modified by Jeffrey Long to convert from Function to Sub calling convention to
' be able to return two values (using ByRef parameters), add necessary changes
' to prevent errors on MacOS which does not support these WinAPI functions,
' and remediate RubberDuckVBA static code analysis messages.

#If Mac Then
    Public Sub ExecuteAndCapture(ByVal CommandLine As String, ByRef stdOut As String, ByRef stdErr As String)
        Debug.Print CommandLine
        stdOut = vbNullString
        stdErr = "ExecuteAndCapture subroutine is not implemented on MacOS"
    End Sub
#Else
'Declaring the necessary API functions and types based on Excel version.
#If Win64 Then  'For 64 bit Excel.
    ' Creates an anonymous pipe, and returns handles to the read and write ends of the pipe.
    ' https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-createpipe
    Public Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As LongPtr, _
                                                                phWritePipe As LongPtr, _
                                                                lpPipeAttributes As Any, _
                                                                ByVal nSize As Long) As Long
    
    ' Reads data from the specified file or input/output (I/O) device. Reads occur at the position specified by the file pointer if supported by the device.
    ' https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-readfile
    Public Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, _
                                                             lpBuffer As Any, _
                                                             ByVal nNumberOfBytesToRead As Long, _
                                                             lpNumberOfBytesRead As Long, _
                                                             lpOverlapped As Any) As Long
    
    ' Creates a new process and its primary thread. The new process runs in the security context of the calling process.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessa
    Public Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                                                                         ByVal lpCommandLine As String, _
                                                                                         lpProcessAttributes As Any, _
                                                                                         lpThreadAttributes As Any, _
                                                                                         ByVal bInheritHandles As Long, _
                                                                                         ByVal dwCreationFlags As Long, _
                                                                                         lpEnvironment As Any, _
                                                                                         ByVal lpCurrentDirectory As String, _
                                                                                         lpStartupInfo As STARTUPINFO, _
                                                                                         lpProcessInformation As PROCESS_INFORMATION) As Long
    ' Closes an open object handle.
    ' https://learn.microsoft.com/en-us/windows/win32/api/handleapi/nf-handleapi-closehandle
    Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    
    ' Retrieves the termination status of the specified process.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getexitcodeprocess
    Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
    
    ' Contains the security descriptor for an object and specifies whether the handle retrieved by specifying this structure is inheritable.
    ' https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/aa379560(v=vs.85)
    Public Type SECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As LongPtr
        bInheritHandle          As Long
    End Type
    
    ' Specifies the window station, desktop, standard handles, and appearance of the main window for a process at creation time.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-startupinfoa
    Public Type STARTUPINFO
        cb                  As Long
        lpReserved          As String
        lpDesktop           As String
        lpTitle             As String
        dwX                 As Long
        dwY                 As Long
        dwXSize             As Long
        dwYSize             As Long
        dwXCountChars       As Long
        dwYCountChars       As Long
        dwFillAttribute     As Long
        dwFlags             As Long
        wShowWindow         As Integer
        cbReserved2         As Integer
        lpReserved2         As LongPtr
        hStdInput           As LongPtr
        hStdOutput          As LongPtr
        hStdError           As LongPtr
    End Type
    
    ' Contains information about a newly created process and its primary thread.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-process_information
    Public Type PROCESS_INFORMATION
        hProcess        As LongPtr
        hThread         As LongPtr
        dwProcessId     As Long
        dwThreadId      As Long
    End Type

#Else 'For 32 bit Excel.
    ' Creates an anonymous pipe, and returns handles to the read and write ends of the pipe.
    ' https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-createpipe
    Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, _
                                                       phWritePipe As Long, _
                                                       lpPipeAttributes As Any, _
                                                       ByVal nSize As Long) As Long
    
    ' Reads data from the specified file or input/output (I/O) device. Reads occur at the position specified by the file pointer if supported by the device.
    ' https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-readfile
    Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
                                                     lpBuffer As Any, _
                                                     ByVal nNumberOfBytesToRead As Long, _
                                                     lpNumberOfBytesRead As Long, _
                                                     lpOverlapped As Any) As Long
    
    ' Creates a new process and its primary thread. The new process runs in the security context of the calling process.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessa
    Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                                                                 ByVal lpCommandLine As String, _
                                                                                 lpProcessAttributes As Any, _
                                                                                 lpThreadAttributes As Any, _
                                                                                 ByVal bInheritHandles As Long, _
                                                                                 ByVal dwCreationFlags As Long, _
                                                                                 lpEnvironment As Any, _
                                                                                 ByVal lpCurrentDirectory As String, _
                                                                                 lpStartupInfo As STARTUPINFO, _
                                                                                 lpProcessInformation As PROCESS_INFORMATION) As Long
    ' Closes an open object handle.
    ' https://learn.microsoft.com/en-us/windows/win32/api/handleapi/nf-handleapi-closehandle
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    ' Retrieves the termination status of the specified process.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getexitcodeprocess
    Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    
    'Types.
    ' Contains the security descriptor for an object and specifies whether the handle retrieved by specifying this structure is inheritable.
    ' https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/aa379560(v=vs.85)
    Public Type SECURITY_ATTRIBUTES
        nLength                                As Long
        lpSecurityDescriptor                   As Long
        bInheritHandle                         As Long
    End Type
    
    ' Specifies the window station, desktop, standard handles, and appearance of the main window for a process at creation time.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-startupinfoa
    Public Type STARTUPINFO
        cb                                     As Long
        lpReserved                             As String
        lpDesktop                              As String
        lpTitle                                As String
        dwX                                    As Long
        dwY                                    As Long
        dwXSize                                As Long
        dwYSize                                As Long
        dwXCountChars                          As Long
        dwYCountChars                          As Long
        dwFillAttribute                        As Long
        dwFlags                                As Long
        wShowWindow                            As Integer
        cbReserved2                            As Integer
        lpReserved2                            As Long
        hStdInput                              As Long
        hStdOutput                             As Long
        hStdError                              As Long
    End Type
    
    ' Contains information about a newly created process and its primary thread.
    ' https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-process_information
    Public Type PROCESS_INFORMATION
        hProcess                               As Long
        hThread                                As Long
        dwProcessId                            As Long
        dwThreadId                             As Long
    End Type
#End If

'Contants.
Public Const STARTF_USESHOWWINDOW  As Long = &H1
Public Const STARTF_USESTDHANDLES  As Long = &H100
Public Const SW_HIDE               As Integer = 0
Public Const BUFSIZE               As Long = 1024 * 10
Public Const STILL_ACTIVE As Long = &H103

Public Sub ExecuteAndCapture(ByVal CommandLine As String, ByRef stdOut As String, ByRef stdErr As String)

    #If Win64 Then
        Dim hStdOutRead As LongPtr
        Dim hStdOutWrite As LongPtr
        Dim hStdErrRead As LongPtr
        Dim hStdErrWrite As LongPtr
        Dim hProcess As LongPtr
    #Else
        Dim hStdOutRead As Long
        Dim hStdOutWrite As Long
        Dim hStdErrRead As Long
        Dim hStdErrWrite As Long
        Dim hProcess As Long
    #End If
        
    Dim sa As SECURITY_ATTRIBUTES
    With sa
        .nLength = LenB(sa)
        .bInheritHandle = 1
        .lpSecurityDescriptor = 0
    End With
    
    'Create pipes for stdout and stderr
    If CreatePipe(hStdOutRead, hStdOutWrite, sa, 0) = 0 Then Exit Sub
    If CreatePipe(hStdErrRead, hStdErrWrite, sa, 0) = 0 Then
        CloseHandle hStdOutRead
        CloseHandle hStdOutWrite
        Exit Sub
    End If
    
    Dim si As STARTUPINFO
    With si
        .cb = LenB(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = hStdOutWrite
        .hStdError = hStdErrWrite
        .hStdInput = 0
    End With
    
    Dim sStdOut As String
    Dim sStdErr As String
    
    Dim pi As PROCESS_INFORMATION
    
    If CreateProcess(vbNullString, CommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, vbNullString, si, pi) Then
        hProcess = pi.hProcess
        CloseHandle hStdOutWrite
        CloseHandle hStdErrWrite
        
        Dim baStdOut(BUFSIZE) As Byte
        Dim baStdErr(BUFSIZE) As Byte
        
        Dim lBytesReadOut As Long
        Dim lBytesReadErr As Long
        
        Dim lExitCode As Long
        
        Do
            ' Read from stdout pipe
            If ReadFile(hStdOutRead, baStdOut(0), BUFSIZE, lBytesReadOut, ByVal 0&) <> 0 Then
                If lBytesReadOut > 0 Then
                    sStdOut = sStdOut & Left$(StrConv(baStdOut(), vbUnicode), lBytesReadOut)
                End If
            End If
            
            ' Read from stderr pipe.
            If ReadFile(hStdErrRead, baStdErr(0), BUFSIZE, lBytesReadErr, ByVal 0&) <> 0 Then
                If lBytesReadErr > 0 Then
                    sStdErr = sStdErr & Left$(StrConv(baStdErr(), vbUnicode), lBytesReadErr)
                End If
            End If
            
            ' Check if process has exited.
            GetExitCodeProcess hProcess, lExitCode
            If lExitCode <> STILL_ACTIVE Then Exit Do
        Loop
        
        'Read any remaining data.
        Do While ReadFile(hStdOutRead, baStdOut(0), BUFSIZE, lBytesReadOut, ByVal 0&) <> 0 And lBytesReadOut > 0
            sStdOut = sStdOut & Left$(StrConv(baStdOut(), vbUnicode), lBytesReadOut)
        Loop
        Do While ReadFile(hStdErrRead, baStdErr(0), BUFSIZE, lBytesReadErr, ByVal 0&) <> 0 And lBytesReadErr > 0
            sStdErr = sStdErr & Left$(StrConv(baStdErr(), vbUnicode), lBytesReadErr)
        Loop
        
        CloseHandle hProcess
        CloseHandle pi.hThread
    End If
    
    CloseHandle hStdOutRead
    CloseHandle hStdErrRead
    
    stdOut = Trim$(sStdOut)
    stdErr = Trim$(sStdErr)
    
End Sub

#End If


