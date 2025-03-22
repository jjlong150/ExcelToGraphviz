Attribute VB_Name = "modExecuteAndCapture"
'@Folder("Open Source")
'@IgnoreModule UseMeaningfulName, HungarianNotation, VariableNotAssigned, IntegerDataType, DefaultMemberRequired, UnassignedVariableUsage

Option Explicit

' ********************************************************************************************
' Module: modExecuteAndCapture
' Description: This module contains the `ExecuteAndCapture` function, which uses the `CreateProcessA`
'              API to execute command-line commands and capture the standard output and error
'              asynchronously. The function is designed to work on both 32-bit and 64-bit
'              versions of Excel.
'
' Acknowledgements:
' - Originally based off of "ExecuteAndCapture" Written by Christos Samaras
'   https://myengineeringworld.net/2020/01/call-csharp-console-app-vba.html
'   Copyright 2024, Christos Samaras
'   MIT License
'
' - The code has been adapted and refined by Jeffrey Long to address specific issues such as preventing
'   deadlocks when reading from pipes, handling large output, and ensuring compatibility with different
'   versions of Excel.
'   - Modified to convert from Function to Sub calling convention to be able to return two values
'    (using ByRef parameters).
'   - Modified to add necessary changes to prevent errors on macOS which does not support the
'     WinAPI functions used.
'   - Modified to remediate RubberDuckVBA static code analysis messages.
'   - Modified to fix a deadlock occurring in original implementation if more than 4096 bytes
'     was written to standard out, or standard error by the process which was invoked.
'     In this situation the process would pause waiting for data to be read from
'     and removed from the pipe before it could write additional data to the pipe. Now the pipes
'     are read as the process runs, as opposed to when the process completes.
'
' - The function utilizes Windows API calls such as `CreateProcessA`, `WaitForSingleObject`,
'   `CloseHandle`, `ReadFile`, `CreatePipe`, and `PeekNamedPipe` to execute and manage
'   command-line processes.
'
' - Special thanks to the VBA community and various online resources for providing insights
'   and code snippets related to handling Windows API calls and asynchronous I/O operations
'   in VBA.
'
' ********************************************************************************************

#If Mac Then
    Public Sub ExecuteAndCapture(ByVal CommandLine As String, ByRef stdOut As String, ByRef stdErr As String)
        Debug.Print CommandLine
        stdOut = vbNullString
        stdErr = "ExecuteAndCapture subroutine is not implemented on MacOS"
    End Sub
#Else
'Declaring the necessary API functions and types based on Excel version.
#If Win64 Then  'For 64 bit Excel.
    ' ----------------------------------
    ' 64-bit Windows API signatures
    ' ----------------------------------

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
    
    ' Copies data from a named or anonymous pipe into a buffer without removing it from the pipe. It also returns information about data in the pipe.
    ' https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-peeknamedpipe
    Public Declare PtrSafe Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByVal lpBuffer As String, ByVal nBufferSize As Long, ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByRef lpBytesLeftThisMessage As Long) As Long
    
    ' ----------------------------------
    ' 64-bit Data Types
    ' ----------------------------------
    
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
    ' ----------------------------------
    ' 32-bit Windows API signatures
    ' ----------------------------------

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
    
    ' Copies data from a named or anonymous pipe into a buffer without removing it from the pipe. It also returns information about data in the pipe.
    ' https://learn.microsoft.com/en-us/windows/win32/api/namedpipeapi/nf-namedpipeapi-peeknamedpipe
    Public Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByVal lpBuffer As String, ByVal nBufferSize As Long, ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByRef lpBytesLeftThisMessage As Long) As Long

    ' ----------------------------------
    ' 32-bit Data Types
    ' ----------------------------------
    
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

' ----------------------------------
' Contants
' ----------------------------------

' StartupInfo constants
Public Const STARTF_USESHOWWINDOW  As Long = &H1
Public Const STARTF_USESTDHANDLES  As Long = &H100
Public Const SW_HIDE               As Integer = 0

' GetExitProcessCode() constants
Public Const STILL_ACTIVE As Long = &H103

' ReadFile() constants
Public Const PIPE_BUFFER_SIZE  As Long = 1024 * 4

Public Sub ExecuteAndCapture(ByVal CommandLine As String, ByRef stdOut As String, ByRef stdErr As String)

    ' Declare pipe handles
    #If Win64 Then
        Dim hStdOutRead As LongPtr
        Dim hStdOutWrite As LongPtr
        Dim hStdErrRead As LongPtr
        Dim hStdErrWrite As LongPtr
    #Else
        Dim hStdOutRead As Long
        Dim hStdOutWrite As Long
        Dim hStdErrRead As Long
        Dim hStdErrWrite As Long
    #End If
        
    ' Initialize security attributes
    Dim sa As SECURITY_ATTRIBUTES
    With sa
        .nLength = LenB(sa)
        .bInheritHandle = 1
        .lpSecurityDescriptor = 0
    End With
    
    'Create pipes for standard output and standard error
    If CreatePipe(hStdOutRead, hStdOutWrite, sa, 0) = 0 Then Exit Sub
    If CreatePipe(hStdErrRead, hStdErrWrite, sa, 0) = 0 Then
        CloseHandle hStdOutRead
        CloseHandle hStdOutWrite
        Exit Sub
    End If
    
    ' Set up startup info
    Dim si As STARTUPINFO
    With si
        .cb = LenB(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = hStdOutWrite
        .hStdError = hStdErrWrite
        .hStdInput = 0
    End With
    
    ' Set up process information
    Dim pi As PROCESS_INFORMATION
    
    ' Execute the command
    If CreateProcess(vbNullString, CommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, vbNullString, si, pi) Then
        ' Close the write ends of the pipes
        CloseHandle hStdOutWrite
        CloseHandle hStdErrWrite
        
        ' Read output and error during process execution
        Dim lExitCode As Long
        Do
            ' Read from standard output pipe
            stdOut = stdOut & ReadPipe(hStdOutRead)
            
            ' Read from standare error pipe.
            stdErr = stdErr & ReadPipe(hStdErrRead)

            ' Check if process has exited.
            GetExitCodeProcess pi.hProcess, lExitCode
            If lExitCode <> STILL_ACTIVE Then Exit Do
        Loop
        
        ' Close handles to process and threads
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
    End If
    
    ' Close handles to pipes
    CloseHandle hStdOutRead
    CloseHandle hStdErrRead
End Sub

#If Win64 Then
Function ReadPipe(hPipe As LongPtr) As String
#Else
Function ReadPipe(hPipe As Long) As String
#End If

    Dim peeked As Boolean
    Dim bytesAvail As Long
    Dim bytesRead As Long
    Dim buffer(PIPE_BUFFER_SIZE) As Byte
    
    ' PeekNamedPipe returns information about data in the pipe without removing it from the pipe.
    peeked = PeekNamedPipe(hPipe, ByVal 0&, 0, ByVal 0&, bytesAvail, ByVal 0&)
    Do
        ' Determine if any data was written to the pipe
        If bytesAvail > 0 Then
            ' The pipe can only accept a maximum of 4096 bytes from the process, at which point
            ' the process will pause until data is read & removed from the pipe, allowing the
            ' process to proceed. Read the pipe in chunks of 4096 bytes.
            Do
                bytesRead = 0
                If ReadFile(hPipe, buffer(0), PIPE_BUFFER_SIZE, bytesRead, ByVal 0&) = 0 Then Exit Do
                If bytesRead > 0 Then
                    ReadPipe = ReadPipe & Left$(StrConv(buffer(), vbUnicode), bytesRead)
                End If
                peeked = PeekNamedPipe(hPipe, ByVal 0&, 0, ByVal 0&, bytesAvail, ByVal 0&)
            Loop While bytesRead > 0
        End If
    Loop While peeked = True And bytesAvail > 0
End Function

#End If






