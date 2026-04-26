Attribute VB_Name = "modExecuteAndCapture"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modExecuteAndCapture
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Bootstrap / Win32 Execution Subsystem
'
' ROLE:
'   High-performance Win32 process-execution and pipe-capture engine. Provides
'   asynchronous, deadlock-free StdOut/StdErr retrieval for external tools
'   (primarily Graphviz) with full 32/64-bit compatibility and macOS stubbing.
'
' RESPONSIBILITIES:
'   - Spawn external processes silently:
'       • CreateProcessA with hidden window (SW_HIDE)
'       • Inheritable pipe handles for StdOut and StdErr
'   - Real-time stream capture:
'       • Non-blocking PeekNamedPipe polling
'       • Chunked 4096-byte reads to prevent pipe saturation
'       • Continuous draining during process execution to avoid 4KB deadlock
'   - Cross-platform resilience:
'       • macOS stub implementation to avoid unsupported WinAPI calls
'       • PtrSafe/LongPtr parity for 32/64-bit Office
'   - Resource hygiene:
'       • Deterministic closing of process, thread, and pipe handles
'       • Safe ByRef return of both output streams
'
' ARCHITECTURAL NOTES:
'   - Based on Christos Samaras' MIT-licensed implementation; heavily hardened
'     for long-running Graphviz workloads and large console output.
'   - Uses SECURITY_ATTRIBUTES, STARTUPINFO, PROCESS_INFORMATION, and the full
'     Win32 pipe/handle lifecycle.
'   - Designed to eliminate UI blocking and buffer back-pressure stalls.
'   - Integrated into the Graphviz execution pipeline via modCreateGraph.
'
' USAGE:
'   - Called by the Graphviz class to execute dot.exe and capture console
'     diagnostics, warnings, and error streams.
'   - Suitable for any external CLI tool requiring silent, asynchronous
'     execution with complete output capture.
'
' RELATED WIKI PAGES:
'   - Graphviz Execution Pipeline
'   - Win32 Process & Pipe Architecture
'   - Deadlock Prevention in External Tool Integration
' =============================================================================

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

' ==========================================================================
' PROCEDURE: ExecuteAndCapture
' PURPOSE:
'   Spawns an external process and captures its output streams in real-time.
'
' TECHNICAL WORKFLOW:
'   1. PIPE INITIALIZATION: Creates two anonymous pipes (StdOut and StdErr)
'      using Win32 Security Attributes to allow handle inheritance.
'   2. STARTUP CONFIGURATION: Forces the external process window to remain
'      hidden (SW_HIDE) to provide a seamless "integrated" feel.
'   3. PROCESS CREATION: Launches the command-line (e.g., dot.exe) and
'      immediately closes the write-end handles to prevent pipe-locks.
'   4. NON-BLOCKING READ LOOP:
'      - Continuously polls both Output and Error pipes while the process
'        is still active.
'      - Prevents the 4KB buffer deadlock by reading data as it is produced.
'   5. RESOURCE RECLAMATION: Systematically closes all process, thread, and
'      pipe handles to ensure no lingering system artifacts remain.
' ==========================================================================
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

' ==========================================================================
' FUNCTION: ReadPipe
' PURPOSE:
'   Drains data from a Windows pipe and converts it into a VBA string.
'
' TECHNICAL WORKFLOW:
'   1. PEEKING: Uses 'PeekNamedPipe' to check if data is waiting without
'      blocking the execution thread.
'   2. CHUNKED READING: Implements a 'Do...Loop' to read the buffer in
'      segments (PIPE_BUFFER_SIZE), typically 4096 bytes.
'   3. DEADLOCK PREVENTION: By clearing the pipe while the process is
'      still running, it prevents the external EXE from stalling when
'      its standard output buffer is full.
'   4. DATA CONVERSION: Converts the raw byte-array buffer into a readable
'      Unicode string via 'UTF8_To_String' and appends it to the result.
'
' USAGE:
'   - Powering 'ExecuteAndCapture' to handle high-volume console feedback
'     from 'dot.exe' (Graphviz).
' ==========================================================================
#If Win64 Then
Private Function ReadPipe(ByVal hPipe As LongPtr) As String
#Else
Private Function ReadPipe(ByVal hPipe As Long) As String
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
                    ReadPipe = ReadPipe & UTF8_To_String(buffer, bytesRead)
                End If
                peeked = PeekNamedPipe(hPipe, ByVal 0&, 0, ByVal 0&, bytesAvail, ByVal 0&)
            Loop While bytesRead > 0
        End If
    Loop While peeked = True And bytesAvail > 0
End Function

Private Function UTF8_To_String(bytes() As Byte, count As Long) As String
    Dim i As Long, c As Long
    Dim result As String

    result = ""
    i = 0

    Do While i < count
        c = bytes(i)

        If c < &H80 Then
            ' 1-byte ASCII
            result = result & ChrW(c)
            i = i + 1

        ElseIf (c And &HE0) = &HC0 Then
            ' 2-byte sequence
            result = result & ChrW(((c And &H1F) * &H40) Or (bytes(i + 1) And &H3F))
            i = i + 2

        ElseIf (c And &HF0) = &HE0 Then
            ' 3-byte sequence
            result = result & ChrW(((c And &HF) * &H1000) _
                                   Or ((bytes(i + 1) And &H3F) * &H40) _
                                   Or (bytes(i + 2) And &H3F))
            i = i + 3

        Else
            ' Unsupported 4-byte UTF-8 (optional to implement)
            result = result & "?"
            i = i + 1
        End If
    Loop

    UTF8_To_String = result
End Function

#End If
