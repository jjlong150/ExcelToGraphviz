Attribute VB_Name = "modUtilityEnvironment"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Environment")

Option Explicit

'@Ignore MoveFieldCloserToUsage
Private username As String
Private tempDir As String

Public Function SearchPathForFile(ByVal filename As String) As Boolean

    SearchPathForFile = False
    
    Dim path As String
    
    ' Pull the PATH environment variable setting into a string, and split it into
    ' an array of directory names
    Dim splitPath() As String
    splitPath = split(Environ$("path"), SEMICOLON)
    
    ' Iterate through the array
    Dim index As Long
    For index = LBound(splitPath) To UBound(splitPath)
        path = Trim$(splitPath(index))
        If path <> vbNullString Then
            'Ensure path is not enclosed in quotes before concatenating the filename
            path = GetStringBetweenDelimiters(path, Chr$(34), Chr$(34))
            
            ' Add a directory delimiter to the end of the directory if not already present
            If Not EndsWith(path, Application.pathSeparator) Then
                path = path & Application.pathSeparator
            End If
            
            ' Determine if the file exists in this directory
            If FileExists(path & filename) Then
                SearchPathForFile = True
                Exit For
            End If
        End If
    Next index
    
End Function

Public Function FindFileOnPath(ByVal filename As String) As String

    FindFileOnPath = vbNullString
    
    Dim path As String
    
    ' Pull the PATH environment variable setting into a string, and split it into
    ' an array of directory names
    Dim splitPath() As String
    splitPath = split(Environ$("path"), SEMICOLON)
    
    ' Iterate through the array
    Dim index As Long
    For index = LBound(splitPath) To UBound(splitPath)
        path = Trim$(splitPath(index))
        If path <> vbNullString Then
            'Ensure path is not enclosed in quotes before concatenating the filename
            path = GetStringBetweenDelimiters(path, Chr$(34), Chr$(34))
            
            ' Add a directory delimiter to the end of the directory if not already present
            If Not EndsWith(path, Application.pathSeparator) Then
                path = path & Application.pathSeparator
            End If
            
            ' Determine if the file exists in this directory
            If FileExists(path & filename) Then
                FindFileOnPath = path & filename
                Exit For
            End If
        End If
    Next index
    
End Function

Public Sub SetTempDirectory()
#If Mac Then
    tempDir = "/Users/" & GetUsername & "/Library/Containers/com.microsoft.Excel/Data"
#Else
    tempDir = Environ$("temp")
#End If
End Sub

Public Function GetTempDirectory() As String
    GetTempDirectory = tempDir
End Function

Public Function GetUsername() As String
    If username = vbNullString Then
#If Mac Then
        ' Get and then cache the username so we are not continously using a shell command to get this value
        username = MacScript("set userName to short user name of (system info)" & vbNewLine & "return userName")
#Else
        username = Application.username
#End If
        username = Trim$(username)
    End If
    GetUsername = username
End Function

Public Function GetEnvVarSeparator() As String
#If Mac Then
    GetEnvVarSeparator = COLON
#Else
    GetEnvVarSeparator = SEMICOLON
#End If
End Function
