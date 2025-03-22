VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Open Source")
'@IgnoreModule

#If Not Mac Then

#If VBA7 Then
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
Private Declare Function getFrequency Lib "kernel32" Alias _
"QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function getTickCount Lib "kernel32" Alias _
"QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If

'This content is released under the (http://opensource.org/licenses/MIT) MIT License.
'Copyright (c) 2016 Daniel Hubmann (hubisan@gmail.com)

'*************************************************************************************

'Class Module: stopwatch
'Author: Daniel Hubmann <hubisan@gmail.com>
'Url: https://github.com/hubisan/excel-vba-stopwatch.cls

'Start a timer and get elapsed time in ms or seconds
'Use start again if you want to restart the timer.
'It's also possible to stop the timer, but there are not many use cases for that.

'Example usage:
'Sub example()
'    Dim x As New Stopwatch
'    x.start
'    'do your stuff here
'    Debug.Print x.Elapsed_ms
'    Debug.Print x.Elapsed_sec
'    Debug.Print x.Elapsed_sec(0)
'End Sub

'METHODS & PROPERTIES
'-------------------------------------------------------------------------------------
'Public Sub start()
'Public Sub stop()
'Public Function Elapsed_ms() As Double
'Public Function Elapsed_sec(Optional number_of_digits_after_decimal As Integer = 3) As Double
'Public Function get_time() As Double

'*************************************************************************************

Option Explicit

Private start_time As Double
Private end_time As Double

'-------------------------------------------------------------------------------------
'Start the timer
'Sets the start time to used later on to calculate elapsed time
'-------------------------------------------------------------------------------------
Public Sub start()
    start_time = get_time
    end_time = 0
End Sub

'-------------------------------------------------------------------------------------
'Stop the timer
'Sets the end time
'-------------------------------------------------------------------------------------
Public Sub stop_it()
    end_time = get_time
End Sub

'-------------------------------------------------------------------------------------
'Get the elapsed time in milliseconds
'@return Dbl: the elapsed time in milliseconds
'-------------------------------------------------------------------------------------
Public Function Elapsed_ms() As Double
    If end_time = 0 Then
        Elapsed_ms = Round((get_time - start_time) * 1000, 0)
    Else
        Elapsed_ms = Round((end_time - start_time) * 1000, 0)
    End If
End Function

'-------------------------------------------------------------------------------------
'get the elapsed time in seconds, define the number of digits after decimal (default 3)
'@number_of_digits_after_decimal Int optional default=3: number of digits after decimal
'@return Dbl: the elapsed time in seconds
'-------------------------------------------------------------------------------------
Public Function Elapsed_sec(Optional number_of_digits_after_decimal As Integer = 3) As Double
    If end_time = 0 Then
        Elapsed_sec = Round(get_time - start_time, number_of_digits_after_decimal)
    Else
        Elapsed_sec = Round(end_time - start_time, number_of_digits_after_decimal)
    End If
End Function

'-------------------------------------------------------------------------------------
'get time in seconds, accuracy is microseconds
'@return: time in seconds, accuracy microseconds AS Double
'-------------------------------------------------------------------------------------
Public Function get_time() As Double
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    get_time = 0
    If cyFrequency = 0 Then getFrequency cyFrequency
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    If cyFrequency Then get_time = cyTicks2 / cyFrequency
End Function

#End If
