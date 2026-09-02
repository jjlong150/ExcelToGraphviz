VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================
' Class: Stack
'===========================

Private items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

' Push any value or object reference onto the stack
Public Sub Push(ByVal item As Variant)
    items.Add item
End Sub

' Pop the top item off the stack
Public Function Pop() As Variant
    If items.Count = 0 Then
        Pop = Empty
        Exit Function
    End If

    If IsObject(items(items.Count)) Then
        Set Pop = items(items.Count)
    Else
        Pop = items(items.Count)
    End If
    items.Remove items.Count
End Function

' Peek at the top item without removing it
Public Function Peek() As Variant
    If items.Count = 0 Then
        Peek = Empty
        Exit Function
    End If

    If IsObject(items(items.Count)) Then
        Set Peek = items(items.Count)
    Else
        Peek = items(items.Count)
    End If
End Function

' Return number of items on the stack
Public Function Count() As Long
    Count = items.Count
End Function

' True if stack is empty
Public Function IsEmpty() As Boolean
    IsEmpty = (items.Count = 0)
End Function

