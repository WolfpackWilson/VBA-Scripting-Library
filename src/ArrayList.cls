VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

''=======================================================
'' Class:       ArrayList
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 6/23/2020    TheEric960      Written
''=======================================================

' ===========================================
' VARIABLES AND INITIALIZATION
' ===========================================

Private list As Collection
Private size As Long

''initialize the list
Private Sub Class_Initialize()
    size = 0
    Set list = New Collection
End Sub

' ===========================================
' GETTERS AND SETTERS
' ===========================================

''get the size of the list
Public Property Get Count() As Long
    Count = list.Count
End Property

' ===========================================
' METHODS
' ===========================================

''adds an item to the end of the list
Public Sub Add(Item As Variant)
    list.Add Item
    size = list.Count
End Sub

''get an item at a specified position
Public Function Item(Pos As Long) As Variant
    If (Pos < 0 Or Pos >= size) Then
        Err.Raise 9, "ArrayList.Item", "Out of bounds"
    Else
        If (IsObject(list(Post + 1))) Then
            Set Item = list(Pos + 1)
        Else
            Item = list(Pos + 1)
        End If
    End If
End Function

''get a copy of the list
Public Function Clone() As Variant
    Dim tmp As ArrayList
    Set tmp = Me
    Set Clone = tmp
End Function

''get an array from the list
Public Function ToArray() As Variant
    Dim arr() As Variant
    ReDim arr(size - 1)
    
    For i = 1 To size
        If (IsObject(list(i))) Then
            Set arr(i - 1) = list(i)
        Else
            arr(i - 1) = list(i)
        End If
    Next
End Function

''test if something exists in the list
Public Function Contains(Item As Variant) As Variant
    Dim flag As Boolean
    flag = False
    
    For Each listItem In list
        If (Item = listItem) Then
            flag = True
            Exit For
        End If
    Next
    
    Contains = flag
End Function

''find the position of an item
Public Function IndexOf(Item As Variant) As Variant
    Dim Pos As Long
    Pos = 0
    
    For Each listItem In list
        If (Item = listItem) Then
            IndexOf = Pos
            Exit Function
        End If
        
        Pos = Pos + 1
    Next
    
    Err.Raise 448, "ArrayList.IndexOf", "Item not found"
End Function

''add and item at a position
Public Sub Insert(Pos As Long, Item As Variant)
    If (Pos < 0 Or Pos > size) Then
        Err.Raise 9, "ArrayList.Insert", "Out of bounds"
    End If
    
    list.Add Item, Pos + 1
    size = list.Count
End Sub

''clear the entire list
Public Sub Clear()
    Set list = New Collection
    size = list.Count
End Sub

''remove a specified entry via position
Public Sub RemoveAt(Pos As Long)
    If (Pos < 0 Or Pos >= size) Then
        Err.Raise 9, "ArrayList.RemoveAt", "Out of bounds"
    End If
    
    list.remove Pos + 1
    size = list.Count
End Sub

''remove a specified entry
Public Sub remove(Item As Variant)
    RemoveAt IndexOf(Item)
End Sub

''remove a range
Public Sub RemoveRange(Pos1 As Long, Pos2 As Long)
    If (Pos1 < 0 Or Pos1 >= size Or Pos2 < 0 Or Pos2 >= size) Then
        Err.Raise 9, "ArrayList.RemoveRange", "Out of bounds"
    ElseIf (Pos1 >= Pos2) Then
        Err.Raise 380, "ArrayList.RemoveRange", "Upper bound is >= to lower bound"
    End If
    
    Dim repeat As Long
    repeat = Pos2 - Pos1 + 1
    
    For i = 1 To repeat
        RemoveAt Pos1
    Next
End Sub

''reverse the list ordering
Public Sub Reverse()
    Dim newList As Collection
    Set newList = New Collection
    
    For i = 1 To size
        If (IsObject(list(i))) Then
            Set newList(size - i + 1) = list(i)
        Else
            newList(size - i + 1) = list(i)
        End If
    Next
End Sub

''sort the list [using insertion sort]
Public Sub Sort()
    Dim tmp As Variant
    Dim j As Long
    
    For i = 1 To size
        j = i
        
        While (j > 1 And list(j - 1) > list(j))
            Swap j, j - 1
            j = j - 1
        Wend
    Next
End Sub

''swap two items in a list
Private Sub Swap(Pos1 As Long, Pos2 As Long)
    Dim tmp As Variant
    
    If (IsObject(Pos1)) Then
        Set tmp = list(Pos1)
    Else
        tmp = list(Pos1)
    End If
    
    If (IsObject(Pos2)) Then
        Set list(Pos1) = list(Pos2)
    Else
        list(Pos1) = list(Pos2)
    End If
    
    If (IsObject(Pos1)) Then
        Set list(Pos2) = list(Pos1)
    Else
        list(Pos2) = list(Pos1)
    End If
End Sub
