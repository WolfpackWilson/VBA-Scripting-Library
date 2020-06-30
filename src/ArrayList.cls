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
'' Version:     0.1.0
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

    If (size = 0) Then
        ToArray = arr
        Exit Function
    End If
    
    ReDim arr(size - 1)
    
    For i = 1 To size
        If (IsObject(list(i))) Then
            Set arr(i - 1) = list(i)
        Else
            arr(i - 1) = list(i)
        End If
    Next
    
    ToArray = arr
End Function

''test if something exists in the list
Public Function Contains(Item As Variant) As Variant
    Dim flag As Boolean
    flag = False
    
    For Each listItem In list
        If (IsObject(Item) And IsObject(listItem)) Then
            If (Item Is listItem) Then
                flag = True
                Exit For
            End If
        ElseIf (IsObject(Item) Or IsObject(listItem)) Then
            ''can't compare; continue
        Else
            If (Item = listItem) Then
                flag = True
                Exit For
            End If
        End If
    Next
    
    Contains = flag
End Function

''find the position of an item
Public Function IndexOf(Item As Variant) As Variant
    Dim Pos As Long
    Pos = 0
    
    For Each listItem In list
        If (IsObject(Item) And IsObject(listItem)) Then
            If (Item Is listItem) Then
                IndexOf = Pos
                Exit Function
            End If
        ElseIf (IsObject(Item) Or IsObject(listItem)) Then
            ''can't compare; continue
        Else
            If (Item = listItem) Then
                IndexOf = Pos
                Exit Function
            End If
        End If
        
        Pos = Pos + 1
    Next
    
    Err.Raise 448, "ArrayList.IndexOf", "Item not found"
End Function

''add and item at a position
Public Sub Insert(Item As Variant, Pos As Long)
    If (Pos < 0 Or Pos > size) Then
        Err.Raise 9, "ArrayList.Insert", "Out of bounds"
    End If
    
    If (Pos = 0) Then
        list.Add Item, before:=Pos + 1
    Else
        list.Add Item, after:=Pos
    End If
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

''remove an inclusive range
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
        newList.Add list(size - i + 1)
    Next
    
    Set list = newList
End Sub

''sort the list - does not support objects
Public Sub Sort()
    Dim tmp As Variant
    
    ''bubble sort
    For i = 1 To size - 1
        For j = i + 1 To size
            If (list(i) > list(j)) Then
                tmp = list(j)
                list.Remove j
                list.Add tmp, before:=i
            End If
        Next
    Next
End Sub
