VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

''=======================================================
'' Class:       Dictionary
'' Version:     0.4.0
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 6/23/2020    TheEric960      Written
'' 6/30/2020    TheEric960      Expanded Scope
''=======================================================


' ===========================================
' VARIABLES AND INITIALIZATION
' ===========================================

Private dict As Collection
Private size As Long

''initialize the class variables
Private Sub Class_Initialize()
    Set dict = New Collection
    size = 0
End Sub

' ===========================================
' GETTERS AND SETTERS
' ===========================================

''size of dict
Public Property Get Count() As Long
    Count = size
End Property

''get value from key
Public Property Get Item(Key As Variant)
    For Each pair In dict
        If (IsObject(Pair.Key) Or IsObject(Key)) Then
            If (IsObject(Pair.Key) And IsObject(Key)) Then
                If (Pair.Key Is Key) Then
                    If (IsObject(Pair.Value)) Then
                        Set Item = Pair.Value
                    Else
                        Item = Pair.Value
                    End If
                    Exit Property
                End If
            End If
        Else
            If (Pair.Key = Key) Then
                If (IsObject(Pair.Value)) Then
                    Set Item = Pair.Value
                Else
                    Item = Pair.Value
                End If
                Exit Property
            End If
        End If
    Next
    
    Err.Raise 380, "Dictionary.Item", "Key not found"
End Property

''set value
Public Property Set Item(Key As Variant, Value As Variant)
    For Each pair In dict
        If (IsObject(Pair.Key) Or IsObject(Key)) Then
            If (IsObject(Pair.Key) And IsObject(Key)) Then
                If (Pair.Key Is Key) Then
                    Set Pair.Value = Value
                    Exit Property
                End If
            End If
        Else
            If (Pair.Key = Key) Then
                Set Pair.Value = Value
                Exit Property
            End If
        End If
    Next
    
    Add Key, Value
End Property

''let value
Public Property Let Item(Key As Variant, Value As Variant)
    For Each Pair In dict
        If (IsObject(Pair.Key) Or IsObject(Key)) Then
            If (IsObject(Pair.Key) And IsObject(Key)) Then
                If (Pair.Key Is Key) Then
                    Pair.Value = Value
                    Exit Property
                End If
            End If
        Else
            If (Pair.Key = Key) Then
                Pair.Value = Value
                Exit Property
            End If
        End If
    Next
    
    Add Key, Value
End Property

' ===========================================
' METHODS
' ===========================================

''add a key-value pair to the dict
Public Sub Add(Key As Variant, Value As Variant)
    If (Exists(Key)) Then
        Err.Raise 380, "Dictionary.Add", "Key exists"
    End If
    
    Dim Pair As New KeyValue
    If (IsObject(Key)) Then
        Set Pair.Key = Key
    Else
        Pair.Key = Key
    End If
    
    If (IsObject(Value)) Then
        Set Pair.Value = Value
    Else
        Pair.Value = Value
    End If
    
    dict.Add Pair
    size = dict.Count
End Sub

''test if something exists in the dict
Public Function Exists(Key As Variant) As Boolean
    Dim flag As Boolean
    flag = False
    
    For Each Pair In dict
        If (IsObject(Pair.Key) Or IsObject(Key)) Then
            If (IsObject(Pair.Key) And IsObject(Key)) Then
                If (Key Is Pair.Key) Then
                    flag = True
                    Exit For
                End If
            End If
        Else
            If (Key = Pair.Key) Then
                flag = True
                Exit For
            End If
        End If
    Next
    
    Exists = flag
End Function

''remove a key-value pair based on a key
Public Sub Remove(Key As Variant)
    For i = 1 To size
        If (IsObject(Key) Or IsObject(dict(i).Key)) Then
            If (IsObject(Key) And IsObject(dict(i).Key)) Then
                If (Key Is dict(i).Key) Then
                    dict.Remove i
                    size = size - 1
                    Exit Sub
                End If
            End If
        Else
            If (Key = dict(i).Key) Then
                dict.Remove i
                size = size - 1
                Exit Sub
            End If
        End If
    Next
    
    Err.Raise 380, "Dictionary.Remove", "Key not found"
End Sub

''clear the dict
Public Sub RemoveAll()
    Set dict = New Collection
    size = 0
End Sub

''returns a collection of keys
Public Function Keys() As Collection
    Dim coll As Collection
    Set coll = New Collection
    
    For Each pair In dict
        coll.Add pair.Key
    Next
    
    Set Keys = coll
End Function

''omitting case sensitivity
