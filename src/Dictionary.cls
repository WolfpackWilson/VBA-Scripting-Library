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
'' Version:     0.2.0
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 6/23/2020    TheEric960      Written
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
Public Property Get Item(Key As String)
    For Each pair In dict
        If (pair.Key = Key) Then
            If (IsObject(pair.Value)) Then
                Set Item = pair.Value
            Else
                Item = pair.Value
            End If
            Exit Property
        End If
    Next
    
    Err.Raise 380, "Dictionary.Item", "Key not found"
End Property

''set value
Public Property Set Item(Key As String, Value As Variant)
    For Each pair In dict
        If (pair.Key = Key) Then
            If (IsObject(Value)) Then
                Set pair.Value = Value
            Else
                pair.Value = Value
            End If
            Exit Property
        End If
    Next
    
    Err.Raise 380, "Dictionary.Item", "Key not found"
End Property

' ===========================================
' METHODS
' ===========================================

''add a key-value pair to the dict
Public Sub Add(Key As String, Value As Variant)
    Dim Pair As New KeyValue
    Pair.Key = Key
    
    If (IsObject(Value)) Then
        Set Pair.Value = Value
    Else
        Pair.Value = Value
    End If
    
    dict.Add Pair
    size = dict.Count
End Sub

''test if something exists in the dict
Public Function Exists(Key As String) As Boolean
    Dim flag As Boolean
    flag = False
    
    For Each keystr In Keys
        If (Key = keystr) Then
            flag = True
            Exit For
        End If
    Next
    
    Exits = flag
End Function

''remove a key-value pair based on a key
Public Sub Remove(Key As String)
    For i = 1 To size
        If (Key = dict(i).Key) Then
            dict.Remove i
            size = size - 1
            Exit Sub
        End If
    Next
    
    Err.Raise 380, "Dictionary.Item", "Key not found"
End Sub

''clear the dict
Public Sub RemoveAll()
    dict = New Collection
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
