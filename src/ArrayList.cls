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
''
''=======================================================

' ===========================================
' VARIABLES AND INITIALIZATION
' ===========================================

Private list() As Variant
Private size As Long

''initialize the list
Private Sub Class_Initialize()
    size = 0
End Sub

' ===========================================
' GETTERS AND SETTERS
' ===========================================

''get the size of the list
Public Property Get Count() As Long
    Count = size
End Property

''get an array from the list
Public Property Get ToArray() As Variant
    ToArray = list
End Property

' ===========================================
' METHODS
' ===========================================

''get an item at a specified position
Public Function Item() As Variant

End Function

''get a copy of the list
Public Function Clone() As Variant

End Function

''test if something exists in the list
Public Function Contains() As Variant

End Function

''find the position of an item
Public Function IndexOf() As Variant

End Function

''add and item at a position
Public Sub Insert()

End Sub

''clear the entire list
Public Sub Clear()

End Sub

''remove a specified entry via position
Public Sub RemoveAt()

End Sub

''remove a specified entry
Public Sub Remove()

End Sub

''remove a range
Public Sub RemoveRange()

End Sub

''reverse the list ordering
Public Function Reverse()

End Function

''sort the list
Public Function Sort()

End Function
