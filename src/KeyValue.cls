VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValue"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
''=======================================================
'' Class:       KeyValue
'' Version:     0.2.0
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 6/23/2020    TheEric960      Written
''=======================================================


' ===========================================
' VARIABLES AND INITIALIZATION
' ===========================================

Public Key As String
Public Value As Variant

''initialize the class variables
Private Sub Class_Initialize()
    Key = ""
    Value = ""
End Sub
