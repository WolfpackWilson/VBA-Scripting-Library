Attribute VB_Name = "Test"

''=======================================================
'' Program:     Test
'' Desc:        Test cases for scripting objects
'' Version:     0.1.0
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 6/23/2020    TheEric960      Written
''=======================================================

''main method
Sub RunTests()
    Debug.Print vbCrLf + vbCrLf + "-----------------"
    TestArrayList
End Sub

''a series of ArrayList (AL) tests
Sub TestArrayList(Optional ByVal Cancel As Boolean = False)
    If (Cancel) Then
        Exit Sub
    End If
    
    ''create empty AL
    Dim arrList As New ArrayList
    Debug.Assert (arrList.Count = 0)
    
    Dim list(), list2() As Variant
    list = arrList.ToArray

    
    ''add numbers and objects to AL
    arrList.Add "banana"
    arrList.Add 24
    
    Dim obj As Object
    Set obj = Excel.Worksheets
    arrList.Add obj
    
    ''insert items
    arrList.Insert "milk", 3
    arrList.Insert "cheese", 0
    arrList.Insert "goat", 5
    list = arrList.ToArray
    Debug.Assert (list(0) = "cheese")
    Debug.Assert (list(5) = "goat")
    
    ''test contains, indexOf
    Debug.Assert arrList.Contains("goat")
    Debug.Assert (arrList.IndexOf("goat") = 5)
    
    ''clear the list
    arrList.Clear
    list2 = arrList.ToArray
    
    ''test sorting and reverse
    For i = 0 To 5
        arrList.Add i
    Next
    
    arrList.Reverse
    arrList.Sort
    
    ''test remove methods
    arrList.Remove 3
    Debug.Assert (Not arrList.Contains(3))
    
    arrList.RemoveAt 2
    Debug.Assert (Not arrList.Contains(2))
    
    arrList.RemoveRange 1, 3
    Debug.Assert (arrList.Count = 1)
    
    Debug.Print "ArrayList Passed"
End Sub
