Attribute VB_Name = "Test"

''=======================================================
'' Program:     Test
'' Desc:        Test cases for scripting objects
'' Version:     0.4.0
'' Changes----------------------------------------------
'' Date         Programmer      Change
'' 6/23/2020    TheEric960      Written
'' 6/23/2020    TheEric960      Added Dictionary Class
'' 6/30/2020    TheEric960      Expanded Dict. key support
''=======================================================

''main method
Sub RunTests()
    Debug.Print vbCrLf + vbCrLf + "-----------------"
    TestArrayList
    TestDictionary
    Debug.Print "All tests passed!"
    'MsgBox "All tests passed!"
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

''a series of Dictionary (Dict) tests
Sub TestDictionary(Optional ByVal Cancel As Boolean = False)
    If (Cancel) Then
        Exit Sub
    End If
    
    Dim dict As New Dictionary
    
    ''add items and objects
    For i = 0 To 8
        dict.Add CStr(i), i + 1
    Next
    Debug.Assert (dict.Count = 9)
    Debug.Assert (dict.Item("7") = 8)
    
    dict.Add "bug", Excel.Worksheets
    Debug.Assert (dict.Count = 10)
        
    ''check the existance of a key
    Debug.Assert (dict.Exists("3"))
    Debug.Assert (dict.Exists("bug"))
    Debug.Assert Not (dict.Exists("fly"))
    
    ''change the value of an item
    dict.Item("bug") = 35
    Debug.Assert (dict.Item("bug") = 35)
    
    ''see keys
    Dim coll As New Collection
    Set coll = dict.Keys
    
    ''remove items
    dict.Remove "bug"
    dict.Remove "3"
    Debug.Assert (Not dict.Exists("bug"))
    Debug.Assert (Not dict.Exists("3"))
    Debug.Assert (dict.Count = 8)
    
    dict.RemoveAll
    Debug.Assert (dict.Count = 0)
    
    ''----- begin testing objects -----
    dict.Add coll, "this is a test"
    Debug.Assert (dict.Item(coll) = "this is a test")
    
    dict.Item(coll) = "overwritten"
    Debug.Assert (dict.Item(coll) = "overwritten")
    Debug.Assert (dict.Count = 1)
    
    Dim coll2 As New Collection
    coll2.Add "A Thing!"
    
    Set dict.Item(coll2) = coll
    Debug.Assert (dict.Item(coll2) Is coll)
    Debug.Assert (dict.Count = 2)
    
    Debug.Assert (dict.Exists(coll2))
    
    dict.Remove coll
    Debug.Assert (Not dict.Exists(coll))
    Debug.Assert (dict.Count = 1)
    
    Set coll = dict.Keys
    
    dict.RemoveAll
    Debug.Assert (dict.Count = 0)
    
    Debug.Print "Dictionary Passed"
End Sub
