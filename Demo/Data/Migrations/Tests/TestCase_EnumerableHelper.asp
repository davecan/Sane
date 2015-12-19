<%
Class EnumerableHelper_Tests
    Public Sub Setup       : End Sub
    Public Sub Teardown    : End Sub
    
    Public Function TestCaseNames
        TestCaseNames = Array("Test_Enumerable_Function_Returns_EnumerableHelper_Instance", _
                              "Test_Map_Returns_EnumerableHelper_Instance", _
                              "Test_Map_Lambda", _
                              "Test_Map_Proc", _
                              "Test_Max_Lambda", _
                              "Test_Max_Proc", _
                              "Test_Min_Lambda", _
                              "Test_Min_Proc", _
                              "Test_Take", _
                              "Test_TakeWhile_Lambda", _
                              "Test_TakeWhile_Proc", _
                              "Test_Sum", _
                              "Test_All", _
                              "Test_Any", _
                              "Test_Where_Lambda", _
                              "Test_Where_Proc", _
                              "Test_Count", _
                              "Test_First", _
                              "Test_Last", _
                              "Test_Chained")
    End Function

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Enumerable_Function_Returns_EnumerableHelper_Instance(T)
        dim list : set list = new LinkedList_Class
        dim E : set E = Enumerable(list)
        T.Assert typename(E) = "EnumerableHelper_Class", "did not return correct instance"
        T.Assert typename(E.Data) = "LinkedList_Class", "Data() is of type " & typename(E.Data)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Map_Returns_EnumerableHelper_Instance(T)
        dim list : set list = new LinkedList_Class
        dim E : set E = Enumerable(list).Map("")
        T.Assert typename(E) = "EnumerableHelper_Class", "returned type " & typename(E)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Map_Lambda(T)
        dim list : set list = OrderedList(10)
        dim E : set E = Enumerable(list).Map("dim x : x = item_ + 1 : V_ = x")

        T.Assert 10 = E.Data.Count, "count: " & E.Data.Count

        dim A : A = list.To_Array()
        dim B : B = E.Data.To_Array()

        T.Assert UBound(A) = UBound(B), "A = " & UBound(A) & ", B = " & UBound(B)

        dim i
        For i = 0 to UBound(A)
            T.Assert B(i) = A(i) + 1, "B(" & i & ") = " & B(i) & ", A(" & i & ") = " & A(i)
        Next
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Map_Proc(T)
        dim list : set list = OrderedList(10)
        dim E : set E = Enumerable(list).Map(GetRef("AddOne"))

        T.Assert 10 = E.Data.Count, "count: " & E.Data.Count

        dim A : A = list.To_Array()
        dim B : B = E.Data.To_Array()

        T.Assert UBound(A) = UBound(B), "A = " & UBound(A) & ", B = " & UBound(B)

        dim i
        For i = 0 to UBound(A)
            T.Assert B(i) = A(i) + 1, "B(" & i & ") = " & B(i) & ", A(" & i & ") = " & A(i)
        Next
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Max_Lambda(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim val : val = Enumerable(list).Max("len(item_)")
        T.Assert 7 = val, "val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Max_Proc(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim val : val = Enumerable(list).Max(GetRef("GetLength"))
        T.Assert 7 = val, "val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Min_Lambda(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim val : val = Enumerable(list).Min("len(item_)")
        T.Assert 3 = val, "val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Min_Proc(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim val : val = Enumerable(list).Min(GetRef("GetLength"))
        T.Assert 3 = val, "val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Take(T)
        dim list : set list = OrderedList(3)

        dim firstTwo : set firstTwo = Enumerable(list).Take(2)

        T.Assert typename(firstTwo) = "EnumerableHelper_Class", "typename = " & typename(firstTwo)

        dim L : set L = firstTwo.Data
        T.AssertEqual 2, L.Count, "Count = " & L.Count
        T.AssertEqual 1, L.Front, "Front = " & L.Front
        T.AssertEqual 2, L.Back, "Back = " & L.Back
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_TakeWhile_Lambda(T)
        dim list : set list = OrderedList(5)
        dim firstTwo : set firstTwo = Enumerable(list).TakeWhile("item_ < 3")

        T.Assert typename(firstTwo) = "EnumerableHelper_Class", "typename = " & typename(firstTwo)

        dim L : set L = firstTwo.Data
        T.AssertEqual 2, L.Count, "Count = " & L.Count
        T.AssertEqual 1, L.Front, "Front = " & L.Front
        T.AssertEqual 2, L.Back, "Back = " & L.Back
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_TakeWhile_Proc(T)
        dim list : set list = OrderedList(5)
        dim firstTwo : set firstTwo = Enumerable(list).TakeWhile(GetRef("IsLessThanThree"))

        T.Assert typename(firstTwo) = "EnumerableHelper_Class", "typename = " & typename(firstTwo)

        dim L : set L = firstTwo.Data
        T.AssertEqual 2, L.Count, "Count = " & L.Count
        T.AssertEqual 1, L.Front, "Front = " & L.Front
        T.AssertEqual 2, L.Back, "Back = " & L.Back
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Sum(T)
        dim list : set list = OrderedList(5)
        dim val : val = Enumerable(list).Sum("item_")
        T.AssertEqual val, 15, "simple: val = " & val

        val = Enumerable(list).Sum("item_ * item_")
        T.AssertEqual val, 55, "squares: val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_All(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim val
        
        val = Enumerable(list).All("len(item_) >= 3")
        T.Assert val, "Len >= 3: val = " & val

        val = Enumerable(list).All("len(item_) = 3")
        T.AssertFalse val, "Len = 3: val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Any(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim val
        
        val = Enumerable(list).Any("len(item_) >= 3")
        T.Assert val, "Len >= 3: val = " & val

        val = Enumerable(list).Any("len(item_) < 3")
        T.AssertFalse val, "Len < 3: val = " & val
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Where_Lambda(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim list2

        set list2 = Enumerable(list).Where("len(item_) > 4")
        T.AssertEqual 2, list2.Data.Count, "list2.Count = " & list2.Data.Count

        T.AssertEqual "Alice", list2.Data.Front(), "list2.Data.Front = " & list2.Data.Front()
        T.AssertEqual "Charlie", list2.Data.Back(), "list2.Data.Front = " & list2.Data.Back()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Where_Proc(T)
        dim list : set list = new LinkedList_Class
        list.Push "Alice"
        list.Push "Bob"
        list.Push "Charlie"
        list.Push "Doug"

        dim E : set E = Enumerable(list).Where(GetRef("IsMoreThanFourChars"))
        T.AssertEqual 2, E.Data.Count, "E.Count = " & E.Data.Count
        T.AssertEqual "Alice", E.Data.Front(), "E.Data.Front = " & E.Data.Front()
        T.AssertEqual "Charlie", E.Data.Back(), "E.Data.Front = " & E.Data.Back()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Count(T)
        dim list : set list = OrderedList(10)
        dim E : set E = Enumerable(list)
        T.AssertEqual 10, E.Count(), empty
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_First(T)
        dim list : set list = OrderedList(10)
        dim E : set E = Enumerable(list)
        T.AssertEqual 1, E.First(), empty
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Last(T)
        dim list : set list = OrderedList(10)
        dim E : set E = Enumerable(list)
        T.AssertEqual 10, E.Last(), empty
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Chained(T)
        dim list : set list = OrderedList(10)
        dim it, item

        dim A : set A = Enumerable(list).Take(5).Where("item_ mod 2 <> 0")
        T.AssertEqual 3, A.Data.Count, "A.Count = " & A.Data.Count
        set it = A.Data.Iterator
        item = it.GetNext()
        T.AssertEqual 1, item, "A 1st item = " & item
        item = it.GetNext()
        T.AssertEqual 3, item, "A 2nd item = " & item
        item = it.GetNext()
        T.AssertEqual 5, item, "A 3rd item = " & item

        dim B : set B = Enumerable(list).Take(5).Select("V_ = item_ * item_")
        T.AssertEqual 5, B.Data.Count, "B.Count = " & B.Data.Count
        set it = B.Data.Iterator
        item = it.GetNext()
        T.AssertEqual 1, item, "B 1st item = " & item
        item = it.GetNext()
        T.AssertEqual 4, item, "B 2nd item = " & item
        item = it.GetNext()
        T.AssertEqual 9, item, "B 3rd item = " & item
        item = it.GetNext()
        T.AssertEqual 16, item, "B 4th item = " & item
        item = it.GetNext()
        T.AssertEqual 25, item, "B 5th item = " & item

        dim list2 : set list2 = new LinkedList_Class
        list2.Push "Alice"
        list2.Push "Bob"
        list2.Push "Charlie"
        list2.Push "Doug"
        list2.Push "Edward"
        list2.Push "Franklin"
        list2.Push "George"
        list2.Push "Hal"
        list2.Push "Isaac"
        list2.Push "Jeremy"

        dim C : C = Enumerable(list2) _
                        .Where("len(item_) > 5") _
                        .Map("set V_ = new ChainedExample_Class : V_.Data = item_ : V_.Length = len(item_)") _
                        .Max("item_.Length")
        
        T.AssertEqual 8, C, "C"
    End Sub


End Class


Class ChainedExample_Class
    Public Data
    Public Length
End Class


Function OrderedList(maxnum)
    dim list : set list = new LinkedList_Class
    dim i
    For i = 1 to maxnum
        list.Push i
    Next
    set OrderedList = list
End Function

Function AddOne(x)
    AddOne = x + 1
End Function

Function GetLength(s)
    GetLength = Len(s)
End Function

Function IsLessThanThree(x)
    IsLessThanThree = x < 3
End Function

Function IsMoreThanFourChars(s)
    IsMoreThanFourChars = Len(s) > 4
End Function


%>