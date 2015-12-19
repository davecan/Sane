<%
Class Test_DynMap_Class
    Public SomeString, SomeInt, SomeDate
    Public Property Get Class_Get_Properties : Class_Get_Properties = Array("SomeString", "SomeInt", "SomeDate") : End Property
End Class

Class DynMap_Tests
    Public Sub Setup       : End Sub
    Public Sub Teardown    : End Sub
    
    Public Function TestCaseNames
        TestCaseNames = Array("Test_From_Recordset_To_Existing_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields", _
                              "Test_From_Recordset_To_Existing_Class_Instance_Maps_Only_Specified_Fields", _
                              "Test_From_Recordset_To_New_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields", _
                              "Test_From_Recordset_To_New_Class_Instance_Maps_Only_Specified_Fields", _
                              "Test_From_Class_Instance_To_Existing_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields", _
                              "Test_From_Class_Instance_To_Existing_Class_Instance_Maps_Only_Specified_Fields", _
                              "Test_From_Class_Instance_To_New_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields", _
                              "Test_From_Class_Instance_To_New_Class_Instance_Maps_Only_Specified_Fields", _
                              "Test_From_Class_Instance_To_Existing_Class_Instance_Maps_One_Expression", _
                              "Test_From_Class_Instance_To_Existing_Class_Instance_Maps_Multiple_Expressions")
    End Function
    
    Private Sub Destroy(o)
        on error resume next
            o.close
        on error goto 0
        set o = nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Recordset_To_Existing_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields(T)
        dim src : set src = Server.CreateObject("ADODB.Recordset")
        with src.Fields
            .Append "SomeString", adVarChar, 100
            .Append "SomeInt", adInteger
            .Append "SomeDate", adDate
        end with
        
        dim dtm : dtm = Now
        
        src.Open
        src.AddNew
        src("SomeString")    = "Some string here"
        src("SomeInt")       = 12345
        src("SomeDate")      = dtm
        src.Update
        
        src.MoveFirst
        
        dim target : set target = new Test_DynMap_Class
        dim result : set result = Automapper().DynMap(src, target, empty, empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Recordset_To_Existing_Class_Instance_Maps_Only_Specified_Fields(T)
        dim src : set src = Server.CreateObject("ADODB.Recordset")
        with src.Fields
            .Append "SomeString", adVarChar, 100
            .Append "SomeInt", adInteger
            .Append "SomeDate", adDate
        end with
        
        dim dtm : dtm = Now
        
        src.Open
        src.AddNew
        src("SomeString")    = "Some string here"
        src("SomeInt")       = 12345
        src("SomeDate")      = dtm
        src.Update
        
        src.MoveFirst
        
        dim target : set target = new Test_DynMap_Class
        target.SomeInt = 54321
        
        dim result : set result = Automapper().DynMap(src, target, array("SomeString", "SomeDate"), empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 54321, result.SomeInt, "SomeInt should have been left untouched, but was mapped anyway."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Recordset_To_New_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields(T)
        dim src : set src = Server.CreateObject("ADODB.Recordset")
        with src.Fields
            .Append "SomeString", adVarChar, 100
            .Append "SomeInt", adInteger
            .Append "SomeDate", adDate
        end with
        
        dim dtm : dtm = Now
        
        src.Open
        src.AddNew
        src("SomeString")    = "Some string here"
        src("SomeInt")       = 12345
        src("SomeDate")      = dtm
        src.Update
        
        src.MoveFirst
        
        dim result : set result = Automapper().DynMap(src, "Test_DynMap_Class", empty, empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Recordset_To_New_Class_Instance_Maps_Only_Specified_Fields(T)
        dim src : set src = Server.CreateObject("ADODB.Recordset")
        with src.Fields
            .Append "SomeString", adVarChar, 100
            .Append "SomeInt", adInteger
            .Append "SomeDate", adDate
        end with
        
        dim dtm : dtm = Now
        
        src.Open
        src.AddNew
        src("SomeString")    = "Some string here"
        src("SomeInt")       = 12345
        src("SomeDate")      = dtm
        src.Update
        
        src.MoveFirst
        
        dim result : set result = Automapper().DynMap(src, "Test_DynMap_Class", array("SomeString", "SomeDate"), empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertNotExists result.SomeInt, "SomeInt should have been left uninitialized, but was mapped anyway."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_Existing_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_DynMap_Class
        src.SomeString = "Some string here"
        src.SomeInt    = 12345
        src.SomeDate   = dtm
        
        dim target    : set target    = new Test_DynMap_Class
        
        dim result : set result = Automapper.DynMap(src, target, empty, empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_Existing_Class_Instance_Maps_Only_Specified_Fields(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_DynMap_Class
        src.SomeString = "Some string here"
        src.SomeInt    = 12345
        src.SomeDate   = dtm
        
        dim target    : set target    = new Test_DynMap_Class
        target.SomeInt = 54321
        
        dim result : set result = Automapper.DynMap(src, target, array("SomeString", "SomeDate"), empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 54321, result.SomeInt, "SomeInt should have been left alone but was mapped anyway."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_New_Class_Instance_With_Empty_Fields_Array_Maps_All_Fields(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_DynMap_Class
        src.SomeString = "Some string here"
        src.SomeInt    = 12345
        src.SomeDate   = dtm
        
        dim result : set result = Automapper.DynMap(src, "Test_DynMap_Class", empty, empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_New_Class_Instance_Maps_Only_Specified_Fields(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_DynMap_Class
        src.SomeString = "Some string here"
        src.SomeInt    = 12345
        src.SomeDate   = dtm
        
        dim result : set result = Automapper.DynMap(src, "Test_DynMap_Class", array("SomeString", "SomeDate"), empty)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertNotExists result.SomeInt, "SomeInt should have been left uninitialized, but was mapped anyway."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_Existing_Class_Instance_Maps_One_Expression(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_DynMap_Class
        src.SomeString = "Some string here"
        src.SomeInt    = 12345
        src.SomeDate   = dtm
        
        dim target : set target = new Test_DynMap_Class
        
        dim result : set result = Automapper.DynMap(src, target, empty, "target.SomeString = ucase(src.SomeString)")
        
        T.AssertEqual "SOME STRING HERE", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_Existing_Class_Instance_Maps_Multiple_Expressions(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_DynMap_Class
        src.SomeString = "Some string here"
        src.SomeInt    = 12345
        src.SomeDate   = dtm
        
        dim target : set target = new Test_DynMap_Class
        
        dim result : set result = Automapper.DynMap(src, target, empty, _
                                                    array("target.SomeString = ucase(src.SomeString)", _
                                                          "target.SomeInt = src.SomeInt * 2", _
                                                          "target.SomeDate = Year(src.SomeDate)"))
        
        T.AssertEqual "SOME STRING HERE", result.SomeString, "Failed to map SomeString."
        T.AssertEqual (12345 * 2), result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual Year(dtm), result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
End Class
%>
