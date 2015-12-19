<%
Class Test_AutoMap_Class
    Public SomeString, SomeInt, SomeDate
    Public Property Get Class_Get_Properties : Class_Get_Properties = Array("SomeString", "SomeInt", "SomeDate") : End Property
End Class

Class AutoMap_Tests
    Public Sub Setup       : End Sub
    Public Sub Teardown    : End Sub
    
    Public Function TestCaseNames
        TestCaseNames = Array("Test_From_Recordset_To_Existing_Class_Instance", _
                              "Test_From_Recordset_To_New_Class_Instance", _
                              "Test_From_Class_Instance_To_Existing_Class_Instance", _
                              "Test_From_Class_Instance_To_New_Class_Instance")
    End Function
    
    Private Sub Destroy(o)
        on error resume next
            o.close
        on error goto 0
        set o = nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Recordset_To_Existing_Class_Instance(T)
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
        
        dim target : set target = new Test_Automap_Class
        dim result : set result = Automapper().Automap(src, target)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Recordset_To_New_Class_Instance(T)
        dim src : set src = Server.CreateObject("ADODB.Recordset")
        with src.Fields
            .Append "SomeString", 200, 100
            .Append "SomeInt", adInteger
            .Append "SomeDate", adDate
        end with
        
        dim dtm : dtm = Now
        
        src.Open
        src.AddNew
        src("SomeString")   = "Some string here"
        src("SomeInt")      = 12345
        src("SomeDate")     = dtm
        src.Update
        
        src.MoveFirst
        
        dim result : set result = Automapper().Automap(src, "Test_Automap_Class")
        
        T.AssertEqual "Test_AutoMap_Class", typename(result), "AutoMap should have returned an instance of Test_AutoMap_Class."
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_Existing_Class_Instance(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_Automap_Class
        src.SomeString  = "Some string here"
        src.SomeInt     = 12345
        src.SomeDate    = dtm
        
        dim target    : set target    = new Test_Automap_Class
        
        dim result : set result = Automapper().Automap(src, target)
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy target
        Destroy result
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_From_Class_Instance_To_New_Class_Instance(T)
        dim dtm : dtm = Now
        dim src : set src = new Test_Automap_Class
        src.SomeString  = "Some string here"
        src.SomeInt     = 12345
        src.SomeDate    = dtm
        
        dim result : set result = Automapper().Automap(src, "Test_Automap_Class")
        
        T.AssertEqual "Some string here", result.SomeString, "Failed to map SomeString."
        T.AssertEqual 12345, result.SomeInt, "Failed to map SomeInt."
        T.AssertEqual dtm, result.SomeDate, "Failed to map SomeDate."
        
        Destroy src
        Destroy result
    End Sub
End Class
%>
