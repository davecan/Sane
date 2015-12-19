<%
'=======================================================================================================================
' Automapper Function Test Cases
'=======================================================================================================================
Class Automapper_Function_Tests
    Public Sub Setup       : End Sub
    Public Sub Teardown    : End Sub
    
    Public Function TestCaseNames
        TestCaseNames = Array("Test_Automap_Function_Returns_Instance", "Test_Automap_Function_Returns_New_Instance")
    End Function
    
    Public Sub Test_Automap_Function_Returns_Instance(T)
        T.AssertEquals "Automapper_Class", typename(Automapper()), "Function Automapper() should return a new instance of Automapper_Class"
    End Sub
    
    Public Sub Test_Automap_Function_Returns_New_Instance(T)
        dim a : set a = Automapper()
        dim b : set b = Automapper()
        T.AssertFalse (a is b), "Automapper() function should return distinct instances of the Automapper_Class"
    End Sub
    
End Class
%>
