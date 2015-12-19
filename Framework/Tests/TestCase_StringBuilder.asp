<%
Class StringBuilder_Tests
    Public Sub Setup         : End Sub
    Public Sub Teardown    : End Sub
    
    Public Function TestCaseNames
        TestCaseNames = Array("Test_Initialized_Object_Should_Be_Empty", _
                                                    "Test_StringBuilder_Function_Should_Return_Initialized_Object", _
                                                    "Test_Default_Property_Should_Be_String", _
                                                    "Test_Default_String_Should_Not_Have_Spaces_Between_Entries", _
                                                    "Test_Join_Should_Allow_Custom_Delimiter_Between_Entries")
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Initialized_Object_Should_Be_Empty(T)
        dim SB : set SB = new StringBuilder_Class
        T.AssertEqual "", SB.TO_String, "Initialized object does not have an empty string."
        set SB = Nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_StringBuilder_Function_Should_Return_Initialized_Object(T)
        dim SB : set SB = new StringBuilder_Class
        T.AssertEqual "", SB.TO_String, "StringBuilder() function did not return initialized object."
        set SB = Nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Default_Property_Should_Be_String(T)
        dim SB : set SB = StringBuilder()
        T.AssertType "String", typename( (SB) ), "Object should default to string output."
        set SB = Nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Default_String_Should_Not_Have_Spaces_Between_Entries(T)
        dim SB : set SB = StringBuilder()
        SB.Add "foo"
        SB.Add "bar"
        SB.Add "baz"
        T.AssertEqual "foobarbaz", SB.TO_String, "Default string should not have spaces between entries."
        set SB = Nothing
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_Join_Should_Allow_Custom_Delimiter_Between_Entries(T)
        dim SB : set SB = StringBuilder()
        SB.Add "foo"
        SB.Add "bar"
        SB.Add "baz"
        T.AssertEqual "foo bar baz", SB.Get(" "), "Get() should allow a space between entries."
        T.AssertEqual "foo---bar---baz", SB.Get("---"), "Get() should allow --- between entries."
        T.AssertEqual "foo" & Chr(27) & "bar" & Chr(27) & "baz", SB.Get(Chr(27)), "Get() should allow non-standard ASCII character between entries."
        set SB = Nothing
    End Sub
    
    
    
End Class
%>
