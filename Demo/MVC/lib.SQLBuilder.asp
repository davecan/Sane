<%
'=======================================================================================================================
' SQL Builder Class
'=======================================================================================================================
Class SQLBuilder_Class
    Private m_sql
    Private m_SB    'StringBuilder
    
    Private Sub Class_Initialize
        'm_sql = ""
        set m_SB = StringBuilder()
    End Sub
    
    Public Default Property Get TO_String
        'TO_String = m_sql
        TO_String = Trim(m_SB.Get(" "))
    End Property
    
    'given an array of field names, generates the SELECT clause
    Public Function [Select](ByVal field_names)
        AddFragment "SELECT"
        if typename(field_names) = "Variant()" then
            dim n : n = ubound(field_names)
            dim i
            for i = 0 to n
                AddFragment field_names(i)
                if (i < n) then AddFragment ","
            next
        else    'assume string or string-like default value
            AddFragment field_names
        end if
        set [Select] = Me
    End Function
    
    Public Function From(ByVal table_names)
        AddFragment "FROM"
        if typename(table_names) = "Variant()" then
            dim n : n = ubound(table_names)
            dim i
            for i = 0 to n
                AddFragment table_names(i)
                'AddFragment "tab" & i
                if (i < n) then AddFragment ","
            next
        else    'assume string or string-like default value
            AddFragment table_names
        end if
        set From = Me
    End Function
    
    Public Function Where
        AddFragment "WHERE"
        set Where = Me
    End Function
    
    Public Function [And]
        AddFragment "AND"
        set [And] = Me
    End Function
    
    Public Function [Or]
        AddFragment "OR"
        set [Or] = Me
    End Function
    
    Public Function Eq(key, val)
        AddFragment key
        AddFragment "="
        AddFragment Quote(val)
        set Eq = Me
    End Function
    
    Public Function EachIsEq(kv_pairs)
        if IsArray(kv_pairs) then
            dim i, key, val
            dim n : n = ubound(kv_pairs)
            for i = 0 to n step 2
                KeyVal kv_pairs, i, key, val
                Eq key, val
                if (i < n - 1) then Me.And    'add AND only if not the last key/val pair, avoids dangling ANDs in the SQL
            next
        end if
        set EachIsEq = Me
    End Function
    
    Public Function BeginsWith(key, val)
        AddFragment key
        AddFragment "LIKE '" & val & "%'"
        set BeginsWith = Me
    End Function
    
    Public Function EachBeginsWith(kv_pairs)
        if IsArray(kv_pairs) then
            dim i, key, val
            dim n : n = ubound(kv_pairs)
            for i = 0 to n step 2
                KeyVal kv_pairs, i, key, val
                BeginsWith key, val
                if (i < n - 1) then Me.And    'add AND only if not the last key/val pair, avoids dangling ANDs in the SQL
            next
        end if
        set EachBeginsWith = Me
    End Function
    
    Public Function EndsWith(key, val)
        AddFragment key & " LIKE '%" & val & "'"
        set EndsWith = Me
    End Function
    
    Public Function Contains(key, val)
        AddFragment " " & key & " LIKE '%" & val & "%'"
        set Contains = Me
    End Function
    
    Public Function [In](key, in_string)
        AddFragment " " & key & " IN (" & in_string & ")"
        set [In] = Me
    End Function
    
    Public Function OrderBy
        AddFragment "ORDER BY"
        set OrderBy = Me
    End Function
    
    Public Function [Asc](field_names)
        if typename(field_names) = "Variant()" then
            dim n : n = ubound(field_names)
            dim i
            for i = 0 to n
                AddFragment field_names(i)
                AddFragment "ASC"
                If i < n then AddFragment(",")
            next
        else    'assume string or string-like default
            AddFragment field_names
            AddFragment "ASC"
        end if
        set [Asc] = Me
    End Function
    
    Public Function Desc(field_names)
        if typename(field_names) = "Variant()" then
            dim n : n = ubound(field_names)
            dim i
            for i = 0 to n
                AddFragment field_names(i)
                AddFragment "DESC"
                If i < n then AddFragment(",")
            next
        else    'assume string or string-like default
            AddFragment field_names
            AddFragment "DESC"
        end if
        set Desc = Me
    End Function
    
    'appends raw sql
    Public Function Raw(raw_sql)
        AddFragment raw_sql
        set Raw = Me
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Private Methods
    '---------------------------------------------------------------------------------------------------------------------
    Private Sub AddFragment(fragment)
        'm_sql = m_sql & fragment
        m_SB.Add fragment
        'put "AddFragment: [-" & fragment & "-]"
        'put "AddFragment: m_SB.Get(' ') := " & m_SB.Get(" ")
    End Sub
    
    Private Function Quote(val)
        'Typename( (val) ) forces val to evaluate and return a value. This accomodates string-like objects that can be
        'output by Response.Write streams but are not themselves Strings (but return a String when eval'd).
        Quote = Choice(typename((val)) = "String", "'" & val & "'", val)    
    End Function
    
    Private Function CommaIf(expr)
        CommaIf = Choice(expr, ", ", "")
    End Function
    
End Class

Function SQLBuilder()
    set SQLBuilder = new SQLBuilder_Class
End Function
%>
