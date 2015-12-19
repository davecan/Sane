<%
'=======================================================================================================================
' AUTOMAPPER CLASS
'=======================================================================================================================

'Side Effects: Since src and target are passed ByRef to reduce unnecessary copying, if src is a recordset then the
'              current record pointer is modified using src.MoveFirst and src.MoveNext. The end result is the current
'              record pointer ends the operation at src.EOF.

Class Automapper_Class
    Private m_src
    Private m_target
    Private m_statements
    Private m_statements_count
    
    Private Property Get Src    : set Src       = m_src       : End Property
    Private Property Get Target : set Target    = m_target    : End Property
    
    Private Sub Class_Initialize
        m_statements = Array()
        m_statements_count = 0
        redim m_statements(m_statements_count)
    End Sub
    
    Private Sub ResetState
        m_statements_count = 0
        redim m_statements(m_statements_count)
        set m_src = Nothing
        set m_target = Nothing
    End Sub

    'Maps all rs or object fields to corresponding fields in the specified class.
    Public Function AutoMap(src_obj, target_obj)
        Set AutoMap = FlexMap(src_obj, target_obj, empty)
    End Function
    
    'Only maps fields specified in the field_names array (array of strings).
    'If field_names is empty, attempts to map all fields from the passed rs or object.
    Public Function FlexMap(src_obj, target_obj, field_names)
        Set FlexMap = DynMap(src_obj, target_obj, field_names, empty)
    End Function
    
    'Only maps fields specified in the field_names array (array of strings).
    'If field_names is empty then src MUST be a recordset as it attempts to map all fields from the recordset.
    'Since there is no reflection in vbscript, there is no way around this short of pseudo-reflection.
    Public Function DynMap(src_obj, target_obj, field_names, exprs)
        SetSource src_obj
        SetTarget target_obj
    
        dim field_name
        dim field_idx 'loop counter
        
        if IsEmpty(field_names) then 'map everything
            if typename(src_obj) = "Recordset" then
                for field_idx = 0 to src_obj.Fields.Count - 1
                    field_name = src_obj.Fields.Item(field_idx).Name
                    'AddStatement field_name
                    AddStatement BuildStatement(field_name)
                next
                
            elseif InStr(typename(src_obj), "Dictionary") > 0 then    'enables Scripting.Dictionary and IRequestDictionary for Request.Querystring and Request.Form
                for each field_name in src_obj
                    AddStatement BuildStatement(field_name)
                next
                
            elseif not IsEmpty(src_obj.Class_Get_Properties) then
                dim props : props = src_obj.Class_Get_Properties
                for field_idx = 0 to ubound(props)
                    field_name = props(field_idx)
                    'AddStatement field_name
                    AddStatement BuildStatement(field_name)
                next
                
            else    'some invalid type of object
                Err.Raise 9, "Automapper.DynMap", "Cannot automatically map this source object. Expected recordset or object implementing Class_Get_Properties reflection, got: " & typename(src_obj)        
            end if
            
        else 'map only specified fields
            for field_idx = lbound(field_names) to ubound(field_names)
                field_name = field_names(field_idx)
                'AddStatement field_name
                AddStatement BuildStatement(field_name)
            next
        end if
        
        dim exprs_idx
        
        if not IsEmpty(exprs) then
            if typename(exprs) = "Variant()" then
                for exprs_idx = lbound(exprs) to ubound(exprs)
                    'field_name = exprs(exprs_idx)
                    'AddStatement field_name
                    AddStatement exprs(exprs_idx)
                next
            else    'assume string or string-like default value
                AddStatement exprs
            end if
        end if
        
        'Can't pre-join the statements because if one fails the rest of them fail too... :(
        'dim joined_statements : joined_statements = Join(m_statements, " : ")
        'put joined_statements
        
        'suspend errors to prevent failing when attempting to map a field that does not exist in the class
        on error resume next
            dim stmt_idx
            for stmt_idx = 0 to ubound(m_statements)
                Execute m_statements(stmt_idx)
            next
        on error goto 0
        
        set DynMap = m_target
        
        ResetState
    End Function
    
    
    Private Sub SetSource(ByVal src_obj)
        set m_src = src_obj
    End Sub
    
    Private Sub SetTarget(ByVal target_obj)
        if typename(target_obj) = "String" then
            set m_target = eval("new " & target_obj)
        else
            set m_target = target_obj
        end if
    End Sub
    
    
    'Builds a statement and adds it to the internal statements array
    Private Sub AddStatement(ByVal stmt)
        redim preserve m_statements(m_statements_count + 1)
        m_statements(m_statements_count) = stmt
        m_statements_count = m_statements_count + 1
    End Sub
    
    Private Function BuildStatement(ByVal field_name)
        dim result
        if typename(m_src) = "Recordset" or InStr(typename(m_src), "Dictionary") > 0 then
            result = "m_target." & field_name & " = m_src(""" & field_name & """)"
        else
            'Funky magic...
            'If src.field_name is an object, ensure the set statement is used
            if IsObject(eval("m_src." & field_name)) then
                result = "set "
            else
                'result = "m_target." & field_name & " = m_src." & field_name
            end if
            result = result & " m_target." & field_name & " = m_src." & field_name
        end if
        BuildStatement = result
    End Function
End Class


Function Automapper()
    Set Automapper = new Automapper_Class
End Function
%>
