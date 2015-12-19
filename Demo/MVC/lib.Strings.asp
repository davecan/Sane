<%
'=======================================================================================================================
' StringBuilder Class
'=======================================================================================================================
Class StringBuilder_Class
    dim m_array
    dim m_array_size
    dim m_cur_pos
    
    Private Sub Class_Initialize
        m_array = Array
        m_array_size = 100
        redim m_array(m_array_size)
        m_cur_pos = -1
    End Sub
    
    Private Sub Extend
        m_array_size = m_array_size + 100
        redim preserve m_array(m_array_size)
    End Sub
    
    Public Sub Add(s)
        m_cur_pos = m_cur_pos + 1
        m_array(m_cur_pos) = s
        if m_cur_pos = m_array_size then Extend
    End Sub
    
    Public Function [Get](delim)
        'have to create a new array containing only the slots actually used, otherwise Join() happily adds delim
        'for *every* slot even the unused ones...
        dim new_array : new_array = Array()
        redim new_array(m_cur_pos)
        dim i
        for i = 0 to m_cur_pos
            new_array(i) = m_array(i)
        next
        [Get] = Join(new_array, delim)
    End Function
    
    Public Default Property Get TO_String
        TO_String = Join(m_array, "")
    End Property
End Class

Function StringBuilder()
    set StringBuilder = new StringBuilder_Class
End Function


'=======================================================================================================================
' Misc
'=======================================================================================================================
Function Excerpt(text, length)
    Excerpt = Left(text, length) & " ..."
End Function

Function IsBlank(text)
    If IsObject(text) then
        If text Is Nothing then
            IsBlank = true
        Else
            IsBlank = false
        End If
    Else
        If IsEmpty(text) or IsNull(text) or Len(text) = 0 then
            IsBlank = true
        Else
            IsBlank = false
        End If
    End If
End Function

%>
