<%
'=======================================================================================================================
' IO Helpers
'=======================================================================================================================
Sub put(v)
    Select Case typename(v)
        Case "LinkedList_Class"       : response.write join(v.TO_Array, ", ")
        Case "DynamicArray_Class"     : response.write JoinList(v)
        Case "Variant()"              : response.write join(v, ", ")
        Case else                     : response.write v
    End Select
End Sub

Sub put_
    put "<br>"
End Sub

Sub putl(v)
    put v
    put_
End Sub

'accepts anything that can have an iterator, including lists, arrays, and recordsets
Sub putlist(col, prefix, suffix)
    dim it : set it = IteratorFor(col)
    Do While it.HasNext
        put prefix & it.GetNext & suffix
    Loop
End Sub

'same as join() for arrays, but for any arbitrary collection
Function JoinList(col)
    dim s : s = ""
    dim it : set it = IteratorFor(col)
    Do While it.HasNext
        s = s & ", "
    Loop
    JoinList = Left(s, Len(s) - 2)
End Function

'---------------------------------------------------------------------------------------------------------------------
'Wrapper for Server.HTMLEncode() -- makes it easier on the eyes when reading the HTML code
Function H(s)
    If Not IsEmpty(s) and Not IsNull(s) then
        H = Server.HTMLEncode(s)
    Else
        H = ""
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
'allows tracing of output on demand without interfering with layout
Sub trace(s)
    comment s
End Sub

'-----------------------------------------------------------------------------------------------------------------------
'outputs an HTML comment, useful for tracing etc
Sub Comment(text)
    response.write vbcrlf & vbcrlf & "<!--" & vbcrlf & H(text) & vbcrlf & "-->" & vbcrlf & vbcrlf
End Sub

'-----------------------------------------------------------------------------------------------------------------------
'pseudo-design-by-contract capability, allows strong-typing of methods and views
Sub ExpectType(obj_type, obj)
    if typename(obj) <> obj_type then Err.Raise 1, "lib.Helpers:ExpectType", "View expected object of type '" & obj_type & "' but received type '" & typename(obj) & "'."
End Sub


'=======================================================================================================================
' Dump* functions for dumping variables, objects, lists, etc for debugging purposes
'=======================================================================================================================
Class DataDumper_Class
    Public Sub Dump(V)
        put "<pre>"
        DumpIt V
        put "</pre>"
    End Sub
    
    Private m_indent
    
    Private Sub Indent
        m_indent = m_indent + 1
        'putl "INDENT: " & m_indent
        'puti m_indent
        'put_
    End Sub
    
    Private Sub Dedent
        m_indent = m_indent - 1
        'putl "INDENT: " & m_indent
    End Sub
    
    Private Sub Class_Initialize
        m_indent = -1     'first indent takes it to 0
    End Sub
    
    'prepends indents
    Private Sub puti(v)
        put Spaces(m_indent) & v
    End Sub
    
    Private Sub DumpIt(V)
        If Instr(Typename(V), "_Class") > 0 then 
            DumpClass V
        ElseIf Typename(V) = "Variant()" then
            DumpArray V
        ElseIf Typename(V) = "Recordset" then
            DumpRecordset V
        Else
            put "&laquo;" & H(V) & "&raquo;" 
        End If
    End Sub
    
    Private Sub DumpArray(V)
        Indent
        dim i
        put_
        puti "[Array:" & vbCR
            Indent
            For i = 0 to UBound(V)
                puti i & " => " 
                DumpIt V(i)
                put_
            Next
            Dedent
        puti "]"
        Dedent
    End Sub
    
    Private Sub DumpClass(C)
        Indent
        dim i
        put_
        puti "{" & Typename(C) & ": " & vbCR
            Indent
        
            On Error Resume Next
                If Ubound(C.Class_Get_Properties) > 0 then
                    dim property_name, the_property
                    For i = 0 to UBound(C.Class_Get_Properties)
                        property_name = C.Class_Get_Properties(i)
                        Execute "Assign the_property, C." & C.Class_Get_Properties(i)
                        'put "property_name: " & property_name & " (" & typename(the_property) & ")" & vbCR
                        
                        If InStr(typename(the_property), "_Class") then
                            puti " " & property_name & " : " & typename(the_property) & " => "
                            DumpClass(the_property)
                        Else
                            puti "    " & property_name & " : " & typename(the_property) & " => " '& Eval("C." & property_name)
                            DumpIt(the_property)
                            If i <> UBound(C.Class_Get_Properties) then put ", "
                            put vbCR
                        End If
                    Next
                Else
                    
                End If
            On Error Goto 0
            
            Dedent
        
        puti "}" & vbCR & vbCR
        Dedent
    End Sub
    
    Sub DumpRecordset(R)
        Indent
        dim field
        put "<table border='1' cellpadding='5' >"
            put "<tr style='background-color: #333; color: white'>"
                For each field in R.Fields
                    put "<th>" & field.Name & "</th>"
                Next
            put "</tr>"
            Do until R.EOF
                put "<tr style='background-color: white'>"
                    For each field in R.Fields
                        put "<td>" & H(R(field.Name)) & "</td>"
                    Next
                put "</tr>"
                R.MoveNext
            Loop
        put "</table>"
    Dedent    
    End Sub
    
    Private Function Spaces(num)
        dim s : s = ""
        dim i
        For i = 1 to num
            s = s & "        "
        Next
        Spaces = s
    End Function
End Class

dim DataDumper_Class__Singleton
Sub Dump(V)
    If IsEmpty(DataDumper_Class__Singleton) then
        set DataDumper_Class__Singleton = new DataDumper_Class
    End If
    DataDumper_Class__Singleton.Dump V
End Sub




'=======================================================================================================================
' Strings
'=======================================================================================================================
'Capitalizes first word of the_string, leaves rest as-is
Function Capitalize(the_string)
    Capitalize = ucase(left(the_string, 1)) & mid(the_string, 2)
End Function

'-----------------------------------------------------------------------------------------------------------------------
Function Wrap(s, prefix, suffix)
    Wrap = prefix & s & suffix
End Function


'=======================================================================================================================
' Logic (i.e. decisions, searches, etc)
'=======================================================================================================================

'TODO: Expand this to accept arbitrary sets, e.g. string, recordset, dictionary, list, etc.
Function Contains(data, value)
    Contains = false
    dim idx
    select case typename(data)
        case "String"
            Contains = Choice(instr(data, value) > 0, true, false)
        
        case "Variant()"
            for idx = lbound(data) to ubound(data)
                if value = data(idx) then
                    Contains = true
                    exit for
                end if
            next
            
        case else
            Err.Raise 9, "mvc.helpers#Contains", "Unexpected type 'data', received: " & typename(data)
    end select
End Function

'-----------------------------------------------------------------------------------------------------------------------
'Boolean type checkers
'Don't forget IsArray is built-in!
Function IsString(value)
    IsString = Choice(typename(value) = "String", true, false)
End Function

Function IsDict(value)
    IsDict = Choice(typename(value) = "Dictionary", true, false)
End Function

Function IsRecordset(value)
    IsRecordset = Choice(typename(value) = "Recordset", true, false)
End Function

Function IsLinkedList(value)
    IsLinkedList = Choice(typename(value) = "LinkedList_Class", true, false)
End Function


'-----------------------------------------------------------------------------------------------------------------------
Sub Destroy(o)
    if isobject(o) then
        if not o is nothing then
            on error resume next
            o.close
            on error goto 0
            set o = nothing
        end if
    end if
End Sub

'-----------------------------------------------------------------------------------------------------------------------
Sub Quit
    response.end
End Sub

Sub Die(msg)
    put "<span style='color: #f00'>" & msg & "</span>"
    Quit
End Sub

'-----------------------------------------------------------------------------------------------------------------------
Sub DumpSession
    put "SESSION" & "<br>"
    dim session_item
    for each session_item in session.contents
        put "<b>" & session_item & "</b> : " & session.contents(session_item) & "<br>"
    next
End Sub



'=======================================================================================================================
' Adapted from Tolerable library
'=======================================================================================================================
' This subroutine allows us to ignore the difference
' between object and primitive assignments.  This is
' essential for many parts of the engine.
Public Sub Assign(ByRef var, ByVal val)
    If IsObject(val) Then
        Set var = val
    Else
        var = val
    End If
End Sub

' This is similar to the  ? :  operator of other languages.
' Unfortunately, both the  if_true  and  if_false  "branches"
' will be evalauted before the condition is even checked. So,
' you'll only want to use this for simple expressions.
Public Function Choice(ByVal cond, ByVal if_true, ByVal if_false)
    If cond Then
        Assign Choice, if_true
    Else
        Assign Choice, if_false
    End If
End Function

' Allows single-quotes to be used in place of double-quotes.
' Basically, this is a cheap trick that can make it easier
' to specify Lambdas.
Public Function Q(ByVal input)
    Q = Replace(input, "'", """")
End Function



%>
