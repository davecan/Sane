<%
Class EnumerableHelper_Class
    Private m_list

    Public Sub Init(list)
        set m_list = list
    End Sub

    Public Sub Class_Terminate
        set m_list = Nothing
    End Sub

    Public Default Function Data()
        set Data = m_list
    End Function


    '---------------------------------------------------------------------------------------------------------------------
    ' Convenience wrappers
    '---------------------------------------------------------------------------------------------------------------------
    
    Public Function Count()
        Count = m_list.Count()
    End Function

    Public Function First()
        Assign First, m_list.Front()
    End Function

    Public Function Last()
        Assign Last, m_list.Back()
    End Function

    '---------------------------------------------------------------------------------------------------------------------
    ' Methods that return a single value
    '---------------------------------------------------------------------------------------------------------------------

    'true if all elements of the list satisfy the condition
    Public Function All(condition)
        dim item_, all_matched : all_matched = true
        dim it : set it = m_list.Iterator
        Do While it.HasNext
            Assign item_, it.GetNext()
            If "String" = typename(condition) then
                If Not eval(condition) then 
                    all_matched = false
                End If
             Else
                If Not condition(item_) then
                    all_matched = false
                End If
            End If
            If Not all_matched then Exit Do
        Loop
        All = all_matched
    End Function

    'true if any element of the list satisfies the condition
    Public Function Any(condition)
        Any = Not All("Not " & condition)
    End Function
    
    Public Function Max(expr)
        dim V_, item_, maxval
        dim it : set it = m_list.Iterator
        If "String" = typename(expr) then
            While it.HasNext
                Assign item_, it.GetNext()
                Assign V_, eval(expr)
                If V_ > maxval then maxval = V_
            Wend
        Else
            While it.HasNext
                Assign item_, it.GetNext()
                Assign V_, expr(item_)
                If V_ > maxval then maxval = V_
            Wend
        End If
        Max = maxval
    End Function

    Public Function Min(expr)
        dim V_, item_, minval
        dim it : set it = m_list.Iterator
        If "String" = typename(expr) then
            While it.HasNext
                Assign item_, it.GetNext()
                If IsEmpty(minval) then  ' empty is always less than everything so set it on first pass
                    Assign minval, item_
                End If
                Assign V_, eval(expr)
                If V_ < minval then minval = V_
            Wend
        Else
            While it.HasNext
                Assign item_, it.GetNext()
                If IsEmpty(minval) then
                    Assign minval, item_
                End If
                V_ = expr(item_)
                If V_ < minval then minval = V_
            Wend
        End If
        Min = minval
    End Function

    Public Function Sum(expr)
        dim V_, item_
        dim it : set it = m_list.Iterator
        While it.HasNext
            Assign item_, it.GetNext()
            execute "V_ = V_ + " & expr
        Wend
        Sum = V_
    End Function


    '---------------------------------------------------------------------------------------------------------------------
    ' Methods that return a new instance of this class
    '---------------------------------------------------------------------------------------------------------------------

    'returns a list that results from running lambda_or_proc once for every element in the list
    Public Function Map(lambda_or_proc)
        dim list2 : set list2 = new LinkedList_Class
        dim it : set it = m_list.Iterator
        dim item_
        If "String" = typename(lambda_or_proc) then
            dim V_
            While it.HasNext
                Assign item_, it.GetNext()
                execute lambda_or_proc
                list2.Push V_
            Wend
        Else
            While it.HasNext
                Assign item_, it.GetNext()
                list2.Push lambda_or_proc(item_)
            Wend
        End If
        set Map = Enumerable(list2)
    End Function

    'alias to match IEnumerable for convenience
    Public Function [Select](lambda_or_proc)
        set [Select] = Map(lambda_or_proc)
    End Function

    'returns list containing first n items
    Public Function Take(n)
        dim list2 : set list2 = new LinkedList_Class
        dim it : set it = m_list.Iterator
        dim i : i = 1
        While it.HasNext And i <= n
            list2.Push it.GetNext()
            i = i + 1
        Wend
        set Take = Enumerable(list2)
    End Function

    'returns list containing elements as long as the condition is true, and skips the remaining elements
    Public Function TakeWhile(condition)
        dim list2 : set list2 = new LinkedList_Class
        dim item_, V_, bln
        dim it : set it = m_list.Iterator
        Do While it.HasNext
            Assign item_, it.GetNext()
            If "String" = typename(condition) then
                'execute condition
                If Not eval(condition) then Exit Do
            Else
                If Not condition(item_) then Exit Do
            End If
            list2.Push item_
        Loop
        set TakeWhile = Enumerable(list2)
    End Function
    
    'returns a list containing only elements that satisfy the condition
    Public Function Where(condition)
        dim list2 : set list2 = new LinkedList_Class
        dim it : set it = m_list.Iterator
        dim item_
        While it.HasNext
            Assign item_, it.GetNext()
            If "String" = typename(condition) then
                If eval(condition) then list2.Push item_
            Else
                If condition(item_) then list2.Push item_
            End If
        Wend
        set Where = Enumerable(list2)
    End Function

End Class


Function Enumerable(list)
    dim E : set E = new EnumerableHelper_Class
    E.Init list
    set Enumerable = E
End Function
%>