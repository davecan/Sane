<%
'=======================================================================================================================
' KVArray
' Relatively painless implementation of key/value pair arrays without requiring a full Scripting.Dictionary COM instance.
' A KVArray is a standard array where element i is the key and element i+1 is the value. Loops must step by 2.
'=======================================================================================================================
'given a KVArray and key index, returns the key and value
'pre:     kv_array has at least key_idx and key_idx + 1 values
'post:    key and val are populated
Sub KeyVal(kv_array, key_idx, ByRef key, ByRef val)
    if (key_idx + 1 > ubound(kv_array)) then err.raise 1, "KeyVal", "expected key_idx < " & ubound(kv_array) - 1 & ", got: " & key_idx
    key = kv_array(key_idx)
    val = kv_array(key_idx + 1)
End Sub

'---------------------------------------------------------------------------------------------------------------------
'Given a KVArray, a key and a value, appends the key and value to the end of the KVArray
Sub KVAppend(ByRef kv_array, key, val)
    dim i : i = ubound(kv_array)
    redim preserve kv_array(i + 2)
    kv_array(i + 1) = key
    kv_array(i + 2) = val
End Sub

'-----------------------------------------------------------------------------------------------------------------------
'Given a KVArray and two variants, populates the first variant with all keys and the second variant with all values.
'If 
'Pre:   kv_array has at least key_idx and key_idx + 1 values
'Post:  key_array contains all keys in kvarray.
'         val_array contains all values in kvarray.
'         key_array and val_array values are in corresponding order, i.e. key_array(i) corresponds to val_array(i).
Sub KVUnzip(kv_array, key_array, val_array)
    dim kv_array_size             : kv_array_size             = ubound(kv_array)
    dim num_pairs                     : num_pairs                     = (kv_array_size + 1) / 2
    dim result_array_size     : result_array_size     = num_pairs - 1
    
    'Extend existing key_array or create new array to hold the keys
    If IsArray(key_array) then
        redim preserve key_array(ubound(key_array) + result_array_size)
    Else
        key_array = Array()
        redim key_array(result_array_size)
    End If
    
    'Extend existing val array or create new array to hold the values
    If IsArray(val_array) then
        redim preserve val_array(ubound(val_array) + result_array_size)
    Else
        val_array = Array()
        redim val_array(num_pairs - 1)
    End If
    
    'Unzip the KVArray into the two output arrays
    dim i, key, val
    dim key_val_arrays_idx : key_val_arrays_idx = 0    ' used to sync loading the key_array and val_array
    For i = 0 to ubound(kv_array) step 2
        KeyVal kv_array, i, key, val
        key_array(key_val_arrays_idx) = key
        val_array(key_val_arrays_idx) = val
        key_val_arrays_idx = key_val_arrays_idx + 1    ' increment by 1 because loop goes to next pair in kv_array
    Next
End Sub

'---------------------------------------------------------------------------------------------------------------------
'Given a KVArray, dumps it to the screen. Useful for debugging purposes.
Sub DumpKVArray(kv_array)
    dim i, key, val
    For i = 0 to ubound(kv_array) step 2
        KeyVal kv_array, i, key, val
        put key & " => " & val & "<br>"
    Next
End Sub


'=======================================================================================================================
' Pair Class
' Holds a pair of values, i.e. a key value pair, recordset field name/value pair, etc. 
' Similar to the C++ STL std::pair class. Useful for some iteration and the like.
'
' This was an interesting idea but so far has not really been used, oh well......
'=======================================================================================================================
Class Pair_Class
    Private m_first, m_second

    Public Property Get First       : First       = m_first     : End Property
    Public Property Get [Second]    : [Second]    = m_second    : End Property
    
    Public Default Property Get TO_String
        TO_String = First & " " & [Second]
    End Property
    
    Public Sub Initialize(ByVal firstval, ByVal secondval)
        Assign m_first, firstval
        Assign m_second, secondval
    End Sub
    
    'Swaps the two values
    Public Sub Swap
        dim tmp
        Assign tmp, m_second
        Assign m_second, m_first
        Assign m_first, tmp
    End Sub
End Class

Function MakePair(ByVal firstval, ByVal secondval)
    dim P : set P = new Pair_Class
    P.Initialize firstval, secondval
    set MakePair = P
End Function



'=======================================================================================================================
' Linked List - From the Tolerable lib
'=======================================================================================================================
' This is just here for reference
Class Iterator_Class
        Public Function HasNext()
        End Function
        
        Public Function PeekNext()
        End Function
        
        Public Function GetNext()
        End Function
        
        
        Public Function HasPrev()
        End Function
        
        Public Function PeekPrev()
        End Function
        
        Public Function GetPrev()
        End Function
End Class


Class Enumerator_Source_Iterator_Class
    Private m_iter
        
    Public Sub Initialize(ByVal iter)
            Set m_iter = iter
    End Sub
        
    Private Sub Class_Terminate()
            Set m_iter = Nothing
    End Sub
        
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_iter.HasNext Then
            Assign retval, m_iter.GetNext
            successful = True
        Else
            successful = False
        End If
    End Sub
End Class


Public Function En_Iterator(ByVal iter)
        Dim retval
        Set retval = New Enumerator_Source_Iterator_Class
        retval.Initialize iter
        Set En_Iterator = Enumerator(retval)
End Function


Class LinkedList_Node_Class
    Public    m_prev
    Public    m_next
    Public    m_value

    Private Sub Class_Initialize()
        Set m_prev = Nothing
        Set m_next = Nothing
    End Sub
        
    Private Sub Class_Terminate()
            Set m_prev    = Nothing
            Set m_next    = Nothing
            Set m_value = Nothing
    End Sub
        
    Public Sub SetValue(ByVal value)
            Assign m_value, value
    End Sub
End Class


Class Iterator_LinkedList_Class
    Private m_left
    Private m_right
        
    Public Sub Initialize(ByVal r)
        Set m_left    = Nothing
        Set m_right = r
    End Sub
        
    Private Sub Class_Terminate()
        Set m_Left    = Nothing
        Set m_Right = Nothing
    End Sub
        
    Public Function HasNext()
        HasNext = Not(m_right Is Nothing)
    End Function
        
    Public Function PeekNext()
        Assign PeekNext, m_right.m_value
    End Function
        
    Public Function GetNext()
        Assign GetNext, m_right.m_value
        Set m_left    = m_right
        Set m_right = m_right.m_next
    End Function
        
    Public Function HasPrev()
        HasPrev = Not(m_left Is Nothing)
    End Function
        
    Public Function PeekPrev()
        Assign PeekPrev, m_left.m_value
    End Function
        
    Public Function GetPrev()
        Assign GetPrev, m_left.m_value
        Set m_right = m_left
        Set m_left  = m_left.m_prev
    End Function
End Class


'-----------------------------------------------------------------------------------------------------------------------
Class LinkedList_Class
    Private m_first
    Private m_last
    Private m_size
        
    Private Sub Class_Initialize()
        Me.Reset
    End Sub
        
    Private Sub Class_Terminate()
        Me.Reset
    End Sub
        
    Public Function Clear()
        Set m_first = Nothing
        Set m_last    = Nothing
        m_size            = 0
        Set Clear     = Me
    End Function
        
    Private Function NewNode(ByVal value)
        Dim retval
        Set retval = New LinkedList_Node_Class
        retval.SetValue value
        Set NewNode = retval
    End Function
        
    Public Sub Reset()
        Set m_first = Nothing
        Set m_last    = Nothing
        m_size            = 0
    End Sub
        
    Public Function IsEmpty()
        IsEmpty = (m_last Is Nothing)
    End Function
        
    Public Property Get Count
        Count = m_size
    End Property
        
    'I just like .Size better than .Count sometimes, sue me
    Public Property Get Size
        Size = m_size
    End Property
        
    Public Function Iterator()
        Dim retval
        Set retval = New Iterator_LinkedList_Class
        retval.Initialize m_first
        Set Iterator = retval
    End Function
        
    Public Function Push(ByVal value)
        Dim temp
        Set temp = NewNode(value)
        If Me.IsEmpty Then
                Set m_first = temp
                Set m_last  = temp
        Else
                Set temp.m_prev   = m_last
                Set m_last.m_next = temp
                Set m_last        = temp
        End If
        m_size = m_size + 1
        Set Push = Me
    End Function


    Public Function Peek()
        ' TODO: Error handling
        Assign Peek, m_last.m_value
    End Function

    ' Alias for Peek
    Public Function Back()
        ' TODO: Error handling
        Assign Back, m_last.m_value
    End Function
        
    Public Function Pop()
        Dim temp
                
        ' TODO: Error Handling
        Assign Pop, m_last.m_value

        Set temp            = m_last
        Set m_last          = temp.m_prev
        Set temp.m_prev     = Nothing
        If m_last Is Nothing Then
            Set m_first = Nothing
        Else
            Set m_last.m_next = Nothing
        End If
        m_size = m_size - 1
    End Function



    Public Function Unshift(ByVal value)
        Dim temp
        Set temp = NewNode(value)
        If Me.IsEmpty Then
            Set m_first = temp
            Set m_last  = temp
        Else
            Set temp.m_next    = m_first
            Set m_first.m_prev = temp
            Set m_first        = temp
        End If
        m_size = m_size + 1
        Set Unshift = Me
    End Function



    ' Alias for Peek
    Public Function Front()
        ' TODO: Error handling
        Assign Front, m_first.m_value
    End Function
        
    Public Function Shift()
        Dim temp
                
        ' TODO: Error Handling
        Assign Shift, m_first.m_value

        Set temp         = m_first
        Set m_first      = temp.m_next
        Set temp.m_next  = Nothing
        If m_first Is Nothing Then
            Set m_last = Nothing
        Else
            Set m_first.m_prev = Nothing
        End If
        m_size = m_size - 1
    End Function

    Public Function TO_Array()
        Dim i, iter
                
        ReDim retval(Me.Count - 1)
        i = 0
        Set iter = Me.Iterator
        While iter.HasNext
            retval(i) = iter.GetNext
            i = i + 1
        Wend
        TO_Array = retval
    End Function
        
    Public Function TO_En()
        Set TO_En = En_Iterator(Iterator)
    End Function

End Class




'=======================================================================================================================
' Dynamic Array - From the Tolerable lib
'=======================================================================================================================
Class DynamicArray_Class
    Private m_data
    Private m_size
        
    Public Sub Initialize(ByVal d, ByVal s)
        m_data = d
        m_size = s
    End Sub
        
    Private Sub Class_Terminate()
        Set m_data = Nothing
    End Sub
        
        
    Public Property Get Capacity
        Capacity = UBOUND(m_data) + 1
    End Property
        
    Public Property Get Count
        Count = m_size
    End Property
        
    ' Alias for Count
    Public Property Get Size
        Size = m_size
    End Property
        
    Public Function IsEmpty()
        IsEmpty = (m_size = 0)
    End Function
        
    Public Function Clear()
        m_size        = 0
        Set Clear = Me
    End Function
        
    Private Sub Grow
        ' TODO: There's probably a better way to
        '       do this.  Doubling might be excessive
        ReDim Preserve m_data(UBOUND(m_data) * 2)
    End Sub
        
    Public Function Push(ByVal val)
        If m_size >= UBOUND(m_data) Then
                Grow
        End If
        Assign m_data(m_size), val
        m_size = m_size + 1
        Set Push = Me
    End Function
        
        
    ' Look at the last element
    Public Function Peek()
        Assign Peek, m_data(m_size - 1)
    End Function
        
    ' Look at the last element and
    ' pop it off of the list
    Public Function Pop()
        Assign Pop, m_data(m_size - 1)
        m_size = m_size - 1
    End Function
        
        
    ' If pseudo_index < 0, then we assume we're counting
    ' from the back of the Array.
    Private Function CalculateIndex(ByVal pseudo_index)
        If pseudo_index >= 0 Then
            CalculateIndex = pseudo_index
        Else
            CalculateIndex = m_size + pseudo_index
        End If
    End Function
        
    Public Default Function Item(ByVal i)
        Assign Item, m_data(CalculateIndex(i))
    End Function
        
        
    ' This does not treat negative indices as wrap-around.
    ' Thus, it is slightly faster.
    Public Function FastItem(ByVal i)
        Assign FastItem, m_data(i)
    End Function


    Public Function Slice(ByVal s, ByVal e)
        s = CalculateIndex(s)
        e = CalculateIndex(e)
        If e < s Then
            Set Slice = DynamicArray()
        Else
            ReDim retval(e - s)
            Dim i, j
            j = 0
            For i = s to e
                    Assign retval(j), m_data(i)
                    j = j + 1
            Next
            Set Slice = DynamicArray1(retval)
        End If
    End Function
        
        
    Public Function Iterator()
        Dim retval
        Set retval = New Iterator_DynamicArray_Class
        retval.Initialize Me
        Set Iterator = retval
    End Function
        
    Public Function TO_En()
        Set TO_En = En_Iterator(Me.Iterator)
    End Function
        
    Public Function TO_Array()
        Dim i
        ReDim retval(m_size - 1)
        For i = 0 to UBOUND(retval)
                Assign retval(i), m_data(i)
        Next
        TO_Array = retval
    End Function

End Class


Public Function DynamicArray()
    ReDim data(3)
    Set DynamicArray = DynamicArray2(data, 0)
End Function

Public Function DynamicArray1(ByVal data)
    Set DynamicArray1 = DynamicArray2(data, UBOUND(data) + 1)
End Function

Private Function DynamicArray2(ByVal data, ByVal size)
    Dim retval
    Set retval = New DynamicArray_Class
    retval.Initialize data, size
    Set DynamicArray2 = retval
End Function


Class Iterator_DynamicArray_Class
    Private m_dynamic_array
    Private m_index
        
    Public Sub Initialize(ByVal dynamic_array)
        Set m_dynamic_array = dynamic_array
        m_index = 0
    End Sub
        
    Private Sub Class_Terminate
        Set m_dynamic_array = Nothing
    End Sub
        
    Public Function HasNext()
        HasNext = (m_index < m_dynamic_array.Size)
    End Function
        
    Public Function PeekNext()
        Assign PeekNext, m_dynamic_array.FastItem(m_index)
    End Function
        
    Public Function GetNext()
        Assign GetNext, m_dynamic_array.FastItem(m_index)
        m_index = m_index + 1
    End Function
        
    Public Function HasPrev()
        HasPrev = (m_index > 0)
    End Function
        
    Public Function PeekPrev()
        Assign PeekPrev, m_dynamic_array.FastItem(m_index - 1)
    End Function
        
    Public Function GetPrev()
        Assign GetPrev, m_dynamic_array.FastItem(m_index - 1)
        m_index = m_index - 1
    End Function
End Class



'=======================================================================================================================
' Other Iterators
'=======================================================================================================================
'!!! EXPERIMENTAL !!!    May not be very useful, oh well...
Class Iterator_Recordset_Class
    Private m_rs
    Private m_record_count
    Private m_current_index
    Private m_field_names        'cached array
    
    Public Sub Initialize(ByVal rs)
        Set m_rs = rs
        m_rs.MoveFirst
        m_rs.MovePrevious
        m_record_count    = rs.RecordCount
        m_current_index = 0
        
        'cache field names
        m_field_names = array()
        redim m_field_names(m_rs.Fields.Count)
        
        dim field
        dim i : i = 0
        for each field in m_rs.Fields
            m_field_names(i) = field.Name
        next
    End Sub
    
    Private Sub Class_Terminate
        Set m_rs = Nothing
    End Sub
    
    Public Function HasNext()
        HasNext = (m_current_index < m_record_count)
        put "m_current_index := " & m_current_index
        put "m_record_count    := " & m_record_count
    End Function
    
    Public Function PeekNext
        if HasNext then
            m_rs.MoveNext
            Assign PeekNext, GetPairs
            m_rs.MovePrevious
        else
            set PeekNext = Nothing
        end if
    End Function
    
    Private Function GetPairs
        
    End Function
    
    Public Function GetNext
        if m_current_index < m_record_count then
            Assign GetNext, m_rs
            m_rs.MoveNext
            m_current_index = m_current_index + 1
        else
            set GetNext = Nothing
        end if
    End Function
    
    Public Function HasPrev()
        if m_rs.BOF then
            HasPrev = false
        else
            m_rs.MovePrevious
            HasPrev = Choice(m_rs.BOF, false, true)
            m_rs.MoveNext
        end if
    End Function
    
    Public Function PeekPrev
        m_rs.MovePrevious
        if m_rs.BOF then
            set PeekPrev = Nothing
        else
            Assign PeekPrev, m_rs
        end if
        m_rs.MoveNext
    End Function
    
    Public Function GetPrev
        m_rs.MovePrevious
        if m_rs.BOF then
            set GetPrev = Nothing
        else
            Assign GetPrev, m_rs
        end if
    End Function
End Class


Class Iterator_Dictionary_Class
    Private m_dic
    Private m_keys                    'array
    Private m_idx                     'current array index
    Private m_keys_ubound     'cached ubound(m_keys)
    
    Public Sub Initialize(ByVal dic)
        set m_dic         = dic
        m_keys                = m_dic.Keys()
        m_idx                 = -1
        m_keys_ubound = ubound(m_keys)
    End Sub
    
    Private Sub Class_Terminate
        set m_dic = Nothing
    End Sub
    
    Public Function HasNext()
        HasNext = (m_idx < m_keys_ubound)
    End Function
    
    Public Function PeekNext()
        Assign PeekNext, m_dic(m_keys(m_idx + 1))
    End Function
    
    Public Function GetNext()
        Assign GetNext, m_dic(m_keys(m_idx + 1))
        m_idx = m_idx + 1
    End Function
    
    Public Function HasPrev()
        HasPrev = (m_idx > 0)
    End Function
    
    Public Function PeekPrev()
        Assign PeekPrev, m_dic(m_keys(m_idx - 1))
    End Function
    
    Public Function GetPrev()
        Assign GetPrev, m_dic(m_keys(m_idx - 1))
        m_idx = m_idx - 1
    End Function
End Class


'=======================================================================================================================
' Iterator Factory
'=======================================================================================================================
'Returns the appropriate iterator for the passed-in collection. Errors if unknown collection.
Function IteratorFor(col)
    dim result
    select case typename(col)
        case "LinkedList_Class"     : set result = new Iterator_LinkedList_Class
        case "Dictionary"           : set result = new Iterator_Dictionary_Class
        case "Recordset"            : set result = new Iterator_Recordset_Class
    end select
    result.Initialize col
    set IteratorFor = result
End Function
%>
