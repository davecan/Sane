<%
Const SMALL_SIZE      = 10
Const MEDIUM_SIZE     = 100
Const LARGE_SIZE      = 1000
Const EXTREME_SIZE    = 10000

Class Performance_Tests
    Public Sub Setup         : End Sub
    Public Sub Teardown    : End Sub
    
    Private m_output
    
    Private Sub Out(s)
        m_output = m_output & s
    End Sub
    
    Private Sub OutLn(s)
        Out s & "<br>"
    End Sub
    
    Private Sub Class_Initialize
        m_output = ""
        OutLn "<table width='30%'>"
    End Sub
    
    Private Sub Class_Terminate
        Out "</table>"
        response.write m_output
    End Sub
    
    Private Sub Header(s)
        Out "<tr><th colspan='2'><h1>" & s & "</h1></th></tr>"
        Out "<tr><th style='text-align: left'>Name</th><th style='text-align: left'>Duration</th></tr>"
    End Sub
    
    Private Sub Blank
        Out "<tr><td>&nbsp;</td></tr>"
    End Sub
    
    Private Sub Profile(name, size, start, finish)
        Out "<tr><td>" & name & " (" & size & " nodes)</td><td>" & Round(finish - start, 2) * 1000 & "ms</td></tr>"
    End Sub
    
    Public Function TestCaseNames
        TestCaseNames = Array("Test_LinkedList", "Test_DynamicArray", "Test_ArrayList")
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    'Linked List Performance Tests
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_LinkedList(T)
        Header "LinkedList_Class"
        LL_Push       SMALL_SIZE
        LL_Get        SMALL_SIZE
        LL_Remove     SMALL_SIZE                     
                                
        Blank             
                                
        LL_Push       MEDIUM_SIZE
        LL_Get        MEDIUM_SIZE
        LL_Remove     MEDIUM_SIZE
                                
        Blank             
                                
        LL_Push       LARGE_SIZE
        LL_Get        LARGE_SIZE
        LL_Remove     LARGE_SIZE
                                
        Blank             
                                
        LL_Push       EXTREME_SIZE
        LL_Get        EXTREME_SIZE
        LL_Remove     EXTREME_SIZE
        
    End Sub
    
    Private Sub LL_Push(num_nodes)
        dim start, finish
        dim list : set list = new LinkedList_Class
        dim i
        
        start = Timer
        for i = 1 to num_nodes
            list.Push i
        next
        finish = Timer
        Profile "Add all nodes", num_nodes, start, finish
    End Sub
    
    Private Sub LL_Get(num_nodes)
        dim start, finish
        dim list : set list = new LinkedList_Class
        dim i
        
        start = Timer
        dim it : set it = list.Iterator
        while it.HasNext
            it.GetNext
        wend
        finish = Timer
        Profile "Get all nodes", num_nodes, start, finish
    End Sub
    
    Private Sub LL_Remove(num_nodes)
        dim start, finish
        dim list : set list = new LinkedList_Class
        dim i
        
        start = Timer
        while list.Count > 0
            list.Pop
        wend
        finish = Timer
        Profile "Remove all nodes", num_nodes, start, finish
    End Sub
    
    
    '---------------------------------------------------------------------------------------------------------------------
    'ArrayList Performance Tests
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_ArrayList(T)
        Header ".NET ArrayList"
        AL_Add        SMALL_SIZE
        AL_Get        SMALL_SIZE
        AL_Remove     SMALL_SIZE                     
                                
        Blank             
                                
        AL_Add        MEDIUM_SIZE
        AL_Get        MEDIUM_SIZE
        AL_Remove     MEDIUM_SIZE
                                
        Blank             
                                
        AL_Add        LARGE_SIZE
        AL_Get        LARGE_SIZE
        AL_Remove     LARGE_SIZE
                                
        Blank             
                                
        AL_Add        EXTREME_SIZE
        AL_Get        EXTREME_SIZE
        AL_Remove     EXTREME_SIZE
    End Sub
    
    Private Sub AL_Add(num_nodes)
        dim start, finish
        dim list : set list = Server.CreateObject("System.Collections.ArrayList")
        dim i
        
        start = Timer
        for i = 1 to num_nodes
            list.Add i
        next
        finish = Timer
        Profile "Add all nodes", num_nodes, start, finish
    End Sub
    
    Private Sub AL_Get(num_nodes)
        dim start, finish
        dim list : set list = Server.CreateObject("System.Collections.ArrayList")
        dim i
        
        start = Timer
        dim elt
        for each elt in list
            'nop
        next
        finish = Timer
        Profile "Get all nodes", num_nodes, start, finish
    End Sub
    
    Private Sub AL_Remove(num_nodes)
        dim start, finish
        dim list : set list = Server.CreateObject("System.Collections.ArrayList")
        dim i
        
        start = Timer
        list.RemoveRange 0, list.Count
        finish = Timer
        Profile "Remove all nodes", num_nodes, start, finish
    End Sub
    
    
    '---------------------------------------------------------------------------------------------------------------------
    ' DynamicArray Tests
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Test_DynamicArray(T)
        Header "DynamicArray_Class"
        DA_Push         SMALL_SIZE
        DA_Get          SMALL_SIZE
        DA_Pop          SMALL_SIZE                     
                                
        Blank             
                                
        'DA_Push         MEDIUM_SIZE
        'DA_Get            MEDIUM_SIZE
        'DA_Pop            MEDIUM_SIZE
        '                        
        'Blank             
        '                        
        'DA_Push         LARGE_SIZE
        'DA_Get            LARGE_SIZE
        'DA_Pop            LARGE_SIZE
        '                        
        'Blank             
        '                        
        'DA_Push         EXTREME_SIZE
        'DA_Get            EXTREME_SIZE
        'DA_Pop            EXTREME_SIZE
    End Sub
    
    Private Sub DA_Push(num_nodes)
        dim start, finish
        'dim arr : arr = array()
        'redim arr(100)
        'dim list : set list = DynamicArray(arr)
        dim list : set list = DynamicArray()
        dim i
        
        start = Timer
        for i = 1 to num_nodes
            list.Push i
        next
        finish = Timer
        Profile "Add all nodes", num_nodes, start, finish
    End Sub
    
    Private Sub DA_Get(num_nodes)
        dim start, finish
        dim list : set list = new DynamicArray_Class
        dim arr : arr = array()
        list.Initialize arr, 100
        dim i
        
        start = Timer
        dim it : set it = list.Iterator
        while it.HasNext
            it.GetNext
        wend
        finish = Timer
        Profile "Get all nodes", num_nodes, start, finish
    End Sub
    
    Private Sub DA_Pop(num_nodes)
        dim start, finish
        dim list : set list = new DynamicArray_Class
        dim arr : arr = array()
        list.Initialize arr, 100
        dim i
        
        start = Timer
        while list.Count > 0
            list.Pop
        wend
        finish = Timer
        Profile "Remove all nodes", num_nodes, start, finish
    End Sub
    
End Class
%>
