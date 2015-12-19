<%
Class Database_Class
    Private m_connection
    Private m_connection_string
    
    Private m_trace_enabled
    Public Sub set_trace(bool) : m_trace_enabled = bool : End Sub
    Public Property Get is_trace_enabled : is_trace_enabled = m_trace_enabled : End Property
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Initialize(connection_string)
        m_connection_string = connection_string
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Function Query(sql, params)
        dim cmd : set cmd = server.createobject("adodb.command")
        set cmd.ActiveConnection = Connection
        cmd.CommandText = sql
        
        dim rs
        
        If IsArray(params) then
            set rs = cmd.Execute(, params)
        ElseIf Not IsEmpty(params) then    ' one parameter
            set rs = cmd.Execute(, Array(params))
        Else
            set rs = cmd.Execute()
        End If
        
        set Query = rs
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Function PagedQuery(sql, params, per_page, page_num)
        dim cmd : set cmd = server.createobject("adodb.command")
        set cmd.ActiveConnection = Connection
        cmd.CommandText = sql
        
        cmd.CommandType = 1                         'adCmdText
        cmd.ActiveConnection.CursorLocation = 3     'adUseClient
        
        dim rs
        
        If IsArray(params) then
            set rs = cmd.Execute(, params)
        ElseIf Not IsEmpty(params) then    ' one parameter
            set rs = cmd.Execute(, Array(params))
        Else
            set rs = cmd.Execute()
        End If
        
        If Not rs.EOF then
            rs.PageSize = 1
            rs.CacheSize = 1
            rs.AbsolutePage = 1
        End If
        
        set PagedQuery = rs
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub [Execute](sql, params)
        me.query sql, params
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub BeginTransaction
        Connection.BeginTrans
    End Sub
    
    Public Sub RollbackTransaction
        Connection.RollbackTrans
    End Sub
    
    Public Sub CommitTransaction
        Connection.CommitTrans
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Private Methods
    '---------------------------------------------------------------------------------------------------------------------
    Private Sub Class_terminate
        Destroy m_connection
    End Sub
    
    Private Function Connection
        if not isobject(m_connection) then 
            set m_connection = Server.CreateObject("adodb.connection")
            m_connection.open m_connection_string
        end if
        set Connection = m_connection
    End Function
end Class
%>
