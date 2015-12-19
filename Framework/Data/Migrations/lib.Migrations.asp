<%
'Represents a single migration (either up or down)
Class Migration_Class
    Private m_name
    Private m_migration_instance
    Private m_sql_array
    Private m_sql_array_size
    Private m_connection
    Private m_has_errors
    
    Public Tracing    'bool
    
    Public Property Get Name
        Name = m_name
    End Property
    
    Public Property Get Migration
        set Migration = m_migration_instance
    End Property
    
    Public Property Set Migration(obj)
        set m_migration_instance = obj
    End Property
    
    Public Property Get HasErrors
        HasErrors = m_has_errors
    End Property
    
    Private Sub Class_Initialize
        m_sql_array = array()
        redim m_sql_array(-1)
        
        m_has_errors = false
    End Sub
    
    Public Sub Initialize(name, migration_instance, connection)
        m_name = name
        set m_migration_instance = migration_instance
        set m_migration_instance.Migration = Me     'how circular can we get? ...
        set m_connection = connection
    End Sub
    
    Private Sub Class_Terminate
        'm_connection.Close
        'set m_connection = Nothing
    End Sub
    
    Public Sub [Do](sql)
        dim new_size : new_size = ubound(m_sql_array) + 1
        redim preserve m_sql_array(new_size)
        m_sql_array(new_size) = sql
        'put "Added command: " & sql
    End Sub
    
    Public Sub Irreversible
        Err.Raise 1, "Migration_Class:Irreversible", "Migration cannot proceed because this migration is irreversible."
    End Sub
    
    Public Sub DownDataWarning
        put "Migration can be downversioned but data changes cannot take place due to the nature of the Up migration in this set."
    End Sub
    
    Public Function Query(sql)
        put "Query: " & sql
        set Query = m_connection.Execute(sql)
    End Function
    
    Public Sub DoUp
        Migration.Up
        ShowCommands
        ExecuteCommands
    End Sub
    
    Public Sub DoDown
        Migration.Down
        ShowCommands
        ExecuteCommands
    End Sub
    
    Private Sub ShowCommands
        put ""
        put "Commands:"
        dim i : i = 0
        For i = 0 to ubound(m_sql_array)
            put "&nbsp;&nbsp;&nbsp;&nbsp;Command: " & m_sql_array(i)
        Next
        put ""
    End Sub
    
    Private Sub ExecuteCommands
        dim i : i = 0
        dim sql
        m_connection.BeginTrans         'wrap entire process in transaction, rollback If error encountered on any statement
        For i = 0 to ubound(m_sql_array)
            sql = m_sql_array(i)
            If not m_has_errors then    'avoid further processing If errors exist
                On Error Resume Next
                    put "Executing: " & sql
                    m_connection.Execute sql
                    If Err.Number <> 0 then        'something went wrong, rollback the transaction and display an error
                        m_has_errors = true
                        m_connection.RollbackTrans
                        put_error "Error during migration: " & Err.Description
                        put_error "SQL: " & sql
                        exit sub
                    End If
                On Error Goto 0
            End If
        Next
        m_connection.CommitTrans     'surprisingly no errors were encountered, so commit the entire transaction
    End Sub
    
    'force font color dIfference
    Private Sub put(s)
        If Me.Tracing then response.write "<div style='color: #999'>" & s & "</div>"
    End Sub
End Class


'---------------------------------------------------------------------------------------------------------------------
'Represents the collection of migrations to be performed
Class Migrations_Class
    Private m_migrations_array            ' 1-based to match migration naming scheme, ignore the first element
    Private m_migrations_array_size
    Private m_version
    Private m_connection
    Private m_connection_string
    Private m_has_errors
    
    Public Tracing    'bool
    
    Private Sub Class_Initialize
        m_migrations_array = array()
        m_migrations_array_size = 0     ' 1-based, ignore the first element
        redim m_migrations_array(m_migrations_array_size)

        m_has_errors = false
    End Sub
    
    Private Sub Class_Terminate
        On Error Resume Next
            m_connection.Close
            set m_connection = Nothing
        On Error Goto 0
    End Sub
    
    'force font color dIfference
    Private Sub put(s)
        If Me.Tracing then response.write "<div style='color: #999'>" & s & "</div>"
    End Sub
    
    Public Sub Initialize(connection_string)
        m_connection_string = connection_string
        set m_connection = Server.CreateObject("ADODB.Connection")
        m_connection.Open m_connection_string
        put "Initialized: " & typename(m_connection)
    End Sub
    
    Public Sub Add(name)
        m_migrations_array_size = m_migrations_array_size + 1
        redim preserve m_migrations_array(m_migrations_array_size)
        dim M : set M = new Migration_Class
        dim migration_instance : set migration_instance = eval("new " & name)
        M.Initialize name, migration_instance, m_connection
        M.Tracing = Me.Tracing
        set m_migrations_array(m_migrations_array_size) = M
    End Sub
    
    Public Sub MigrateUp
        MigrateUpTo m_migrations_array_size
    End Sub
    
    Public Sub MigrateUpBy(num)
        MigrateUpTo Version + num
    End Sub
    
    Public Sub MigrateUpTo(requested_version)
        requested_version = CInt(requested_version)
        put "Migrating Up To Version " & requested_version
        dim M, class_name
        
        If Version >= requested_version then
            put_error "DB already at higher version than requested up migration."
        ElseIf requested_version > m_migrations_array_size then
            put_error "Requested version exceeds available migrations. Only " & m_migrations_array_size & " migrations are available."
        Else
            While (NextVersion <= requested_version) and (not m_has_errors)
                set M = m_migrations_array(NextVersion)
                put ""
                put "<b>Up: " & M.name & "</b>"
                M.DoUp
                m_has_errors = M.HasErrors
                If not m_has_errors then IncrementVersion
            Wend
        End If
    End Sub
    
    Public Sub MigrateDown
        MigrateDownTo 0
    End Sub
    
    Public Sub MigrateDownBy(num)
        MigrateDownTo Version - num
    End Sub
    
    Public Sub MigrateDownTo(requested_version)
        requested_version = CInt(requested_version)
        put "Migrating Down To Version: " & requested_version
        dim M, class_name
        
        If requested_version < 0 then
            put_error "Cannot migrate down to a version less than 0."
        ElseIf requested_version > Version then
            put_error "Cannot migrate down to a version higher than the current version."
        ElseIf requested_version = Version then
            put_error "Cannot migrate down to the current version, already there."
        Else
            While (Version > requested_version) and (not m_has_errors)
                set M = m_migrations_array(Version)
                put ""
                put "<b>Down: " & M.Name & "</b>"
                M.DoDown
                m_has_errors = M.HasErrors
                If not m_has_errors then DecrementVersion
            Wend
        End If
    End Sub
    
    
    
    Public Property Get Version
        If IsEmpty(m_version) then
            m_version = GetDBVersion()
        End If
        Version = m_version
    End Property
    
    
    Private Property Let Version(val)
        m_version = val
        m_connection.Execute "update meta_migrations set version = " & m_version
    End Property
    
    Public Property Get NextVersion
        NextVersion = Version + 1
    End Property
    
    Private Function GetDBVersion()
        dim rs : set rs = m_connection.Execute("select version from meta_migrations")
        If rs.BOF or rs.EOF then
            GetDBVersion = NULL
        Else
            GetDBVersion = rs("version")
        End If
        rs.Close
        set rs = Nothing
    End Function
    
    Private Sub IncrementVersion
        If not m_has_errors then Version = Version + 1
    End Sub
    
    Private Sub DecrementVersion
        If not m_has_errors then Version = Version - 1
    End Sub
    
End Class


dim Migrations_Class__Singleton

Function Migrations()
    If IsEmpty(Migrations_Class__Singleton) then set Migrations_Class__Singleton = new Migrations_Class
    set Migrations = Migrations_Class__Singleton
End Function
%>
