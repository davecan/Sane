<%
' This class encapsulates database access into one location, isolating database details from the rest of the app.
' Multiple databases can be handled in one of two ways:
'
'   Option 1.   Use a single DAL_Class instance with separate public properties for each database.
'               Ex: To access Orders use DAL.Orders and to access Employees use DAL.Employees.
'
'   Option 2.   Use a separate DAL_Class instance for each database.
'               Ex:
'                   dim OrdersDAL : set OrdersDAL = new DAL_Class
'                   OrdersDAL.ConnectionString = "..."  <-- you would have to create this property to use this approach
'
' If you only access one database it is easier to just set the global DAL singleton to an instance of the
' Database_Class and use it directly. See the example project for details.

'=======================================================================================================================
' DATA ACCESS LAYER Class
'=======================================================================================================================
Class DAL_Class
    Public Database1, Database2
    
    ' could also lazy load these if desired
    Private Sub Class_Initialize
        set Database1 = new Database_Class
        Database1.Initialize "Provider=SQLOLEDB.1;Data Source=...;Initial Catalog=...;uid=...;pwd=...;"
        
        set Database2 = new Database_Class
        Database2.Initialize "Provider=SQLOLEDB.1;Data Source=...;Initial Catalog=...;uid=...;pwd=...;"
    End Sub
End Class



dim DAL__Singleton : set DAL__Singleton = Nothing

Function DAL()
    If DAL__Singleton is Nothing then
        set DAL__Singleton = new DAL_Class
    End If
    set DAL = DAL__Singleton
End Function
%>