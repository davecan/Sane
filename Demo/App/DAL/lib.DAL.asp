<%
' Since we only access one database in this example we set the global DAL singleton to
' an instance of the Database_Class in lib.Data. To see how to easily handle multiple
' databases in a single DAL see lib.DAL in the core framework.

dim DAL__Singleton : set DAL__Singleton = Nothing

Function DAL()
    If DAL__Singleton is Nothing then
        set DAL__Singleton = new Database_Class
        DAL__Singleton.Initialize "Provider=SQLOLEDB.1;Data Source=WIN-SRV-2012;Initial Catalog=NORTHWND;uid=sa;pwd=password;"
    End If
    set DAL = DAL__Singleton
End Function
%>