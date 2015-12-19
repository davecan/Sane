<%
Class Example_01_Create_Example_Table
    Public Migration

    Public Sub Up
        Migration.Do "create table Example_Table (id int not null, name varchar(100) not null)"
    End Sub
    
    Public Sub Down
        Migration.Do "drop table Example_Table"
    End Sub
End Class

Migrations.Add "Example_01_Create_Example_Table"
%>
