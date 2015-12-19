<%
Class Example_02_Insert_Records_To_Example_Table
    Public Migration

    Public Sub Up
        dim i
        For i = 1 to 100
            Migration.Do "INSERT INTO Example_Table (id, name) VALUES (" & i & ", 'Name " & i & "');"
        Next
    End Sub
    
    Public Sub Down
        Migration.Do "DELETE FROM Example_Table WHERE id >= 1 and id <= 100"
    End Sub
End Class

Migrations.Add "Example_02_Insert_Records_To_Example_Table"
%>
