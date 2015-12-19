<%
Class Example_03_Insert_More_Records_To_Example_Table
    Public Migration

    Public Sub Up
        dim i
        For i = 101 to 200
            Migration.Do "INSERT INTO Example_Table (id, name) VALUES (" & i & ", 'ANOTHER Name " & i & "');"
        Next
    End Sub
    
    Public Sub Down
        Migration.Do "DELETE FROM Example_Table WHERE id >= 101 and id <= 200"
    End Sub
End Class

Migrations.Add "Example_03_Insert_More_Records_To_Example_Table"
%>
