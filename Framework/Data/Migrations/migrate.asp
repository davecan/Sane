<%
Option Explicit

Sub put(v)
    response.write v & "<br>"
End Sub

Sub put_
    put ""
End Sub

Sub put_error(s)
    put "<span style='color: red; font-weight: bold;'>" & s & "</span>"
End Sub
%>

<!--#include file="../../MVC/lib.all.asp"-->
<!--#include file="../../App/DAL/lib.DAL.asp"-->

<!--#include file="lib.Migrations.asp"-->

<%
'Have to initialize Migrations_Class before including any actual migrations, because they each automatically append themselves to the Migrations class for convenience.
'TODO: This can be refactored by not having the individual migration files auto-add themselves, but then this file must manually add each one using a slightly dIfferent
'            naming convention, i.e. given include file 01_Create_Users.asp the command would be Migrations.Add "Migration_01_Create_Users" or such. At least this way is automated.

Migrations.Initialize "Provider=SQLOLEDB.1;Data Source=WIN-SRV-2012;Initial Catalog=Nitro_Example;uid=...;pwd=...;"

Migrations.Tracing = false
%>

<!--#include file="Example_01_create_example_table.asp"-->
<!--#include file="Example_02_Insert_Records_To_Example_Table.asp"-->
<!--#include file="Example_03_Insert_More_Records_To_Example_Table.asp"-->


<%
Sub HandleMigration
    putl "<b>Starting Version: " & Migrations.Version & "</b>"
    If Request.Form("mode") = "direct" then
        If Request.Form("direction") = "Up" then
            If Len(Request.Form("to")) > 0 then
                Migrations.MigrateUpTo(Request.Form("to"))
            Else
                Migrations.MigrateUp
            End If
        ElseIf Request.Form("direction") = "Down" then
            If Len(Request.Form("to")) > 0 then
                Migrations.MigrateDownTo(Request.Form("to"))
            Else
                Migrations.MigrateDown
            End If
        End If
    ElseIf Request.Form("mode") = "up_one" then
        Migrations.MigrateUpBy 1
    ElseIf Request.Form("mode") = "down_one" then
        Migrations.MigrateDownBy 1
    End If
    putl "<b style='color: darkgreen'>Final Version: " & Migrations.Version & "</b>"
End Sub

Sub ShowForm
%>
    <form action="migrate.asp" method="POST">
        <input type="hidden" name="mode" value="direct">
        <p>
            <b>Direction: </b>
            <select name="direction">
                <option value="Up">Up</option>
                <option value="Down">Down</option>
            </select>
            &nbsp;&nbsp;
            <b>To: </b>
            <input type="text" size="5" name="to">
            &nbsp;&nbsp;
            <input type="Submit" value="Migrate!">
        </p>
    </form>
    
    <form action="migrate.asp" method="POST" style="display: inline">
        <input type="hidden" name="mode" value="up_one">
        <input type="Submit" value="Up 1">
    </form>
    
    <form action="migrate.asp" method="POST">
        <input type="hidden" name="mode" value="down_one">
        <input type="Submit" value="Down 1">
    </form>
    
    <hr>
<%
End Sub

Sub Main
    ShowForm
    
    If Len(Request.Form("mode")) > 0 then
        HandleMigration
    Else
        putl "<b>Version: " & Migrations.Version & "</b>"
    End If
End Sub
%>

<!doctype html>
<html>
<head>
    <style>
    body { font-family: calibri; }
    </style>
</head>
<body>
    <% Main %>
</body>
</html>
