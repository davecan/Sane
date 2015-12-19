<%
Option Explicit

Sub Main
    ShowFileList
    
    If Len(Request.Form("filename")) > 0 then
        ShowDocs
    End If
End Sub


Sub ShowFileList
%>
    <h1>Select File</h1>
        
        <form action="docs.asp" method="POST">
            <select name="filename">
                <%
                    dim fso : set fso = Server.CreateObject("Scripting.FileSystemObject")
                    dim files : set files = fso.GetFolder(Server.MapPath(".")).Files
                    dim file
                    For Each file in files
                        If fso.GetExtensionName(file.Path) = "asp" and file.Name <> "docs.asp" and InStr(file.Name, "_old") = 0 then
                            response.write "<option value='" & file.Name & "'>" & file.Name & "</option>"
                        End If
                    Next
                %>
            </select>
            <input type="submit" value="Get Docs">
        </form>
        <hr>
<%
End Sub


Sub ShowDocs
    dim fso : set fso = Server.CreateObject("Scripting.FileSystemObject")
    dim path : path = fso.GetFolder(Server.MapPath(".")).Path
    dim file : set file = fso.OpenTextFile(path & "\" & Request.Form("filename"))
    
    dim re : set re = new RegExp
    With re
        .Pattern = "Public Property|Public Sub|Public Function"
        .Global = true
        .IgnoreCase = true
    End With
    
    dim line, matches, result
    
    Do Until file.AtEndOfStream
        line = file.ReadLine()
        set matches = re.Execute(line)
        If matches.Count > 0 then
            result = line
            result = Replace(result, "Public Property", "<span class='subdued'>Property</span>")
            result = Replace(result, "Public Sub",            "<span class='subdued'>Sub</span>")
            result = Replace(result, "Public Function", "<span class='subdued'>Function</span>")
            response.write "<p>" & result & "</p>"
        End If
    Loop
    
End Sub


%>

<!doctype html>
<html>
    <head>
        <style>
            body { font-family: calibri; }
            
            p { font-weight: bold; }
            .subdued { font-weight: normal; color: #999; }
        </style>
    </head>
    <body>
        
        <% Call Main %>
    
    </body>
</html>
