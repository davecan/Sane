<% 
Class FormCache_Class
    'given a form name and IRequestDictionary params (request.form) object, caches form values 
    Public Sub SerializeForm(form_name, params)
        dim form_key, form_val, serialized_key
        For Each form_key in params
            form_val = params(form_key)
            serialized_key = CachedFormKeyName(form_name, form_key)
            'put "serialize<br>"
            'put "--form_key := " & form_key & "<br>"
            'put "--form_val := " & form_val & "<br>"
            'put "--serialized_key := " & serialized_key & "<br>"
            Session(serialized_key) = form_val
        Next
    End Sub
  
    'given a form name, returns a dict with the form's stored values
    Public Function DeserializeForm(form_name)
        dim dict : set dict = Nothing
        dim serialized_key, serialized_val, form_key, form_val
    
        For Each serialized_key in Session.Contents
            'put "serialized_key: " & serialized_key & "<br>"
            If InStr(serialized_key, "mvc.form." & form_name) > 0 then
                'put "--match" & "<br>"
        
                If dict Is Nothing then
                    set dict = Server.CreateObject("Scripting.Dictionary")
                    'put "dict created<br>"
                End If
        
                form_val = Session(serialized_key)
                form_key = Replace(serialized_key, "mvc.form." & form_name & ".", "")
                dict(form_key) = form_val
                'Session.Contents.Remove serialized_key
        
                'put "--serialized_val: " & serialized_val & "<br>"
                'put "--form_val: " & form_val & "<br>"
            End If
        Next
        set DeserializeForm = dict
    End Function
  
    'given a form name, clears the keys for that form
    Public Sub ClearForm(form_name)
        dim key
        For Each key in Session.Contents
            If InStr(key, CachedFormKeyName(form_name, key)) > 0 then
                Session.Contents.Remove key
            End If
        Next
    End Sub
  
    Private Function CachedFormKeyName(form_name, key)
        CachedFormKeyName = "mvc.form." & form_name & "." & key
    End Function
End Class


dim FormCache__Singleton
Function FormCache()
  if IsEmpty(FormCache__Singleton) then
    set FormCache__Singleton = new FormCache_Class
  end if
  set FormCache = FormCache__Singleton
End Function
%>