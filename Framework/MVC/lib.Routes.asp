<%
'=======================================================================================================================
' ROUTING HELPER
'=======================================================================================================================
Class Route_Helper_Class
    Public Property Get NoCacheToken
        NoCacheToken = Timer() * 100
    End Property

    Public Sub Initialize(app_url)
        m_app_url         = app_url
        m_content_url     = m_app_url & "Content/"
        m_stylesheets_url = m_content_url & "Styles/"
        m_controllers_url = m_app_url & "Controllers/"
    End Sub
    
    Public Property Get AppURL
        AppURL = m_app_url
    end Property
    
    Public Property Get ContentURL
        ContentURL = m_content_url
    end Property
    
    Public Property Get ControllersURL
        ControllersUrl = m_controllers_url
    end Property
    
    Public Property Get StylesheetsURL
        StylesheetsURL = m_stylesheets_url
    end Property
    
    ''
     ' Generates a URL to the specified controller + action combo, with querystring parameters appended if included.
     ' 
     ' @param         controller_name     String          name of the controller
     ' @param         action_name         String          name of the controller action
     ' @param         params_array        KV Array        key/value pair array, to be converted to &key1=val1&key2=val2&...&keyn=valn 
     ' @returns 
     ''
    Public Function UrlTo(controller_name, action_name, params_array)
        dim qs : qs = TO_Querystring(params_array)
        if len(qs) > 0 then qs = "&" & qs
        UrlTo = Me.ControllersURL & controller_name & "/" & controller_name & "Controller.asp?_A=" & action_name & qs & "&_NC=" & NoCacheToken
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    ' PRIVATE
    '---------------------------------------------------------------------------------------------------------------------
    Private m_app_url
    Private m_content_url
    Private m_stylesheets_url
    Private m_controllers_url
        
    Private Function TO_Querystring(the_array)
        dim result : result = ""
        if not isempty(the_array) then
            dim idx
            for idx = lbound(the_array) to ubound(the_array) step 2
                result = result & GetParam(the_array, idx)
                'append & between parameters, but not on the last parameter
                if not (idx = ubound(the_array) - 1) then result = result & "&"
            next
        end if
        TO_Querystring = result
    End Function
    
    Private Function GetParam(params_array, key_idx)
        dim key, val    
        KeyVal params_array, key_idx, key, val
        GetParam = key & "=" & val
    End Function
end class


dim Route_Helper__Singleton : set Route_Helper__Singleton = Nothing
Function Routes()
    if Route_Helper__Singleton is Nothing then
        set Route_Helper__Singleton = new Route_Helper_Class
    end if
    set Routes = Route_Helper__Singleton
End Function

%>
