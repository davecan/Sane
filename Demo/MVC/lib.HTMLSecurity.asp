<%
'=======================================================================================================================
' HTML SECURITY HELPER
'=======================================================================================================================
Class HTML_Security_Helper_Class

    '---------------------------------------------------------------------------------------------------------------------
    'Uses Scriptlet.TypeLib to generate a GUID. There may be a better/faster way than this to generate a nonce.
    Public Function Nonce()
        dim TL : set TL = CreateObject("Scriptlet.TypeLib") 
        Nonce = Left(CStr(TL.Guid), 38)    'avoids issue w/ strings appended after this token not being displayed on screen, MSFT bug
        set TL = Nothing
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    'Name is probably the combined ControllerName and ActionName of the form generator by convention
    Public Sub SetAntiCSRFToken(name)
        Session(name & ".anti_csrf_token") = Nonce()
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    'Returns the CSRF token nonce from the session corresponding to the passed name
    Public Function GetAntiCSRFToken(name)
        dim token : token = Session(name & ".anti_csrf_token")
        If Len(token) = 0 then
            SetAntiCSRFToken name
        End If
        GetAntiCSRFToken = token
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    'Removes the current CSRF token nonce for the passed name
    Public Sub ClearAntiCSRFToken(name)
        Session.Contents.Remove(name & ".anti_csrf_token")
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    'Returns true if passed nonce matches the stored CSRF token nonce for the specified name, false if not
    Public Function IsValidAntiCSRFToken(name, nonce)
        IsValidAntiCSRFToken = (GetAntiCSRFToken(name) = nonce)
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    'If an invalid CSRF nonce is passed, sets the flash and redirects using the appropriate MVC.Redirect* method.
    'If a valid CSRF nonce is passed, clears it from the cache to reset the state to the beginning.
    Public Sub OnInvalidAntiCSRFTokenRedirectToAction(token_name, token, action_name)
        OnInvalidAntiCSRFTokenRedirectToExt token_name, token, MVC.ControllerName, action_name, empty
    End Sub
    
    Public Sub OnInvalidAntiCSRFTokenRedirectToActionExt(token_name, token, action_name, params)
        OnInvalidAntiCSRFTokenRedirectToExt token_name, token, MVC.ControllerName, action_name, params
    End Sub
    
    Public Sub OnInvalidAntiCSRFTokenRedirectTo(token_name, token, controller_name, action_name)
        OnInvalidAntiCSRFTokenRedirectToExt token_name, token, controller_name, action_name
    End Sub
    
    Public Sub OnInvalidAntiCSRFTokenRedirectToExt(token_name, token, controller_name, action_name, params)
        If IsValidAntiCSRFToken(token_name, token) then
            ClearAntiCSRFToken token_name
        Else
            ClearAntiCSRFToken token_name
            Flash.AddError "Invalid form state. Please try again."
            MVC.RedirectToExt controller_name, action_name, params
        End If
    End Sub
End Class


dim HTML_Security_Helper__Singleton
Function HTMLSecurity()
    If IsEmpty(HTML_Security_Helper__Singleton) Then
        set HTML_Security_Helper__Singleton = new HTML_Security_Helper_Class
    End If
    set HTMLSecurity = HTML_Security_Helper__Singleton
End Function
%>
