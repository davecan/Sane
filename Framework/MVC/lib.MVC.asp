<%
Response.ExpiresAbsolute = "2000-01-01" 
Response.AddHeader "pragma", "no-cache" 
Response.AddHeader "cache-control", "private, no-cache, must-revalidate"


'=======================================================================================================================
' MVC Dispatcher
'=======================================================================================================================
Class MVC_Dispatcher_Class
    Private m_controller_name
    Private m_action_name
    Private m_default_action_name
    Private m_action_params
    Private m_controller_instance
    Private m_is_partial
    
    Private Sub Class_initialize
        m_default_action_name = "Index"
        SetControllerActionNames
        'SetActionParams
    End Sub
    
    Public Property Get ControllerName
        ControllerName = m_controller_name
    end Property
    
    Public Property Get ActionName
        ActionName = m_action_name
    end Property
    
    Public Property Get IsPartial
        IsPartial = m_is_partial
    End Property
    
    'Public Property Get ActionParams
    '    if IsEmpty(m_action_params) then set m_action_params = Server.CreateObject("Scripting.Dictionary")
    '    set ActionParams = m_action_params
    'end Property
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Instantiates the controller and executes the requested action on the controller.
    Public Sub Dispatch
        dim class_name : class_name = m_controller_name & "Controller"
        
        'set the global controller reference
        executeglobal "dim Controller : set Controller = new " & class_name
        
        If Request.Querystring("_P").Count = 1 then ' = 1 Or Request.Querystring("_P") = "true" then
            m_is_partial = true
        Else
            m_is_partial = false
        End If
        
        If Not IsPartial then
        %> 
            <!--#include file="../App/Views/Shared/layout.header.asp"-->
        <%
        End If
        
        ExecuteAction ActionName
        
        If Not IsPartial then
        %>    
            <!--#include file="../App/Views/Shared/layout.footer.asp"-->
        <%
        End If
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Executes the requested action on the current already-instantiated controller
    Private Sub ExecuteAction(action_name)
        ' no longer want to pass Request.Form as a parameter to actions -- allows actions to be called any way, more flexibility
        Execute "Controller." & action_name
    End Sub
    
 
    '---------------------------------------------------------------------------------------------------------------------
    ' Ensures an action request comes in via HTTP POST only. Raises error if not.
    Public Sub RequirePost
        If Request.Form.Count = 0 Then Err.Raise 1, "MVC_Helper_Class:RequirePost", "Action only responds to POST requests."
    End Sub
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub RedirectTo(controller_name, action_name)
        RedirectToExt controller_name, action_name, empty
    End Sub
    
    ' Redirects the browser to the specified action on the specified controller with the specified querystring parameters.
    ' params is a KVArray of querystring parameters.
    Public Sub RedirectToExt(controller_name, action_name, params)
        Response.Redirect Routes.UrlTo(controller_name, action_name, params)
    End Sub
    
    ' Shortcut for RedirectToActionExt that does not require passing a parameters argument.
    Public Sub RedirectToAction(ByVal action_name)
        RedirectToActionExt action_name, empty
    End Sub
    
    ' Redirects the browser to the specified action in the current controller, passing the included parameters. 
    ' params is a KVArray of querystring parameters.
    Public Sub RedirectToActionExt(ByVal action_name, ByVal params)
        RedirectToExt ControllerName, action_name, params
    End Sub
    
    ' Redirects to the specified action using a form POST.
    Public Sub RedirectToActionPOST(action_name)
        RedirectToActionExtPOST action_name, empty
    End Sub
    
    ' Redirects to the specified action name on the current controller using a form post
    Public Sub RedirectToActionExtPOST(action_name, params)
        put "<form id='mvc_redirect_to_action_post' action='" & Routes.UrlTo(ControllerName, action_name, empty) & "' method='POST'>"
            put "<input type='hidden' name='mvc_redirect_to_action_post_flag' value='1'>"
            if Not IsEmpty(params) then
                dim i, key, val
                for i = 0 to ubound(params) step 2
                    KeyVal params, i, key, val
                    put "<input type='hidden' name='" & key & "' value='" & val & "'>"
                next
            end if
        put "</form>"
        put "<script type='text/javascript'>"
            put "$('#mvc_redirect_to_action_post').submit();"
        put "</script>"
    End Sub
     
    
    '---------------------------------------------------------------------------------------------------------------------
    ' PRIVATE 
    '---------------------------------------------------------------------------------------------------------------------
    Private Sub SetControllerActionNames
        dim full_path           : full_path           = request.servervariables("path_info")
        dim part_path           : part_path           = split(full_path, Routes.ControllersUrl)(1)
        dim part_path_split     : part_path_split     = split(part_path, "/")
        
        m_controller_name = part_path_split(0)
        m_action_name     = Choice(request("_A") <> "", request("_A"), m_default_action_name)
    End Sub
    
    ' This is deprecated to avoid creating a Dictionary object with every request.
    ' Hasn't been used in forever anyway.
    'Private Sub SetActionParams
    '    dim key, val
    '    'set m_action_params = Server.CreateObject("scripting.dictionary")
    '    for each key in request.querystring
    '        val = request.querystring(key)
    '        'ignore service keys
    '        if instr(1, "_A", key, 1) = 0 then
    '            ActionParams.add key, CStr(val)
    '        end if
    '    next
    'End Sub
    
end Class



dim MVC_Dispatcher_Class__Singleton
Function MVC()
    if IsEmpty(MVC_Dispatcher_Class__Singleton) then
        set MVC_Dispatcher_Class__Singleton = new MVC_Dispatcher_Class
    end if
    set MVC = MVC_Dispatcher_Class__Singleton
End Function


%>
