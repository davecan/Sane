<%
'=======================================================================================================================
' HTML HELPER
'=======================================================================================================================
Class HTML_Helper_Class
    'Duplicate of Routes.NoCacheToken, copied to avoid extra lookup into the Routes object for such a trivial function.
    'Allows caller to reference HTML.NoCacheToken in cases where it seems to feel right.
    Public Property Get NoCacheToken
        NoCacheToken = Timer() * 100
    End Property
    
    'Ensures safe output
    Public Function Encode(ByVal value)
        If Not IsEmpty(value) and Not IsNull(value) then
            Encode = Server.HtmlEncode(value)
        End If
    End Function

    '---------------------------------------------------------------------------------------------------------------------
    'LinkTo and its relatives DO NOT HTMLEncode the link_text! This allows use of HTML within the link, especially
    'useful for Bootstrap icons and the like.
    '
    'Bottom Line: If you need to HTMLEncode the link text YOU MUST DO IT YOURSELF! The H() method makes this easy!
    Public Function LinkTo(link_text, controller_name, action_name)
        LinkTo = LinkToExt(link_text, controller_name, action_name, empty, empty)
    End Function

    Public Function LinkToExt(link_text, controller_name, action_name, params_array, attribs_array)
        LinkToExt = "<a href='" & Encode(Routes.UrlTo(controller_name, action_name, params_array)) & "'" &_ 
                    HtmlAttribs(attribs_array) & ">" & link_text & "</a>" & vbCR
    End Function
    
    Public Function LinkToIf(condition, link_text, controller_name, action_name)
        if condition then
            LinkToIf = LinkToExt(link_text, controller_name, action_name, empty, empty)
        end if
    End Function
    
    Public Function LinkToExtIf(condition, link_text, controller_name, action_name, params_array, attribs_array)
        if condition then
            LinkToExtIf = LinkToExt(link_text, controller_name, action_name, params_array, attribs_array)
        end if
    End Function
    
    Public Function LinkToUnless(condition, link_text, controller_name, action_name)
        if not condition then
            LinkToIf = LinkToExt(link_text, controller_name, action_name, empty, empty)
        end if
    End Function
    
    Public Function LinkToExtUnless(condition, link_text, controller_name, action_name, params_array, attribs_array)
        if not condition then
            LinkToExtUnless = LinkToExt(link_text, controller_name, action_name, params_array, attribs_array)
        end if
    End Function
    
    
    ''
     ' Creates a form button and a hidden form to enforce POST submissions. Params are in hidden fields.
     ''
    'Public Function PostButtonLinkTo(controller_name, action_name, params)
    '    dim id : id = "post_button__" & controller_name & action_name
    '    dim s
    '    s = "<form id='" & id & "' action='" & Routes.UrlTo(controller_name, action_name, empty) & "' method='POST'>"
    '    dim i, key, val
    '    for i = 0 to ubound(params) step 2
    '        KeyVal params, i, key, val
    '        s = s & "<input type='hidden' name='" & key & "' value='" & val & "'>"
    '    next
    '    s = s & "<input type='submit' value='&gt;&gt;'>"
    '    s = s & "</form>"
    '    PostButtonLinkTo = s
    'End Function
    
    Public Function PostButtonTo(button_contents, controller_name, action_name, form_fields)
        PostButtonTo = PostButtonToExt(button_contents, controller_name, action_name, form_fields, empty)
    End Function
    
    Public Function PostButtonToExt(button_contents, controller_name, action_name, form_fields, url_params)
        dim s : s = "<form action='" & Routes.UrlTo(controller_name, action_name, url_params) & "' method='POST' style='margin: 0;'>"
            dim i, key, val
            for i = 0 to ubound(form_fields) step 2
                KeyVal form_fields, i, key, val
                s = s & HTML.Hidden(key, val)
            next
            s = s & HTML.SubmitButton(button_contents)
        s = s & "</form>" & vbCR
        PostButtonToExt = s
    End Function
    
    Public Function AppStylesheetTag
        AppStylesheetTag = StylesheetTag(Routes.StylesheetsURL & "App.css")
    End Function
    
    Public Function ControllerStylesheetTag
        ControllerStylesheetTag = StylesheetTag(Routes.StylesheetsUrl & MVC.ControllerName & "Controller.css")
    End Function
    
    Public Function StylesheetTag(url)
        StylesheetTag = "<link rel='stylesheet' href='" & Encode(url) & "?" & Year(now) & Month(now) & Day(now) & Hour(now) & Minute(now) & Second(now) & "'>" & vbCR
    End Function
    
    Public Function JSTag(url)
        JSTag = "<script type='text/javascript' src='" & Encode(url) & "'></script>" & vbCR
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Form Helpers
    '---------------------------------------------------------------------------------------------------------------------
    Public Function FormTag(controller_name, action_name, route_attribs, form_attribs)
        FormTag = "<form action='" & Routes.UrlTo(controller_name, action_name, route_attribs) & "' method='POST' " & HtmlAttribs(form_attribs) & ">" & vbCR
    End Function
    
    Public Function Label(name, for_name)
        Label = LabelExt(name, for_name, empty)
    End Function
    
    Public Function LabelExt(name, for_name, attribs)
        LabelExt = "<label for='" & Encode(for_name) & "' " & HtmlAttribs(attribs) & ">" & Encode(name) & "</label>" & vbCR
    End Function
    
    Public Function Hidden(id, value)
        Hidden = HiddenExt(id, value, empty)
    End Function
    
    Public Function HiddenExt(id, value, attribs)
        HiddenExt = "<input type='hidden' id='" & Encode(id) & "' name='" & Encode(id) & "' value='" & Encode(value) & "' " & HtmlAttribs(attribs) & " >" & vbCR
    End Function
    
    Public Function TextBox(id, value)
        TextBox = TextBoxExt(id, value, empty)
    End Function
    
    Public Function TextBoxExt(id, value, attribs)
        TextBoxExt = "<input type='text'    id='" & Encode(id) & "' name='" & Encode(id) & "' value='" & Encode(value) & "' " & HtmlAttribs(attribs) & " >" & vbCR
    End Function
    
    Public Function TextArea(id, value, rows, cols)
        TextArea = TextAreaExt(id, value, rows, cols, empty)
    End Function
    
    Public Function TextAreaExt(id, value, rows, cols, attribs)
        TextAreaExt = "<textarea id='" & Encode(id) & "' name='" & Encode(id) & "' cols='" & Encode(cols) & "' rows='" & Encode(rows) & "' " & HtmlAttribs(attribs) & " >" &_
                                    Encode(value) & "</textarea>" & vbCR
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    'If list is a recordset then option_value_field and option_text_field are required.
    'If list is an array the method assumes it is a KVArray and those parameters are ignored.
    Public Function DropDownList(id, selected_value, list, option_value_field, option_text_field)
        DropDownList = DropDownListExt(id, selected_value, list, option_value_field, option_text_field, empty)
    End Function
    
    Public Function DropDownListExt(id, selected_value, list, option_value_field, option_text_field, attribs)
        If IsNull(selected_value) then
            selected_value = ""
        Else
            selected_value = CStr(selected_value)
        End If
        
        dim item, options, opt_val, opt_txt 
        options = "<option value=''>"    ' first value is "non-selected" blank state
        select case typename(list)
            case "Recordset"
                do until list.EOF
                    If IsNull(list(option_value_field)) then
                        opt_val = ""
                    Else
                        opt_val = CStr(list(option_value_field))
                    End If
                        
                    opt_txt = list(option_text_field)
                    If Not IsNull(opt_val) And Not IsEmpty(opt_val) then
                        options = options & "<option value='" & Encode(opt_val) & "' " & Choice((CStr(opt_val) = CStr(selected_value)), "selected='selected'", "") & ">" & Encode(opt_txt) & "</option>" & vbCR
                    End If
                    
                    list.MoveNext
                loop
            case "Variant()"        'assumes KVArray
                dim i
                for i = 0 to ubound(list) step 2
                    KeyVal list, i, opt_val, opt_txt
                    options = options & "<option value='" & Encode(opt_val) & "' " & Choice((CStr(opt_val) = CStr(selected_value)), "selected='selected'", "") & ">" & Encode(opt_txt) & "</option>" & vbCR
                next
        end select
        DropDownListExt = "<select id='" & Encode(id) & "' name='" & Encode(id) & "' " & HtmlAttribs(attribs) & " >" & vbCR & options & "</select>" & vbCR
    End Function
    
    Public Function Checkbox(id, value)
        Checkbox = CheckboxExt(id, value, empty)
    End Function
    
    Public Function CheckboxExt(id, value, attribs)
        CheckBoxExt = "<input type='checkbox' id='" & Encode(id) & "' name='" & Encode(id) & "' " & Choice( (value = 1) or (value = true) or (LCase(value) = "true") or (LCase(value) = "on"), "checked='checked'", "") & " " & HtmlAttribs(attribs) & ">" & vbCR
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    'Button text IS NOT ENCODED! As with LinkTo, this allows use of Bootstrap icons and other arbitrary HTML in the 
    'button. If you need to HTMLEncode the text you MUST do it yourself!
    Public Function SubmitButton(text)
        SubmitButton = "<button type='submit' class='btn'>" & text & "</button>" & vbCR
    End Function
    
    Public Function Button(button_type, text, class_name)
        Button = "<button type='" & Encode(button_type) & "' class='btn " & Encode(class_name) & "'>" & text & "</button>" & vbCR
    End Function
    
    Public Function ButtonExt(button_type, text, attribs_array)
        ButtonExt = "<button type='" & Encode(button_type) & "' " & HtmlAttribs(attribs_array) & ">" & text & "</button>" & vbCR
    End Function
    
    
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Function Tag(Tag_name, attribs_array)
        Tag = "<" & Encode(tag_name) & " " & HtmlAttribs(attribs_array) & ">"
    End Function
    
    Public Function Tag_(Tag_name)
        Tag_ = "</" & Encode(tag_name) & ">"
    End Function
    
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Function HtmlAttribs(attribs)
        dim result : result = ""
        if not IsEmpty(attribs) then
            if IsArray(attribs) then
                dim idx
                for idx = lbound(attribs) to ubound(attribs) step 2
                    result = result & " " & HtmlAttrib(attribs, idx) & " "
                next
            else    ' assume string or string-like default value
                result = attribs
            end if
        end if
        HtmlAttribs = result
    End Function
    
    Public Function HtmlAttrib(attribs_array, key_idx)
        dim key, val
        KeyVal attribs_array, key_idx, key, val
        HtmlAttrib = Encode(key) & "='" & Encode(val) & "'"
    End Function
    
End Class


dim HTML_Helper__Singleton : set HTML_Helper__Singleton = Nothing
Function HTML()
    if HTML_Helper__Singleton Is Nothing then
        set HTML_Helper__Singleton = new HTML_Helper_Class
    End if
    set HTML = HTML_Helper__Singleton
End Function




%>
