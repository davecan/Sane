<%
'=======================================================================================================================
' Bootstrap Helper - Provides some convenience methods for Bootstrap use
'=======================================================================================================================
Class Bootstrap_Helper_Class
    '---------------------------------------------------------------------------------------------------------------------
    ' Forms
    '---------------------------------------------------------------------------------------------------------------------
    'Creates .control-label for Bootstrap styling
    Public Function ControlLabel(name, for_name)
        ControlLabel = HTML.LabelExt(name, for_name, Array("class", "control-label"))
    End Function
    
    Public Function Control(label_text, control_name, controls_html)
        Control = "<div class='control-group'>" & ControlLabel(label_text, control_name) & "<div class='controls'>" & controls_html & "</div></div>"
    End Function
    
    '---------------------------------------------------------------------------------------------------------------------
    ' Modals
    '---------------------------------------------------------------------------------------------------------------------
    Public Function ModalLinkTo(link_text, controller_name, action_name, modal_name, html_class_attrib)
        ModalLinkTo = ModalLinkToExt(link_text, controller_name, action_name, params_array, modal_name, html_class_attrib, empty)
    End Function
    
    Public Function ModalLinkToExt(link_text, controller_name, action_name, params_array, modal_name, html_class_attrib, html_attribs_array)
        ModalLinkToExt = "<a href='" & Server.HTMLEncode(Routes.UrlTo(controller_name, action_name, params_array)) & "'" &_
                       "     class='" & html_class_attrib & " modal-link' data-modal='" & modal_name & "' " &_
                       HTML.HtmlAttribs(html_attribs_array) & ">" & link_text & "</a>"
    End Function
    
    'Generates the jQuery handler for the modal dialogs
    Public Function ModalHandlerScript()
        dim s
        s = " <script type='text/javascript'>                                                                      " & vbCR &_
                "     // when a .modal-link is clicked, display the href target in the popup modal                 " & vbCR &_
                "     $(function() {                                                                               " & vbCR &_
                "         $('a.modal-link').click(function(event) {                                                " & vbCR &_
                "             event.preventDefault();                                                              " & vbCR &_
                "             var modalName = $(this).attr('data-modal');                                          " & vbCR &_
                "             $('#' + modalName).removeData('modal').modal( { remote: $(this).attr('href') } );    " & vbCR &_
                "         });                                                                                                                                                                    " & vbCR &_
                "     });                                                                                                                                                                        " & vbCR &_
                " </script>                                                                                                                                                                "
        ModalHandlerScript = s
    End Function
End Class


dim Bootstrap_Helper_Class__Singleton
Function Bootstrap()
    If IsEmpty(Bootstrap_Helper_Class__Singleton) then set Bootstrap_Helper_Class__Singleton = new Bootstrap_Helper_Class
    set Bootstrap = Bootstrap_Helper_Class__Singleton
End Function



%>
