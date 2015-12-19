<%
'=======================================================================================================================
' Validation Classes
'=======================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------
' Exists Validation
'-----------------------------------------------------------------------------------------------------------------------
Class ExistsValidation_Class
    Private m_instance
    Private m_field_name
    Private m_message
    Private m_ok
    
    Public Function Initialize(instance, field_name, message)
        set m_instance  = instance
        m_field_name    = field_name
        m_message       = message
        m_ok            = true
        set Initialize  = Me
    End Function
    
    Public Sub Check
        If Len(eval("m_instance." & m_field_name)) = 0 then
            m_ok = false
        End If
    End Sub
    
    Public Property Get OK
        OK = m_ok
    End Property
    
    Public Property Get Message
        Message = m_message
    End Property
End Class

Sub ValidateExists(instance, field_name, message)
    if not IsObject(instance.Validator) then set instance.Validator = new Validator_Class
    instance.Validator.AddValidation new ExistsValidation_Class.Initialize(instance, field_name, message)
End Sub


'-----------------------------------------------------------------------------------------------------------------------
' Minimum Length Validation
'-----------------------------------------------------------------------------------------------------------------------
Class MinLengthValidation_Class
    Private m_instance
    Private m_field_name
    Private m_size
    Private m_message
    Private m_ok
    
    Public Function Initialize(instance, field_name, size, message)
        set m_instance  = instance
        m_field_name    = field_name
        m_size          = size
        m_message       = message
        m_ok            = true
        set Initialize  = Me
    End Function
    
    Public Sub Check
        If Len(eval("m_instance." & m_field_name)) < m_size then m_ok = false
    End Sub
    
    Public Property Get OK
        OK = m_ok
    End Property
    
    Public Property Get Message
        Message = m_message
    End Property
End Class

Sub ValidateMinLength(instance, field_name, size, message)
    if not IsObject(instance.Validator) then set instance.Validator = new Validator_Class
    instance.Validator.AddValidation new MinLengthValidation_Class.Initialize(instance, field_name, size, message)
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Max Length Validation
'-----------------------------------------------------------------------------------------------------------------------
Class MaxLengthValidation_Class
    Private m_instance
    Private m_field_name
    Private m_size
    Private m_message
    Private m_ok
    
    Public Function Initialize(instance, field_name, size, message)
        set m_instance  = instance
        m_field_name    = field_name
        m_size          = size
        m_message       = message
        m_ok            = true
        set Initialize  = Me
    End Function
    
    Public Sub Check
        If Len(eval("m_instance." & m_field_name)) > m_size then m_ok = false
    End Sub
    
    Public Property Get OK
        OK = m_ok
    End Property
    
    Public Property Get Message
        Message = m_message
    End Property
End Class

Sub ValidateMaxLength(instance, field_name, size, message)
    if not IsObject(instance.Validator) then set instance.Validator = new Validator_Class
    instance.Validator.AddValidation new MaxLengthValidation_Class.Initialize(instance, field_name, size, message)
End Sub

'-----------------------------------------------------------------------------------------------------------------------
' Numeric Validation
'-----------------------------------------------------------------------------------------------------------------------
Class NumericValidation_Class
    Private m_instance
    Private m_field_name
    Private m_message
    Private m_ok
    
    Public Function Initialize(instance, field_name, message)
        set m_instance  = instance
        m_field_name    = field_name
        m_message       = message
        m_ok            = true
        set Initialize  = Me
    End Function
    
    Public Sub Check
        If Not IsNumeric(eval("m_instance." & m_field_name)) then m_ok = false
    End Sub
    
    Public Property Get OK
        OK = m_ok
    End Property
    
    Public Property Get Message
        Message = m_message
    End Property
End Class

Sub ValidateNumeric(instance, field_name, message)
    if not IsObject(instance.Validator) then set instance.Validator = new Validator_Class
    instance.Validator.AddValidation new NumericValidation_Class.Initialize(instance, field_name, message)
End Sub


'-----------------------------------------------------------------------------------------------------------------------
' Regular Expression Pattern Validation
'-----------------------------------------------------------------------------------------------------------------------
Class PatternValidation_Class
    Private m_instance
    Private m_field_name
    Private m_pattern
    Private m_message
    Private m_ok
    
    Public Function Initialize(instance, field_name, pattern, message)
        set m_instance  = instance
        m_field_name    = field_name
        m_pattern       = pattern
        m_message       = message
        m_ok            = true
        set Initialize  = Me
    End Function
    
    Public Sub Check
        dim re : set re = new RegExp
        With re
            .Pattern    = m_pattern
            .Global     = true
            .IgnoreCase = true
        End With
        dim matches : set matches = re.Execute(eval("m_instance." & m_field_name))
        if matches.Count = 0 then
            m_ok = false
        end if
    End Sub
    
    Public Property Get OK
        OK = m_ok
    End Property
    
    Public Property Get Message
        Message = m_message
    End Property
End Class

Sub ValidatePattern(instance, field_name, pattern, message)
    if not IsObject(instance.Validator) then set instance.Validator = new Validator_Class
    instance.Validator.AddValidation new PatternValidation_Class.Initialize(instance, field_name, pattern, message)
End Sub



'-----------------------------------------------------------------------------------------------------------------------
' Validator Class
' This class is not intended to be used directly. Models should use the Validate* subs instead.
'-----------------------------------------------------------------------------------------------------------------------
Class Validator_Class
    Private m_validations
    Private m_errors
    
    Private Sub Class_Initialize
        m_validations = Array()
        redim m_validations(-1)
    
        m_errors = Array()
        redim m_errors(-1)
    End Sub
    
    Public Property Get Errors
        Errors = m_errors
    End Property
    
    Public Sub AddValidation(validation)
        dim n : n = ubound(m_validations) + 1
        redim preserve m_validations(n)
        set m_validations(n) = validation
    End Sub
    
    Public Sub Validate
        dim n : n = ubound(m_validations)
        dim i, V
        for i = 0 to n
            set V = m_validations(i)
            V.Check
            if not V.OK then
                AddError V.Message
            end if
        next
    End Sub
    
    Public Property Get HasErrors
        HasErrors = (ubound(m_errors) > -1)
    End Property
    
    'Public to allow other errors to be added by the controller for circumstances not accounted for by the validators
    Public Sub AddError(msg)
        redim preserve m_errors(ubound(m_errors) + 1)
        m_errors(ubound(m_errors)) = msg
    End Sub
End Class
%>
