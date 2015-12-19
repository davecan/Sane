<%

'********************************************************************
' Name: ASPUnit.asp
'
' Purpose: Contains the main ASPUnit classes
'********************************************************************

Class TestSuite
	Private m_oTestCases

	Private Sub Class_Initialize()
		Set m_oTestCases = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		Set m_oTestCases = Nothing
	End Sub

	Public Sub AddTestCase(oTestContainer, sTestMethod)
		Dim oTestCase
		Set oTestCase = New TestCase
		Set oTestCase.TestContainer = oTestContainer
		oTestCase.TestMethod = sTestMethod

		m_oTestCases.Add oTestCase, oTestCase
	End Sub

	Public Sub AddAllTestCases(oTestContainer)
		Dim oTestCase, sTestMethod

		For Each sTestMethod In oTestContainer.TestCaseNames()
			AddTestCase oTestContainer, sTestMethod
		Next
	End Sub

	Public Function Count()
		Count = m_oTestCases.Count
	End Function

	Public Sub Run(oTestResult)
		Dim oTestCase
		For Each oTestCase In m_oTestCases.Items
			oTestCase.Run oTestResult
		Next
	End Sub
End Class

Class TestCase
	Private m_oTestContainer
	Private m_sTestMethod

	Public Property Get TestContainer()
		Set TestContainer = m_oTestContainer
	End Property

	Public Property Set TestContainer(oTestContainer)
		Set m_oTestContainer = oTestContainer
	End Property

	Public Property Get TestMethod()
		TestMethod = m_sTestMethod
	End Property

	Public Property Let TestMethod(sTestMethod)
		m_sTestMethod = sTestMethod
	End Property

	Public Sub Run(oTestResult)

                Dim iOldFailureCount
                Dim iOldErrorCount

                iOldFailureCount = oTestResult.Failures.Count
                iOldErrorCount = oTestResult.Errors.Count

		On Error Resume Next
		oTestResult.StartTest Me

		m_oTestContainer.SetUp()

		If (Err.Number <> 0) Then
			oTestResult.AddError Err
		Else
			' Response.Write("m_oTestContainer." & m_sTestMethod & "(oTestResult)<br>")
			Execute("m_oTestContainer." & m_sTestMethod & "(oTestResult)")

			If (Err.Number <> 0) Then
				' Response.Write(Err.Description & "<br>")
				oTestResult.AddError Err
			End	If
		End If
		Err.Clear()

		m_oTestContainer.TearDown()

                If (Err.Number <> 0) Then
			oTestResult.AddError Err
		End If

		'Log success if no failures or errors occurred
		If oTestResult.Failures.Count = iOldFailureCount And oTestResult.Errors.Count = iOldErrorCount Then
		        oTestResult.AddSuccess
		End If
		
                oTestResult.EndTest
                
		On Error Goto 0
	End Sub

End Class

Class TestResult

	Private m_dicErrors
	Private m_dicFailures
	Private m_dicSuccesses
	Private m_dicObservers
	Private m_iRunTests
	Private m_oCurrentTestCase

	Private Sub Class_Initialize
		Set m_dicErrors = Server.CreateObject("Scripting.Dictionary")
		Set m_dicFailures = Server.CreateObject("Scripting.Dictionary")
                Set m_dicSuccesses = Server.CreateObject("Scripting.Dictionary")
		Set m_dicObservers = Server.CreateObject("Scripting.Dictionary")
                m_iRunTests = 0		
	End Sub

	Private Sub Class_Terminate
		Set m_dicErrors = Nothing
		Set m_dicFailures = Nothing
                Set m_dicSuccesses = Nothing
		Set m_dicObservers = Nothing
		Set m_oCurrentTestCase = Nothing
	End Sub

	Public Property Get Errors()
		Set Errors = m_dicErrors
	End Property

	Public Property Get Failures()
		Set Failures = m_dicFailures
	End Property

        Public Property Get Successes()
                Set Successes = m_dicSuccesses
        End Property

	Public Property Get RunTests()
		RunTests = m_iRunTests
	End Property

	Public Sub StartTest(oTestCase)
		Set m_oCurrentTestCase = oTestCase

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnStartTest
		Next
	End Sub

	Public Sub EndTest()
		m_iRunTests = m_iRunTests + 1

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnEndTest
		Next
	End Sub

	Public Sub AddObserver(oObserver)
		m_dicObservers.Add oOserver, oObserver
	End Sub

	Public Function AddError(oError)
		Dim oTestError
		Set oTestError = New TestError
		oTestError.Initialize m_oCurrentTestCase, oError.Number, oError.Source, oError.Description
		m_dicErrors.Add oTestError, oTestError

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnError
		Next

		Set AddError = oTestError
	End Function

	Public Function AddFailure(sMessage)
	    Dim oTestError
		Set oTestError = New TestError
		oTestError.Initialize m_oCurrentTestCase, 0, " ", sMessage
		m_dicFailures.Add oTestError, oTestError

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnFailure
		Next

		Set AddFailure = oTestError
	End Function

        Public Function AddSuccess
		Dim oTestError
		Set oTestError = New TestError
		oTestError.Initialize m_oCurrentTestCase, 0, " ", "Test completed without failures"
		m_dicSuccesses.Add oTestError, oTestError

		Dim oObserver
		For Each oObserver In m_dicObservers.Items
			oObserver.OnSuccess
		Next
        End Function

	Public Sub Assert(bCondition, sMessage)
	        If Not bCondition Then
		        AddFailure sMessage
		End If
	End Sub

	Public Sub AssertEquals(vExpected, vActual, sMessage)
		If vExpected <> vActual Then
			AddFailure NotEqualsMessage(sMessage, vExpected, vActual)
		End	If
	End Sub

	'Build a message about a failed equality check
	Function NotEqualsMessage(sMessage, vExpected, vActual)
		'NotEqualsMessage = sMessage & " expected: " & CStr(vExpected) & " but was: " & CStr(vActual)
		NotEqualsMessage = sMessage & "<br>" &_
		                                     "<table><tr><th class='expected'>Expected</th><td class='expected'><span class='left-delim'>(" & typename(vExpected) & ") [</span>" & CStr(vExpected) & "<span class='right-delim'>]</span></td></tr><tr><th class='actual'>Actual</th><td class='actual'><span class='left-delim'>(" & typename(vActual) & ") [</span>" & CStr(vActual) & "<span class='right-delim'>]</span></td></tr></table>"
	End Function

	Public Sub AssertNotEquals(vExpected, vActual, sMessage)
		If vExpected = vActual Then
			AddFailure sMessage & " expected: " & CStr(vExpected) & " and actual: " & CStr(vActual) & " should not be equal."
		End	If
	End Sub

	Public Sub AssertExists(vVariable, sMessage)
		If IsObject(vVariable) Then
			If (vVariable Is Nothing) Then
				AddFailure sMessage & " - Variable of type " & TypeName(vVariable) & " is Nothing."
			End If
		ElseIf IsEmpty(vVariable) Then
			AddFailure sMessage & " - Variable " & TypeName(vVariable) & " is Empty (Uninitialized)."
		End If
	End Sub
	
	'---------------------------------------------------------------------------------------------------------------------
	' CUSTOM ASSERTIONS
	'---------------------------------------------------------------------------------------------------------------------
	Public Sub AssertFalse(bCondition, sMessage)
	    If bCondition then
	        AddFailure sMessage
	    End If
	End Sub
	
	Public Sub AssertEqual(vExpected, vActual, sMessage)
	    AssertEquals vExpected, vActual, sMessage
	End Sub
	
	Public Sub AssertNotEqual(vExpected, vActual, sMessage)
	    AssertNotEqual vExpected, vActual, sMessage
	End Sub
	
	Public Sub AssertNotExists(vVariable, sMessage)
	    If IsObject(vVariable) then
	        If (vVariable Is Not Nothing) then
	            AddFailure sMessage & " - Variable of type " & TypeName(vVariable) & " should be Nothing."
	        End If
	    ElseIf Not IsEmpty(vVariable) then
	        AddFailure sMessage & " - Variable " & TypeName(vVariable) & " should be Empty (Uninitialized)."
	    End If
	End Sub
	
	'Ensures (obj1 Is obj2) = true 
	Public Sub AssertSame(obj1, obj2, sMessage)
	    Assert (obj1 Is obj2), sMessage
	End Sub
	
	'Ensures (obj1 Is obj2) = false
	Public Sub AssertDifferent(obj1, obj2, sMessage)
	    Assert (not (obj1 Is obj2)), sMessage
	End Sub
	
	'Forces a test failure
	Public Sub Fail(sMessage)
	    AddFailure "Forced Failure: " & sMessage
	End Sub
	
	'Flags a test as needing implementation, otherwise an empty test will silently pass
	Public Sub NotImplemented
	    AddFailure "Test not implemented."
	End Sub
	
	Public Sub AssertType(sTypeName, vVariable, sMessage)
	    AssertEqual sTypeName, typename(vVariable), sMessage
	End Sub

End Class

Class TestError

	Private m_oTestCase
	Private m_lErrNumber
	Private m_sSource
	Private m_sDescription

	Public Sub Initialize(oTestCase, lErrNumber, sSource, sDescription)
		Set m_oTestCase = oTestCase
		m_lErrNumber = lErrNumber
		m_sSource = sSource
		m_sDescription = sDescription
	End Sub

	Public Property Get TestCase
		Set TestCase = m_oTestCase
	End Property

	Public Property Get ErrNumber
		ErrNumber = m_lErrNumber
	End Property

	Public Property Get Source
		Source = m_sSource
	End Property

	Public Property Get Description
		Description = m_sDescription
	End Property

End Class
%>