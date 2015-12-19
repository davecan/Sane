<%
Class TCTest

	Public Function TestCaseNames()
		TestCaseNames = Array("test", "test2", "test3")
	End Function

	Public Sub SetUp()
		'Response.Write("SetUp<br>")
	End Sub

	Public Sub TearDown()
		'Response.Write("TearDown<br>")
	End Sub

	Public Sub test(oTestResult)
		'Response.Write("test<br>")
		'Err.Raise 5, "hello", "error"
	End Sub

	Public Sub test2(oTestResult)
		'Response.Write("test2<br>")
		oTestResult.Assert False, "Assert False!"

		oTestResult.AssertEquals 4, 4, "4 = 4, Should not fail!"
		oTestResult.AssertEquals 4, 5, "4 != 5, Should fail!"
		oTestResult.AssertNotEquals 5, 5, "AssertNotEquals(5, = 5) should fail!"

        oTestResult.AssertExists new TestResult, "new TestResult Should not fail!"
        oTestResult.AssertExists Nothing, "Nothing: Should not exist!"
        oTestResult.AssertExists 4, "4 Should exist?!"
	End Sub

	Public Sub test3(oTestResult)
		oTestResult.Assert True, "Success"
	End Sub
End Class
%>