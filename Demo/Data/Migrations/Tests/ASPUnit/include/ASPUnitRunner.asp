<%
'********************************************************************
' Name: ASPUnitRunner.asp
'
' Purpose: Contains the UnitRunner class which is used to render the unit testing UI
'********************************************************************

'********************************************************************
' Include Files
'********************************************************************
%>
<!-- #include file="ASPUnit.asp"-->
<%

Const ALL_TESTCONTAINERS = "All Test Containers"
Const ALL_TESTCASES = "All Test Cases"

Const FRAME_PARAMETER = "UnitRunner"
Const FRAME_SELECTOR = "selector"
Const FRAME_RESULTS = "results"

Const STYLESHEET = "aspunit/include/ASPUnit.css"
Const SCRIPTFILE = "aspunit/include/ASPUnitRunner.js"

Class UnitRunner

	Private m_dicTestContainer

	Private Sub Class_Initialize()
		Set m_dicTestContainer = CreateObject("Scripting.Dictionary")
	End Sub

	Public Sub AddTestContainer(oTestContainer)
		m_dicTestContainer.Add TypeName(oTestContainer), oTestContainer
	End Sub

	Public Function Display()
		If (Request.QueryString(FRAME_PARAMETER) = FRAME_SELECTOR) Then
			DisplaySelector
		ElseIf (Request.QueryString(FRAME_PARAMETER) = FRAME_RESULTS) Then
			DisplayResults
		Else
			ShowFrameSet
		End if
	End Function

'********************************************************************
' Frameset
'********************************************************************
	Private Function ShowFrameSet()
%>
<HTML>
<HEAD>
<TITLE>ASPUnit Test Runner</TITLE>
</HEAD>
<FRAMESET ROWS="70, *" BORDER=0 FRAMEBORDER=0 FRAMESPACING=0>
	<FRAME NAME="<% = FRAME_SELECTOR %>" src="<% = GetSelectorFrameSrc %>" marginwidth="0" marginheight="0" scrolling="auto" border=0 frameborder=0 noresize>
	<FRAME NAME="<% = FRAME_RESULTS %>" src="<% = GetResultsFrameSrc %>" marginwidth="0" marginheight="0" scrolling="auto" border=0 frameborder=0 noresize>
</FRAMESET>
<%
	End Function

	Private Function GetSelectorFrameSrc()
		GetSelectorFrameSrc = Request.ServerVariables("SCRIPT_NAME") & "?" & FRAME_PARAMETER & "=" & FRAME_SELECTOR
	End Function

	Private Function GetResultsFrameSrc()
		GetResultsFrameSrc = Request.ServerVariables("SCRIPT_NAME") & "?" & FRAME_PARAMETER & "=" & FRAME_RESULTS
	End Function

'********************************************************************
' Selector Frame
'********************************************************************
	Private Function DisplaySelector()
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="<% = STYLESHEET %>" MEDIA="screen" TYPE="text/css">
<SCRIPT>
function ComboBoxUpdate(strSelectorFrameSrc, strSelectorFrameName)
{
	document.frmSelector.action = strSelectorFrameSrc;
	document.frmSelector.target = strSelectorFrameName;
	document.frmSelector.submit();
}
</SCRIPT>
</HEAD>
<BODY>
		<FORM NAME="frmSelector" ACTION="<% = GetResultsFrameSrc %>" TARGET="<% = FRAME_RESULTS %>" METHOD=POST>
			<TABLE class='Form'>
				<TR>
				    <TD>
						<INPUT TYPE="Submit" NAME="cmdRun" class="Submit" VALUE="Run Tests">
					</TD>
					<TD ALIGN="right">Test:</TD>
					<TD>
						<SELECT NAME="cboTestContainers" OnChange="ComboBoxUpdate('<% = GetSelectorFrameSrc %>', '<% = FRAME_SELECTOR %>');">
						<OPTION><% = ALL_TESTCONTAINERS %>
<%
							AddTestContainers
%>
						</SELECT>
					</TD>
					<TD ALIGN="right">Test Method:</TD>
					<TD>
						<SELECT NAME="cboTestCases">
						<OPTION><% = ALL_TESTCASES %>
<%
							AddTestMethods
%>
						</SELECT>
					<TD>
						<INPUT TYPE="checkbox" NAME="chkShowSuccess"> Show Passing Tests</INPUT>
					</TD>
					</TD>
				</TR>
			</TABLE>
		</FORM>
</BODY>
</HTML>
<%
	End Function

	Private Function AddTestContainers()
		Dim oTestContainer, sTestContainer
		For Each oTestContainer In m_dicTestContainer.Items()
			sTestContainer = TypeName(oTestContainer)
			If (sTestContainer = Request.Form("cboTestContainers")) Then
				Response.Write("<OPTION SELECTED>" & sTestContainer)
			Else
				Response.Write("<OPTION>" & sTestContainer)
			End If
		Next
	End Function

	Private Function AddTestMethods()
		Dim oTestContainer, sContainer, sTestMethod

		If (Request.Form("cboTestContainers") <> ALL_TESTCONTAINERS And Request.Form("cboTestContainers") <> "") Then
			sContainer = CStr(Request.Form("cboTestContainers"))
			Set oTestContainer = m_dicTestContainer.Item(sContainer)
			For Each sTestMethod In oTestContainer.TestCaseNames()
				Response.Write("<OPTION>" & sTestMethod)
			Next
		End If
	End Function

	Private Function TestName(oResult)
		If (oResult.TestCase Is Nothing) Then
			TestName = ""
		Else
			TestName = TypeName(oResult.TestCase.TestContainer) & "." & oResult.TestCase.TestMethod
		End If
	End Function

'********************************************************************
' Results Frame
'********************************************************************
	Private Function DisplayResults()
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="<% = STYLESHEET %>" MEDIA="screen" TYPE="text/css">
</HEAD>
<BODY>
<%
		Dim oTestResult, oTestSuite
		Set oTestResult = New TestResult

		' Create TestSuite
		Set oTestSuite = BuildTestSuite()

		' Run Tests
		oTestSuite.Run oTestResult

		' Display Results
		DisplayResultsTable oTestResult
%>
</BODY>
</HTML>
<%
	End Function

	Private Function BuildTestSuite()

		Dim oTestSuite, oTestContainer, sContainer
		Set oTestSuite = New TestSuite

		If (Request.Form("cmdRun") <> "") Then
			If (Request.Form("cboTestContainers") = ALL_TESTCONTAINERS) Then
				For Each oTestContainer In m_dicTestContainer.Items()
					If Not(oTestContainer Is Nothing) Then
						oTestSuite.AddAllTestCases oTestContainer
					End If
				Next
			Else
				sContainer = CStr(Request.Form("cboTestContainers"))
				Set oTestContainer = m_dicTestContainer.Item(sContainer)

				Dim sTestMethod
				sTestMethod = Request.Form("cboTestCases")

				If (sTestMethod = ALL_TESTCASES) Then
					oTestSuite.AddAllTestCases oTestContainer
				Else
					oTestSuite.AddTestCase oTestContainer, sTestMethod
				End If
			End If
		End If

		Set BuildTestSuite = oTestSuite
	End Function

	Private Function DisplayResultsTable(oTestResult)
%>
			<TABLE BORDER="1" class='Results'>
				<TR><TH WIDTH="10%" class="Type">Type</TH><TH WIDTH="20%" class="Test">Test</TH><TH WIDTH="70%" class="Desc">Description</TH></TR>
<%
		If Not(oTestResult Is Nothing) Then
			Dim oResult
			If (Request.Form("chkShowSuccess") <> "") Then
	                        For Each oResult in oTestResult.Successes
					Response.Write("	<TR CLASS=""success""><TD class='Type'>Success</TD><TD class='Test'>" & TestName(oResult) & "</TD><TD class='Desc'>" & oResult.Source & oResult.Description & "</TD></TR>")
	                        Next
	                End If

			For Each oResult In oTestResult.Errors
				Response.Write("	<TR CLASS=""error""><TD class='Type'>Error</TD><TD class='Test'>" & TestName(oResult) & "</TD><TD class='Desc'>" & oResult.Source & " (" & Trim(oResult.ErrNumber) & "): " & oResult.Description & "</TD></TR>")
			Next

			For Each oResult In oTestResult.Failures
				Response.Write("	<TR CLASS=""warning""><TD class='Type'>Failure</TD><TD class='Test'>" & TestName(oResult) & "</TD><TD class='Desc'>" & oResult.Description & "</TD></TR>")
			Next

			Response.Write "	<TR><TD ALIGN=""center"" COLSPAN=3 class='" & ResultRowClass(oTestResult) & "'>" & "Tests: " & oTestResult.RunTests & ", Errors: " & oTestResult.Errors.Count & ", Failures: " & oTestResult.Failures.Count & "</TD></TR>"
		End If
%>
			</TABLE>
<%
	End Function
	
	Private Function ResultRowClass(oTestResult)
	    if oTestResult.Errors.Count > 0 then
	        ResultRowClass = "error"
	    elseif oTestResult.Failures.Count > 0 then
	        ResultRowClass = "warning"
	    elseif oTestResult.Successes.Count > 0 then
	        ResultRowClass = "success"
	    end if
	End Function

	Public Sub OnStartTest()

	End Sub

	Public Sub OnEndTest()

	End Sub

	Public Sub OnError()

	End Sub

	Public Sub OnFailure()

	End Sub

        Public Sub OnSuccess()

        End Sub
End Class
%>

