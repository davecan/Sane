<%
Option Explicit
%>
<!-- #include file="../include/ASPUnitRunner.asp"-->
<!-- #include file="TCTest.asp"-->
<%
	Dim oRunner
	Set oRunner = New UnitRunner
	oRunner.AddTestContainer New TCTest
	
	oRunner.Display()
%>
