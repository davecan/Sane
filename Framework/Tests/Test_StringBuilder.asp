<%
Option Explicit
%>
<!--#include file="../MVC/lib.Strings.asp"-->
<!--#include file="ASPUnit/include/ASPUnitRunner.asp"-->
<!--#include file="TestCase_StringBuilder.asp"-->
<%
dim Runner
set Runner = new UnitRunner
Runner.AddTestContainer new StringBuilder_Tests
Runner.Display
%>
