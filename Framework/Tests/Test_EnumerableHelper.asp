<%
Option Explicit
%>
<!--#include file="../MVC/lib.all.asp"-->
<!--#include file="ASPUnit/include/ASPUnitRunner.asp"-->
<!--#include file="TestCase_EnumerableHelper.asp"-->
<%
dim Runner
set Runner = new UnitRunner
Runner.AddTestContainer new EnumerableHelper_Tests
Runner.Display
%>
