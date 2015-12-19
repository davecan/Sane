<%
Option Explicit
%>
<!--#include file="../MVC/lib.all.asp"-->
<!--#include file="ASPUnit/include/ASPUnitRunner.asp"-->
<!--#include file="TestCase_Performance.asp"-->
<%
dim Runner
set Runner = new UnitRunner
Runner.AddTestContainer new Performance_Tests
Runner.Display
%>
