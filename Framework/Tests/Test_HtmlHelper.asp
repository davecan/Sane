<%
Option Explicit
%>
<!--#include file="../MVC/lib.Strings.asp"-->
<!--#include file="ASPUnit/include/ASPUnitRunner.asp"-->
<!--#include file="TestCase_HtmlHelperDropdownLists.asp"-->
<%
dim Runner
set Runner = new UnitRunner
Runner.AddTestContainer new HtmlHelper_Tests
Runner.Display
%>
