<%
Option Explicit
%>
<!--#include file="../MVC/lib.Automapper.asp"-->
<!--#include file="ASPUnit/include/ASPUnitRunner.asp"-->
<!--#include file="TestCase_Automapper_Function.asp"-->
<!--#include file="TestCase_Automapper_AutoMap.asp"-->
<!--#include file="TestCase_Automapper_FlexMap.asp"-->
<!--#include file="TestCase_Automapper_DynMap.asp"-->
<%
'Used in each of the test case classes included into this file
'could remove and just use inline numbers in the classes and comment what they mean when called
'Ref: http://www.w3schools.com/ado/ado_datatypes.asp
dim adVarChar : adVarChar = CInt(200)
dim adInteger : adInteger = CInt(3)
dim adBoolean : adBoolean = CInt(11)
dim adDate    : adDate    = CInt(7)


dim Runner : set Runner = new UnitRunner
Runner.AddTestContainer new Automapper_Function_Tests
Runner.AddTestContainer new AutoMap_Tests
Runner.AddTestContainer new FlexMap_Tests
Runner.AddTestContainer new DynMap_Tests
Runner.Display
%>
