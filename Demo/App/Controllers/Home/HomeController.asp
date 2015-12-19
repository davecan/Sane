<% Option Explicit %>
<!--#include file="../../include_all.asp"-->
<!--#include file="../../ViewModels/DashboardViewModels.asp"-->


<%
Class HomeController
    Public Model

    Public Sub Index
        set Model = new DashboardViewModel_Class
        set Model.TopTenCategories = ReportRepository.TopTenCategories()
        set Model.TopTenProducts = ReportRepository.TopTenProducts()
        set Model.LastTenShippedOrders = ReportRepository.LastTenShippedOrders()
        set Model.LastTenUnshippedOrders = ReportRepository.LastTenUnshippedOrders()

        %> <!--#include file="../../Views/Home/Index.asp"--> <% 
    End Sub
End Class

MVC.Dispatch
%>