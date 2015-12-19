<% Option Explicit %>
<!--#include file="../../include_all.asp"-->
<!--#include file="../../DomainModels/OrderRepository.asp"-->
<!--#include file="../../ViewModels/OrdersViewModels.asp"-->


<%
Class OrdersController
    Public Model

    Public Sub Index
        set Model = new Index_ViewModel_Class
        set Model.Orders = OrderRepository.RecentOrdersSummary()
        Model.SalesTotal = Enumerable(Model.Orders).Sum("item_.Subtotal")

        %> <!--#include file="../../Views/Orders/Index.asp"--> <% 
    End Sub

    Public Sub Show
        dim id : id = Request("Id")
        set Model = OrderRepository.FindById(id)
        Model.Subtotal = Enumerable(Model.LineItems).Sum("item_.ExtendedPrice")

        %> <!--#include file="../../Views/Orders/Show.asp"--> <% 
    End Sub
End Class

MVC.Dispatch
%>