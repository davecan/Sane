# The friendly Classic ASP framework
Sane is a full-featured MVC framework that brings sanity to Classic ASP.

# Example

Domain model:

```asp
<%
Class OrderModel_Class
    Public Validator
    Public OrderNumber, DateOrdered, CustomerName, LineItems
    
    Public Property Get SaleTotal
        SaleTotal = Enumerable(LineItems).Sum("item_.Subtotal")  ' whaaaa?
    End Property
    
    Public Sub Class_Initialize
        ValidatePattern   Me, OrderNumber, "^\d{9}[\d|X]$", "Order number format is incorrect."
        ValidateExists    Me, DateOrdered,    "DateOrdered cannot be blank."
        ValidateExists    Me, CustomerName,   "Customer name cannot be blank."
    End Sub
End Class

Class OrderLineItemModel_Class
    Public ProductName, Price, Quantity, Subtotal
End Class
%>
```

*(the repository takes a little bit more work but not much)*

Controller:

```asp
<%
Class OrdersController
    Public Model

    Public Sub Show
        MVC.RequirePost
        dim id : id = Request("Id")
        set Model = new Show_ViewModel_Class
        set Model.Order = OrderRepository.FindById(id)

        %> <!--#include file="../../Views/Orders/Index.asp"--> <% 
    End Sub
End Class

MVC.Dispatch
%>
```

View:

```asp
<h2>Order Summary</h2>

<div class="row">
    <div class="col-md-2">
        Order #<%= Model.Order.OrderNumber %>
    </div>
    <div class="col-md-10">
        Ordered on <%= Model.Order.DateOrdered %> 
                by <%= Model.Order.CustomerName %> 
                for <%= FormatCurrency(Model.Order.SaleTotal) %>
    </div>
</div>

<table class="table">
    <thead>
        <tr>
            <th>Product</th>
            <th>Price</th>
            <th>Qty</th>
            <th>Subtotal</th>
        </tr>
        <% dim it : set it = Model.Order.LineItems.Iterator %>
        <% dim item %>
        <% While it.HasNext %>
          <% set item = it.GetNext() %>
          <tr>
            <td><%= item.ProductName %></td>
            <td><%= item.Price %></td>
            <td><%= item.Quantity %></td>
            <td><%= item.Subtotal %></td>
          </tr>
        <% Wend %>
    </thead>
</table>
```

# But... why??

Mostly to prove it can be done. The vast majority of developers hate on VBScript and Classic ASP, mostly with good reason.
Many of the issues that plague Classic ASP stem from the constraints of the time in which it was developed, the mid-1990s.
Developers were unable to use what today are considered fundamental practices (widely using classes, etc) because the
language was not designed to execute in a manner we would call "fast" and using these practices would cause the application
to bog down and crash. Because of this the ASP community was forced to use ASP the same way PHP was used -- as an inline
page-based template processor, not a full application framework in its own right. Plus, let's be honest, Microsoft
marketed ASP at everyone regardless of skill level, and most of the tutorials found online were horrible and encouraged 
horribly bad practices.

Today we know better, and thanks to Moore's Law computing power has risen roughly 500-fold since the mid-90s, so we can
afford to do things that were unthinkable a few years back.

This framework was extracted from a real project that was built in this fashion. It worked quite well, and there should be
no reason (functionally speaking) that it shouldn't work as a viable application framework. That said, realistically if we
need to develop an application today we would use a modern framework such as .NET MVC or one of its competitors, so this is
really just here in case it is helpful to someone else. Plus it was fun to build. :)

# Features
* Useful Collections
  * LinkedList_Class adapted from the excellent Tolerable library
  * KVArray, a unique data structure I wish I'd built 15+ years ago that makes dynamic named parameters ridiculously easy
* Database Migrations
    `Migration.Do "create table Example_Table (id int not null, name varchar(100) not null)"`
    
TODO: List more features after getting some sleep.........
