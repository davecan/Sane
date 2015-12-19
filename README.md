# Sane, the friendly Classic ASP framework
Sane is a relatively full-featured MVC framework that brings sanity to Classic ASP.

## Example

### Database Migrations

```vb
Class Migration_01_Create_Orders_Table
    Public Migration

    Public Sub Up
        Migration.Do "create table Orders " &_
                     "(OrderNumber varchar(10) not null, DateOrdered datetime, CustomerName varchar(50))"
    End Sub
    
    Public Sub Down
        Migration.Do "drop table Orders"
    End Sub
End Class

Migrations.Add "Migration_01_Create_Orders_Table"
```

Migrations can be stepped up and down via web interface located at `migrate.asp`. `Migration.Do` executes SQL commands. Migrations are processed in the order loaded. Recommend following a structured naming scheme as shown above for easy ordering. There are a few special commands, such as
`Migration.Irreversible` that let you stop a down migration from proceeding, etc.

*Note: the migrations web interface is very basic and non-pretty*

### Domain Models

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

### Domain Repositories

Domain models can be built by converting an ADO Recordset into a linked list of domain models via Automapper-style functions:

```vb
Class OrderRepository_Class
    Public Function GetAll()
        dim sql : sql = "select OrderNumber, DateOrdered, CustomerName from Orders"
        dim rs : set rs = DAL.Query(sql, empty)  'optional second parameter, can be single scalar or array of binds
        
        dim list : set list = new LinkedList_Class
        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new OrderModel_Class)      ' keanuwhoa.jpg
            rs.MoveNext
        Loop
        
        set GetAll = list
        Destroy rs        ' no passing around recordsets, no open connections to deal with
    End Function
End Class

' Convenience wrapper lazy-loads the repository
dim OrderRepository__Singleton
Function OrderRepository()
    If IsEmpty(OrderRepository__Singleton) then
        set OrderRepository__Singleton = new OrderRepository_Class
    End If
    set OrderRepository = OrderRepository__Singleton
End Function
```

The use of the `empty` keyword is a common approach taken by this framework. A common complaint of VBScript is that it does not
allow optional parameters. While this is technically true it is easy to work around, yet virtually every example found online
involves passing empty strings, or null values, or a similar approach. Using the built-in VBScript keyword `empty` is a semantically-meaningful way to handle optional parameters, making it clear that we specifically intended to ignore the optional parameter. In this case the `DAL.Query` method accepts two parameters, the SQL query and a second parameter containing bind values. The second parameter can be either a single value as in `DAL.Query("select a from b where a = ?", "foo")` or an array of binds e.g. 
`DAL.Query("select a from b where a = ? and c = ?", Array("foo", "bar")`. In this example the `DAL` variable is simply an instance
of the `Database_Class` from `lib.Data.asp`.

The `Automapper` object is a VBScript class that attempts to map each field in the source object to a corresponding field in the
target object. The source object can be a recordset or a custom class. The function can map to a new or existing object. The
`Automapper` object contains three methods: `AutoMap` which attempts to map all properties; `FlexMap` which allows you to choose
a subset of properties to map, e.g. `Automapper.FlexMap(rs, new OrderModel_Class, array("DateOrdered", "CustomerName"))`
will only copy the two specified fields from the source recordset to the new model instance; and `DynMap` which allows you to
dynamically remap values, for a contrived example see:

```vb
Automapper.DynMap(rs, new OrderModel_Class, _
                  array("target.CustomerName = UCase(src.CustomerName)", _
                        "target.LikedOrder = src.CustomerWasHappy"))
```

### Controllers

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

The controller only knows about the domain, not the database. Joy!

Actions are functions without parameters (like Rails and unlike .NET MVC) -- params are pulled from the `Request` object as in traditional ASP. `MVC.RequirePost` allows you to restrict the action to only respond to `POST` requests -- errors out otherwise.

`MVC.Dispatch` is the magic sauce entry point to the framework. Since views are `#include`d into the controller, and since an
app can have `1..n` controllers each with `1..n` actions, having a central monolithic MVC dispatcher is not feasible beyond
a few simple controllers. This is because ASP loads and compiles the entire `#include`d file for every page view. Given that,
the framework instead delegates instantiation out to the controllers themselves, making them responsible for kicking off the
framework instead of making the framework responsible for loading and instantiating all controllers and loading and compiling all
views for every request. The framework does instantiate the controller dynamically based on naming conventions, but only one
controller is parsed and loaded per request instead of all controllers. By only loading one controller we can use the savings to
instead load up a lot of helpful libraries that make development much more developer friendly.

Because of this approach the "routing engine" is really just a class that knows how to build URLs to controller files, so
`Routes.UrlTo("Orders", "Show", Array("Id", order.Id))` generates the URL `/App/Controllers/OrdersController.asp?_A=Show&Id=123`
(for order.Id = 123). URLs point to controllers and provide the name of the action to be executed via the `_A` parameter. Action
parameters are passed via a `KVArray` data structure, which is simply an array of key/value pairs that is used extensively
throughout the framework. For example, here are two `KVArray`s used in one of the many HTML helpers:

```asp
<%= HTML.LinkToExt("View Orders", _
                   "Orders", _
                   "List", _
                   array("param1", "value1", "param2", "value2"), _
                   array("class", "btn btn-primary", "id", "orders-button")) %>
```

Behind the scenes this method builds an anchor that routes to the correct controller/action combo, passes the specified parameters
via the querystring, and has the specified HTML `class` and `id` attributes. `KVArray`s are easily handled thanks to a few helper
methods such as `KeyVal` and `KVUnzip`.

### Views

Because it is `#include`d into the controller action, the view has full access to the controller's `Model` instance. Here it accesses
the `Order` property of the view model and iterates over the `LineItems` property (which would be a `LinkedList_Class` instance built
inside the repository) to build the view. Using view models you can create rich views that are not tied to a specific recordset
structure.

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

## Features
### Useful Collections
* LinkedList_Class adapted from the excellent Tolerable library to work in an ASP environment

* `Enumerable(list)` builder that provides chainable lambda-style calls on a list. From the unit tests:

```vb
Enumerable(list) _
    .Where("len(item_) > 5") _
    .Map("set V_ = new ChainedExample_Class : V_.Data = item_ : V_.Length = len(item_)") _
    .Max("item_.Length")
```

* KVArray, a unique data structure that makes dynamic forward-only named parameters ridiculously easy, and can be passed around through the session / etc unlike Dictionary COM objects which are apartment-threaded

... TODO: List more features ...

## But... why??

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

## Thanks

* To Brian Lauber for his excellent [Tolerable library](https://github.com/briandamaged/tolerable) which pushes VBScript far beyond what anyone thought it was capable of. Several years ago I implemented lambdas in VBScript using dynamically built `Proc` and `Func` classes (intending to directly address criticism from [this guy](http://thom.org.uk/2006/02/23/adventures-in-vbscript-volume-i/)) but never published it. Brian pushed the boundaries way beyond that. The `LinkedList_Class` and its iterators are adapted from his work, with his extremely powerful lambda capabilities tailored out to avoid bogging down ASP too much. The framework also adopts some of his coding conventions, such as `_Class` suffixes and the use of lazy-loaded global-scope singleton functions.
* To Emmet M. who coincidentally was building [a similar MVC framework](http://www.codeproject.com/Articles/585883/Classic-ASP-and-MVC) the same time I was working on this one. Some of the MVC engine internals were inspired by and adapted from his work.
