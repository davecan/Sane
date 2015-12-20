# Sane, the friendly Classic ASP framework
Sane is a relatively full-featured MVC framework that brings sanity to Classic ASP. It has some similarities in style to both .NET
MVC and Rails, but doesn't exactly resemble either one. It is opinionated in that it assumes controllers will exist in a specific
folder location, but that location is somewhat configurable.

*Note:* This framework was extracted from a real-world internal workflow routing project, so it has a few rough edges.

**But aren't there other MVC-style frameworks?**

"There are many like it, but this one is mine."

![alt text](http://www.gunaxin.com/wp-content/uploads/2010/02/FMJ-M14-560x420.jpg "")

## Features

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

Migrations can be stepped up and down via web interface located at [`migrate.asp`](Framework/Data/Migrations/migrate.asp). `Migration.Do` executes SQL commands. Migrations are processed in the order loaded. Recommend following a structured naming scheme as shown above for easy ordering. There are a few special commands, such as
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

Domain models can be built by converting an ADO Recordset into a linked list of domain models via Automapper-style transforms. [Say Whaaat?](https://www.youtube.com/watch?v=bdxVcv6M6kE&t=2)

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

Because both source and target can be any object with instance methods this is a very useful way to manage model binding in CRUD 
methods, for example:

```vb
Public Sub CreatePost
    dim new_product_model : set new_product_model = Automapper.AutoMap(Request.Form, new ProductModel_Class)
    ... etc
End Sub
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

### Database Class

Wraps connection details and access to the database. In addition to the examples already shown it can also handle:

* Execution of SQL without return values: `DAL.Execute "delete from Orders where OrderId = ?", id`
* Paged queries: `set rs = DAL.PagedQuery(sql, params, per_page, page_num)` 
    * See the demo [ProductsController.asp](Demo/App/Controllers/Products/ProductsController.asp) for a usage example
    * Note: This uses recordset paging, you must implement your own server-side paging if needed.
* Transactions: `DAL.BeginTransaction`, `DAL.CommitTransaction`, and `DAL.RollbackTransaction`

The class also automatically closes and destroys the wrapped connection via the `Class_Terminate` method which is called when the
class is ready for destruction.

### Enumerable Builder

Provides chainable lambda-style calls on a list. From the unit tests:

```vb
Enumerable(list) _
    .Where("len(item_) > 5") _
    .Map("set V_ = new ChainedExample_Class : V_.Data = item_ : V_.Length = len(item_)") _
    .Max("item_.Length")
```
```V_``` is a special instance variable used by the `Map` method to represent the result of the "lambda" expression. `item_` is
another special instance variable that represents the current item being processed. So in this case, `Map` iterates over each
item in the list and executes the passed "lambda" expression. The result of `Map` is a new instance of `EnumerableHelper_Class`
containing a list of `ChainedExample_Class` instances built by the expression. This enumerable is then processed by `Max` to
return a single value, the maximum length.

### Debugging Helpers

Since the use of step-through debugging is not always possible in Classic ASP, these make debugging and tracing much easier.

`Dump` outputs objects in a meaningful way:

```vb
dim a : a = GetSomeArray()
Dump a
```

Output:

```
[Array:
        0 => «elt1»
        1 => «elt2»
        2 => «elt3»
]
```

It even handles custom classes, using the `Class_Get_Properties` field:

`Dump Product`

Output:

```
{ProductModel_Class: 
            Id : Long => «17», 
            Name : String => «Alice Mutton», 
            CategoryId : Long => «6», 
            Category : Empty => «», 
            CategoryName : String => «Meat/Poultry», 
            SupplierId : Long => «7», 
            Supplier : Empty => «», 
            SupplierName : String => «Pavlova, Ltd.», 
            UnitPrice : Currency => «250», 
            UnitsInStock : Integer => «23», 
            UnitsOnOrder : Integer => «0», 
            ReorderLevel : Integer => «0», 
            Discontinued : Boolean => «True»
}
```

`quit` immediately halts execution. `die "some message"` halts execution and outputs an "some message" to the screen. `trace "text"` 
and `comment "text"` both write HTML comments containing "text", useful for tracing behind the scenes without disrupting layout.

### Rails-style "Flash" Messages

`Flash.Success = "Product updated."`, `Flash.Errors = model.Validator.Errors`, etc.

### Form Serialization

If errors are encountered when creating a model we should be able to re-display the form with the user's content still filled in. To
simplify this the framework provides the `FormCache` object which serializes/deserializes form data via the session.

For example, in a `Create` action we can have:

```vb
Public Sub Create
    dim form_params : set form_params = FormCache.DeserializeForm("NewProduct")
    If Not form_params Is Nothing then
        set Model = Automapper.AutoMap(form_params, new Create_ViewModel_Class)
    Else
        set Model = new Create_ViewModel_Class
    End If
    
    %> <!--#include file="../../Views/Products/Create.asp"--> <%
End Sub
```

And in ```CreatePost```:

```vb
Public Sub CreatePost
    dim new_product_model : set new_product_model = Automapper.AutoMap(Request.Form, new ProductModel_Class)
    new_product_model.Validator.Validate
    
    If new_product_model.Validator.HasErrors then
        FormCache.SerializeForm "NewProduct", Request.Form
        Flash.Errors = new_product_model.Validator.Errors
        MVC.RedirectToAction "Create"
    Else
        ProductRepository.AddNew new_product_model
        FormCache.ClearForm "NewProduct"
        Flash.Success = "Product added."
        MVC.RedirectToAction "Index"
    End If
End Sub
```

### Model Validations

Validate models by calling the appropriate `Validate*` helper method from within the model's `Class_Initialize` constructor:

```vb
Private Sub Class_Initialize
    ValidateExists      Me, "Name", "Name must exist."
    ValidateMaxLength   Me, "Name", 10, "Name cannot be more than 10 characters long."
    ValidateMinLength   Me, "Name", 2,  "Name cannot be less than 2 characters long."
    ValidateNumeric     Me, "Quantity", "Quantity must be numeric."
    ValidatePattern     Me, "Email", "[\w-]+@([\w-]+\.)+[\w-]+", "E-mail format is invalid."
End Sub
```

Currently only `ValidateExists`, `ValidateMinLength`, `ValidateMaxLength`, `ValidateNumeric`, and `ValidatePattern` are included.
What these helper methods actually do is create a new instance of the corresponding validation class and attach it to the model's
`Validator` property. For example, when a model declares a validation using `ValidateExists Me, "Name", "Name must exist."` the
following is what actually happens behind the scenes:

```vb
Sub ValidateExists(instance, field_name, message)
    if not IsObject(instance.Validator) then set instance.Validator = new Validator_Class
    instance.Validator.AddValidation new ExistsValidation_Class.Initialize(instance, field_name, message)
End Sub
```

Here `Me` is the domain model instance. The `Validator_Class` is then used (via `YourModel.Validator`) to validate all registered
validation rules, setting the `Errors` and `HasErrors` fields if errors are found. This is similar to the Observer pattern.

Adding new validations is easy, just add a new validation class and helper `Sub`. For example, to add a validation that requires
that a string start with the letter "A" you would create a `StartsWithLetterAValidation_Class` and helper method 
`Sub ValidateStartsWithA(instance, field_name, message)`, then call it via `ValidateStartsWithA Me, "MyField", "Field must start with A."`

### Numerous Helpers

* `put` wraps `Response.Write` and varies its output based on the passed type, with special output for lists and arrays.
* `H(string)` HTMLEncodes a string
* Adapted from Tolerable:
    * `Assign(target, src)` abstracts away the need to use `set` for objects in cases where we deal with variables of arbitrary type
    * `Choice(condition, trueval, falseval)` is a more functiony `iif`
* HTML Helpers
    * `HTML.FormTag(controller_name, action_name, route_attribs, form_attribs)`
    * `HTML.TextBox(id, value)`
    * `HTML.TextArea(id, value, rows, cols)`
    * `HTML.DropDownList(id, selected_value, list, option_value_field, option_text_field)`
    * And more, many with `*Ext` variants
* Cross-Site Request Forgery protection (nonce can be set for whole site or per-form)
    * In `Edit` action: `HTMLSecurity.SetAntiCSRFToken "ProductEditForm"`
    * In view: `<%= HTML.Hidden("nonce", HTMLSecurity.GetAntiCSRFToken("ProductEditForm")) %>`
    * In `EditPost` action: `HTMLSecurity.OnInvalidAntiCsrfTokenRedirectToActionExt "ProductEditForm", Request.Form("nonce"), "Edit", Array("Id", Request.Form("Id"))`
* Exposed access to current controller/action names: `MVC.ControllerName`, `MVC.ActionName`
* Redirect via GET or POST: `MVC.RedirectTo(controller_name, action_name)` or `MVC.RedirectToActionPOST(action_name)` with `*Ext` variants
    * POST generates dynamic client-side form automatically submitted via JQuery

### "Generally Consistent API"

One idiom used throughout the framework is a workaround for VBScript not allowing method overloads. Generally speaking there are two 
cases, one where the method is full-featured with several parameters and one where the method signature is simplified. This is handled
by having the full-featured method append `Ext` to the end to denote it as an "extended" version of the simplified method. 

For example, this is from the `HTML_Helper_Class`:

```vb
Public Function LinkTo(link_text, controller_name, action_name)
    LinkTo = LinkToExt(link_text, controller_name, action_name, empty, empty)
End Function

Public Function LinkToExt(link_text, controller_name, action_name, params_array, attribs_array)
    LinkToExt = "<a href='" & Encode(Routes.UrlTo(controller_name, action_name, params_array)) & "'" &_ 
                HtmlAttribs(attribs_array) & ">" & link_text & "</a>" & vbCR
End Function
```

And this is from the `MVC_Dispatcher_Class`:

```vb
Public Sub RedirectTo(controller_name, action_name)
    RedirectToExt controller_name, action_name, empty
End Sub

' Redirects the browser to the specified action on the specified controller with the specified querystring parameters.
' params is a KVArray of querystring parameters.
Public Sub RedirectToExt(controller_name, action_name, params)
    Response.Redirect Routes.UrlTo(controller_name, action_name, params)
End Sub
```

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
