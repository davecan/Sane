# Sane, the dev-friendly Classic ASP MVC framework
Sane is a relatively full-featured MVC framework that brings sanity to Classic ASP. It has some similarities in style to both .NET
MVC and Rails, but doesn't exactly resemble either one. It is opinionated in that it assumes controllers will exist in a specific
folder location, but that location is somewhat configurable.

Major features distinguishing this framework include:

* Extensive use of domain repositories with domain model classes, and separate view model classes
* Automapper functionality, allowing mapping recordsets to domain objects (and objects to objects etc)
* Extensive use of linked lists and iterators
* Enumerable methods like `All/Any` boolean tests, `Min/Max/Sum`, `Map/Select` for projections, and `Where` for filters, with basic lambda-style expressions supported
* Database migrations with `Up` and `Down` steppable migrations -- version controlling database changes
* OWASP Top 10 mitigations
* Use of a key-value array data structure to easily pass key-value pairs between methods without creating a `Scripting.Dictionary` object each time
* Tons of HTML helpers, many of which use the `KVArray` and its helper methods to make building HTML easy
* Form state serialization baked in -- easily serialize/deserialize an entire form to/from the session in one line, for preserving state when returning forms to the user to fix validation errors
* Chocolate gravy

**All of this was written in VBScript. Really.**

Sane is licensed under the terms of GPLv3.

*Note:* This framework was extracted from a real-world internal workflow routing project, so it has a few rough edges.

## Code Examples: Products

This gives a quick overview of the code flow for one controller and the models and views it uses:

* [ProductsController](Demo/App/Controllers/Products/ProductsController.asp)
* [ProductsRepository with domain model](Demo/App/DomainModels/ProductRepository.asp)
* [Products view models](Demo/App/ViewModels/ProductsViewModels.asp)
* [Products Index view](Demo/App/Views/Products/Index.asp)
* [Product Edit view](Demo/App/Views/Products/Edit.asp)

## Aren't there other MVC-style frameworks?

"There are many like it, but this one is mine."

![alt text](http://www.gunaxin.com/wp-content/uploads/2010/02/FMJ-M14-560x420.jpg "")

## But... why??

Mostly because it was an interesting project that pushes the limits of Classic ASP. The vast majority of developers hate on VBScript and Classic ASP, mostly with good reason.
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

## Installation

**Dependency:** The demo was built against the Microsoft Northwind sample database. [Download the SQL Server BAK file here](https://northwinddatabase.codeplex.com/) and restore it into a SQL Server instance. [SQL Scripts and MDF file are also available](http://www.microsoft.com/en-us/download/details.aspx?id=23654).

1. Download the release to your directory of choice.
2. In Visual Studio, `File -> Open Web Site...` and select the `Demo` directory
3. Open file `App\DAL\lib.DAL.asp` and modify the connection string to point to your database.
4. Start the website (F5 or CTRL-F5) at `/index.asp`

The file `index.asp` will automatically redirect you to the Home controller, `/App/Controllers/HomeController.asp` and will load the
default action, `Index`.

## Features

A few of the features below have corresponding ASPUnit tests. Find them in the [Tests directory](Tests).

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

#### More about KVArrays

The `KVArray` data structure is fundamental to much of the framework and greatly simplifies coding. Fundamentally a `KVArray` is nothing more than a standard VBScript array that should always be consumed in groups of two. In other words, to build a `KVArray` we just need to build an array with element 0 being the first key and element 1 its value, element 2 the second key and element 3 its value, etc.

In essence you can imagine a `KVArray` as a way to use `System.Object` style calls as is done in .NET's `Html.ActionLink`.

For example:

```vb
dim kvarray : kvarray = Array(6)
'Element 1: Name = Bob
kvarray(0) = "Name"
kvarray(1) = "Bob"
'Element 2: Age = 35
kvarray(2) = "Age"
kvarray(3) = 35
'Element 3: FavoriteColor = Blue
kvarray(4) = "FavoriteColor"
kvarray(5) = "Blue"
```

But in reality you would never write it like that, instead you would use the inline `Array` constructor like this:

```vb
dim params : params = Array("Name", "Bob", "Age", 35, "FavoriteColor", "Blue")
```

Or for more readability:

```vb
dim params : params = Array( _
    "Name", "Bob",           _
    "Age", 35,               _
    "FavoriteColor", "Blue"  _
)
```

To iterate over this array step by 2 and use `KeyVal` to get the current key and value:

```vb
dim idx, the_key, the_val
For idx = 0 to UBound(kvarray) step 2
    KeyVal kvarray, idx, the_key, the_val
Next
```

On each iteration, `the_key` will contain the current key (e.g. "Name", "Age", or "FavoriteColor") and `the_val` will contain the key's corresponding value.

**But why not use a Dictionary?**

Dictionaries are great, but they are COM components and were at least historically expensive to instantiate and [because of threading should not be placed in the session](https://msdn.microsoft.com/en-us/library/ms972335.aspx#asptips_topic5). They are also cumbersome to work with for the use cases in this framework and there is no easy way to instantiate them inline with a dynamic number of parameters.

What we actually need is a fast, forward-only key-value data structure that allows us to iterate over the values and pluck out each key and value to build something like an HTML tag with arbitrary attributes or SQL `where` clause with arbitrary columns, not fast lookup of individual keys. So we need a hybrid of the array and dictionary that meets our specific needs and allows inline declaration of an arbitrary number of parameters. The `KVArray` allows us to very naturally write code like the `LinkToExt` example above, or manually building URLs using `Routes.UrlTo()`:

```asp
<%
<a href="<%= Routes.UrlTo("Users", "Edit", array("Id", user.Id)) %>">
    <i class="glyphicon glyphicon-user"></a>
</a>
%>
```

We can also create generic repository `Find` methods that can be used like this:

```vb
set expensive_products_starting_with_C = ProductRepository.Find( _
    array("name like ?", "C%", _
          "price > ?", expensive_price _
    ) _
)

set cheap_products_ending_with_Z = ProductRepository.Find( _
    array("name like ?", "%Z", _
          "price < ?", cheap_price _
    ) _
)
```

There are examples of this in the demo repositories, where `KVUnzip` is also used very effectively to help easily build the sql `where` clause. The below example is from the [`ProductRepository.Find()` method](Demo/App/DomainModels/ProductRepository.asp) which accepts a `KVArray` containing predicate key-value pairs and *unzips* it into two separate arrays that are used to build the query:

```vb
If Not IsEmpty(where_kvarray) then 
    sql = sql & " WHERE " 
    dim where_keys, where_values 
    KVUnzip where_kvarray, where_keys, where_values 

    dim i 
    For i = 0 to UBound(where_keys) 
        If i > 0 then sql = sql & " AND " 
        sql = sql & " " & where_keys(i) & " " 
    Next 
End If 

...

dim rs : set rs = DAL.Query(sql, where_values) 
set Find = ProductList(rs) 
```

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
validation rules, setting the `Errors` and `HasErrors` fields if errors are found. This is similar to the Observer pattern. The
reason we pass `Me` is because this allows us to have a conveniently-worded method for each validation that has strong semantic
meaning, e.g. `ValidateExists`. It takes a bit of code-jutsu but its worth it.

Adding new validations is easy, just add a new validation class and helper `Sub`. For example, to add a validation that requires
that a string start with the letter "A" you would create a `StartsWithLetterAValidation_Class` and helper method 
`Sub ValidateStartsWithA(instance, field_name, message)`, then call it via `ValidateStartsWithA Me, "MyField", "Field must start with A."`

### Domain Repositories

Domain models can be built by converting an ADO Recordset into a linked list of domain models via Automapper-style transforms. [Say Whaaat?](https://www.youtube.com/watch?v=bdxVcv6M6kE&t=2)

```vb
Class OrderRepository_Class
    Public Function GetAll()
        dim sql : sql = "select OrderNumber, DateOrdered, CustomerName from Orders"
        dim rs : set rs = DAL.Query(sql, empty)  'optional second parameter, can be scalar or array of binds
        
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
involves passing empty strings, or null values, or a similar approach. Using the built-in VBScript keyword `empty` is a semantically-meaningful way to handle optional parameters, making it clear that we specifically intended to ignore the optional parameter. In this case the `DAL.Query` method accepts two parameters, the SQL query and an optional second parameter containing bind values. The second parameter can be either a single value as in `DAL.Query("select a from b where a = ?", "foo")` or an array of binds e.g. 
`DAL.Query("select a from b where a = ? and c = ?", Array("foo", "bar")`. In the above example it is explicitly ignored since there are no bind variables in the SQL.

In this example the `DAL` variable is simply an instance
of the `Database_Class` from `lib.Data.asp`. In the original project the DAL was a custom class that acted as an entry point for
a set of lazy-loaded `Database_Class` instances, allowing data to be shared and moved between databases during the workflow.

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

### Views

Because it is `#include`d into the controller action, the view has full access to the controller's `Model` instance. Here it accesses
the `Order` property of the view model and iterates over the `LineItems` property (which would be a `LinkedList_Class` instance built
inside the repository) to build the view. Using view models you can create rich views that are not tied to a specific recordset
structure. See the `HomeController` in the demo for an example view model that contains four separate lists of domain objects to
build a dashboard summary view.

The `MVC.RequireModel` method provides the ability to strongly-type the view, mimicking the `@model` directive in .NET MVC.

```asp
<% MVC.RequireModel Model, "Show_ViewModel_Class" %>

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

### Database Class

Wraps connection details and access to the database. In addition to the examples already shown it can also handle:

* Execution of SQL without return values: `DAL.Execute "delete from Orders where OrderId = ?", id`
* Paged queries: `set rs = DAL.PagedQuery(sql, params, per_page, page_num)` 
    * See the demo [ProductsController.asp](Demo/App/Controllers/Products/ProductsController.asp) for a usage example
    * Note: This uses recordset paging, you must implement your own server-side paging if needed.
* Transactions: `DAL.BeginTransaction`, `DAL.CommitTransaction`, and `DAL.RollbackTransaction`

The class also automatically closes and destroys the wrapped connection via the `Class_Terminate` method which is called when the
class is ready for destruction.

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

The real-world project from which the framework was extracted contained approximately 3 dozen migrations, so it worked very well
for versioning the DB during development.

*Note: the migrations web interface is very basic and non-pretty*

**Dependency:** To use the migrations feature you must first create the table `meta_migrations` using the script [`! Create Migrations Table.sql`](Sane/Framework/Data/Migrations/! Create Migrations Table.sql).

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

And it handles nesting, as seen here when a call to `Dump Model` was placed in the `Show` action of `OrdersController.asp` in the Demo:

```vb
{OrderModel_Class: 
            Id : Long => «11074», 
            CustomerId : String => «SIMOB», 
            OrderDate : Date => «5/6/1998», 
            RequiredDate : Date => «6/3/1998», 
            ShippedDate : Null => «», 
            ShipName : String => «Simons bistro», 
            ShipAddress : String => «Vinbæltet 34», 
            ShipCity : String => «Kobenhavn», 
            ShipCountry : String => «Denmark», 
         LineItems : LinkedList_Class => 
                [List:
                        1 => 
                                {OrderLineItemModel_Class: 
                                            ProductId : Long => «16», 
                                            ProductName : String => «Pavlova», 
                                            UnitPrice : Currency => «17.45», 
                                            Quantity : Integer => «14», 
                                            Discount : Single => «0.05», 
                                            ExtendedPrice : Currency => «232.09»
                                }
                ]}
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

### OWASP Top 10 Mitigations

The framework provides tools to help mitigate the following three items items from the OWASP Top 10:

* Injection: The `Database_Class` support parameterized queries.
* Cross-Site Scripting: Several helper methods automatically encode output, and the `H()` method is provided for simple encoding of all other output.
* Cross-Site Request Forgery: The [`HtmlSecurity` helper](Framework/MVC/lib.HTMLSecurity.asp) provides per-form and per-site nonce checks to mitigate this threat.

The remaining seven vulnerabilities are mostly or wholly the responsibility of the developer and/or administrator.

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

## Thanks

* To Brian Lauber for his excellent [Tolerable library](https://github.com/briandamaged/tolerable) which pushes VBScript far beyond what anyone thought it was capable of. Several years ago I implemented lambdas in VBScript using dynamically built `Proc` and `Func` classes (intending to directly address criticism from [this guy](http://thom.org.uk/2006/02/23/adventures-in-vbscript-volume-i/)) but never published it. Brian pushed the boundaries way beyond that. The `LinkedList_Class` and its iterators are adapted from his work, with his extremely powerful lambda capabilities tailored out to avoid bogging down ASP too much. The framework also adopts some of his coding conventions, such as `_Class` suffixes and the use of lazy-loaded global-scope singleton functions.
* To Emmet M. who coincidentally was building [a similar MVC framework](http://www.codeproject.com/Articles/585883/Classic-ASP-and-MVC) the same time I was working on this one. Some of the MVC engine internals were inspired by his work.
