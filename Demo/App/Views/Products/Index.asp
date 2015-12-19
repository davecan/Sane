<h2>
    <%= H(Model.Title) %>
    
</h2>

<p>
    <%= H(Model.RecordCount) %> products found. Showing <%= H(Model.PageSize) %> records per page.
    
    <%= HTML.LinkToExt("<i class='glyphicon glyphicon-plus-sign'></i> New", "Products", "Create", empty, Array("class", "btn btn-xs btn-primary")) %>
    
    
</p>



<table class="table table-striped">
    <thead>
        <tr>
            <th></th>
            <th>Name</th>
            <th>Category</th>
            <th>Supplier</th>
            <th style="text-align: right">Unit Price</th>
            <th style="text-align: right">Units In Stock</th>
            <th style="text-align: right">Units On Order</th>
            <th style="text-align: right">Reorder Level</th>
            <th style="text-align: center">Discontinued?</th>
        </tr>
    </thead>
    <tbody>
        <% dim it : set it = Model.Products.Iterator %>
        <% dim product %>
        <% While it.HasNext %>
            <% set product = it.GetNext() %>
            <tr>
                <td>
                    <%= HTML.LinkToExt("<i class='glyphicon glyphicon-pencil'></i>", "Products", "Edit", Array("Id", product.Id), Array("class", "btn btn-primary")) %>
                </td>
                <td><%= H(product.Name) %></td>
                <td><%= H(product.CategoryName) %></td>
                <td><%= H(product.SupplierName) %></td>
                <td style="text-align: right"><%= H(FormatCurrency(product.UnitPrice)) %></td>
                <td style="text-align: right"><%= H(product.UnitsInStock) %></td>
                <td style="text-align: right"><%= H(product.UnitsOnOrder) %></td>
                <td style="text-align: right"><%= H(product.ReorderLevel) %></td>
                <td style="text-align: center">
                    <% If product.Discontinued then %>
                        <i class="glyphicon glyphicon-ok"></i>
                    <% End If %>
                </td>
            </tr>
        <% Wend %>
    </tbody>
</table>

<div>
    <% If Model.CurrentPageNumber <> 1 then %>
        <%= HTML.LinkToExt("<i class='glyphicon glyphicon-chevron-left'></i><i class='glyphicon glyphicon-chevron-left'></i>", MVC.ControllerName, MVC.ActionName, Array("page_num", 1), Array("class", "btn btn-default")) %>
        &nbsp;
        <%= HTML.LinkToExt("<i class='glyphicon glyphicon-chevron-left'></i>", MVC.ControllerName, MVC.ActionName, Array("page_num", Model.CurrentPageNumber - 1), Array("class", "btn btn-default")) %>
        &nbsp;
    <% Else %>
        <a class='btn btn-default disabled'><i class='glyphicon glyphicon-chevron-left'></i><i class='glyphicon glyphicon-chevron-left'></i></a>
        &nbsp;
        <a class='btn btn-default disabled'><i class='glyphicon glyphicon-chevron-left'></i></a>
        &nbsp; 
    <% End If %>

    <% If CInt(Model.CurrentPageNumber) < CInt(Model.PageCount) then %>
        <%= HTML.LinkToExt("<i class='glyphicon glyphicon-chevron-right'></i>", MVC.ControllerName, MVC.ActionName, Array("page_num", Model.CurrentPageNumber + 1), Array("class", "btn btn-default")) %>
        &nbsp;
        <%= HTML.LinkToExt("<i class='glyphicon glyphicon-chevron-right'></i><i class='glyphicon glyphicon-chevron-right'></i>", MVC.ControllerName, MVC.ActionName, Array("page_num", Model.PageCount), Array("class", "btn btn-default")) %>
        &nbsp;
    <% Else %>
        <a class='btn btn-default disabled'><i class='glyphicon glyphicon-chevron-right'></i><i class='glyphicon glyphicon-chevron-right'></i></a>
        &nbsp;
        <a class='btn btn-default disabled'><i class='glyphicon glyphicon-chevron-right'></i></a>
        &nbsp; 
    <% End If %>
</div>