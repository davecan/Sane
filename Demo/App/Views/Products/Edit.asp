<h2><%= H(Model.Title) %></h2>

<%= HTML.FormTag("Products", "EditPost", empty, empty) %>
    <%= HTML.Hidden("nonce", HTMLSecurity.GetAntiCSRFToken("ProductEditForm")) %>
    <%= HTML.Hidden("Id", Model.Product.Id) %>

    <div class="row">
        <div class="col-md-4">
            <div class="form-group">
                <label for="Name">Name</label>
                <%= HTML.TextboxExt("Name", Model.Product.Name, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-4">
            <div class="form-group">
                <label for="Category">Category</label>
                <%= HTML.DropDownListExt("CategoryId", Model.Product.CategoryId, Model.Categories, empty, empty, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-4">
            <div class="form-group">
                <label for="Supplier">Supplier</label>
                <%= HTML.DropDownListExt("SupplierId", Model.Product.SupplierId, Model.Suppliers, empty, empty, Array("class", "form-control")) %>
            </div>
        </div>
    </div>

    <div class="row">
        <div class="col-md-1">
            <div class="form-group">
                <label for="UnitPrice">Unit Price</label>
                <%= HTML.TextboxExt("UnitPrice", Model.Product.UnitPrice, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="UnitsInStock">Units In Stock</label>
                <%= HTML.TextboxExt("UnitsInStock", Model.Product.UnitsInStock, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="UnitsOnOrder">Units On Order</label>
                <%= HTML.TextboxExt("UnitsOnOrder", Model.Product.UnitsOnOrder, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="ReorderLevel">Reorder Level</label>
                <%= HTML.TextboxExt("ReorderLevel", Model.Product.ReorderLevel, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="Discontinued">
                    Discontinued?
                    <%= HTML.CheckboxExt("Discontinued", Model.Product.Discontinued, Array("class", "form-control")) %>
                </label>
            </div>
        </div>
    </div>

    <hr />

    <div class="form-group">
        <%= HTML.Button("submit", "<i class='glyphicon glyphicon-ok'></i> Save", "btn-primary") %>
        &nbsp;&nbsp;
        <%= HTML.LinkToExt("<i class='glyphicon glyphicon-remove'></i> Delete", "Products", "Delete", Array("id", Model.Product.Id), Array("class", "btn btn-danger")) %>
        &nbsp;&nbsp;
        <%= HTML.LinkToExt("Cancel", "Products", "Index", empty, Array("class", "btn btn-default")) %>
    </div>

</form>