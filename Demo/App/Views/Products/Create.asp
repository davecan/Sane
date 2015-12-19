<h2>Create Product</h2>

<%= HTML.FormTag("Products", "CreatePost", empty, empty) %>
    <%= HTML.Hidden("nonce", HTMLSecurity.GetAntiCSRFToken("ProductsCreateForm")) %>
    
    <div class="row">
        <div class="col-md-4">
            <div class="form-group">
                <label for="Name">Name</label>
                <%= HTML.TextboxExt("Name", Model.Name, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-4">
            <div class="form-group">
                <label for="Category">Category</label>
                <%= HTML.DropDownListExt("CategoryId", empty, Model.Categories, empty, empty, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-4">
            <div class="form-group">
                <label for="Supplier">Supplier</label>
                <%= HTML.DropDownListExt("SupplierId", empty, Model.Suppliers, empty, empty, Array("class", "form-control")) %>
            </div>
        </div>
    </div>

    <div class="row">
        <div class="col-md-1">
            <div class="form-group">
                <label for="UnitPrice">Unit Price</label>
                <%= HTML.TextboxExt("UnitPrice", Model.UnitPrice, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="UnitsInStock">Units In Stock</label>
                <%= HTML.TextboxExt("UnitsInStock", Model.UnitsInStock, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="UnitsOnOrder">Units On Order</label>
                <%= HTML.TextboxExt("UnitsOnOrder", Model.UnitsOnOrder, Array("class", "form-control")) %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="ReorderLevel">Reorder Level</label>
                <%= HTML.TextboxExt("ReorderLevel", Model.ReorderLevel, Array("class", "form-control")) %>
            </div>
        </div>
    </div>

    <hr />

    <div class="form-group">
        <%= HTML.Button("submit", "<i class='glyphicon glyphicon-ok'></i> Create", "btn-primary") %>
        &nbsp;&nbsp;
        <%= HTML.LinkToExt("<i class='glyphicon glyphicon-remove'></i> Cancel", "Products", "Index", empty, Array("class", "btn btn-default")) %>
    </div>

</form>