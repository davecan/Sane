<h2><%= H(Model.Title) %></h2>

<p class="alert alert-danger">Are you sure you want to delete this product?</p>

<%= HTML.FormTag("Products", "DeletePost", empty, Array("class", "form-horizontal")) %>
    <%= HTML.Hidden("nonce", HTMLSecurity.GetAntiCSRFToken("ProductsDeleteForm")) %>
    <%= HTML.Hidden("Id", Model.Product.Id) %>

    
    <div class="row">
        <div class="col-md-4">
            <div class="form-group">
                <label for="Name">Name</label>
                <br />
                <%= Model.Product.Name %>
            </div>
        </div>

        <div class="col-md-4">
            <div class="form-group">
                <label for="Category">Category</label>
                <br />
                <%= Model.Product.CategoryName %>
            </div>
        </div>

        <div class="col-md-4">
            <div class="form-group">
                <label for="Supplier">Supplier</label>
                <br />
                <%= Model.Product.SupplierName %>
            </div>
        </div>
    </div>

    <div class="row">
        <div class="col-md-1">
            <div class="form-group">
                <label for="UnitPrice">Unit Price</label>
                <br />
                <%= Model.Product.UnitPrice %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="UnitsInStock">Units In Stock</label>
                <br />
                <%= Model.Product.UnitsInStock %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="UnitsOnOrder">Units On Order</label>
                <br />
                <%= Model.Product.UnitsOnOrder %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="ReorderLevel">Reorder Level</label>
                <br />
                <%= Model.Product.ReorderLevel %>
            </div>
        </div>

        <div class="col-md-2">
            <div class="form-group">
                <label for="Discontinued">Discontinued?</label>
                <br />
                <%= Model.Product.Discontinued %>
            </div>
        </div>
    </div>


    <div class="col-md-10">
        <div class="form-group">
            <%= HTML.Button("submit", "<i class='glyphicon glyphicon-remove'></i> Confirm Delete", "btn-danger") %>
            &nbsp;&nbsp;
            <%= HTML.LinkToExt("Cancel", "Products", "Index", empty, Array("class", "btn btn-success")) %>
        </div>
    </div>
</form>