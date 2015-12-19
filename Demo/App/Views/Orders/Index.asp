<h2>Recent Orders</h2>

<h3>Total Sales: <%= FormatCurrency(Model.SalesTotal) %></h3>

<table class="table table-striped">
    <thead>
        <tr>
            <th></th>
            <th>Ordered</th>
            <th>Required</th>
            <th>Shipped</th>
            <th>Customer</th>
            <th>Total</th>
        </tr>
    </thead>
    <tbody>
        <% dim it : set it = Model.Orders.Iterator %>
        <% dim order %>
        <% While it.HasNext %>
            <% set order = it.GetNext() %>
            <tr>
                <td>
                    <a class="btn btn-primary" href="<%= Routes.UrlTo("Orders", "Show", Array("Id", order.Id)) %>">
                        <i class="glyphicon glyphicon-search"></i>
                    </a>
                </td>
                <td><%= order.OrderDate %></td>
                <td><%= order.RequiredDate %></td>
                <td><%= order.ShippedDate %></td>
                <td>
                    <%= order.CustomerId %>
                    <br />
                    <%= order.ShipCity %>, <%= order.ShipCountry %>
                </td>
                <td>
                    <%= FormatCurrency(order.Subtotal) %>
                </td>
            </tr>
        <% Wend %>
    </tbody>
</table>