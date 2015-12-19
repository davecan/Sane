
<h2>Order Details</h2>

<div class="panel panel-default">
    <div class="panel-heading">
        Order #<%= Model.Id %>
    </div>
    <div class="panel-body">
        <div class="row">
            <div class="col-md-2">
                <b>Customer</b><br />
                <%= Model.CustomerId %>
            </div>
            <div class="col-md-2">
                <b>Ordered On</b><br />
                <%= Model.OrderDate %>
            </div>
            <div class="col-md-2">
                <b>Required By</b><br />
                <%= Model.RequiredDate %>
            </div>
            <% If IsBlank(Model.ShippedDate) then %>
                <div class="col-md-2">
                    <span class='label label-danger'>Not yet shipped</span>
                </div>
            <% Else %>
                <div class="col-md-2">
                    <b>Shipped On</b><br />
                    <%= Model.ShippedDate %>
                </div>
            <% End If %>
            <div class="col-md-4 text-right">
                <span class="label label-success" style="font-size: 150%"><%= FormatCurrency(Model.Subtotal) %></span>
            </div>
        </div>
    </div>
</div>

<div class="panel panel-default">
    <div class="panel-heading">
        Line Items <span class="badge"><%= Model.LineItems.Count %></span>
    </div>
    <table class="table">
        <thead>
            <tr>
                <th></th>
                <th class='text-right'>Unit Price</th>
                <th class='text-right'>Quantity</th>
                <th class='text-right'>Discount</th>
                <th class='text-right'>Sale Price</th>
            </tr>
        </thead>
        <tbody>
            <% dim it : set it = Model.LineItems.Iterator %>
            <% dim item %>
            <% While it.HasNext %>
                <% set item = it.GetNext() %>
                <tr>
                    <td><%= item.ProductName %></td>
                    <td class='text-right'><%= item.UnitPrice %></td>
                    <td class='text-right'><%= item.Quantity %></td>
                    <td class='text-right'><%= item.Discount %></td>
                    <td class='text-right'><%= item.ExtendedPrice %></td>
                </tr>
            <% Wend %>
            <tr>
                <td colspan="5" class="text-right text-success">
                    <%= FormatCurrency(Model.Subtotal) %>
                </td>
            </tr>
        </tbody>
    </table>
</div>