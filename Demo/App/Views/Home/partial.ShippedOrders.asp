        <div class="panel panel-default">
            <div class="panel-heading">
                Shipped Orders
            </div>
            <table class="table table-striped table-condensed">
                <thead>
                    <tr>
                        <th>Shipped</th>
                        <th>Company</th>
                        <th style="text-align: right">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    <% dim shipped_it : set shipped_it = Model.LastTenShippedOrders.Iterator %>
                    <% dim shipped_order %>
                    <% While shipped_it.HasNext %>
                        <% set shipped_order = shipped_it.GetNext() %>
                        <tr>
                            <td><%= H(shipped_order.ShippedDate) %></td>
                            <td><%= H(shipped_order.CompanyName) %></td>
                            <td style="text-align: right"><%= H(FormatNumber(shipped_order.SaleAmount)) %></td>
                        </tr>
                    <% Wend %>
                </tbody>
            </table>
        </div>