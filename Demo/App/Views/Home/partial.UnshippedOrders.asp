        <div class="panel panel-default">
            <div class="panel-heading">
                Unshipped Orders
            </div>
            <table class="table table-striped table-condensed">
                <thead>
                    <tr>
                        <th>Required</th>
                        <th>Ordered</th>
                        <th>Customer</th>
                        <th>Fulfillment Days</th>
                    </tr>
                </thead>
                <tbody>
                    <% dim unshipped_it : set unshipped_it = Model.LastTenUnshippedOrders.Iterator %>
                    <% dim unshipped_order %>
                    <% While unshipped_it.HasNext %>
                        <% set unshipped_order = unshipped_it.GetNext() %>
                        <% 
                        dim unshipped_row_class
                        If unshipped_order.IsTooLateToFulfill then 
                            unshipped_row_class = "danger"
                        ElseIf unshipped_order.IsAlmostTooLateToFulfill then
                            unshipped_row_class = "warning"
                        Else
                            unshipped_row_class = ""
                        End If
                        %>
                        <tr class="<%= unshipped_row_class %>">
                            <td><%= H(unshipped_order.RequiredDate) %></td>
                            <td><%= H(unshipped_order.OrderDate) %></td>
                            <td><%= H(unshipped_order.CustomerId & ", " & unshipped_order.ShipCity & ", " & unshipped_order.ShipCountry) %></td>
                            <td><%= H(unshipped_order.DaysToFulfill) %></td>
                        </tr>
                    <% Wend %>
                </tbody>
            </table>
        </div>