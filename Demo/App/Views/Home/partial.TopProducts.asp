        <div class="panel panel-default">
            <div class="panel-heading">
                Top Products
            </div>
            <table class="table table-striped table-condensed">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th style="text-align: right">Sales</th>
                    </tr>
                </thead>
                <tbody>
                    <% dim product_it : set product_it = Model.TopTenProducts.Iterator %>
                    <% dim product %>
                    <% While product_it.HasNext %>
                        <% set product = product_it.GetNext() %>
                        <% 
                        dim product_row_class
                        If product.Sales > 30000 then 
                            product_row_class = "success"
                        ElseIf product.Sales > 20000 then
                            product_row_class = "warning"
                        Else
                            product_row_class = "danger"
                        End If
                        %>
                        <tr class="<%= product_row_class %>">
                            <td><%= H(product.Name) %></td>
                            <td style="text-align: right"><%= H(FormatNumber(product.Sales)) %></td>
                        </tr>
                    <% Wend %>
                </tbody>
            </table>
        </div>