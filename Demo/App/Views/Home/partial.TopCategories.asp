        <div class="panel panel-default">
            <div class="panel-heading">
                Top Categories
            </div>
            <table class="table table-striped table-condensed">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th style="text-align: right">Sales</th>
                    </tr>
                </thead>
                <tbody>
                    <% dim category_it : set category_it = Model.TopTenCategories.Iterator %>
                    <% dim category %>
                    <% While category_it.HasNext %>
                        <% set category = category_it.GetNext() %>
                        <% 
                        dim category_row_class
                        If category.Sales > 100000 then 
                            category_row_class = "success"
                        ElseIf category.Sales > 60000 then
                            category_row_class = "warning"
                        Else
                            category_row_class = "danger"
                        End If
                        %>
                        <tr class="<%= category_row_class %>">
                            <td><%= H(category.Name) %></td>
                            <td style="text-align: right"><%= H(FormatNumber(category.Sales)) %></td>
                        </tr>
                    <% Wend %>
                </tbody>
            </table>
        </div>