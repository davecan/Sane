<%
'=======================================================================================================================
' Order Model
'=======================================================================================================================

Class OrderModel_Class
    Public Validator
    Public Class_Get_Properties

    Public Id
    Public CustomerId
    Public OrderDate
    Public RequiredDate
    Public ShippedDate
    Public ShipName
    Public ShipAddress
    Public ShipCity
    Public ShipCountry
    Public Subtotal

    Public LineItems

    Private Sub Class_Initialize
        Class_Get_Properties = Array("Id", "CustomerId", "OrderDate", "RequiredDate", "ShippedDate", _
                                     "ShipName", "ShipAddress", "ShipCity", "ShipCountry")
    End Sub
End Class


Class OrderLineItemModel_Class
    Public Validator
    Public Class_Get_Properties

    Public ProductId
    Public ProductName
    Public UnitPrice
    Public Quantity
    Public Discount
    Public ExtendedPrice

    Private Sub Class_Initialize
        Class_Get_Properties = Array("ProductId", "ProductName", "UnitPrice", "Quantity", "Discount", "ExtendedPrice")
    End Sub
End Class


'=======================================================================================================================
' Order Repository
'=======================================================================================================================

Class OrderRepository_Class
    Public Function FindById(id)
        dim sql : sql = "SELECT OrderId Id, CustomerId, OrderDate, RequiredDate, ShippedDate, " &_
                        "       ShipName, ShipAddress, ShipCity, ShipCountry " &_
                        "FROM Orders WHERE OrderId = ? "
        dim rs : set rs = DAL.Query(sql, id)

        dim order : set order = Automapper.AutoMap(rs, new OrderModel_Class)
        set order.LineItems = LineItemsForOrderId(id)

        set FindById = order
        Destroy rs
    End Function

    Private Function LineItemsForOrderId(id)
        dim sql : sql = "SELECT ProductId, ProductName, UnitPrice, Quantity, Discount, ExtendedPrice " &_
                        "FROM [Order Details Extended] WHERE OrderId = ? ORDER BY ProductName"
        dim rs : set rs = DAL.Query(sql, id)

        dim list : set list = new LinkedList_Class

        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new OrderLineItemModel_Class)
            rs.MoveNext
        Loop

        set LineItemsForOrderId = list
        Destroy rs
    End Function

    Public Function Find(where_kvarray, order_string_or_array)
        dim sql : sql = "SELECT TOP 50 o.OrderId Id, o.CustomerId, o.OrderDate, o.RequiredDate, s.Subtotal, " &_
                        "       o.ShippedDate, o.ShipName, o.ShipAddress, o.ShipCity, o.ShipCountry " &_
                        "FROM Orders o " &_
                        "INNER JOIN [Order Subtotals] s ON o.OrderId = s.OrderId "

        If Not IsEmpty(where_kvarray) then
            sql = sql & " WHERE "
            dim where_keys, where_values
            KVUnzip where_kvarray, where_keys, where_values

            dim i
            For i = 0 to UBound(where_keys)
                If i > 0 then sql = sql & " AND "
                sql = sql & " " & where_keys(i) & " "
            Next
        End If

        If Not IsEmpty(order_string_or_array) then
            sql = sql & "ORDER BY "
            If IsArray(order_string_or_array) then
                dim order_array : order_array = order_string_or_array
                For i = 0 to UBound(order_array)
                    If i > 0 then sql = sql & ", "
                    sql = sql & " " & order_array(i)
                Next
            Else
                sql = sql & order_string_or_array & " "
            End If
        End If

        dim rs : set rs = DAL.Query(sql, where_values)
        dim list : set list = new LinkedList_Class

        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new OrderModel_Class)
            rs.MoveNext
        Loop

        set Find = list
        Destroy rs
    End Function

    Public Function RecentOrdersSummary()
        set RecentOrdersSummary = Find(empty, "OrderDate DESC")
    End Function
End Class



dim OrderRepository__Singleton
Function OrderRepository()
    If IsEmpty(OrderRepository__Singleton) then
        set OrderRepository__Singleton = new OrderRepository_Class
    End If
    set OrderRepository = OrderRepository__Singleton
End Function
%>