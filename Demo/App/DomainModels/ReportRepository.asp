<%
'=======================================================================================================================
' Report Models
'=======================================================================================================================
Class ReportModel_SalesByName_Class
    Public Name
    Public Sales
End Class

Class ReportModel_ShippedOrder_Class
    Public SaleAmount
    Public CompanyName
    Public ShippedDate
End Class

Class ReportModel_UnshippedOrder_Class
    Public CustomerId
    Public ShipCity
    Public ShipCountry
    Public OrderDate
    Public RequiredDate
    
    Public Property Get DaysToFulfill
        DaysToFulfill = DateDiff("d", OrderDate, RequiredDate)
    End Property

    Public Property Get IsTooLateToFulfill
        IsTooLateToFulfill = DateDiff("d", RequiredDate, #5/15/1998#) > 1
    End Property

    Public Property Get IsAlmostTooLateToFulfill
        IsAlmostTooLateToFulfill = DateDiff("d", RequiredDate, #5/25/1998#) > 1 And Not IsTooLateToFulfill
    End Property
End Class


'=======================================================================================================================
' Report Repository
'=======================================================================================================================

Class ReportRepository_Class
    Public Function TopTenCategories()
        dim sql : sql = "SELECT TOP 10 CategoryName Name, CategorySales Sales FROM [Category Sales for 1997] ORDER BY CategorySales DESC"
        dim rs : set rs = DAL.Query(sql, empty)

        dim list : set list = new LinkedList_Class
        dim model

        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new ReportModel_SalesByName_Class)
            rs.MoveNext
        Loop

        set TopTenCategories = list
        Destroy rs
    End Function

    Public Function TopTenProducts()
        dim sql : sql = "SELECT TOP 10 ProductName Name, ProductSales Sales FROM [Product Sales for 1997] ORDER BY ProductSales DESC"
        dim rs : set rs = DAL.Query(sql, empty)

        dim list : set list = new LinkedList_Class
        dim model

        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new ReportModel_SalesByName_Class)
            rs.MoveNext
        Loop

        set TopTenProducts = list
        Destroy rs
    End Function

    Public Function LastTenShippedOrders()
        dim sql : sql = "SELECT TOP 10 SaleAmount, CompanyName, ShippedDate FROM [Sales Totals by Amount] ORDER BY ShippedDate Desc"
        dim rs : set rs = DAL.Query(sql, empty)

        dim list : set list = new LinkedList_Class
        dim model

        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new ReportModel_ShippedOrder_Class)
            rs.MoveNext
        Loop

        set LastTenShippedOrders = list
        Destroy rs
    End Function

    Public Function LastTenUnshippedOrders()
        dim sql : sql = "SELECT Top 10 CustomerId, ShipCity, ShipCountry, OrderDate, RequiredDate FROM Orders WHERE ShippedDate IS NULL ORDER BY RequiredDate"
        dim rs : set rs = DAL.Query(sql, empty)

        dim list : set list = new LinkedList_Class
        dim model

        Do until rs.EOF
            list.Push Automapper.AutoMap(rs, new ReportModel_UnshippedOrder_Class)
            rs.MoveNext
        Loop

        set LastTenUnshippedOrders = list
        Destroy rs
    End Function
End Class



dim ReportRepository__Singleton
Function ReportRepository()
    If IsEmpty(ReportRepository__Singleton) then
        set ReportRepository__Singleton = new ReportRepository_Class
    End If
    set ReportRepository = ReportRepository__Singleton
End Function
%>