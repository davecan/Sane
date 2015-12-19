<%
'=======================================================================================================================
' Product Model
'=======================================================================================================================

Class ProductModel_Class
    Public Validator
    Public Class_Get_Properties

    Public Id
    Public Name
    Public UnitPrice
    Public UnitsInStock
    Public UnitsOnOrder
    Public ReorderLevel

    Public CategoryId
    Public Category
    Public CategoryName

    Public SupplierId
    Public Supplier
    Public SupplierName

    private m_discontinued

    Public Property Get Discontinued
        Discontinued = Choice(1 = m_discontinued, true, false)
    End Property

    Public Property Let Discontinued(v)
        'die v
        m_discontinued = Choice(true = v, 1, 0)
    End Property

    Private Sub Class_Initialize
        ValidateExists      Me, "Name",         "Product Name must exist."
        ValidateNumeric     Me, "UnitPrice",    "Unit Price must be numeric."
        ValidateNumeric     Me, "UnitsInStock", "Units In Stock must be numeric."
        ValidateNumeric     Me, "UnitsOnOrder", "Units On Order must be numeric."
        ValidateNumeric     Me, "ReorderLevel", "Reorder Level must be numeric."

        Class_Get_Properties = Array("Id", "Name", "CategoryId", "Category", "CategoryName", _
                                     "SupplierId", "Supplier", "SupplierName", "UnitPrice",  _
                                     "UnitsInStock", "UnitsOnOrder", "ReorderLevel", "Discontinued")
      End Sub
End Class


'=======================================================================================================================
' Product Repository
'=======================================================================================================================

Class ProductRepository_Class

    Public Function FindById(id)
        dim sql : sql = "SELECT Products.ProductId Id, Products.ProductName Name, Products.CategoryId, c.CategoryName, Products.SupplierId, s.CompanyName SupplierName, " &_
                        "       Products.UnitPrice, Products.UnitsInStock, Products.UnitsOnOrder, Products.ReorderLevel, Products.Discontinued " &_
                        "FROM Products " &_
                        "INNER JOIN Categories c ON Products.CategoryId = c.CategoryId " &_
                        "INNER JOIN Suppliers s ON Products.SupplierId = s.SupplierId "  &_
                        "WHERE Products.ProductId = ?"
        dim rs : set rs = DAL.Query(sql, id)

        If rs.EOF then
          Err.Raise 1, "ProductRepository_Class:FindById", ProductNotFoundException("id", id)
        Else
          set FindById = Automapper.AutoMap(rs, "ProductModel_Class")
        End If
    End Function

    
    ' List ProductModels
    '---------------------------------------------------------------------------------------------------------------------

    Public Function GetAll()
        set GetAll = Find(empty, "Name")
    End Function

    Public Function Find(where_kvarray, order_string_or_array)
        dim sql : sql = "SELECT Products.ProductId Id, Products.ProductName Name, Products.CategoryId, c.CategoryName, Products.SupplierId, s.CompanyName SupplierName, " &_
                        "       Products.UnitPrice, Products.UnitsInStock, Products.UnitsOnOrder, Products.ReorderLevel, Products.Discontinued " &_
                        "FROM Products " &_
                        "INNER JOIN Categories c ON Products.CategoryId = c.CategoryId " &_
                        "INNER JOIN Suppliers s ON Products.SupplierId = s.SupplierId "

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
        set Find = ProductList(rs)
        Destroy rs
    End Function

    Public Function FindPaged(where_kvarray, order_string_or_array, per_page, page_num, ByRef page_count, ByRef record_count)
        dim sql : sql = "SELECT Products.ProductId Id, Products.ProductName Name, Products.CategoryId, c.CategoryName, Products.SupplierId, s.CompanyName SupplierName, " &_
                        "       Products.UnitPrice, Products.UnitsInStock, Products.UnitsOnOrder, Products.ReorderLevel, Products.Discontinued " &_
                        "FROM Products " &_
                        "INNER JOIN Categories c ON Products.CategoryId = c.CategoryId " &_
                        "INNER JOIN Suppliers s ON Products.SupplierId = s.SupplierId "

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

        dim list : set list = new LinkedList_Class
        dim rs   : set rs   = DAL.PagedQuery(sql, where_values, per_page, page_num)

        If Not rs.EOF and Not (IsEmpty(per_page) and IsEmpty(page_num) and IsEmpty(page_count) and IsEmpty(record_count)) then
            rs.PageSize     = per_page
            rs.AbsolutePage = page_num
            page_count      = rs.PageCount
            record_count    = rs.RecordCount
        End If

        set FindPaged = PagedProductList(rs, per_page)
        Destroy rs
    End Function

    Private Function ProductList(rs)
        dim list : set list = new LinkedList_Class
        dim model

        Do until rs.EOF
            set model = new ProductModel_Class
            list.Push Automapper.AutoMap(rs, model)
            rs.MoveNext
        Loop

        set ProductList = list
    End Function

    Private Function PagedProductList(rs, per_page)
        dim list : set list = new LinkedList_Class

        dim x : x = 0
        Do While x < per_page and Not rs.EOF
            list.Push Automapper.AutoMap(rs, new ProductModel_Class)
            x = x + 1
            rs.MoveNext
        Loop
        set PagedProductList = list
    End Function

    Private Function ProductNotFoundException(ByVal field_name, ByVal field_val)
        ProductNotFoundException = "Product was not found with " & field_name & " of '" & field_val & "'."
    End Function


    ' Add / Update / Delete ProductModels
    '---------------------------------------------------------------------------------------------------------------------

    'sets the model.Id on successful insert
    Public Sub AddNew(ByRef model)
        dim sql : sql = "INSERT INTO Products " &_
                        "(ProductName, CategoryId, SupplierId, UnitPrice, UnitsInStock, UnitsOnOrder, ReorderLevel) " &_
                        "VALUES (?, ?, ?, ?, ?, ?, ?)"
        DAL.Execute sql, Array(model.Name, model.CategoryId, model.SupplierId, model.UnitPrice, model.UnitsInStock, model.UnitsOnOrder, model.ReorderLevel)

        sql = "SELECT TOP 1 ProductId FROM Products ORDER BY ProductId DESC"
        dim rs : set rs = DAL.Query(sql, empty)
        model.Id = rs("ProductId")
        Destroy rs
    End Sub

    Public Sub Update(model)
        dim sql : sql = "UPDATE Products SET ProductName = ?, CategoryId = ?, SupplierId = ?, " &_
                        "UnitPrice = ?, UnitsInStock = ?, UnitsOnOrder = ?, ReorderLevel = ?, Discontinued = ? " &_
                        "WHERE ProductId = ?"
        DAL.Execute sql, Array(model.Name, model.CategoryId, model.SupplierId, _
                               model.UnitPrice, model.UnitsInStock, model.UnitsOnOrder, _
                               model.ReorderLevel, model.Discontinued, model.Id)
    End Sub

    Public Sub Delete(id)
        dim sql : sql = "DELETE FROM Products WHERE ProductId = ?"
        DAL.Execute sql, id
    End Sub

End Class


dim ProductRepository__Singleton
Function ProductRepository()
    If IsEmpty(ProductRepository__Singleton) then
        set ProductRepository__Singleton = new ProductRepository_Class
    End If
    set ProductRepository = ProductRepository__Singleton
End Function
%>