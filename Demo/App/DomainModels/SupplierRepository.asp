<%
'=======================================================================================================================
' Supplier Model
'=======================================================================================================================
Class SupplierModel_Class
    Public Id
    Public CompanyName
End Class


'=======================================================================================================================
' Supplier Repository
'=======================================================================================================================

Class SupplierRepository_Class
    Public Function SuppliersKVArray()
        dim sql : sql = "SELECT SupplierId, CompanyName FROM Suppliers ORDER BY CompanyName"
        dim rs : set rs = DAL.Query(sql, empty)

        dim kvarray : kvarray = Array()
        dim i

        Do until rs.EOF
            i = UBound(kvarray) + 1
            ReDim Preserve kvarray(i + 1)    ' add 2 slots for kvarray
            kvarray(i) = rs("SupplierId")
            kvarray(i+1) = rs("CompanyName")
            rs.MoveNext
        Loop
        
        SuppliersKVArray = kvarray
        Destroy rs
    End Function
End Class



dim SupplierRepository__Singleton
Function SupplierRepository()
    If IsEmpty(SupplierRepository__Singleton) then
        set SupplierRepository__Singleton = new SupplierRepository_Class
    End If
    set SupplierRepository = SupplierRepository__Singleton
End Function
%>