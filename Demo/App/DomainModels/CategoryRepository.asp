<%
'=======================================================================================================================
' Category Model
'=======================================================================================================================

Class CategoryModel_Class
    Public Validator
    Public Class_Get_Properties

    Public Id
    Public Name

    Private Sub Class_Initialize
        ValidateExists Me, "Name", "Name must exist."

        Class_Get_Properties = Array("Id", "Name")
    End Sub
End Class


'=======================================================================================================================
' Category Repository
'=======================================================================================================================

Class CategoryRepository_Class
    Public Function CategoriesKVArray()
        dim sql : sql = "SELECT CategoryId, CategoryName FROM Categories ORDER BY CategoryName"
        dim rs : set rs = DAL.Query(sql, empty)

        dim kvarray : kvarray = Array()
        dim i

        Do until rs.EOF
            i = UBound(kvarray) + 1  
            ReDim Preserve kvarray(i + 1)    ' add 2 slots for kvarray
            kvarray(i) = rs("CategoryId")
            kvarray(i+1) = rs("CategoryName")
            rs.MoveNext
        Loop

        CategoriesKVArray = kvarray
        Destroy rs
    End Function

    Private Function CategoryList(rs)
        dim list : set list = new LinkedList_Class
        dim model

        Do until rs.EOF
            set model = new CategoryModel_Class
            list.Push Automapper.AutoMap(rs, model)
            rs.MoveNext
        Loop

        set CategoryList = list
    End Function
End Class



dim CategoryRepository__Singleton
Function CategoryRepository()
    If IsEmpty(CategoryRepository__Singleton) then
        set CategoryRepository__Singleton = new CategoryRepository_Class
    End If
    set CategoryRepository = CategoryRepository__Singleton
End Function
%>