<% Option Explicit %>
<!--#include file="../../include_all.asp"-->
<!--#include file="../../ViewModels/ProductsViewModels.asp"-->

<%
Class ProductsController
    Public Model
    
    Public Sub Index
        'setup paging
        dim page_size : page_size = 10
        dim page_num  : page_num  = Choice(Len(Request.Querystring("page_num")) > 0, Request.Querystring("page_num"), 1)
        dim page_count, record_count

        set Model                   = new Index_ViewModel_Class
        Model.Title                 = "Products"
        set Model.Products          = ProductRepository.FindPaged(empty, "Name", page_size, page_num, page_count, record_count)
        Model.CurrentPageNumber     = page_num
        Model.PageSize              = page_size
        Model.PageCount             = page_count
        Model.RecordCount           = record_count
        %> <!--#include file="../../Views/Products/Index.asp"--> <%
    End Sub


    ' Edit
    '---------------------------------------------------------------------------------------------------------------------

    Public Sub Edit
        dim id : id = CInt(Request.QueryString("Id"))
        set Model = new Edit_ViewModel_Class
        set Model.Product = ProductRepository.FindById(id)
        
        Model.Categories = CategoryRepository.CategoriesKVArray()
        Model.Suppliers = SupplierRepository.SuppliersKVArray()

        Model.Title = "Edit Product"

        HTMLSecurity.SetAntiCSRFToken "ProductEditForm"
        %> <!--#include file="../../Views/Products/Edit.asp"--> <%
    End Sub

    Public Sub EditPost
        MVC.RequirePost
        HTMLSecurity.OnInvalidAntiCsrfTokenRedirectToActionExt "ProductEditForm", Request.Form("nonce"), "Edit", Array("Id", Request.Form("Id"))

        dim product_id  : product_id    = CInt(Request.Form("Id"))
        dim model       : set model     = ProductRepository.FindById(product_id)
        set model = Automapper.AutoMap(Request.Form, model)

        model.Validator.Validate

        If model.Validator.HasErrors then
            FormCache.SerializeForm "EditProduct", Request.Form
            Flash.Errors = model.Validator.Errors
            MVC.RedirectToActionExt "Edit", Array("Id", product_id)
        Else
            model.Discontinued = Choice("on" = Request.Form("Discontinued"), true, false)

            ProductRepository.Update model
            FormCache.ClearForm "EditProduct"
            Flash.Success = "Product updated."
            MVC.RedirectToAction "Index"
        End If
    End Sub


    ' Add
    '---------------------------------------------------------------------------------------------------------------------

    Public Sub Create
        dim form_params : set form_params = FormCache.DeserializeForm("NewProduct")
        If Not form_params Is Nothing then
            set Model = Automapper.AutoMap(form_params, new Create_ViewModel_Class)
        Else
            set Model = new Create_ViewModel_Class
        End If

        Model.Categories = CategoryRepository.CategoriesKVArray()
        Model.Suppliers = SupplierRepository.SuppliersKVArray()

        HTMLSecurity.SetAntiCSRFToken "ProductsCreateForm"

        %> <!--#include file="../../Views/Products/Create.asp"--> <%
    End Sub

    Public Sub CreatePost
        MVC.RequirePost
        HTMLSecurity.OnInvalidAntiCSRFTokenRedirectToAction "ProductsCreateForm", Request.Form("nonce"), "Create"

        dim new_product_model : set new_product_model = Automapper.AutoMap(Request.Form, new ProductModel_Class)
        
        new_product_model.Validator.Validate
        
        If new_product_model.Validator.HasErrors then
            FormCache.SerializeForm "NewProduct", Request.Form
            Flash.Errors = new_product_model.Validator.Errors
            MVC.RedirectToAction "Create"
        Else
            ProductRepository.AddNew new_product_model
            FormCache.ClearForm "NewProduct"
            Flash.Success = "Product added."
            MVC.RedirectToAction "Index"
        End If
    End Sub


    ' Delete
    '---------------------------------------------------------------------------------------------------------------------

    Public Sub Delete
        dim id : id = CInt(Request.QueryString("Id"))
        set Model = new Delete_ViewModel_Class
        set Model.Product = ProductRepository.FindById(id)
        Model.Title = "Delete Product"

        HTMLSecurity.SetAntiCSRFToken "ProductsDeleteForm"

        %> <!--#include file="../../Views/Products/Delete.asp"--> <%
    End Sub

    Public Sub DeletePost
        MVC.RequirePost
        HTMLSecurity.OnInvalidAntiCSRFTokenRedirectToAction "ProductsDeleteForm", Request.Form("nonce"), "Create"
        
        dim id : id = CInt(Request.Form("Id"))
        ProductRepository.Delete id

        Flash.Success = "Product deleted."
        MVC.RedirectToAction "Index"
    End Sub

End Class

MVC.Dispatch
        %>