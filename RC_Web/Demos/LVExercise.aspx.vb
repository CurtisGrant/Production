Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Partial Class Demos_LVExercise
   Inherits BasePage
    Public Shared conString As String = "Server=LP-CURTIS;Initial Catalog=TCM;Integrated Security=True"
    Public Shared con As SqlConnection
    Public Shared sql As String
    Protected Sub Item_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "Insert" AndAlso Page.IsValid Then
            SqlDataSource2.Insert()
        End If
    End Sub
    Protected Sub Inserting_New_Items(sender As Object, e As SqlDataSourceCommandEventArgs) _
        Handles SqlDataSource2.Inserting
        Dim newItem As TextBox = GridView1.FooterRow.FindControl("newItem")
        Dim newDescription As TextBox = GridView1.FooterRow.FindControl("newDescription")
        Dim newVendorItem As TextBox = GridView1.FooterRow.FindControl("newVendorItem")
        Dim newCost As TextBox = GridView1.FooterRow.FindControl("newCost")
        Dim newRetail As TextBox = GridView1.FooterRow.FindControl("newRetail")
        Dim newCategory As TextBox = GridView1.FooterRow.FindControl("newCategory")
        Dim newSubCategory As TextBox = GridView1.FooterRow.FindControl("newSubCategory")
        Dim newStockUnits As TextBox = GridView1.FooterRow.FindControl("newStockUnits")


    End Sub

    Protected Sub btnGO_Click(sender As Object, e As EventArgs) Handles btnGO.Click

    End Sub
End Class
