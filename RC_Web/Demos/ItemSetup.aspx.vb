Imports System.Data
Imports System.Data.SqlClient

Partial Class Demos_ItemSetup
    Inherits System.Web.UI.Page
    Public Shared tbl As DataTable
    Public Shared row As DataRow
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            tbl = New DataTable
            tbl.Columns.Add("Item", GetType(System.String))
            tbl.Columns.Add("Description", GetType(System.String))
            tbl.Columns.Add("Vendor", GetType(System.String))
            tbl.Columns.Add("Category", GetType(System.String))
            tbl.Columns.Add("Buyer", GetType(System.String))
            tbl.Columns.Add("SubCategory", GetType(System.String))
            tbl.Columns.Add("UnitMeas", GetType(System.String))
            tbl.Columns.Add("BuyUnit", GetType(System.String))
            tbl.Columns.Add("SellUnit", GetType(System.String))
            lv1.DataSource = tbl.DefaultView
        End If
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        row = tbl.NewRow
        row("Item") = txtItem.ToString
        row("Description") = txtDesc.ToString
        row("Vendor") = txtVendor.ToString
        tbl.Rows.Add(row)

        lv1.DataSource = tbl
        lv1.DataBind()

    End Sub
End Class
