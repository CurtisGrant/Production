Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web.UI
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic

Partial Class Tools_CreateNewPreq
    Inherits BasePage
    Public Shared con, con2 As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared rdr As SqlDataReader
    Private Shared sql As String
    Private Shared oTest As Object
    Public Shared dt, hdrtbl, workTbl As DataTable
    Public Shared yourIndex, ourIndex As Integer
    Public Shared yourItem, ourItem, conString, itemStatus, thisBuyer, thisVendor, thisLocation As String
    Public Shared thisPreq As String
    Public Shared orderDateClicked, deliverDateClicked, cancelDateClicked As Boolean

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        conString = "Server=LP-CURTIS;Initial Catalog=TCM;Integrated Security=True"
        con = New SqlConnection(conString)
        con.Open()
        sql = "SELECT PREQ_NO FROM Purchase_Request_Header ORDER BY PREQ_NO"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddlPreq.Items.Add(rdr("PREQ_NO"))
        End While
        con.Close()


    End Sub

    

    Protected Function Fix_Hyphen(ByVal valu As String) As String
        Dim fixedValue As String
        fixedValue = valu.Replace("'", "''")
        Return fixedValue
    End Function

    Protected Sub ddlPreq_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPreq.SelectedIndexChanged
        thisPreq = ddlPreq.SelectedItem.ToString
    End Sub
End Class
