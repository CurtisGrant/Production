Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web.UI
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Partial Class _CreateItems
   Inherits BasePage
    Public Shared con, con2 As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared rdr As SqlDataReader
    Private Shared sql As String
    Private Shared oTest As Object
    Public Shared dt, hdrtbl, workTbl As DataTable
    Public Shared yourIndex, ourIndex As Integer
    Public Shared yourItem, ourItem, conString, itemStatus As String

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        conString = "Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True"
        If IsPostBack Then
            FileUpload1.Attributes("onchange") = "UploadFile(this)"
        End If
    End Sub

    Protected Sub LoadXL_Click(sender As Object, e As EventArgs) Handles LoadXL.Click
        Dim path As String = Server.MapPath("~/UploadedFiles/")
        Dim pathName As String = ""
        Dim fileOK As Boolean = False
        If FileUpload1.HasFile Then
            Dim fileExtension As String
            fileExtension = System.IO.Path. _
                GetExtension(FileUpload1.FileName).ToLower()
            Dim allowedExtensions As String() = _
                {".csv", ".txt", ".xls", ".xlsx"}
            For i As Integer = 0 To allowedExtensions.Length - 1
                If fileExtension = allowedExtensions(i) Then
                    fileOK = True
                End If
            Next
            If fileOK Then
                Try
                    FileUpload1.PostedFile.SaveAs(path & _
                         FileUpload1.FileName)
                    'Label1.Text = "File uploaded!"
                    'Label1.Visible = True
                    pathName = path & FileUpload1.FileName.ToString
                Catch ex As Exception
                    'Label1.Text = "File could not be uploaded."
                End Try
            Else
                'Label1.Text = "Cannot accept files of this type."
            End If

            Call Load_Spreadsheet(pathName)

        End If
    End Sub
    

    Protected Sub Load_Spreadsheet(ByVal pathName As String)
        Try

            Dim path As String = CStr(pathName)
            dt = New DataTable
            hdrtbl = New DataTable
            Dim ws As DataTable
            Dim ds As DataSet
            ds = New DataSet
            Dim sheetname As String
            Dim con As OleDb.OleDbConnection
            con = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" & path & ";Excel 12.0 xml;HDR=Yes;")
            Dim adapter As System.Data.OleDb.OleDbDataAdapter
            con.Open()
            ws = con.GetSchema("Tables")
            sheetname = ws.Rows(0).Item("Table_Name")
            Dim connect As OleDb.OleDbConnection
            connect = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Excel 12.0 xml;HDR=YES;")
            adapter = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & sheetname & "]", con)
            adapter.Fill(dt)
            dt.Columns.Add("STATUS")
            ''
            ''  Get rid of empty rows
            ''
            For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                Dim row As DataRow = dt.Rows(i)
                If row.Item(0) Is Nothing Then
                    dt.Rows.Remove(row)
                ElseIf row.Item(0).ToString = "" Then
                    dt.Rows.Remove(row)
                End If
            Next

            gv1.DataSource = dt
            gv1.DataBind()
            gv1.AutoGenerateEditButton = True
            con.Close()
            con.Dispose()

            CheckStatus.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub CheckStatus_Click(sender As Object, e As EventArgs) Handles CheckStatus.Click
        Try
            If IsPostBack Then
                Dim con As New SqlConnection(conString)
                Dim itemExists As Boolean

                For Each row As DataRow In dt.Rows
                    oTest = row("ITEM")
                    If Not IsDBNull(oTest) Then
                        itemExists = False
                        con.Open()
                        sql = "SELECT DISTINCT Item FROM Item_Master WHERE Item = '" & CStr(oTest) & "'"
                        cmd = New SqlCommand(sql, con)
                        rdr = cmd.ExecuteReader
                        While rdr.Read
                            If Not IsDBNull(rdr("Item")) And Not IsNothing(rdr("Item")) Then itemExists = True
                        End While
                        con.Close()
                    End If
                    If itemExists Then row("STATUS") = "EXISTS" Else row("STATUS") = "NEW"
                Next
            End If
            gv1.DataSource = Nothing
            gv1.DataSource = dt
            gv1.DataBind()
        Catch ex As Exception

        End Try
    End Sub
End Class
