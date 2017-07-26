Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web.UI
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic

Partial Class Tools_ImportPOs:
    Inherits BasePage
    Public Shared con, con2 As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared rdr As SqlDataReader
    Private Shared sql As String
    Private Shared oTest As Object
    ''Public Shared ourFieldTbl As DataTable
    ''Public Shared thisFieldtbl As DataTable
    Public Shared dt As DataTable
    Public Shared yourIndex, ourIndex As Integer
    Public Shared yourItem, ourItem As String

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim con As New SqlConnection("Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True")
        Dim con2 As New SqlConnection("Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True")
        ''Dim con As New SqlConnection("Server=LPS4SQL;Initial Catalog=PARGIF;Integrated Security=True")
        con.Open()
        sql = "SELECT Preq_No, Buyer, Vendor_Name FROM Purchase_Request ORDER BY Preq_No"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            ddPreq.Items.Add(rdr("Preq_No") & "/" & rdr("Buyer") & "/" & rdr("Vendor_Name"))
        End While
        con.Close()
        ''OurFields.Items.Add("ITEM")
        ''OurFields.Items.Add("COLOR")
        ''OurFields.Items.Add("SIZE")
        ''OurFields.Items.Add("SIZE2")
        ''OurFields.Items.Add("DISCRIPTION")
        ''OurFields.Items.Add("COST")
        ''OurFields.Items.Add("RETAIL")
        ''OurFields.Items.Add("ORDER QTY")
        ''OurFields.Items.Add("DEPARTMENT")
        ''OurFields.Items.Add("CLASS")
        ''OurFields.Items.Add("STOCKING UNIT")
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsPostBack Then
            Dim path As String = Server.MapPath("~/UploadedFiles/")
            Dim pathName As String
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
                        Label1.Text = "File uploaded!"
                        pathName = path & FileUpload1.FileName.ToString
                    Catch ex As Exception
                        Label1.Text = "File could not be uploaded."
                    End Try
                Else
                    Label1.Text = "Cannot accept files of this type."
                End If
                ''Call Load_Headers(pathName)
                Call Load_Spreadsheet(pathName)
            End If
        End If
    End Sub
    Protected Sub Load_Spreadsheet(ByVal pathName)
        Try
            Dim path As String = CStr(pathName)
            dt = New DataTable
            ''thisFieldtbl = New DataTable
            ''origTable = New DataTable
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
            For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                Dim row As DataRow = dt.Rows(i)
                If row.Item(0) Is Nothing Then
                    dt.Rows.Remove(row)
                ElseIf row.Item(0).ToString = "" Then
                    dt.Rows.Remove(row)
                End If
            Next
            ''
            ''  Get rid of empty rows
            ''
            For Each column As DataColumn In dt.Columns
                oTest = column.ToString
                If Not IsNothing(oTest) And Not IsDBNull(oTest) Then
                    ''thisFieldtbl.Columns.Add(CStr(oTest), GetType(System.String))
                    ''YourFields.Items.Add(CStr(oTest))
                    YourFields.Items.Add(CStr(oTest))
                End If
            Next

            GridView1.DataSource = dt
            GridView1.DataBind()

            con.Close()
            con.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Load_Headers(ByVal pathName)
        Dim path As String = CStr(pathName)
        Dim dt As DataTable = New DataTable
        ''origTable = New DataTable
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
       
    End Sub

    Protected Sub gridview1_RowDataBound(sender As Object, e As WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim row As GridViewRow = GridView1.SelectedRow
                Dim col As Integer = GridView1.SelectedIndex
                Dim xx As String = "help"
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Sub gridview1_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        Dim rowIndex As Integer = Convert.ToInt16(e.CommandArgument)
        Dim nrow As GridViewRow = GridView1.Rows(rowIndex)
        Dim name As String = GridView1.Rows(rowIndex).Cells(1).Text


    End Sub
    Public Shared Function Check_Sku(ByVal sku As String) As String
        Dim result As String = ""
        Dim daSku As String = ""
        Dim con2 As New SqlConnection("Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True")
        con2.Open()
        sql = "SELECT Sku FROM Item_Master WHERE Sku = '" & sku & "'"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Sku")
            If Not IsDBNull(oTest) Then daSku = CStr(oTest)
        End While
        con2.Close()
        If Not IsDBNull(daSku) And Not IsNothing(daSku) Then
            If daSku <> "" Then result = "EXISTS" Else result = "NEW SKU"
        End If
        Return result

    End Function

    Protected Sub YourFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles YourFields.SelectedIndexChanged
        yourItem = YourFields.SelectedItem.ToString
        yourIndex = YourFields.SelectedIndex
        BoundFields.Items.Add(ourItem)
    End Sub

    Protected Sub RequiredFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RequiredFields.SelectedIndexChanged
        ourItem = RequiredFields.SelectedItem.ToString
        ourIndex = RequiredFields.SelectedIndex
    End Sub
End Class
