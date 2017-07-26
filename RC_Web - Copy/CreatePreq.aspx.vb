Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web.UI
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Partial Class _CreatePreq:
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
        If Not IsPostBack Then
            Dim con As New SqlConnection("Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True")
            Dim con2 As New SqlConnection("Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True")
            ''Dim con As New SqlConnection("Server=LPS4SQL;Initial Catalog=PARGIF;Integrated Security=True")
            conString = "Server=CURTIS-MOBILE;Initial Catalog=TCM;Integrated Security=True"
            con.Open()
            sql = "SELECT Preq_No, Buyer, Vendor_Name FROM Purchase_Request ORDER BY Preq_No"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                ddPreq.Items.Add(rdr("Preq_No") & "/" & rdr("Buyer") & "/" & rdr("Vendor_Name"))
            End While
            con.Close()
            FileUpload1.Attributes("onchange") = "UploadFile(this)"

        End If
    End Sub

    Protected Sub MapFields_Click(sender As Object, e As EventArgs) Handles MapFields.Click
        ''If IsPostBack Then
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

            RequiredFields.Visible = True
            YourFields.Visible = True


            Call Load_Spreadsheet(pathName)

        End If
        ''End If
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

            MapFields.Visible = False
            ResetColumnHeadings.Visible = True
            CheckSkus.Visible = True
            Call Reset_Column_Headings()

            GridView1.DataSource = hdrtbl
            GridView1.DataBind()

            con.Close()
            con.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Sub YourFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles YourFields.SelectedIndexChanged
        yourItem = YourFields.SelectedItem.ToString
        yourIndex = YourFields.SelectedIndex
        GridView1.HeaderRow.Cells(yourIndex).Text = ourItem.ToString
        hdrtbl.Columns(yourIndex).ColumnName = ourItem

    End Sub

    Protected Sub RequiredFields_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RequiredFields.SelectedIndexChanged
        ourItem = RequiredFields.SelectedItem.ToString
        ourItem = ourItem.Replace("*", "")
        ourIndex = RequiredFields.SelectedIndex
    End Sub

    Protected Sub ResetColumnHeadings_Click(sender As Object, e As EventArgs) Handles ResetColumnHeadings.Click
        Call Reset_Column_Headings()
    End Sub
    Protected Sub Reset_Column_Headings()
        
        Try
            hdrtbl.Clear()
            hdrtbl.Columns.Clear()
            YourFields.Items.Clear()
            For Each column As DataColumn In dt.Columns
                oTest = column.ToString
                If Not IsNothing(oTest) And Not IsDBNull(oTest) Then

                    YourFields.Items.Add(CStr(oTest))
                    hdrtbl.Columns.Add(CStr(oTest))
                End If
            Next
            ''
            ''  Grab the first 10 rows and put then in hdrtbl
            ''
            Dim cnt As Integer = 0
            For Each row In dt.Rows
                hdrtbl.ImportRow(row)
                cnt += 1
                If cnt = 10 Then Exit For
            Next

            GridView1.DataSource = hdrtbl
            GridView1.DataBind()
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub CheckSkus_Click(sender As Object, e As EventArgs) Handles CheckSkus.Click
        Try
            Dim x As Integer = hdrtbl.Columns.Count
            If IsPostBack Then
                Dim requiredFieldsOK As Boolean = True
                Dim foundIt As Boolean = False
                Dim requiredField, testField As String

                For Each item In RequiredFields.Items
                    foundIt = False
                    requiredField = UCase(item.ToString)
                    If Left(requiredField, 1) = "*" Then
                        testField = requiredField.Replace("*", "")
                        foundIt = Find_Your_Column(testField)
                        If Not foundIt Then
                            requiredFieldsOK = False
                        End If
                    End If
                Next
                If requiredFieldsOK Then
                    Call Check_SKUs()
                Else
                    mapLabel.Visible = True
                End If

            End If
        Catch ex As Exception
        End Try
    End Sub
    Protected Sub CreateSpreadsheet_Click(sender As Object, e As EventArgs) Handles CreateSpreadsheet.Click
        Try
            
        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Check_SKUs()
        Dim sku, item, dim1, dim2, dim3 As String
        Dim newSKU As Boolean
        Dim wRow As DataRow
        workTbl = New DataTable
        Try
            For Each column As DataColumn In hdrtbl.Columns
                oTest = column.ColumnName.ToString
                workTbl.Columns.Add(oTest.ToString)
            Next
            '' workTbl.Columns.Add("STATUS", GetType(System.String))
            For Each row As DataRow In dt.Rows
                wRow = workTbl.NewRow
                oTest = row("ITEM")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then item = oTest.ToString
                ''oTest = row("COLOR")
                ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dim1 = oTest.ToString
                ''oTest = row("SIZE")
                ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dim2 = oTest.ToString
                ''oTest = row("SIZE2")
                ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dim3 = oTest.ToString

                Call Check_Item_Master(item, dim1, dim2, dim3, conString)

                wRow.ItemArray = row.ItemArray
                wRow("STATUS") = itemStatus
                workTbl.Rows.Add(wRow)
            Next
            If IsPostBack Then
                GridView1.DataSource = Nothing
                ''GridView1.DataBind()
                oTest = workTbl.Rows(0).Item("ITEM")
                oTest = workTbl.Rows(0).Item("STATUS")
                GridView1.DataSource = workTbl
                GridView1.DataBind()
                CreateSpreadsheet.Visible = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub Check_Item_Master(ByVal item As String, ByVal dim1 As Object, ByVal dim2 As Object, ByVal dim3 As Object,
                                    constring As String)
        itemStatus = "NEW ITEM"
        Dim val As String = Nothing
        con = New SqlConnection(constring)
        con.Open()
        ''sql = "SELECT Sku FROM Item_Master WHERE Item_No = '" & item & "'"
        ''If Not String.IsNullOrEmpty(dim1) Then sql &= " AND DIM1 = '" & dim1 & "'"
        ''If Not String.IsNullOrEmpty(dim2) Then sql &= " AND DIM2 = '" & dim2 & "'"
        ''If Not String.IsNullOrEmpty(dim3) Then sql &= " AND DIM3 = '" & dim3 & "'"
        sql = "select distinct item from item_master where item='" & item & "'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If Not IsDBNull(rdr("Item")) Then val = "EXISTS"
        End While
        con.Close()
        If val = "EXISTS" Then itemStatus = "EXISTS"
    End Sub

    Protected Function Find_Your_Column(ByVal ourColumn As String) As Boolean
        Dim yourColumn As String = Nothing
        For Each column As DataColumn In hdrtbl.Columns
            oTest = UCase(column.ColumnName)
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                yourColumn = UCase(column.ColumnName.ToString)
            End If
            If yourColumn = ourColumn Then Return True
        Next
        Return False
    End Function
End Class
