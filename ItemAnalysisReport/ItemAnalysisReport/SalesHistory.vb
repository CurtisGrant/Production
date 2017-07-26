Imports System.Data.SqlClient
Public Class SalesHistory
    Public Shared conString As String
    Public Shared item As String
    Public Shared sql As String
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared oTest As Object
    Private Sub SalesHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conString = Item_Analysis.conString
        item = Preview.selectedItem
        Call Get_Data()
    End Sub
    Private Sub Get_Data()
        Preview.lblProcessing.Visible = False
        Dim today, aYearAgo As Date
        today = CDate(Date.Today)
        aYearAgo = DateAdd(DateInterval.Day, -12, today)
        txtItem.Text = "Item No    " & item
        Dim tbl As New DataTable
        Dim row As DataRow
        tbl.Columns.Add("Location", GetType(System.String))
        tbl.Columns.Add("Week", GetType(System.String))
        tbl.Columns.Add("EndDate", GetType(System.DateTime))
        tbl.Columns.Add("Sales", GetType(System.Decimal))
        tbl.Columns.Add("TurnRate", GetType(System.Decimal))
        tbl.Columns.Add("Score", GetType(System.String))
        tbl.Columns.Add("GM%", GetType(System.String))
        tbl.Columns.Add("QtySold", GetType(System.Int16))
        tbl.Columns.Add("Received", GetType(System.Int16))
        tbl.Columns.Add("MaxOH", GetType(System.Int16))
        tbl.Columns.Add("EndOH", GetType(System.Int16))
        Dim con As New SqlConnection(conString)
        con.Open()
        Dim pct, sales, cost As Decimal
        sql = "SELECT d.Str_Id, c.Week_Num AS Wk, d.eDate, Sales_Retail, Turns, Score, Sales_Cost, Sold, Recvd, " & _
                "i.Max_OH, i.End_OH FROM Item_Detail d " & _
            "LEFT JOIN Item_Inv i ON i.Str_Id = d.Str_Id AND i.Item_No = d.Item_No AND i.eDate = d.eDate " & _
            "LEFT JOIN Calendar c ON c.eDate = d.eDate AND c.Week_Id > 0 " & _
            "WHERE d.Str_Id = '1' AND d.Item_No = '" & item & "' AND Sold <> 0 ORDER BY eDate DESC"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("Location") = rdr("Str_Id")
            row("Week") = rdr("Wk")
            row("EndDate") = rdr("eDate")
            oTest = rdr("Sales_Retail")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                sales = CDec(oTest)
            Else
                sales = 0
            End If
            row("Sales") = Format(sales, "###,###,###.00")
            oTest = rdr("Turns")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then row("TurnRate") = Format(rdr("Turns"), "####.0")
            End If

            row("Score") = rdr("Score")
            oTest = rdr("Sales_Cost")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                cost = CDec(oTest)
            Else : cost = 0
            End If
            If sales <> 0 Then
                pct = (sales - cost) / sales
            Else : pct = 0
            End If
            row("GM%") = Format(pct, "###.0%")
            row("QtySold") = rdr("Sold")
            row("Received") = rdr("Recvd")
            row("MaxOH") = rdr("Max_OH")
            row("EndOH") = rdr("End_OH")
            tbl.Rows.Add(row)
        End While
        dgv1.DataSource = tbl.DefaultView
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        For i As Integer = 0 To 10
            dgv1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        dgv1.Columns(8).Visible = False                                                 ' hide recvd
    End Sub

    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim str, hdr As String
        If rowIdx = 0 And columnIdx > 0 Then
            hdr = dgv1.Columns(columnIdx).HeaderText
            Select Case hdr
                Case "QtySold"
                    str = "Drill down qtys are entire hist"
            End Select
            With dgv1.Rows(0).Cells(columnIdx)
                .ToolTipText = str
            End With
            ' End If
        End If
    End Sub
End Class