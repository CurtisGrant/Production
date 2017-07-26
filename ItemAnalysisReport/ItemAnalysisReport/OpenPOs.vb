Imports System.Data.SqlClient
Public Class OpenPOs
    Public Shared item As String
    Public Shared con As SqlConnection = Item_Analysis.con
    Public Shared sql As String
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared oTest As Object
    Private Sub OpenPOs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Call Get_Data()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub
    Private Sub Get_Data()
        Preview.lblProcessing.Visible = False
        item = Preview.selectedItem
        Dim selectedStores As String = Item_Analysis.selectedStores
        selectedStores = selectedStores.Replace(",", "','")
        Dim today, aYearAgo, ordDate As Date
        today = CDate(Date.Today)
        aYearAgo = DateAdd(DateInterval.Month, -12, today)
        txtItem.Text = "Item No    " & item
        Dim wks As Integer
        con.Open()
        sql = "SELECT Value FROM Controls WHERE ID = 'Extract_Weeks' AND Parameter = 'PurchaseOrders'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            wks = rdr("Value")
        End While
        con.Close()

        ordDate = DateAdd(DateInterval.Day, wks * -7, today)
        Dim tbl As New DataTable
        Dim row As DataRow
        tbl.Columns.Add("Location", GetType(System.String))
        tbl.Columns.Add("PO", GetType(System.String))
        tbl.Columns.Add("Order Date", GetType(System.DateTime))
        tbl.Columns.Add("Due Date", GetType(System.DateTime))
        tbl.Columns.Add("Cancel Date", GetType(System.DateTime))
        tbl.Columns.Add("Qty Ordered", GetType(System.Int16))
        tbl.Columns.Add("Qty Recv'd", GetType(System.Int16))
        tbl.Columns.Add("Qty Cancl'd", GetType(System.Int16))
        tbl.Columns.Add("Receipt Date", GetType(System.DateTime))
        tbl.Columns.Add("Week", GetType(System.Int16))

        'cpConn.Open()
        con.Open()
        'sql = "SELECT LOC_ID AS Location, l.PO_NO AS PO, CONVERT(Date,h.ORD_DAT) AS OrdDate, CONVERT(Date,l.DELIV_DAT) AS DueDate, " & _
        '    "CONVERT(Date,l.CANCEL_DAT) AS CanDate, l.ORD_QTY * l.ORD_QTY_NUMER AS ORD_QTY, " & _
        '    "(SELECT ISNULL(SUM(QTY_RECVD * ISNULL(QTY_RECVD_NUMER,1)),0) FROM PO_RECVR_HIST_LIN r " & _
        '    "WHERE r.PO_NO = l.PO_NO AND QTY_RECVD > 0 AND r.ITEM_NO = l.ITEM_NO) AS Qty_RECVD, " & _
        '    "(SELECT CONVERT(Date, MIN(r.RECVR_DAT)) FROM PO_RECVR_HIST_LIN r " & _
        '    "WHERE r.PO_NO = l.PO_NO AND r.QTY_RECVD > 0 AND r.ITEM_NO = l.ITEM_NO) AS RecvDate " & _
        '    "FROM PO_ORD_LIN l " & _
        '    "JOIN PO_ORD_HDR h ON h.PO_NO = l.PO_NO  WHERE l.ITEM_NO = '" & item & "' AND LOC_ID = '1' " & _
        '    "UNION " & _
        '    "SELECT rl.RTV_LOC_ID AS Location, rl.RTV_NO + 'RTV' AS PO, CONVERT(Date,rl.RTV_DAT) AS OrdDate, CONVERT(Date,rl.RTV_DAT) AS DueDate, " & _
        '    "CONVERT(Date,rl.RTV_DAT) AS CanDate, 0 AS ORD_QTY, rl.QTY_RETD * -1 AS Qty_RECVD, CONVERT(Date,rl.RTV_DAT) AS RecvDate FROM PO_RTV_HIST_LIN rl " & _
        '    "JOIN PO_RTV_HIST h ON h.RTV_NO = rl.RTV_NO " & _
        '    "WHERE h.RTV_LOC_ID = '1' AND rl.ITEM_NO = '" & item & "' " & _
        '    "ORDER BY PO DESC"
        sql = "SELECT d.Str_Id AS Location, d.PO_NO AS PO, h.Order_Date AS OrdDate, h.Due_Date AS DueDate, h.Cancel_Date AS CanDate, " & _
            "d.Qty_Ordered AS ORD_QTY, d.Qty_Recvd AS QTY_RECVD, d.Qty_Cancelled AS QTY_CAN, Last_Recvd_Date AS RecvDate, c.Week_Id FROM PO_Detail d " & _
            "LEFT JOIN Calendar c ON Last_recvd_Date BETWEEN sDate AND eDate AND c.Week_Id > 0 " & _
            "JOIN PO_Header h ON h.PO_NO = d.PO_NO WHERE d.Str_Id IN ('" & selectedStores & "') AND d.Item_No = '" & item & "' " & _
            "UNION " & _
            "SELECT LOCATION, TRANS_ID + 'RTV' AS PO, NULL AS OrdDate, NULL, NULL, NULL, QTY * -1 AS QTY_RECVD, NULL, " & _
            "TRANS_DATE AS RecvDate, Week_Id " & _
            "FROM Daily_Transaction_Log l JOIN Calendar c ON CONVERT(Date,l.TRANS_DATE) BETWEEN sDate AND eDate AND c.Week_Id > 0 " & _
            "WHERE LOCATION IN ('" & selectedStores & "') AND ITEM_NO = '" & item & "' AND [TYPE] = 'RTV' " & _
            "AND TRANS_DATE >= '" & ordDate & "' " & _
            "ORDER BY OrdDate DESC"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row(0) = rdr("Location")
            row(1) = rdr("PO")
            oTest = rdr("OrdDate")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                ordDate = CDate(oTest)
            End If
            row(2) = rdr("OrdDate")
            row(3) = rdr("DueDate")
            row(4) = rdr("CanDate")
            row(5) = rdr("ORD_QTY")
            row(6) = rdr("QTY_RECVD")
            row(7) = rdr("QTY_CAN")
            row(8) = rdr("RecvDate")
            row(9) = rdr("Week_Id")
            ''If ordDate > aYearAgo Then tbl.Rows.Add(row)
            tbl.Rows.Add(row)
        End While

        con.Close()
        dgv1.DataSource = tbl.DefaultView
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        For i As Integer = 0 To 9
            dgv1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        For r As Integer = 0 To tbl.Rows.Count - 1
            oTest = tbl.Rows(r).Item(1)
            If InStr(oTest, "RTV") Then
                For c As Integer = 0 To tbl.Columns.Count - 1
                    dgv1.Rows(r).Cells(c).Style.BackColor = Color.LightSalmon
                Next
            End If
        Next
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
                Case "Qty Ordered"
                    str = "Drill down qtys are entire hist"
                Case "Qty Recv'd"
                    str = "Drill down qtys are entire hist"
            End Select
            'If columnIdx = 5 Then
            With dgv1.Rows(0).Cells(columnIdx)
                .ToolTipText = str
            End With
            ' End If
        End If
    End Sub
End Class