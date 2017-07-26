Imports System.Xml
Imports System.Data.SqlClient
Public Class BuyerPct
    Public Shared con, con2, con3 As SqlConnection
    Public Shared rdr, rdr2, rdr3 As SqlDataReader
    Public Shared conString, sql As String
    Public Shared cmd As SqlCommand
    Public Shared thisYear, thisPeriod, totalPeriods, todaysYear, todaysPeriod As Integer
    Public Shared thisStore, thisDept, thisTable, thisBuyer As String
    Public Shared oTest As Object
    Public Shared formLoaded, somethingChanged As Boolean
    Public Shared tbl As DataTable
    Public Shared row As DataRow
    Public Shared tNew, tRows As Integer
    Public Shared today As Date = CDate(Date.Today)

    Private Sub BuyerPct_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            formLoaded = False
            Me.Location = New Point(300, 200)
            MainMenu.lblProcessing.Visible = False
            serverLabel.Text = MainMenu.serverLabel.Text
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con = New SqlConnection(conString)
            Dim cnt As Integer = 0

            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Sales_Plan ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYear.Items.Add(rdr(0))
            End While
            con.Close()
            thisYear = DatePart(DateInterval.Year, Date.Today())
            cboYear.SelectedIndex = cboYear.FindString(thisYear)

            con.Open()
            totalPeriods = 0
            sql = "SELECT DISTINCT Prd_Id FROM Sales_Plan WHERE Prd_Id > 0 ORDER By Prd_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboPeriod.Items.Add(rdr("Prd_Id"))
                totalPeriods += 1
            End While
            con.Close()
            cboPeriod.SelectedIndex = 0
            thisPeriod = cboPeriod.SelectedItem

            con.Open()
            sql = "SELECT ID FROM Stores WHERE Status = 'Active' ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStore.Items.Add(rdr(0))
            End While
            con.Close()
            cboStore.SelectedIndex = 0
            thisStore = cboStore.SelectedItem

            con.Open()
            sql = "SELECT ID FROM Departments WHERE Status = 'Active' ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr(0)
                cboDept.Items.Add(rdr(0))
            End While
            con.Close()
            cboDept.SelectedIndex = 0
            thisDept = cboDept.SelectedItem

            con.Open()
            sql = "SELECT Year_Id, Prd_Id FROM Calendar WHERE '" & today & "' BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysYear = rdr("Year_Id")
                todaysPeriod = rdr("Prd_Id")
            End While
            con.Close()

            Call Load_Data()
            formLoaded = True
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            ''MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "ERROR LOADING FORM")
        End Try
    End Sub

    Private Sub Load_Data()
        Try
            Dim stopwatch As New Stopwatch
            stopwatch.Start()
            lblProcessing.Visible = True
            Me.Refresh()

            somethingChanged = False
            tbl = New DataTable
            tbl = New DataTable
            tbl.Columns.Clear()
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Buyer"
            tbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = tbl.Columns("Buyer")
            tbl.PrimaryKey = PrimaryKey
            tbl.Columns.Add("Yearly Average")
            tbl.Columns.Add("6 Period Average")
            tbl.Columns.Add("3 period Average")
            tbl.Columns.Add("Period Average")
            tbl.Columns.Add("Current Plan%")
            '' tbl.Columns.Add("New Plan%")
            tbl.Columns.Add("Enter New %")
            tbl.Columns.Add("Last Year Actual%")
            tbl.Columns.Add("Current OnHand %")
            Dim tAvg52, tAvg26, tAvg12, tAvg4, tCurrent, tInv, tlyAvg As Decimal
            Dim tRows As Integer
            con.Open()
            sql = "SELECT Buyer FROM Buyers WHERE Status = 'Active' ORDER BY Buyer"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisBuyer = rdr("Buyer")
                row = tbl.NewRow
                oTest = rdr("Buyer")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row(0) = oTest
                con2.Open()
                '' ''sql = "DECLARE @wk52 date, @wk26 date, @wk12 date, @wk4 date " & _
                '' ''    "SELECT @wk52 = DATEADD(week,-52,'" & today & "') " & _
                '' ''    "SELECT @wk26 = DATEADD(week,-26,'" & today & "') " & _
                '' ''    "SELECT @wk12 = DATEADD(week,-12,'" & today & "') " & _
                '' ''    "SELECT @wk4 = DATEADD(week,-4,'" & today & "') " & _
                '' ''    "CREATE TABLE #t1 (Buyer varchar(20) NOT NULL, Avg52 decimal(18,4)) INSERT INTO #t1 (Buyer, Avg52) " & _
                '' ''    "SELECT Buyer, SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                '' ''       "WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg52 " & _
                '' ''       "FROM Item_Sales WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' " & _
                '' ''           "AND Dept = '" & thisDept & "' " & _
                '' ''           "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                '' ''    "CREATE TABLE #t2 (Buyer varchar(20) NOT NULL, Avg26 decimal(18,4)) INSERT INTO #t2 (Buyer, Avg26) " & _
                '' ''    "SELECT Buyer, SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                '' ''        "WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg26 " & _
                '' ''        "FROM Item_Sales WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' " & _
                '' ''           " AND Dept = '" & thisDept & "' " & _
                '' ''           "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                '' ''    "CREATE TABLE #t3 (Buyer varchar(20) NOT NULL, Avg12 decimal(18,4)) INSERT INTO #t3 (Buyer, Avg12) " & _
                '' ''    "SELECT Buyer, SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                '' ''        "WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg12 " & _
                '' ''        "FROM Item_Sales WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' " & _
                '' ''            "AND Dept = '" & thisDept & "' " & _
                '' ''            "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                '' ''    "CREATE TABLE #t4 (Buyer varchar(20) NOT NULL, Avg4 decimal(18,4)) INSERT INTO #t4 (Buyer, Avg4) " & _
                '' ''    "SELECT Buyer, SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                '' ''        "WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg4 " & _
                '' ''        "FROM Item_Sales WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' " & _
                '' ''            "AND Dept = '" & thisDept & "' " & _
                '' ''            "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                '' ''    "CREATE TABLE #t6 (Buyer varchar(20) NOT NULL, lyAvg decimal(18,4)) INSERT INTO #t6 (Buyer, lyAvg) " & _
                '' ''    "SELECT Buyer, SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                '' ''       "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & thisYear - 1 & " AND Prd_Id = " & thisPeriod & " " & _
                '' ''           "AND Dept = '" & thisDept & "') AS lyAvg " & _
                '' ''           "FROM Item_Sales WHERE Year_Id = " & thisYear - 1 & " AND Prd_Id = " & thisPeriod & " " & _
                '' ''           "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                '' ''           "AND Prd_ID = " & thisPeriod & " GROUP BY Buyer " & _
                '' ''    "CREATE TABLE #t5 (Buyer varchar(20) NOT NULL, pctInv decimal(18,4)) INSERT INTO #t5 (Buyer, pctInv) " & _
                '' ''    "SELECT w.Buyer, SUM(Act_Inv_Retail) / (SELECT SUM(Act_Inv_Retail) FROM Item_Sales " & _
                '' ''        "WHERE CONVERT(VARCHAR,'" & today & "',101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "' " & _
                '' ''        "AND Act_Inv_Retail > 0) AS pctInv " & _
                '' ''        "FROM Item_Sales w INNER JOIN Buyers b ON b.ID = w.Buyer " & _
                '' ''            "WHERE CONVERT(VARCHAR,'" & today & "',101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "' AND Status = 'Active' " & _
                '' ''            "GROUP BY w.Buyer " & _
                '' ''    "SELECT #t1.Buyer, #t1.Avg52, #t2.Avg26, #t3.Avg12, #t4.Avg4, " & _
                '' ''        "(SELECT PCT FROM Buyer_PCT p WHERE Plan_Year = '" & thisYear & "' " & _
                '' ''           "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " " & _
                '' ''           "AND p.Buyer = '" & thisBuyer & "' " & _
                '' ''           "AND Dept = '" & thisDept & "') AS Prd1, #t6.lyAvg, #t5.pctInv " & _
                '' ''           "FROM #t1 JOIN #t2 ON #t2.Buyer = #t1.Buyer " & _
                '' ''           "LEFT JOIN #t3 ON #t3.Buyer = #t1.Buyer " & _
                '' ''           "LEFT JOIN #t4 ON #t4.Buyer = #t1.Buyer " & _
                '' ''           "LEFT JOIN #t5 ON #t5.Buyer = #t1.Buyer " & _
                '' ''           "LEFT JOIN #t6 ON #t6.Buyer = #t1.Buyer " & _
                '' ''           "INNER JOIN Buyers b ON b.Buyer = #t1.Buyer ORDER BY b.Buyer"
                ''sql = "DECLARE @wk52 date, @wk26 date, @wk12 date, @wk4 date " & _
                ''    "SELECT @wk52 = DATEADD(week,-52,'" & today & "') " & _
                ''    "SELECT @wk26 = DATEADD(week,-26,'" & today & "') " & _
                ''    "SELECT @wk12 = DATEADD(week,-12,'" & today & "') " & _
                ''    "SELECT @wk4 = DATEADD(week,-4,'" & today & "') " & _
                ''    "CREATE TABLE #t1 (Buyer varchar(20) NOT NULL, Avg52 decimal(18,4)) " & _
                ''    "INSERT INTO #t1 (Buyer, Avg52) " & _
                ''    "SELECT Buyer, SUM(Sales_Retail) / (SELECT ISNULL(SUM(Sales_Retail),0) FROM Item_Sales d " & _
                ''    "JOIN Item_Master m on m.Sku = d.Sku " & _
                ''       "WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg52 " & _
                ''       "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' " & _
                ''           "AND Dept = '" & thisDept & "' " & _
                ''           "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                ''    "CREATE TABLE #t2 (Buyer varchar(20) NOT NULL, Avg26 decimal(18,4)) " & _
                ''    "INSERT INTO #t2 (Buyer, Avg26) " & _
                ''    "SELECT Buyer, SUM(Sales_Retail) / (SELECT ISNULL(SUM(Sales_Retail),0) FROM Item_Sales d " & _
                ''    "JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg26 " & _
                ''        "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' " & _
                ''           " AND Dept = '" & thisDept & "' " & _
                ''           "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                ''    "CREATE TABLE #t3 (Buyer varchar(20) NOT NULL, Avg12 decimal(18,4)) " & _
                ''    "INSERT INTO #t3 (Buyer, Avg12) " & _
                ''    "SELECT Buyer, SUM(Sales_Retail) / (SELECT ISNULL(SUM(Sales_Retail),0) FROM Item_Sales d " & _
                ''    "JOIN Item_Master m on m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg12 " & _
                ''        "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' " & _
                ''            "AND Dept = '" & thisDept & "' " & _
                ''            "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                ''    "CREATE TABLE #t4 (Buyer varchar(20) NOT NULL, Avg4 decimal(18,4)) " & _
                ''    "INSERT INTO #t4 (Buyer, Avg4) " & _
                ''    "SELECT Buyer, SUM(Sales_Retail) / (SELECT ISNULL(SUM(Sales_Retail),0) FROM Item_Sales d " & _
                ''    "JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg4 " & _
                ''        "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''        "WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' " & _
                ''            "AND Dept = '" & thisDept & "' " & _
                ''            "AND Buyer = '" & thisBuyer & "' GROUP BY Buyer " & _
                ''    "CREATE TABLE #t6 (Buyer varchar(20) NOT NULL, lyAvg decimal(18,4)) " & _
                ''    "INSERT INTO #t6 (Buyer, lyAvg) " & _
                ''    "SELECT Buyer, SUM(Sales_Retail) / (SELECT ISNULL(SUM(Sales_Retail),0) FROM Item_Sales d " & _
                ''    "JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''    "JOIN Calendar c ON c.eDate = d.eDate AND c.Week_Id > 0 " & _
                ''       "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & thisYear - 1 & " AND Prd_Id = " & thisPeriod & " " & _
                ''           "AND Dept = '" & thisDept & "') AS lyAvg " & _
                ''           "FROM Item_Sales d JOIN Item_Master m ON m.Sku = d.Sku " & _
                ''           "JOIN Calendar c ON c.eDate = d.eDate AND c.Week_Id > 0 " & _
                ''           "WHERE Year_Id = " & thisYear - 1 & " AND Prd_Id = " & thisPeriod & " " & _
                ''           "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                ''           "AND Prd_ID = " & thisPeriod & " GROUP BY Buyer " & _
                ''    "CREATE TABLE #t5 (Buyer varchar(20) NOT NULL, pctInv decimal(18,4)) " & _
                ''    "INSERT INTO #t5 (Buyer, pctInv) " & _
                ''    "SELECT m.Buyer, SUM(End_OH * Retail) / (SELECT SUM(End_OH * Retail) FROM Item_Inv i " & _
                ''    "JOIN Item_Master m On m.Sku = i.Sku " & _
                ''        "WHERE CONVERT(VARCHAR,'" & today & "',101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "' " & _
                ''        "AND End_OH > 0) AS pctInv " & _
                ''        "FROM Item_Inv i JOIN Item_Master m ON m.Sku = i.Sku " & _
                ''        "INNER JOIN Buyers b ON b.ID = m.Buyer " & _
                ''            "WHERE CONVERT(VARCHAR,'" & today & "',101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "' " & _
                ''            "AND b.Status = 'Active' " & _
                ''            "GROUP BY m.Buyer " & _
                ''    "SELECT #t1.Buyer, #t1.Avg52, #t2.Avg26, #t3.Avg12, #t4.Avg4, " & _
                ''        "(SELECT PCT FROM Buyer_PCT p WHERE Plan_Year = '" & thisYear & "' " & _
                ''           "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " " & _
                ''           "AND p.Buyer = '" & thisBuyer & "' " & _
                ''           "AND Dept = '" & thisDept & "') AS Prd1, #t6.lyAvg, #t5.pctInv " & _
                ''           "FROM #t1 JOIN #t2 ON #t2.Buyer = #t1.Buyer " & _
                ''           "LEFT JOIN #t3 ON #t3.Buyer = #t1.Buyer " & _
                ''           "LEFT JOIN #t4 ON #t4.Buyer = #t1.Buyer " & _
                ''           "LEFT JOIN #t5 ON #t5.Buyer = #t1.Buyer " & _
                ''           "LEFT JOIN #t6 ON #t6.Buyer = #t1.Buyer " & _
                ''           "INNER JOIN Buyers b ON b.Buyer = #t1.Buyer ORDER BY b.Buyer"

                cmd = New SqlCommand("sp_RCAdmin_BuyerPCT_GetData", con2)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
                cmd.Parameters.Add("@dept", SqlDbType.VarChar).Value = thisDept
                cmd.Parameters.Add("@buyer", SqlDbType.VarChar).Value = thisBuyer
                cmd.Parameters.Add("@year", SqlDbType.Int).Value = thisYear
                cmd.Parameters.Add("@period", SqlDbType.Int).Value = thisPeriod
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    oTest = rdr2(0)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        oTest = rdr2("Avg52")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(1) = Format(rdr2(1), "###.0%")
                            tAvg52 += CDec(oTest)
                        Else : row(1) = "%"
                        End If
                        oTest = rdr2("Avg26")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(2) = Format(rdr2(2), "###.0%")
                            tAvg26 += CDec(oTest)
                        Else : row(2) = "%"
                        End If
                        oTest = rdr2("Avg12")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(3) = Format(rdr2(3), "###.0%")
                            tAvg12 += CDec(oTest)
                        Else : row(3) = "%"
                        End If
                        oTest = rdr2("Avg4")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(4) = Format(rdr2("Avg4"), "###.0%")
                            tAvg4 += rdr2("Avg4")
                        Else : row(4) = "%"
                        End If

                        oTest = rdr2("Prd1")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(5) = Format(oTest, "###.0%")
                            tCurrent += oTest
                        Else : row(5) = "%"
                        End If

                        oTest = rdr2("lyAvg")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(7) = Format(oTest, "###.0%")
                            tlyAvg += oTest
                        Else : row(7) = "%"
                        End If
                        oTest = rdr2("pctInv")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row(8) = Format(oTest, "###.0%")
                            tInv += oTest
                        Else : row(8) = "%"
                        End If

                    End If
                End While
                con2.Close()
                tbl.Rows.Add(row)
            End While
            con.Close()

            Dim tlt As Decimal = tAvg52 + tAvg26 + tAvg12 + tAvg4 + tCurrent + tlyAvg + tInv
            ''If tlt = 7 Then
            ''Else
            ''    row = tbl.NewRow
            ''    row(0) = "OTHER"
            ''    row(1) = Format(1 - tAvg52, "###.#%")
            ''    row(2) = Format(1 - tAvg26, "###.#%")
            ''    row(3) = Format(1 - tAvg12, "###.#%")
            ''    row(4) = Format(1 - tAvg4, "###.#%")
            ''    row(5) = Format(1 - tCurrent, "###.#%")
            ''    row(7) = Format(1 - tlyAvg, "###.0%")
            ''    row(8) = Format(1 - tInv, "###.#%")
            ''    tbl.Rows.Add(row)
            ''    tRows = tbl.Rows.Count - 1
            ''    tAvg52 += (1 - tAvg52)
            ''    tAvg26 += (1 - tAvg26)
            ''    tAvg12 += (1 - tAvg12)
            ''    tAvg4 += (1 - tAvg4)
            ''    tCurrent += (1 - tCurrent)
            ''    tlyAvg += (1 - tlyAvg)
            ''    tInv += (1 - tInv)
            ''End If
            row = tbl.NewRow
            row(0) = "Total"
            row(1) = Format(tAvg52, "###.#%")
            row(2) = Format(tAvg26, "###.#%")
            row(3) = Format(tAvg12, "###.#%")
            row(4) = Format(tAvg4, "###.#%")
            row(5) = Format(tCurrent, "###.#%")
            row(7) = Format(tlyAvg, "###.0%")
            row(8) = Format(tInv, "###.#%")
            tbl.Rows.Add(row)
            tRows = tbl.Rows.Count - 1
            row = tbl.NewRow
            row(0) = "Difference"

            row(1) = Format(1 - tAvg52, "###.0%")
            row(2) = Format(1 - tAvg26, "###.0%")
            row(3) = Format(1 - tAvg12, "###.0%")
            row(4) = Format(1 - tAvg4, "###.0%")
            row(5) = Format(1 - tCurrent, "###.0%")
            row(7) = Format(1 - tlyAvg, "###.0%")
            row(8) = Format(1 - tInv, "###.0%")
            tbl.Rows.Add(row)
            lblProcessing.Visible = False

            dgv1.DataSource = tbl.DefaultView
            dgv1.AutoResizeColumns()
            dgv1.Columns(0).ReadOnly = True
            dgv1.Rows(tRows).ReadOnly = True
            dgv1.Rows(tRows + 1).ReadOnly = True
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Format = "p2"
                If i = 6 Then dgv1.Columns(6).DefaultCellStyle.BackColor = Color.Cornsilk
                dgv1.Columns(i).ReadOnly = True
            Next
            dgv1.Columns(6).ReadOnly = False
            dgv1.Refresh()
            stopwatch.Stop()
            Dim ts As TimeSpan = stopwatch.Elapsed
            Dim et As String = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
            '' MsgBox("Elapsed time = " & et)
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            ''MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "ERROR LOADING DATA")
        End Try
    End Sub

    'Restrict Draft Variance to numbers only
    Private Sub dgv1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then
            Dim tb As TextBox = TryCast(e.Control, TextBox)
            If Me.dgv1.CurrentCell.ColumnIndex = 6 Then
                AddHandler tb.KeyPress, AddressOf tb_KeyPress
            End If
        End If
    End Sub

    Private Sub tb_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) And Not Char.IsDigit(e.KeyChar) And e.KeyChar <> "." And e.KeyChar <> "-" Then
            e.Handled = True
        End If
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Try
            Dim rowIndex, columnIndex As Integer
            If e.RowIndex > -1 Then rowIndex = e.RowIndex
            If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
            ''  MsgBox("you're in cell enter row " & rowIndex & " column " & columnIndex)
            Dim change, totalChange As Decimal
            row = tbl(rowIndex)
            oTest = dgv1.Rows(rowIndex).Cells("Enter New %").Value
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                somethingChanged = True
                If thisYear <= todaysYear And thisPeriod < todaysPeriod Then
                    MessageBox.Show("PERCENTAGES CANNOT BE CHANGED FOR PRIOR PERIODS!", "ENTER NEW PLAN PERCENT")
                    row(columnIndex) = Nothing
                    Exit Sub
                End If
                change = oTest * 0.01
                For Each row As DataRow In tbl.Rows
                    oTest = row(0)
                    If oTest <> "Total" And oTest <> "Difference" Then
                        oTest = row("Enter New %")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            oTest = Replace(oTest, "%", "")
                            If oTest <> "" And IsNumeric(oTest) Then
                                change = CDec(oTest * 0.01)
                                totalChange += change
                            End If
                        End If
                    End If
                Next
                Dim foundRow As DataRow = tbl.Rows.Find("Total")
                foundRow("Enter New %") = Format(totalChange, "###.0%")
                foundRow = tbl.Rows.Find("Difference")
                foundRow("Enter New %") = Format(1 - totalChange, "###.0%")
            End If
            Dim dgvrow As DataGridViewRow = Me.dgv1.Rows(rowIndex)
            Me.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ENTER NEW PLAN PERCENT")
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        Try
            Dim foundRow As DataRow = tbl.Rows.Find("Total")
            If Not IsNothing(foundRow) Then
                oTest = foundRow("Enter New %")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    Dim pct As Decimal = CDec(Replace(oTest, "%", ""))
                    If pct <> 100 Then
                        MessageBox.Show("ADJUSTMENTS MUST EQUAL 100% BEFORE THEY CAN BE SAVED!", "SAVE CHANGES")
                        Exit Sub
                    End If
                End If
            End If
            Dim message, title As String
            Dim theValue As Decimal
            Dim r As Integer = totalPeriods - thisPeriod
            Dim endPeriod As Integer = totalPeriods
            message = "CLICK OK TO SAVE THIS PERIOD ONLY. OPTIONALLY, ENTER A NUMBER UP TO " & r & " TO COPY THESE PERCENTAGES TO SUBSEQUENT PERIODS"
            title = "SAVE PERCENTAGES TO PERIOD(S)"
            oTest = InputBox(message, title)

            If IsNumeric(oTest) Then
                If CInt(oTest) > r Then
                    MessageBox.Show("NUMBER CAN'T EXCEED " & r & "!", "SAVE CHANGES")
                    Exit Sub
                Else
                    endPeriod = thisPeriod + CInt(oTest)
                End If
                If CInt(oTest) < 0 Then
                    MessageBox.Show("NUMBER CAN'T BE LESS THAN ZERO!", "SAVE CHANGES")
                    Exit Sub
                End If
            Else
                If oTest <> "" Then
                    MessageBox.Show("INVALID ENTRY!", "SAVE CHANGES")
                    Exit Sub
                Else : endPeriod = thisPeriod
                End If
            End If

            Dim cnt As Integer = 0

            For Each row As DataRow In tbl.Rows

                oTest = row("Enter New %")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If Not IsNothing(oTest) And Not IsDBNull(oTest) And row(0) <> "Difference" And row(0) <> "Total" Then
                        thisBuyer = row(0).ToString
                        theValue = CDec(Replace(oTest, "%", "") * 0.01)
                        con.Open()
                        For i As Integer = thisPeriod To endPeriod
                            sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Plan_Year = '" & thisYear & "' " & _
                                    "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                    "AND Buyer = '" & thisBuyer & "') " & _
                                "INSERT INTO Buyer_Pct (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, PCT) " & _
                                    "SELECT '" & thisYear & "', '" & thisStore & "', " & i & ", '" & thisDept & "', '" & _
                                    thisBuyer & "', " & theValue & " " & _
                                    "ELSE " & _
                                "UPDATE Buyer_Pct SET PCT = " & theValue & " WHERE Plan_Year = '" & thisYear & "' " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                "AND Buyer = '" & thisBuyer & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                            '
                            '              Update same record for future years if there are any              
                            '
                            sql = "UPDATE Buyer_PCT SET PCT = " & theValue & " WHERE Plan_Year > " & thisYear & " " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                "AND Buyer = '" & thisBuyer & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                        Next
                        con.Close()
                        row("Current Plan%") = Format(oTest * 0.01, "###%")
                        dgv1.Item(5, cnt).Style.BackColor = Color.Yellow
                        dgv1.Item(5, cnt).Style.Font = New Font("Sans Serif", 8.25, FontStyle.Bold)
                    End If
                End If
                cnt += 1
            Next
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
        End Try

    End Sub

    Private Sub cboYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYear.SelectedIndexChanged
        thisYear = cboYear.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboPeriod_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPeriod.SelectedIndexChanged
        thisPeriod = cboPeriod.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                e.Cancel = False
            Case CloseReason.UserClosing
                If somethingChanged Then
                    Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        Case DialogResult.Yes
                            Call Save_Changes()
                            e.Cancel = False
                        Case DialogResult.No
                            e.Cancel = False
                        Case Windows.Forms.DialogResult.Cancel
                            e.Cancel = True
                            Exit Sub
                    End Select
                End If
            Case Else
                e.Cancel = False
        End Select
    End Sub
End Class