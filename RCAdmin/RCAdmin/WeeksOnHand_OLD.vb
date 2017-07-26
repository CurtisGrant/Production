Imports System.Data.SqlClient
Imports System.Xml
Public Class WeeksOnHand_OLD
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared wkTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = CDate(Date.Today)
    Public Shared lastYear As Integer = DatePart(DateInterval.Year, Today) - 1
    Public Shared thisYear, thisStore, thisPlan, thisDept, thisBuyer, thisClass, thisTable As String
    Public Shared currentYear, thisPeriod As Integer
    Public Shared thisDate, thisSdate, thisEdate, prevPeriodEndDate, prevYearStartDate, prevPeriodStartDate As Date
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, rowIndex, tRows As Int16
    Public Shared tNew As Decimal
    Public Shared periodDates(50) As String
    Public Shared theValue As String
    Public Shared thisUser = Environment.UserName
    Public Shared planBy As String
    Public Shared periodSelected, numberPeriods As Integer
    Public Shared buyers(50) As String
    Public Shared formLoaded As Boolean = False
    Public Shared changesMade As Boolean

    Private Sub WeeksSupply_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            Dim conString As String = MainMenu.conString
            Dim havePlan As Boolean = False
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            Dim cnt As Integer = 0

            con.Open()
            sql = "SELECT Year_Id FROM Calendar WHERE '" & today & "' BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                currentYear = rdr("Year_Id")
                thisYear = rdr("Year_Id")
            End While
            con.Close()

            con.Open()
            sql = "select dateadd(week,-52,(select max(edate) from calendar where edate < getdate() and prd_id > 0 and week_id=0)) as startdate, " & _
                "max(sdate) as prdstartdate, max(edate) as enddate from calendar where edate < getdate() and prd_id > 0 and week_id=0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                prevYearStartDate = rdr("startdate")
                prevPeriodStartDate = rdr("prdstartdate")
                prevPeriodEndDate = rdr("enddate")
            End While
            con.Close()

            Dim daysinperiod As Integer = DateDiff(DateInterval.Day, prevPeriodStartDate, prevPeriodEndDate) + 1

            con.Open()
            sql = "SELECT DISTINCT Str_Id FROM Buyer_PCT WHERE Year_Id = " & thisYear & " ORDER BY Str_Id"
            If con.State = ConnectionState.Closed Then con.Open()

            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStr.Items.Add(rdr("Str_Id"))
            End While
            con.Close()
            cboStr.SelectedIndex = 0


            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Buyer_PCT ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYr.Items.Add(rdr("Year_Id"))
            End While
            con.Close()
            Dim idx As Integer = cboYr.FindString(thisYear)
            cboYr.SelectedIndex = idx


            con.Open()
            sql = "SELECT DISTINCT p.Buyer FROM Buyer_PCT p JOIN Buyers b ON p.Buyer = b.Buyer " & _
                "WHERE Year_Id = " & thisYear & " AND b.Status = 'Active' ORDER BY p.Buyer"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboBuyer.Items.Add(rdr("Buyer"))
            End While
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT ID as Dept FROM Departments WHERE Status = 'Active' ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboDept.Items.Add(rdr("Dept"))
            End While
            cboDept.SelectedIndex = 0
            thisDept = cboDept.SelectedItem
            con.Close()

            Array.Resize(buyers, cnt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD FORM")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Load_Data()
        Try
            lblProcessing.Visible = True
            Me.Refresh()
            changesMade = False
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            Dim tAvg52 As Decimal = 0
            Dim tAvg26 As Decimal = 0
            Dim tAvg12 As Decimal = 0
            Dim tAvg4 As Decimal = 0
            Dim tCurrent As Decimal = 0
            Dim tInv As Decimal = 0
            Dim inv, sales As Decimal
            Dim firstRecord As Boolean = True
            Dim numBuyers As Integer = 0
            Dim period, wksoh As Integer
            Dim last6, last3, last1, prdEdate As Date
            Dim daysinperiod As Integer = DateDiff(DateInterval.Day, prevPeriodStartDate, prevPeriodEndDate) + 1
            pb1.Visible = True
            pb1.Step = 1
            pb1.Value = 1
            pb1.Minimum = 1
            Dim row As DataRow
            wkTbl = New DataTable
            wkTbl.Clear()
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Period"
            wkTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = wkTbl.Columns("Period")
            wkTbl.PrimaryKey = PrimaryKey
            wkTbl.Columns.Add("Yearly Average")
            wkTbl.Columns.Add("6 Period Average")
            wkTbl.Columns.Add("3 Period Average")
            wkTbl.Columns.Add("Period Average")
            wkTbl.Columns.Add("Plan WksOH")
            wkTbl.Columns.Add("Enter New Wks OnHand")
            wkTbl.Columns.Add("Sales Plan")

            con.Open()
            '' sql = "SELECT Prd_Id, sDate, eDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id > 0 AND Week_Id = 0 ORDER BY Prd_Id"
            sql = "SELECT Prd_Id, sDate, eDate FROM Calendar WHERE sDate >= '" & prevYearStartDate & "' AND eDate <= '" & prevPeriodEndDate & "' " & _
                "AND Prd_Id > 0 AND Week_Id = 0 ORDER BY Prd_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            Dim cnt As Integer = 0
            While rdr.Read
                prdEdate = rdr("eDate")
                last6 = DateAdd(DateInterval.Day, daysinperiod * -6, prdEdate)
                last3 = DateAdd(DateInterval.Day, daysinperiod * -3, prdEdate)
                last1 = DateAdd(DateInterval.Day, daysinperiod * -1, prdEdate)
                row = wkTbl.NewRow
                row(0) = rdr("Prd_Id")
                period = rdr("Prd_Id")
                periodDates(cnt) = rdr("sDate") & " - " & rdr("eDate")
                cnt += 1
                con2.Open()
                sql = "select isnull(sum(act_inv_retail),0) as inv, isnull(sum(act_sales),0) as sales from Item_Sales " & _
                    "where edate between dateadd(week,-52,'" & prdEdate & "') and '" & prdEdate & "' " & _
                    "and str_id='" & thisStore & "' and dept='" & thisDept & "' and buyer='" & thisBuyer & "' " & _
                    "and act_sales > 0"
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    inv = rdr2("inv")
                    sales = rdr2("sales")
                    If inv > 0 And sales <> 0 Then
                        wksoh = inv / sales
                        row("Yearly Average") = Format(wksoh, "####")
                    End If

                End While
                con2.Close()

                con2.Open()
                sql = "select isnull(sum(act_inv_retail),0) as inv, isnull(sum(act_sales),0) as sales from Item_Sales " & _
                   "where edate between '" & last6 & "' AND '" & prdEdate & "' " & _
                   "and str_id='" & thisStore & "' and dept='" & thisDept & "' and buyer='" & thisBuyer & "' " & _
                   "and act_sales > 0"
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    inv = rdr2("inv")
                    sales = rdr2("sales")
                    If inv > 0 And sales <> 0 Then
                        wksoh = inv / sales
                        row("6 Period Average") = Format(wksoh, "####")
                    End If
                End While
                con2.Close()

                con2.Open()
                sql = "select isnull(sum(act_inv_retail),0) as inv, isnull(sum(act_sales),0) as sales from Item_Sales " & _
                   "where edate between '" & last3 & "' AND '" & prdEdate & "' " & _
                   "and str_id='" & thisStore & "' and dept='" & thisDept & "' and buyer='" & thisBuyer & "' " & _
                   "and act_sales > 0"
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    inv = rdr2("inv")
                    sales = rdr2("sales")
                    If inv > 0 And sales <> 0 Then
                        wksoh = inv / sales
                        row("3 Period Average") = Format(wksoh, "####")
                    End If
                End While
                con2.Close()

                con2.Open()
                sql = "select isnull(sum(act_inv_retail),0) as inv, isnull(sum(act_sales),0) as sales from Item_Sales " & _
                   "where edate between '" & last1 & "' and '" & prdEdate & "' " & _
                   "and str_id='" & thisStore & "' and dept='" & thisDept & "' and buyer='" & thisBuyer & "' " & _
                   "and act_sales > 0"
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    inv = rdr2("inv")
                    sales = rdr2("sales")
                    If inv > 0 And sales <> 0 Then
                        wksoh = inv / sales
                        row("Period Average") = Format(wksoh, "####")
                    End If
                End While
                con2.Close()

                con2.Open()
                sql = "select isnull(plan_wksoh,0) from buyer_pct where str_id='" & thisStore & "' and Plan_year = " & thisYear & " " & _
                    "and dept='" & thisDept & "' and buyer='" & thisBuyer & "' and prd_id=" & period & " "
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    oTest = rdr2(0)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row("Plan WksOH") = Format(oTest, "####")
                End While
                con2.Close()

                con2.Open()
                sql = "SELECT ISNULL(p.Amt * b.Pct,0) FROM Sales_Plan p JOIN Buyer_PCT b ON b.Year_Id = p.Year_Id AND b.Prd_Id = p.Prd_Id " & _
                    "AND b.Str_Id = p. Str_Id AND b.Dept = p.Dept " & _
                    "WHERE p.Str_Id = '" & thisStore & "' AND p.Dept = '" & thisDept & "' AND b.Buyer = '" & thisBuyer & "' AND p.Year_Id = " & thisYear & " " & _
                    "AND p.Prd_Id = " & period & " AND p.Week_Id = 0 AND p.Status = 'Active'"
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    oTest = rdr2(0)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row("Sales Plan") = Format(oTest, "###,###,###")
                End While
                con2.Close()

                wkTbl.Rows.Add(row)

            End While
            Array.Resize(periodDates, cnt)
            con.Close()

            tRows = wkTbl.Rows.Count - 1
            pb1.PerformStep()
            pb1.Visible = False
            dgv1.DataSource = wkTbl.DefaultView
            dgv1.AutoResizeColumns()
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Format = "p2"
                If i = 6 Then dgv1.Columns(6).DefaultCellStyle.BackColor = Color.Cornsilk
                dgv1.Columns(i).ReadOnly = True
            Next
            tNew = 0
            dgv1.Columns(6).ReadOnly = False
            dgv1.Refresh()
            formLoaded = True
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD DATA")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
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
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        If columnIndex = 6 And e.RowIndex < tRows Then
            oTest = dgv1.Rows(rowIndex).Cells(0).Value
            Dim row As DataRow = wkTbl.Rows.Find(CInt(oTest))
            If thisYear <= currentYear And (rowIndex + 1) <= thisPeriod Then
                MessageBox.Show("CURRENT OR PRIOR PERIODS CANNOT BE CHANGED!", "CHANGE WEEKS SUPPLY")
                row(6) = Nothing
                Exit Sub
            End If
            oTest = dgv1.Rows(rowIndex).Cells(6).Value
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then
                    dgv1.Rows(rowIndex).Cells(6).Value = Format(CInt(oTest), "####")
                    Me.Refresh()
                    changesMade = True
                End If
            End If
        End If
        Me.Refresh()
    End Sub

    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim x As Integer = dgv1.Rows.Count
        Dim str As String
        If columnIdx = 0 Then
            If rowIdx > -1 And rowIdx < x Then
                str = periodDates(rowIdx)
                With dgv1.Rows(rowIdx).Cells(0)
                    .ToolTipText = str
                End With
            End If
        End If
    End Sub

    Private Sub cboYr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYr.SelectedIndexChanged
        thisYear = cboYr.SelectedItem
        Dim str, mm, dd As String
        mm = DatePart(DateInterval.Month, Date.Today)
        dd = DatePart(DateInterval.Day, Date.Today)
        str = "" + thisYear + "-" + mm + "-" + dd + ""
        thisDate = Date.Parse(str)
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboStr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStr.SelectedIndexChanged
        thisStore = cboStr.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboBuyer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboBuyer.SelectedIndexChanged
        thisBuyer = cboBuyer.SelectedItem
        ''If con.State = ConnectionState.Closed Then MsgBox("Buyer changed con is closed")
        ''If con.State = ConnectionState.Open Then con.Close()
        Call Load_Data()
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        lblProcessing.Visible = True
        Me.Refresh()
        Call Save_Changes()
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub

    Private Sub Save_Changes()
        Try
            Dim period As Integer = 0
            Dim cnt As Integer = 0
            For Each row In wkTbl.Rows
                period += 1
                oTest = row(6)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then
                        thisBuyer = cboBuyer.SelectedItem.ToString
                        theValue = CInt(oTest)
                        con.Open()
                        sql = "UPDATE Item_Sales SET Plan_wksOH = " & oTest & " WHERE Year_Id = " & thisYear & " " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & period & " AND Dept = '" & thisDept & "' " & _
                                "AND Buyer = '" & thisBuyer & "' "
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " " & _
                            "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "') " & _
                            "INSERT INTO Buyer_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Plan_WksOH) " & _
                            "SELECT " & thisYear & ",'" & thisStore & "'," & period & ",'" & thisDept & "','" & thisBuyer & "'," & oTest & " " & _
                            "ELSE " & _
                            "UPDATE Buyer_Pct SET Plan_WksOH = " & oTest & " WHERE Plan_Year = " & thisYear & " " & _
                            "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                            "AND Prd_Id = " & period & " "
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        con.Close()
                        row(5) = oTest
                        dgv1.Item(5, cnt).Style.BackColor = Color.Yellow
                        dgv1.Item(5, cnt).Style.Font = New Font("Sans Serif", 8.25, FontStyle.Bold)
                    End If
                End If
                cnt += 1
            Next
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            MessageBox.Show(ex.Message, "SAVE CHANGES")
        End Try
    End Sub

    Private Sub Backfill()

        Try
            Dim actInv As Int32 = 0
            Dim planSales As Int32 = 0
            Dim numWeeks As Int16
            Dim eDatex As Date
            Dim oTest As Object
            Dim store, dept, buyer, clss As String
            store = cboStr.SelectedItem.ToString
            dept = cboDept.SelectedItem.ToString
            buyer = cboBuyer.SelectedItem.ToString


            txtPlan.Visible = True
            lblPlan.Visible = True


            con.Open()
            sql = "SELECT sDate, eDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisSdate = rdr("sDate")
                thisEdate = rdr("eDate")
            End While
            con.Close()

            con.Open()
            sql = "Select Str_Id, Dept, Buyer, Class, eDate, Act_Inv_Retail FROM Item_Sales " & _
                "WHERE Act_Inv_Retail > 0 AND eDate >= '" & thisSdate & "' AND eDate <= '" & thisEdate & "' " & _
                "AND Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' " & _
                "ORDER BY eDate, Str_Id, Dept, Buyer, Class"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                store = rdr(0)
                dept = rdr(1)
                buyer = rdr(2)
                clss = rdr(3)
                eDatex = rdr(4)
                oTest = rdr(5)

                txtPlan.Text = store & " " & dept & " " & buyer & " " & clss & " " & eDatex
                Me.Refresh()

                If IsNumeric(oTest) Then actInv = oTest Else actInv = 0
                Console.WriteLine(eDatex & " " & store & " " & dept & " " & buyer & " " & clss & " " & oTest)
                planSales = 0
                numWeeks = 0
                con2.Open()
                sql = "SELECT ISNULL(Plan_Sales,0) + ISNULL(Plan_Mkdn,0) As Sales FROM Item_Sales " & _
                    "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND " & _
                    "Class = '" & clss & "' AND eDate > '" & eDatex & "' ORDER by eDate"
                Dim cmd26 As New SqlCommand(sql, con2)
                Dim rdr26 As SqlDataReader = cmd26.ExecuteReader
                While rdr26.Read
                    oTest = rdr26(0)
                    If IsNumeric(oTest) Then
                        planSales += oTest
                        numWeeks += 1
                    End If
                    If planSales >= actInv Then GoTo 10
                End While
10:             con2.Close()
                If planSales > 0 Then
                    con2.Open()
                    sql = "UPDATE Item_Sales SET Projected_WksOH = " & numWeeks & " WHERE Str_Id = '" & store & "' AND " & numWeeks & " > 0 " & _
                        "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND eDate = '" & eDatex & "'"
                    Dim cmd27 As New SqlCommand(sql, con2)
                    cmd27.ExecuteNonQuery()
                    con2.Close()
                End If
            End While
            con.Close()
            txtPlan.Visible = False
            lblPlan.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '' ''    Private Sub btnBackfill_Click(sender As Object, e As EventArgs) Handles btnBackfill.Click
    '' ''        txtPlan.Visible = True
    '' ''        Dim actInv As Int32 = 0
    '' ''        Dim planSales As Int32 = 0
    '' ''        Dim numWeeks As Int16
    '' ''        Dim eDatex, Start_Date, ThisEndDate As Date
    '' ''        Dim oTest As Object
    '' ''        Dim store, dept, buyer, clss As String
    '' ''        Dim finalWeek, sdate, edate As Date
    '' ''        Dim year, period, week, yrprd, weeknum, yrwk, cnt As Int16
    '' ''        Dim sales, mkdn, amt As Decimal
    '' ''        con.Open()
    '' ''        sql = "SELECT MAX(eDate) FROM Item_Sales WHERE Year_Id = DATEPART(Year,GETDATE())"
    '' ''        Dim cmd0 As New SqlCommand(sql, con)
    '' ''        Dim rdr0 As SqlDataReader = cmd0.ExecuteReader
    '' ''        While rdr0.Read
    '' ''            finalWeek = rdr0(0)
    '' ''        End While
    '' ''        con.Close()






    '' ''        ''   GoTo 5





    '' ''        '                   Update Item_Sales records with Plan_Sales & Plan_Mkdn
    '' ''        con.Open()
    '' ''        sql = "SELECT s.Str_Id, s.Year_Id, c.Prd_Id, s.Week_Id, c.Dept, c.Buyer, c.Class, ((s.Amt * c.Pct) * b.Pct) AS amt FROM Sales_Plan s " & _
    '' ''            "JOIN Class_PCT c ON c.Str_Id = s.Str_Id AND c.Plan_Year = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.Dept = s.Dept " & _
    '' ''            "JOIN Buyer_PCT b ON b.Plan_Year = s.Year_Id AND b.Str_Id = s.Str_Id AND b.Prd_Id = s.Prd_Id AND b.Dept = s.Dept AND b.Buyer = c.Buyer " & _
    '' ''            "WHERE Plan_Id = '2015_Plan1' AND s.Str_Id = c.Str_Id AND s.Week_Id > 0 ORDER BY s.Str_Id, s.Dept, s.Prd_Id"
    '' ''        cmd = New SqlCommand(sql, con)
    '' ''        rdr = cmd.ExecuteReader
    '' ''        While rdr.Read
    '' ''            oTest = rdr("Str_Id")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then store = oTest
    '' ''            oTest = rdr("Year_Id")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then year = oTest
    '' ''            oTest = rdr("Prd_Id")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then period = oTest
    '' ''            oTest = rdr("Week_Id")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then week = oTest
    '' ''            oTest = rdr("Dept")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest
    '' ''            oTest = rdr("Buyer")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = oTest
    '' ''            oTest = rdr("Class")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then clss = oTest
    '' ''            oTest = rdr("amt")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
    '' ''                If IsNumeric(oTest) Then amt = oTest Else amt = 0
    '' ''            End If
    '' ''            cnt += 1
    '' ''            If cnt Mod 100 Then
    '' ''                txtPlan.Text = "Updating " & store & " " & dept & " " & buyer & " " & clss & " " & period & " " & week
    '' ''                Me.Refresh()
    '' ''            End If
    '' ''            con2.Open()
    '' ''            sql = "IF NOT EXISTS (SELECT Str_Id FROM Item_Sales WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND " & _
    '' ''                "Year_Id = " & year & " AND Prd_Id = " & period & " AND Week_Id = " & week & ") " & _
    '' ''                "INSERT INTO Item_Sales (Str_Id, Dept, Buyer, Class, Year_Id, Prd_Id, Week_Id, sDate, eDate, YrPrd, Week_Num, Plan_Sales) " & _
    '' ''                "SELECT '" & store & "','" & dept & "','" & buyer & "','" & clss & "'," & year & "," & period & "," & week & ",sDate,eDate,YrPrd,Week_Id," & amt & " FROM Calendar " & _
    '' ''                "WHERE Year_Id = " & year & " AND Prd_Id = " & period & " AND PrdWk = " & week & " " & _
    '' ''                "ELSE " & _
    '' ''                "UPDATE Item_Sales SET Plan_Sales = " & CInt(amt) & " WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND " & _
    '' ''                "Year_Id = " & year & " AND Prd_Id = " & period & " AND Week_Id = " & week & " "
    '' ''            cmd = New SqlCommand(sql, con2)
    '' ''            cmd.ExecuteNonQuery()
    '' ''            con2.Close()
    '' ''        End While
    '' ''        con.Close()

    '' ''5:      Dim dte As String = InputBox("Enter sDate to Begin", "Enter sDate to Begin", "Enter theDate")
    '' ''        Start_Date = CDate(dte)
    '' ''        Dim theEndDate, End_Date As Date
    '' ''        con.Open()
    '' ''        sql = "SELECT MAX(eDate) AS eDate FROM Calendar WHERE Year_Id = DATEPART(year,'" & Start_Date & "')"
    '' ''        cmd = New SqlCommand(sql, con)
    '' ''        rdr = cmd.ExecuteReader
    '' ''        While rdr.Read
    '' ''            theEndDate = rdr("eDate")
    '' ''        End While
    '' ''        con.Close()

    '' ''        '                                                 Update Plan_Inv_Retail
    '' ''        con.Open()
    '' ''        sql = "SELECT DISTINCT eDate FROM Item_Sales WHERE sDate >= '" & Start_Date & "' " & _
    '' ''            "AND eDate <= '" & theEndDate & "' ORDER BY eDate"
    '' ''        cmd = New SqlCommand(sql, con)
    '' ''        cmd.CommandTimeout = 120
    '' ''        rdr = cmd.ExecuteReader
    '' ''        While rdr.Read
    '' ''            End_Date = rdr("eDate")
    '' ''            ThisEndDate = rdr("eDate")
    '' ''            Dim thisWeek As Date = End_Date                                       ' this weeks eDate
    '' ''            Dim prevWeek As Date = DateAdd(DateInterval.Day, -7, thisWeek)
    '' ''            Do Until thisWeek >= theEndDate                                         ' finalWeek is the last eDate in OTB

    '' ''                txtPlan.Text = "Updating Plan_Inv_Retail for week " & End_Date & "  " & thisWeek
    '' ''                Me.Refresh()

    '' ''                sql = "UPDATE Item_Sales SET Item_Sales.Plan_Inv_Retail = ( " & _
    '' ''                    "SELECT SUM(ISNULL(o2.Plan_Sales,0)) + SUM(ISNULL(o2.Plan_Mkdn,0)) FROM Item_Sales AS o2 " & _
    '' ''                    "WHERE o2.Str_Id = Item_Sales.Str_Id AND o2.Dept = Item_Sales.Dept " & _
    '' ''                    "AND o2.Buyer = Item_Sales.Buyer AND o2.Class = Item_Sales.Class " & _
    '' ''                    "AND o2.sDate BETWEEN Item_Sales.sDate AND DATEADD(week,Item_Sales.Plan_WksOH-1,Item_Sales.sDate)) " & _
    '' ''                    "WHERE Item_Sales.Buyer <> '*' AND Item_Sales.Class <> '*' " & _
    '' ''                    "AND Item_Sales.eDate = '" & thisWeek & "'"
    '' ''                con2.Open()
    '' ''                cmd = New SqlCommand(sql, con2)
    '' ''                cmd.ExecuteNonQuery()
    '' ''                If thisweek = thisEndDate Then
    '' ''                    sql = "UPDATE Item_Sales SET Projected_Inv = ISNULL(Act_Inv_Retail,0) " & _
    '' ''                        "+ ISNULL(OnOrder,0) " & _
    '' ''                        "- ISNULL(Plan_Sales,0) " & _
    '' ''                        "+ ISNULL(Act_Sales,0) " & _
    '' ''                        "- ISNULL(Plan_Mkdn,0) " & _
    '' ''                        "+ ISNULL(Act_Mkdn,0) WHERE eDate = '" & ThisEndDate & "'"
    '' ''                Else
    '' ''                    sql = "UPDATE f SET f.Projected_Inv = ISNULL(p.Projected_Inv,0) " & _
    '' ''                        "+ ISNULL(f.OnOrder,0) " & _
    '' ''                        "- ISNULL(f.Plan_Sales,0) " & _
    '' ''                        "+ ISNULL(f.Act_Sales,0) " & _
    '' ''                        "- ISNULL(f.Plan_Mkdn,0) " & _
    '' ''                        "+ ISNULL(f.Act_Mkdn,0) " & _
    '' ''                        "FROM Item_Sales AS f INNER JOIN Item_Sales AS p " & _
    '' ''                        "ON f.Str_Id = p.Str_Id AND p.Dept = f.Dept " & _
    '' ''                        "AND p.Buyer = f.Buyer AND p.Class = f.Class " & _
    '' ''                        "WHERE f.eDate = '" & thisWeek & "' " & _
    '' ''                        "AND p.eDate = '" & prevWeek & "'"
    '' ''                End If
    '' ''                cmd = New SqlCommand(sql, con2)
    '' ''                cmd.CommandTimeout = 120
    '' ''                cmd.ExecuteNonQuery()
    '' ''                con2.Close()
    '' ''                prevWeek = thisWeek
    '' ''                thisWeek = DateAdd(DateInterval.Day, 7, thisWeek)
    '' ''            Loop
    '' ''        End While
    '' ''        con.Close()
    '' ''55:
    '' ''        End_Date = CDate(dte)                                              ' set End_Date to eDate entered above
    '' ''        con2.Open()
    '' ''        sql = "Select Str_Id, Dept, Buyer, Class, eDate, Act_Inv_Retail FROM Item_Sales " & _
    '' ''            "WHERE Act_Inv_Retail > 0 AND eDate >= '" & End_Date & "' " & _
    '' ''            "ORDER BY eDate, Str_Id, Dept, Buyer, Class"
    '' ''        cmd = New SqlCommand(sql, con2)
    '' ''        cmd.CommandTimeout = 120
    '' ''        rdr2 = cmd.ExecuteReader
    '' ''        While rdr2.Read
    '' ''            store = "1"
    '' ''            dept = "Z"
    '' ''            buyer = ""
    '' ''            clss = ""
    '' ''            oTest = rdr2("Str_Id")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then store = oTest
    '' ''            oTest = rdr2("Dept")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest
    '' ''            oTest = rdr2("Buyer")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = oTest
    '' ''            oTest = rdr2("Class")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then clss = oTest
    '' ''            oTest = rdr2("eDate")
    '' ''            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then eDatex = oTest
    '' ''            oTest = rdr2("Act_Inv_Retail")
    '' ''            If IsNumeric(oTest) Then actInv = oTest Else actInv = 0

    '' ''            Console.WriteLine("Updating Projected_WksOH " & eDatex & " " & store & " " & dept & " " & buyer & " " & clss & " " & oTest)
    '' ''            planSales = 0
    '' ''            numWeeks = 0
    '' ''            con3.Open()
    '' ''            sql = "SELECT ISNULL(Plan_Sales,0) + ISNULL(Plan_Mkdn,0) As Sales FROM Item_Sales " & _
    '' ''                "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND " & _
    '' ''                "Class = '" & clss & "' AND eDate > '" & eDatex & "' ORDER by eDate"
    '' ''            Dim cmd26 As New SqlCommand(sql, con3)
    '' ''            cmd26.CommandTimeout = 120
    '' ''            Dim rdr26 As SqlDataReader = cmd26.ExecuteReader
    '' ''            While rdr26.Read And planSales < actInv
    '' ''                oTest = rdr26(0)
    '' ''                If IsNumeric(oTest) Then
    '' ''                    planSales += oTest
    '' ''                    numWeeks += 1
    '' ''                End If
    '' ''                '' If planSales >= actInv Then GoTo 10
    '' ''            End While
    '' ''10:         con3.Close()

    '' ''            If planSales > 0 Then
    '' ''                If con3.State = ConnectionState.Open Then con3.Close()
    '' ''                con3.Open()

    '' ''                sql = "UPDATE Item_Sales SET Projected_WksOH = " & numWeeks & " WHERE Str_Id = '" & store & "' " & _
    '' ''                    "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND eDate = '" & eDatex & "'"
    '' ''                Dim cmd27 As New SqlCommand(sql, con3)
    '' ''                cmd27.CommandTimeout = 120
    '' ''                cmd27.ExecuteNonQuery()
    '' ''                con3.Close()
    '' ''            End If
    '' ''        End While
    '' ''        con2.Close()
    '' ''        If con.State = ConnectionState.Open Then con.Close()

    '' ''    End Sub

    Private Sub btnSaveWeekly_Click(sender As Object, e As EventArgs)
        Try
            '      moved code to Activate Button on PlanMaintenance
            MsgBox(thisPlan & " Saved to Item_Sales")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SAVE CHANGES")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                e.Cancel = False
            Case CloseReason.UserClosing
                If changesMade Then
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