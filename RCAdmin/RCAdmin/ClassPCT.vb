Imports System.Data.SqlClient
Imports System.Xml
Public Class ClassPct
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared tbl, wkTbl, dayTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = CDate(Date.Today)
    Public Shared thisYear, thisStore, thisDept, thisPlan, thisBuyer, thisClass As String
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, lastYear, thisWeek, tRows As Int16
    Public Shared rowIndex As Int16
    Public Shared rnd As Double
    Public Shared tNew As Decimal
    Public Shared theValue As String
    Public Shared thisUser = Environment.UserName
    Public Shared planBy As String
    Public Shared thisPeriod, periodSelected, numberPeriods As Integer
    Public Shared percentagesBy As String
    Public Shared periods(50) As String
    Public Shared todaysPeriod, totalPeriods As Integer
    Public Shared formLoaded As Boolean
    Public Shared somethingChanged As Boolean
    Public Shared initialLoad As Boolean

    Private Sub ClassPct_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            formLoaded = False
            serverLabel.Text = MainMenu.serverLabel.Text
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)

            con.Open()
            sql = "SELECT Prd_Id FROM Calendar WHERE '" & today & "' BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysPeriod = rdr("Prd_Id")
            End While
            con.Close()

            con.Open()
            sql = "SELECT Str_Id FROM Stores WHERE Status = 'Active' ORDER BY Str_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStr.Items.Add(rdr("Str_Id"))
            End While
            con.Close()
            cboStr.SelectedIndex = 0

            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Sales_Plan ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYr.Items.Add(rdr("Year_Id"))
            End While
            con.Close()
            Dim yr As Integer = DatePart(DateInterval.Year, Date.Today)
            cboYr.SelectedIndex = cboYr.FindString(yr)

            con.Open()
            sql = "SELECT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboDept.Items.Add(rdr("Dept"))
            End While
            con.Close()
            cboDept.SelectedIndex = 0

            con.Open()
            sql = "SELECT Buyer FROM Buyers WHERE Status = 'Active' ORDER BY Buyer"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboBuyer.Items.Add(rdr("Buyer"))
            End While
            con.Close()
            formLoaded = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD FORM")
        End Try
    End Sub

    Private Sub Load_Data()
        '                                                         Create the class table columns
        lblProcessing.Visible = True
        Me.Refresh()
        Try
            somethingChanged = False
            Dim clss As String
            Dim tAvg52 As Decimal = 0
            Dim tAvg26 As Decimal = 0
            Dim tAvg12 As Decimal = 0
            Dim tAvg4 As Decimal = 0
            Dim tCurrent As Decimal = 0
            Dim tInv As Decimal = 0
            Dim tLyYrPrd As Decimal = 0
            Dim firstRecord As Boolean = True
            Dim row As DataRow
            wkTbl = New DataTable
            wkTbl.Clear()
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Class"
            wkTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = wkTbl.Columns("Class")
            wkTbl.PrimaryKey = PrimaryKey
            wkTbl.Columns.Add("52 Week Average")
            wkTbl.Columns.Add("26 Week Average")
            wkTbl.Columns.Add("12 Week Average")
            wkTbl.Columns.Add("4 Week Average")
            wkTbl.Columns.Add("Current Plan%")
            wkTbl.Columns.Add("Enter New %")
            wkTbl.Columns.Add("LY Actual")
            wkTbl.Columns.Add("Current On Hand %")
            con.Open()
            '' ''sql = "SELECT DISTINCT Class FROM Class_PCT WHERE Plan_Year = '" & thisPlan & "' " & _
            '' ''        "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' " & _
            '' ''        "AND Buyer = '" & thisBuyer & "' ORDER BY Class"
            sql = "SELECT ID AS Class FROM Classes WHERE Dept = '" & thisDept & "' AND Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                ''pb1.PerformStep()
                clss = rdr("Class")
                row = wkTbl.NewRow
                row(0) = clss
                wkTbl.Rows.Add(row)
                con2.Open()

                sql = "DECLARE @wk52 date, @wk26 date, @wk12 date, @wk4 date, @lyYrPrd int " & _
                "SELECT @wk52 = DATEADD(week,-52,'" & today & "') " & _
                "SELECT @wk26 = DATEADD(week,-26,'" & today & "') " & _
                "SELECT @wk12 = DATEADD(week,-12,'" & today & "') " & _
                "SELECT @wk4 = DATEADD(week,-4,'" & today & "') " & _
                "SET @lyYrPrd = (SELECT YrPrd FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_id = " & thisPeriod & " AND Week_Id = 0) - 100 " & _
                "CREATE TABLE #t1 (Class varchar(20) NOT NULL, Avg52 decimal(18,4)) INSERT INTO #t1 (Class, Avg52) " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0)  AS Avg52 " & _
                    "FROM Sales_Summary WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "CREATE TABLE #t2 (Class varchar(20) NOT NULL, Avg26 decimal(18,4)) INSERT INTO #t2 (Class, Avg26) " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0) AS Avg26 " & _
                    "FROM Sales_Summary WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "CREATE TABLE #t3 (Class varchar(20) NOT NULL, Avg12 decimal(18,4)) INSERT INTO #t3 (Class, Avg12) " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0) AS Avg12 " & _
                    "FROM Sales_Summary WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "CREATE TABLE #t4 (Class varchar(20) NOT NULL, Avg4 decimal(18,4)) INSERT INTO #t4 (Class, Avg4) " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0) AS Avg4 " & _
                    "FROM Sales_Summary WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "CREATE TABLE #t5 (Class varchar(20) NOT NULL, pctInv decimal(18,4)) INSERT INTO #t5 (Class, pctInv) " & _
                 "SELECT Class, COALESCE(SUM(Act_Inv_Retail) / (SELECT NULLIF(SUM(Act_Inv_Retail),0) FROM Inv_Summary i " & _
                    "JOIN Stores s ON s.Inv_Loc = i.Loc_Id " & _
                        "WHERE CONVERT(VARCHAR,'" & today & "',101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "'),0) AS pctInv " & _
                        "FROM Inv_Summary w INNER JOIN Buyers b ON b.ID = w.Buyer JOIN Stores s ON s.Inv_Loc = w.Loc_Id " & _
                            "WHERE CONVERT(VARCHAR,'" & today & "',101) BETWEEN sDate AND eDate AND s.Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & thisDept & "' AND b.Status = 'Active' " & _
                            "GROUP BY Class " & _
                "CREATE TABLE #t6 (Class varchar(20) NOT NULL, lyYrPrd decimal(18,4)) INSERT INTO #t6 (Class, lyYrPrd) " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Sales_Summary " & _
                        "WHERE YrPrd = @lyYrPrd AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "' AND YrPrd = @lyYrPrd),0) AS lyYrPrd " & _
                        "FROM Sales_Summary WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "' AND Class = '" & clss & "' AND YrPrd = @lyYrPrd GROUP BY Class " & _
                "SELECT #t1.Class, #t1.Avg52, #t2.Avg26, #t3.Avg12, #t4.Avg4, " & _
                    "ISNULL((SELECT PCT FROM Class_PCT p WHERE Plan_Year = '" & thisYear & "' " & _
                        "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND p.Class = #t1.Class),0) AS Prd1, #t5.pctInv, #t6.lyYrPrd " & _
                    "FROM #t1 LEFT JOIN #t2 ON #t2.Class = #t1.Class " & _
                        "LEFT JOIN #t3 ON #t3.Class = #t1.Class " & _
                        "LEFT JOIN #t4 ON #t4.Class = #t1.Class " & _
                        "LEFT JOIN #t5 ON #t5.Class = #t1.Class " & _
                        "LEFT JOIN #t6 ON #t6.Class = #t1.Class ORDER BY Class"


                'pb1.Visible = False
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    oTest = rdr2(0)
                    '' ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    '' ''row = wkTbl.NewRow
                    '' ''oTest = rdr2(0)
                    '' ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row(0) = rdr2(0)
                    '' ''row(0) = clss
                    oTest = rdr2("Avg52")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("52 Week Average") = Format(rdr2("Avg52"), "###.0%")
                        tAvg52 += oTest
                    End If
                    oTest = rdr2("Avg26")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("26 Week Average") = Format(rdr2("Avg26"), "###.0%")
                        tAvg26 += oTest
                    End If
                    oTest = rdr2("Avg12")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("12 Week Average") = Format(rdr2("Avg12"), "###.0%")
                        tAvg12 += oTest
                    End If
                    oTest = rdr2("Avg4")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("4 Week Average") = Format(rdr2("Avg4"), "###.0%")
                        tAvg4 += oTest
                    End If
                    oTest = rdr2("Prd1")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("Current Plan%") = Format(rdr2("Prd1"), "###.0%")
                        tCurrent += oTest
                    End If
                    oTest = rdr2("pctInv")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(8) = Format(oTest, "###.0%")
                        tInv += rdr2("pctInv")
                    End If
                    oTest = rdr2("lyYrPrd")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row("LY Actual") = Format(oTest, "###.0%")
                        tLyYrPrd += CDec(oTest)
                    End If
                    '' ''wkTbl.Rows.Add(row)
                    '' ''End If
                End While
                con2.Close()
            End While
            con.Close()
            row = wkTbl.NewRow
            row(0) = "Total"
            row(1) = Format(Math.Round(tAvg52, 1, MidpointRounding.AwayFromZero), "###.#%")
            row(2) = Format(Math.Round(tAvg26, 1, MidpointRounding.AwayFromZero), "###.#%")
            row(3) = Format(Math.Round(tAvg12, 1, MidpointRounding.AwayFromZero), "###.#%")
            row(4) = Format(Math.Round(tAvg4, 1, MidpointRounding.AwayFromZero), "###.#%")
            row(5) = Format(Math.Round(tCurrent, 1, MidpointRounding.AwayFromZero), "###.#%")
            row(7) = Format(Math.Round(tLyYrPrd, 1, MidpointRounding.AwayFromZero), "###.#%")
            row(8) = Format(Math.Round(tInv, 1, MidpointRounding.AwayFromZero), "###.#%")
            wkTbl.Rows.Add(row)
            tRows = wkTbl.Rows.Count - 1
            row = wkTbl.NewRow
            row(0) = "Difference"
            wkTbl.Rows.Add(row)
            dgv1.DataSource = wkTbl.DefaultView
            Dim cnt As Integer = dgv1.RowCount
            If cnt > 1 Then
                dgv1.Rows(cnt - 1).ReadOnly = True
                dgv1.Rows(cnt - 2).ReadOnly = True
            End If
            dgv1.AutoResizeColumns()
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Format = "p2"
                If i = 5 Then dgv1.Columns(6).DefaultCellStyle.BackColor = Color.Cornsilk
                dgv1.Columns(i).ReadOnly = True
            Next
            dgv1.Columns(6).ReadOnly = False
            dgv1.Refresh()
            formLoaded = True
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD FORM")
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

    ''Public Sub dgv1_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        Dim change, totalChange As Decimal
        Dim row As DataRow = wkTbl(rowIndex)
        oTest = dgv1.Rows(rowIndex).Cells("Enter New %").Value
        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
            somethingChanged = True
            If thisYear <= DatePart(DateInterval.Year, Date.Today) And thisPeriod < todaysPeriod Then
                MessageBox.Show("PERCENTAGES CANNOT BE CHANGED FOR PRIOR PERIODS!", "ENTER NEW PLAN PERCENT")
                dgv1.Rows(rowIndex).Cells(columnIndex).Value = Nothing
                Exit Sub
            End If
            For Each row In wkTbl.Rows
                If row(0) <> "Total" And row(0) <> "Difference" Then
                    oTest = row(6)
                    If Not IsNothing(oTest) And Not IsDBNull(oTest) Then
                        oTest = Replace(oTest, "%", "")
                        change = CDec(oTest) * 0.01
                        totalChange += change
                        row(6) = Format(change, "###.0%")
                    End If

                End If
            Next
            Dim foundRow As DataRow = wkTbl.Rows.Find("Total")
            foundRow(6) = Format(totalChange, "###.0%")
            foundRow = wkTbl.Rows.Find("Difference")
            foundRow(6) = Format(1 - totalChange, "###.0%")
        End If
      
    End Sub

    Private Sub Save_Class_Plan()
        lblProcessing.Visible = True
        Me.Refresh()
        Dim row As DataRow = wkTbl.Rows.Find("Total")
        oTest = row(6)
        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
            oTest = Replace(row(6), "%", "")
            If oTest <> 100 Then
                MessageBox.Show("NEW PLAN PERCENT MUST EQUAL 100% BEFORE CHANGES CAN BE SAVED!", "SAVE CHANGES")
                lblProcessing.Visible = False
                Exit Sub
            End If
        End If
        Dim message, title As String
        Dim r As Integer = totalPeriods - thisPeriod
        If r > 6 Then r = 6
        Dim rx As Integer = r
        Dim theValue As Decimal

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
        Dim thisClass As String
        Dim lastYear As Int16 = thisYear - 1
        Dim rowCount As Integer = 0
        For Each row In wkTbl.Rows
            oTest = row(0)
            If oTest <> "Total" And oTest <> "Difference" Then
                oTest = row(6)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "%" Then
                        thisClass = row(0).ToString
                        theValue = CDec(Replace(oTest, "%", "") * 0.01)
                        con.Open()
                        For i As Integer = thisPeriod To endPeriod
                            sql = "IF NOT EXISTS (SELECT * FROM Class_PCT WHERE Plan_Year = '" & thisYear & "' " & _
                                    "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                    "AND BUYER = '" & thisBuyer & "' AND Class = '" & thisClass & "') " & _
                                "INSERT INTO Class_Pct (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Class, PCT) " & _
                                    "SELECT '" & thisYear & "', '" & thisStore & "', " & i & ", '" & thisDept & "', '" & _
                                    thisBuyer & "', '" & thisClass & "', " & theValue & " " & _
                                "ELSE " & _
                                "UPDATE Class_PCT SET PCT = " & theValue & " WHERE Plan_Year = '" & thisYear & "' " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                "AND Buyer = '" & thisBuyer & "' AND Class = '" & thisClass & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                            '
                            '               Update same record for future years if any exists
                            '
                            sql = "UPDATE Class_PCT SET PCT = " & theValue & " WHERE Plan_Year > " & thisYear & " " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                "AND Buyer = '" & thisBuyer & "' AND Class = '" & thisClass & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                        Next
                        con.Close()
                        dgv1.Rows(rowCount).Cells(6).Value = Nothing
                        row(5) = Format(theValue, "###.0%")
                        dgv1.Item(5, rowCount).Style.BackColor = Color.Yellow
                        dgv1.Item(5, rowCount).Style.Font = New Font("Sans Serif", 8.25, FontStyle.Bold)
                    End If
                End If
            End If
            rowCount += 1
        Next
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Class_Plan()
    End Sub

    Private Sub cboStr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStr.SelectedIndexChanged
        Try
            thisStore = cboStr.SelectedItem
            If formLoaded Then Call Load_Data()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SELECT STORE")
        End Try
    End Sub

    Private Sub cboYr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYr.SelectedIndexChanged
        Try
            thisYear = cboYr.SelectedItem
            cboPrd.Items.Clear()
            totalPeriods = 0
            con.Open()
            sql = "SELECT DISTINCT Prd_Id FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboPrd.Items.Add(rdr("Prd_Id"))
                totalPeriods += 1
            End While
            con.Close()
            cboPrd.SelectedIndex = 0
            If formLoaded Then Call Load_Data()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SELECT YEAR")
        End Try
    End Sub

    Private Sub cboPrd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPrd.SelectedIndexChanged
        thisPeriod = cboPrd.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboBuyer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboBuyer.SelectedIndexChanged
        thisBuyer = cboBuyer.SelectedItem
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
                            Call Save_Class_Plan()
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