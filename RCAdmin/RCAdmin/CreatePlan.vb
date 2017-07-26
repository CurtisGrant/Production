Imports System.Data.SqlClient
Imports System.Text
Public Class CreatePlan
    Public Shared planName As String = Nothing
    Public Shared conString, conString2, server, database, sql As String
    Public Shared con, con2, con3, con4, con5, con6 As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Private Shared tblPeriod, dateTbl, buyerTbl As DataTable
    Private Shared oTest As Object
    Private Shared thisDept, thisPlan, thisStore, thisUser As String
    Private Shared lastYear, thisYear, thisPeriod, thisWeek, rnd, yrwk As Integer
    Private Shared loDate, hiDate, thisEdate, lyBeginDate, lyEndDate As Date
    Private Shared today As Date = Date.Today
    Private Shared formLoad As Boolean = False
    Private firstPlan As Boolean = True

    Private Sub CreateParams_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            conString = MainMenu.conString
            conString2 = MainMenu.conString
            con = PlanMaintenance.con
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con4 = New SqlConnection(conString)
            con5 = New SqlConnection(conString)
            con6 = New SqlConnection(conString)
            For Each item In PlanMaintenance.cboStore.Items
                Me.cboStore.Items.Add(item)
            Next
            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Calendar ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYear.Items.Add(rdr(0))
            End While
            con.Close()

            con.Open()
            sql = "SELECT Str_Id FROM Stores WHERE Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStore.Items.Add(rdr("Str_Id"))
            End While
            con.Close()

            con.Open()
            sql = "SELECT TOP 1 Plan_id FROM Sales_Plan"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                firstPlan = False
            End While
            con.Close()

            cboStore.SelectedIndex = 0
            Dim year As Integer = Date.Today.Year
            cboYear.SelectedIndex = cboYear.FindString(year)
            thisUser = Environment.UserName
            thisPlan = Nothing
            cboRnd.Items.Add(1)
            cboRnd.Items.Add(10)
            cboRnd.Items.Add(50)
            cboRnd.Items.Add(100)
            cboRnd.Items.Add(500)
            cboRnd.Items.Add(1000)
            cboRnd.SelectedIndex = 3
            rnd = CDbl(cboRnd.SelectedItem)

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            MessageBox.Show(ex.Message, "CREATE NEW PLAN")
        End Try

    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem.ToString
        If formLoad Then Call Create_New_Plan(planName)
    End Sub

    Private Sub cboYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYear.SelectedIndexChanged
        thisYear = cboYear.SelectedItem
        If formLoad Then Call Create_New_Plan(planName)
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Try
            lblProcessing.Visible = True
            Me.Refresh()

            Dim tablesOK As Boolean = True = True
            Call Verify_Sales_Plan_Tables(tablesOK)
            If Not tablesOK Then
                MessageBox.Show("Use RCSetp to create database tables necessary for sales planning.", "ERROR! TABLE(s) MISSING!")
                Exit Sub
            End If
            
            con.Open()
            cmd = New SqlCommand("sp_RCSetup_CreateFirstSalesPlan", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
            cmd.Parameters.Add("@thisYear", SqlDbType.Int).Value = thisYear
            cmd.ExecuteNonQuery()

            cmd = New SqlCommand("sp_OptionalModule_ProjectedWksOH", con)
            cmd.ExecuteNonQuery()
            con.Close()

            ''Dim itExists As Boolean = False
            ''Dim cDte, dDte As Date
            ''con.Open()
            ''sql = "DECLARE @maxDataDate AS Date, @maxCalendarDate AS Date, @planDataStartDate AS Date " & _
            ''    "SET @maxDataDate = (SELECT MAX(eDate) FROM Item_Sales WHERE ISNULL(Act_Sales,0) <> 0 AND Year_Id = " & thisYear - 1 & ") " & _
            ''    "SET @maxCalendarDate = (SELECT MAX(eDate) FROM Calendar WHERE Year_Id = " & thisYear - 1 & ") " & _
            ''    "SET @planDataStartDate = (SELECT MIN(sDate) FROM Calendar WHERE Year_Id = " & thisYear - 1 & " AND Prd_Id = 0) " & _
            ''    "SELECT @planDataStartDate AS planDataStartDate, @maxDataDate AS dataDate, @maxCalendarDate AS calendarDate"
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''While rdr.Read
            ''    dDte = rdr("dataDate")
            ''    cDte = rdr("calendarDate")
            ''    lyBeginDate = rdr("planDataStartDate")
            ''End While
            ''con.Close()
            ''If cDte <> dDte Then
            ''    con.Open()
            ''    sql = "DECLARE @period AS integer, @planDataEndDate AS Date " & _
            ''        "SET @period = (SELECT MAX(YrPrd) FROM Item_Sales WHERE ISNULL(Act_Sales,0) <> 0 AND Prd_Id > 0) " & _
            ''        "SELECT MAX(eDate) AS eDate FROM Calendar WHERE YrPrd < @period AND Prd_Id > 0"
            ''    cmd = New SqlCommand(sql, con)
            ''    rdr = cmd.ExecuteReader
            ''    While rdr.Read
            ''        hiDate = rdr("eDate")
            ''        lyEndDate = rdr("eDate")
            ''    End While
            ''    con.Close()
            ''Else
            ''    hiDate = cDte
            ''    lyEndDate = cDte
            ''End If


            ''con.Open()
            ''sql = "SELECT Year_Id, Prd_Id, PrdWk, YrWk FROM Calendar WHERE eDate = '" & hiDate & "'"
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''While rdr.Read
            ''    yrwk = rdr("YrWk")
            ''    thisPeriod = rdr("Prd_Id")
            ''End While
            ''con.Close()

            ''con.Open()
            ''sql = "SELECT eDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id = 0"
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''While rdr.Read
            ''    lyEndDate = rdr("eDate")
            ''End While
            ''con.Close()

            ''con.Open()
            ''sql = "SELECT MIN(eDate) AS eDate FROM Calendar WHERE Year_Id = " & thisYear - 1 & " "
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''While rdr.Read
            ''    loDate = rdr("eDate")
            ''End While
            ''con.Close()



            ''con.Open()
            ''sql = "SELECT DISTINCT Plan_Id, Year_Id FROM Sales_Plan WHERE Year_Id = " & thisYear & " AND Status = 'Active' " & _
            ''    "AND Str_Id = '" & thisStore & "'"
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''itExists = False
            ''While rdr.Read
            ''    If Not IsNothing("Plan_Id") Then itExists = True
            ''    planName = rdr("Plan_Id")
            ''End While
            ''con.Close()
            ''If itExists Then
            ''    '' MsgBox("AN ACTIVE PLAN ALREADYS EXISTS!")
            ''    MessageBox.Show("AN ACTIVE PLAN ALREADY EXISTS!", "CREATE NEW PLAN")
            ''    Exit Sub
            ''End If

            ''Call Create_New_Plan(planName)

            ''PlanMaintenance.cboPlan.Items.Add(thisPlan)
            ''Dim idx As Integer = PlanMaintenance.cboPlan.FindString(thisPlan)
            ''PlanMaintenance.cboPlan.SelectedIndex = idx
            ''lblProcessing.Visible = False
            ''Me.Refresh()
            PlanMaintenance.Show()
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR CREATING NEW PLAN")
        End Try
    End Sub

    Private Sub Verify_Sales_Plan_Tables(ByVal tablesOK As Boolean)
        tablesOK = True
        Dim val As String = ""
        con.Open()
        sql = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'Buyer_PCT') " & _
            "BEGIN " & _
            "SELECT 'YES' AS ANS " & _
            "END"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If rdr("ANS") = "YES" Then val = "YES"
        End While
        con.Close()
        If val = "" Then
            tablesOK = False
            Exit Sub
        End If
        con.Open()
        sql = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'Class_PCT') " & _
            "BEGIN " & _
            "SELECT 'YES' AS ANS " & _
            "END"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If rdr("ANS") = "YES" Then val = "YES"
        End While
        con.Close()
        If val = "" Then
            tablesOK = False
            Exit Sub
        End If
        con.Open()
        sql = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'DayOfWeekPct') " & _
            "BEGIN " & _
            "SELECT 'YES' AS ANS " & _
            "END"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If rdr("ANS") = "YES" Then val = "YES"
        End While
        con.Close()
        If val = "" Then
            tablesOK = False
            Exit Sub
        End If
        con.Open()
        sql = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'Day_Sales_Plan') " & _
            "BEGIN " & _
            "SELECT 'YES' AS ANS " & _
            "END"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If rdr("ANS") = "YES" Then val = "YES"
        End While
        con.Close()
        If val = "" Then
            tablesOK = False
            Exit Sub
        End If
        con.Open()
        sql = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'Sales_Plan') " & _
            "BEGIN " & _
            "SELECT 'YES' AS ANS " & _
            "END"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If rdr("ANS") = "YES" Then val = "YES"
        End While
        con.Close()
        If val = "" Then
            tablesOK = False
            Exit Sub
        End If
    End Sub

    Private Sub Create_New_Plan(ByVal planName)
        Try
            con.Close()
            Dim start As Integer
            Dim numbr As Integer = 0
            Dim test As String
            con.Open()
            sql = "SELECT MAX(Plan_Id) FROM Sales_Plan WHERE Plan_Id LIKE '" & thisYear & "_Plan%'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                If Not IsDBNull(rdr(0)) And Not IsNothing(rdr(0)) Then
                    test = rdr(0)
                    start = InStr(test, "Plan")
                    Dim final As String = Mid(test, start)
                    oTest = getNumeric(final)
                    If IsNumeric(oTest) Then numbr = CInt(oTest)
                End If
            End While
            con.Close()

            numbr += 1
            thisPlan = thisYear.ToString & "_Plan" & numbr.ToString
            Dim tblDept As DataTable = New DataTable
            tblPeriod = New DataTable
            tblDept.Columns.Add("Dept")
            tblPeriod.Columns.Add("Period")
            Dim row As DataRow
            con.Open()
            sql = "SELECT ID FROM Departments WHERE Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = tblDept.NewRow
                row(0) = rdr(0)
                tblDept.Rows.Add(row)
            End While
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Prd_Id FROM Calendar WHERE Prd_Id > 0 AND Year_Id = " & thisYear & " " & _
                 "ORDER BY Prd_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = tblPeriod.NewRow
                row("Period") = rdr("Prd_Id")
                tblPeriod.Rows.Add(row)
            End While
            con.Close()



            lblSaving.Visible = True
            pb1.Visible = True
            pb1.Step = 1
            pb1.Value = 1
            pb1.Minimum = 1
            pb1.Maximum = 13
            thisUser = "RCAdmin"

            Call Create_Other_Tables()                         ' Update Class_PCT and DayofWeekPCT tables

            For Each dRow In tblDept.Rows
                thisDept = dRow(0)
                lblProcessing.Text = "Updating " & thisDept
                Me.Refresh()



                ''             Call Compute_Weeks_OnHand()




                For Each pRow In tblPeriod.Rows
                    thisPeriod = pRow(0)
                    con.Open()
                    sql = "SELECT eDate, PrdWK FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = " & pRow(0) & " "
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        thisEdate = rdr("eDate")
                        thisWeek = rdr("PrdWk")
                        lastYear = thisYear - 1
                        ''con2.Open()
                        ''thisStore = cboStore.SelectedItem
                        ''sql = "If NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" &
                        ''        thisStore & "' AND Year_Id = " & thisYear & " AND Prd_Id = " &
                        ''            thisPeriod & " AND Week_Id = " & thisWeek & " AND Dept = '" & thisDept & "') " & _
                        ''        "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Status, Last_Update, Last_User) " & _
                        ''        "SELECT '" & thisPlan & "', '" & thisStore & "', '" & thisDept & "', " & thisYear & ", " &
                        ''            thisPeriod & ", " & thisWeek & ", 'Active','" & today & "','Initial Build'"
                        ''cmd = New SqlCommand(sql, con2)
                        ''cmd.ExecuteNonQuery()
                        ''con2.Close()
                        If thisEdate <= lyEndDate Then
                            Call Save_This_Week()                          ' skip computing actuals for current week and forward
                        End If
                    End While
                    con.Close()
                Next
                pb1.PerformStep()
                Call Save_This_Department()
            Next
            pb1.Visible = False
            formLoad = True
            lblProcessing.Text = Nothing
            Me.Refresh()
            MessageBox.Show(planName & " CREATED", "CREATE NEW PLAN")
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Create_Other_Tables()
        Try
            dateTbl = New DataTable
            Dim column As New DataColumn()
            column.DataType = System.Type.GetType("System.DateTime")
            column.ColumnName = "eDate"
            dateTbl.Columns.Add(column)
            Dim primaryKey(1) As DataColumn
            primaryKey(0) = dateTbl.Columns("eDate")
            dateTbl.PrimaryKey = primaryKey
            dateTbl.Columns.Add("Year", GetType(System.Int16))
            dateTbl.Columns.Add("Period", GetType(System.Int16))
            buyerTbl = New DataTable
            buyerTbl.Columns.Add("Buyer")
            Dim row As DataRow
            lastYear = thisYear - 1
            '
            '                                   Get eDates for last year's periods
            '
            con.Open()
            If firstPlan Then
                sql = "DECLARE @startDate date, @endDate date " & _
                    "SET @startDate = (SELECT MIN(eDate) FROM Item_Sales) " & _
                    "SET @endDate = (SELECT MAX(eDate) FROM Calendar " & _
                        "WHERE eDate < (SELECT eDate FROM Calendar " & _
                        "WHERE CONVERT(Date,GetDate()) BETWEEN sDate AND eDate AND Week_Id > 0)) " & _
                    "SELECT eDate, Year_Id, Prd_Id FROM Calendar WHERE Week_Id = 0 AND Prd_Id > 0 " & _
                        "AND eDate BETWEEN @startDate AND @endDate"
            Else
                sql = "SELECT eDate, Year_Id, Prd_Id  FROM Calendar WHERE Week_Id = 0 AND Prd_Id > 0 " & _
                    "AND Year_Id = " & lastYear & " "
            End If
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = dateTbl.NewRow
                row("eDate") = rdr("eDate")
                row("Year") = rdr("Year_Id")
                row("Period") = rdr("Prd_Id")
                dateTbl.Rows.Add(row)
            End While
            con.Close()

            If dateTbl.Rows.Count < 1 Then
                MessageBox.Show("Could not determine dates for past year. Update Calendar and try again.", "ERROR")
                Exit Sub
            End If

            '
            '                Get first and last date from datetbl
            '
            Dim startDate, endDate As Date
            startDate = dateTbl.Rows(0)("eDate")
            endDate = dateTbl.Rows(dateTbl.Rows.Count - 1)("eDate")

            con.Open()
            sql = "SELECT Buyer FROM Buyers WHERE Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = buyerTbl.NewRow
                row("Buyer") = rdr("Buyer")
                buyerTbl.Rows.Add(row)
            End While
            con.Close()
            '
            '                           Determine each buyer's percentage of total sales and store it in Buyer_PCT
            '
            lblProcessing.Text = "Updating Buyer_PCT"
            Me.Refresh()
            con.Open()
            cmd = New SqlCommand("sp_RCAdmin_CreatePlan_UpdateBuyerPCT", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@year", SqlDbType.Int).Value = lastYear
            cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
            cmd.ExecuteNonQuery()
            ''sql = "CREATE TABLE #t1 (Str_Id varchar(10) NOT NULL, Year_Id Integer, Prd_Id integer, Dept varchar(20), " & _
            ''    "Buyer varchar(20), Sales decimal(18,4), Pct decimal(18,4), WksOH int) " & _
            ''    "INSERT INTO #t1 (Str_Id, Year_Id, Prd_Id, Dept, Buyer, Sales, Pct, WksOH) " & _
            ''    "SELECT Str_Id, Year_Id, Prd_Id, Dept, CASE WHEN w.Buyer IS NULL OR w.Buyer = '' THEN 'OTHER' Else w.Buyer END As Buyer, " & _
            ''    "SUM(ISNULL(Act_Sales,0)) AS Sales, 0, 0 FROM Item_Sales w " & _
            ''    "LEFT JOIN Buyers b ON b.ID = w.Buyer AND b.Status = 'Active' " & _
            ''    "WHERE Str_Id = '" & thisStore & "' AND eDate BETWEEN '" & startDate & "' AND '" & endDate & "' " & _
            ''    "AND Act_Sales <> 0 " & _
            ''    "GROUP BY Str_Id, Year_Id, Prd_Id, Dept, w.Buyer, b.ID " & _
            ''    "CREATE TABLE #t2 (Str_Id varchar(10) NOT NULL, Year_Id integer, Prd_Id integer, Dept varchar(20), sales decimal(18,4)) " & _
            ''    "INSERT INTO #t2 (Str_Id, Year_Id, Prd_Id, Dept, sales) " & _
            ''    "SELECT Str_Id, Year_Id, Prd_Id, Dept, ISNULL(SUM(Sales),0) AS Sales FROM #t1 " & _
            ''    "GROUP BY Str_Id, Year_Id, Prd_Id, Dept " & _
            ''    "UPDATE #t1 SET #t1.PCT = CASE WHEN #t2.Sales > 0 THEN #t1.Sales / #t2.Sales ELSE 0 END FROM #t1 " & _
            ''    "JOIN #t2 ON #t2.Str_Id = #t1.Str_Id AND #t2.Year_Id = #t1.Year_Id " & _
            ''    "AND #t2.Prd_Id = #t1.Prd_Id AND #t2.Dept = #t1.Dept " & _
            ''    "MERGE Buyer_PCT target USING #t1 source ON target.Str_Id = source.Str_Id AND target.Plan_Year = source.Year_Id " & _
            ''        "AND target.Prd_Id = source.Prd_Id AND target.Dept = source.Dept AND target.Buyer = source.buyer " & _
            ''    "WHEN NOT MATCHED BY TARGET THEN " & _
            ''        "INSERT(Str_Id, Plan_Year, Prd_Id, Dept, buyer, Act_PCT, PCT) " & _
            ''        "VALUES(Str_Id, Year_Id, Prd_Id, Dept, buyer, Pct, Pct) " & _
            ''    "WHEN MATCHED THEN " & _
            ''    "UPDATE SET target.Act_PCT = source.Pct;"
            ''cmd = New SqlCommand(sql, con)
            ''cmd.CommandTimeout = 480
            ''cmd.ExecuteNonQuery()
            ''sql = "DECLARE @lastPeriod int, @maxPeriod int, @deptCnt int, @NUMDEPT int, @thisYear int, @lastYear int " & _
            ''    " @deptTbl TABLE(ID int NOT NULL Identity(1,1), dept varchar(10)) " & _
            ''    "DECLARE @dept varchar(10) " & _
            ''    "SET @thisYear = DATEPART(Year,GETDATE()) " & _
            ''    "SET @lastYear = @thisYear - 1 " & _
            ''    "SET @lastPeriod = (SELECT MAX(Prd_Id) FROM Buyer_PCT WHERE Plan_Year = 2016) " & _
            ''    "SET @maxPeriod = (SELECT MAX(Prd_Id) FROM Calendar WHERE Year_Id = 2016) " & _
            ''    "INSERT @deptTbl(dept) " & _
            ''    "SELECT DISTINCT Dept FROM Buyer_Pct WHERE Year_Id = 2016 AND Prd_Id = @lastPeriod " & _
            ''    "SET @NUMDEPT = @@ROWCOUNT " & _
            ''    "SET @deptCnt = @NUMDEPT " & _
            ''    "WHILE @deptCnt > 0 " & _
            ''    "BEGIN " & _
            ''        "SELECT @dept = dept FROM @depttbl WHERE ID = @deptcnt " & _
            ''        "INSERT INTO Buyer_PCT(Plan_Year, Str_Id, Prd_Id, Dept, Buyer, PCT) " & _
            ''        "SELECT @thisYear, Str_Id, Prd_Id, Dept, Buyer, Act_PCT FROM Buyer_Pct " & _
            ''        "WHERE Plan_Year = @lastYear AND Dept = @dept AND Prd_Id > @lastPeriod AND Prd_Id <= @maxPeriod " & _
            ''        "SET @deptCnt = @deptCnt - 1 " & _
            ''    "END"
            ''cmd = New SqlCommand(sql, con)
            ''cmd.ExecuteNonQuery()
            con.Close()

            lblProcessing.Text = "Updating Class_PCT"
            Me.Refresh()
            con.Open()
            cmd = New SqlCommand("sp_RCAdmin_CreatePlan_UpdateClassPCT", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@year", SqlDbType.Int).Value = lastYear
            cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
            cmd.ExecuteNonQuery()
            ''sql = "CREATE TABLE #t1 (Str_Id varchar(10) NOT NULL, Year_Id integer, Prd_Id integer, Dept varchar(20), Buyer varchar(20), " & _
            ''    "Class varchar(20), Sales decimal(18,4), Pct decimal(18,4)) " & _
            ''    "INSERT INTO #t1 (Str_Id, Year_Id, Prd_Id, Dept, Buyer, Class, Sales) " & _
            ''    "SELECT Str_Id, Year_Id, Prd_Id, Dept, CASE WHEN Buyer IS NULL OR Buyer = '' THEN 'OTHER' ELSE Buyer END AS Buyer, " & _
            ''    "Class, ISNULL(SUM(Act_Sales),0) AS Sales FROM Item_Sales WHERE Str_Id = '1' " & _
            ''    "AND eDate BETWEEN '1/28/2012' AND '7/9/2016' AND Act_Sales <> 0 " & _
            ''    "GROUP BY Str_Id, Year_Id, Prd_Id, Dept, Buyer, Class " & _
            ''    "CREATE TABLE #t2 (Str_Id varchar(10) NOT NULL, Year_Id integer, Prd_Id integer, " & _
            ''    "Dept varchar(20), Buyer varchar(20), Sales decimal(18,4)) " & _
            ''    "INSERT INTO #t2 (Str_Id, Year_Id, Prd_Id, Dept, Buyer, Sales) " & _
            ''    "SELECT Str_Id, Year_Id, Prd_Id, Dept, Buyer, SUM(Sales) As Sales FROM #t1 WHERE Sales <> 0 " & _
            ''    "GROUP BY Str_Id, Year_Id, Prd_Id, Dept, Buyer " & _
            ''    "UPDATE #t1 SET #t1.Pct = #t1.Sales / #t2.Sales FROM #t1 JOIN #t2 ON #t2.Str_Id = #t1.Str_Id " & _
            ''    "AND #t2.Year_Id = #t1.Year_Id " & _
            ''    "AND #t2.Prd_Id = #t1.Prd_Id AND #t2.Dept = #t1.Dept  AND #t2.Buyer = #t1.Buyer " & _
            ''    "MERGE Class_PCT target USING #t1 source ON target.Str_Id = source.Str_Id AND target.Year_Id = source.Year_Id " & _
            ''    "AND target.Prd_Id = source.Prd_Id AND target.Dept = source.Dept AND target.Buyer = source.Buyer " & _
            ''    "AND target.Class = source.class " & _
            ''    "WHEN NOT MATCHED BY TARGET THEN INSERT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Class, Act_Pct, PCT) " & _
            ''    "VALUES (Year_Id, Str_Id, Prd_Id, Dept, Buyer, Class, PCT, PCT) " & _
            ''    "WHEN MATCHED THEN UPDATE SET target.PCT = source.PCT;"
            ''cmd = New SqlCommand(sql, con)
            ''cmd.CommandTimeout = 480
            ''cmd.ExecuteNonQuery()
            con.Close()

            lblProcessing.Text = "Updating DayOfWeekPCT"
            Me.Refresh()
            con.Open()
            cmd = New SqlCommand("sp_RCAdmin_CreatePlan_UpdateDayOfWeekPCT", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@year", SqlDbType.Int).Value = lastYear
            cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
            ''sql = "CREATE TABLE #t1 (Year_Id integer, Str_Id varchar(10), Prd_Id integer NOT NULL, Day integer, " & _
            ''   "Sales decimal(18,4), PCT decimal(18,4)) " & _
            ''   "INSERT INTO #t1 (Year_Id, Str_Id, Prd_Id, Day, Sales, PCT) " & _
            ''   "SELECT Year_Id, LOCATION, Prd_Id, DATEPART(dw,TRANS_DATE) AS Day, ISNULL(SUM(QTY * RETAIL),0) AS Sales, 0 " & _
            ''       "FROM Daily_Transaction_Log l " & _
            ''       "JOIN Calendar c ON TRANS_DATE BETWEEN sDate AND eDate AND PrdWk > 0 " & _
            ''       "WHERE QTY <> 0 AND RETAIL > 0 AND LOCATION = '" & thisStore & "' " & _
            ''       "AND TYPE = 'Sold' AND DATEPART(dw,TRANS_DATE) > 1 " & _
            ''       "AND TRANS_DATE BETWEEN '" & startDate & "' AND '" & endDate & "' " & _
            ''       "GROUP BY Year_Id, LOCATION, Prd_Id, DATEPART(dw,TRANS_DATE) " & _
            ''   "CREATE TABLE #t2 (Year_Id integer, Str_Id varchar(10), Prd_Id integer NOT NULL, Sales decimal(18,4)) " & _
            ''   "INSERT INTO #t2 (Year_Id, Str_Id, Prd_Id, Sales) " & _
            ''   "SELECT Year_Id, Str_Id, Prd_Id, SUM(Sales) AS Sales FROM #t1 " & _
            ''   "GROUP BY Year_Id, Str_Id, Prd_Id " & _
            ''   "UPDATE #t1 SET #t1.PCT = #t1.Sales / #t2.Sales FROM #t1 " & _
            ''   "JOIN #t2 ON #t2.Year_Id = #t1.Year_Id AND #t2.Str_Id = #t1.Str_Id AND #t2.Prd_Id = #t1.Prd_Id " & _
            ''   "MERGE DayOfWeekPct target USING #t1 source ON source.Year_Id = target.Year_Id " & _
            ''        "AND source.Str_Id = target.Str_Id AND source.Prd_Id = target.Prd_Id AND source.Day = target.Day " & _
            ''    "WHEN NOT MATCHED BY TARGET THEN " & _
            ''        "INSERT(Year_Id, Str_Id, Prd_Id, Day, PCT) " & _
            ''        "VALUES(Year_Id, Str_Id, Prd_Id, Day, PCT) " & _
            ''    "WHEN MATCHED THEN " & _
            ''        "UPDATE SET target.PCT = source.PCT; "
            ''cmd = New SqlCommand(sql, con)
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "ERROR CREATING OTHER TABLES" & sql)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Compute_Weeks_OnHand()
        '
        '                           Compute inventory actual weeks on hand last year and save it as Plan_Sales for this year
        '
        lblProcessing.Text = "Updating Week On Hand"
        Me.Refresh()
        con2.Open()
        Dim thisBuyer As String
        Dim onHand As Decimal
        Dim totl As Decimal
        Dim cnt, wks As Integer
        Dim selectDate As Date
        For Each row In dateTbl.Rows
            selectDate = row("eDate")
            thisPeriod = row("Period")
            lastYear = row("Year")
            For Each brow In buyerTbl.Rows
                thisBuyer = brow("Buyer")
                selectDate = row("eDate")
                onHand = 0
                totl = 0
                cnt = 0
                con.Open()
                sql = "SELECT ISNULL(SUM(End_OH * Retail),0) As Inv FROM Item_Inv i " & _
                    "JOIN Item_Master m ON m.Sku = i.Sku " & _
                    "WHERE eDate = '" & selectDate & "' " & _
                    "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "'"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    onHand = rdr("Inv")
                End While
                con.Close()

                con.Open()
                sql = "SELECT eDate, ISNULL(SUM(Sales_Retail),0) As Sales FROM Item_Sales d " & _
                   "JOIN Item_Master m ON m.Sku = d.Sku " & _
                   "WHERE Str_Id = '1' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                   "AND eDate >= '" & selectDate & "' " & _
                   "GROUP BY eDate ORDER By eDate"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    cnt += 1
                    totl += rdr("Sales")
                    If totl >= onHand Then
                        wks = cnt
                        Exit While
                    End If
                End While
                con.Close()
                '
                '                    Update or create a new record for last year's actual weeks on hand
                '
                If totl > onHand Then
                    con.Open()              ' Update or create a new records for last year's actual weeks on hand
                    sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "' AND Plan_Year = " & lastYear & " AND Prd_Id = " & thisPeriod & ") " & _
                        "INSERT INTO Buyer_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Act_WksOH) " & _
                        "SELECT " & lastYear & ",'" & thisStore & "'," & thisPeriod & ",'" & thisDept & "','" & thisBuyer & "'," & wks & " " & _
                        "ELSE " & _
                        "UPDATE Buyer_Pct SET Act_WksOH = " & wks & " WHERE Plan_Year= " & lastYear & " " & _
                        "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                    con.Close()
                    '
                    '               Update or create a new record for this year's plan weeks on hand
                    con.Open()
                    sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "' AND Plan_Year = " & thisYear & " AND Prd_Id = " & thisPeriod & ") " & _
                        "INSERT INTO Buyer_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Plan_WksOH) " & _
                        "SELECT " & thisYear & ",'" & thisStore & "'," & thisPeriod & ",'" & thisDept & "','" & thisBuyer & "'," & wks & " " & _
                        "ELSE " & _
                        "UPDATE Buyer_Pct SET Plan_WksOH = " & wks & " WHERE Plan_Year= " & thisYear & " " & _
                        "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                    con.Close()
                End If
            Next
        Next
        con2.Close()

    End Sub
    Private Sub Save_This_Week()
        Try
            lastYear = thisYear - 1
            con2.Open()
            sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint,@amt decimal " & _
                "DECLARE @rnd AS integer, @mpct decimal (18,4), @msales decimal(18,4), @mallsales decimal (18,4), @mamt decimal " & _
                "SELECT @rnd = " & rnd & " " & _
            "SET @sales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & ") " & _
            "SET @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                "AND Prd_Id = " & thisPeriod & ") " & _
            "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
            "SET @amt = CONVERT(int,(@pct * @allsales)) " & _
            "SET @msales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Item_Sales " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & ") " & _
            "SET @mallsales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Item_Sales " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                "AND Prd_Id = " & thisPeriod & ") " & _
            "SET @mpct = (SELECT COALESCE(@msales / NULLIF((@sales + @msales),0),0) WHERE @msales <> 0) " & _
            "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
            "AND Dept = '" & thisDept & "' AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & ") " & _
            "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User, Status, " & _
            "Mkdn_Pct) " & _
            "SELECT '" & thisPlan & "','" & thisStore & "','" & thisDept & "'," & thisYear & "," & thisPeriod & "," & _
                thisWeek & ",@amt,@pct,'" & today & "','" & thisUser & "','Active',@mpct " & _
            "ELSE " & _
            "UPDATE Sales_Plan SET Amt = ROUND(@amt / @rnd,0) * @rnd, Pct = @pct, Last_Update = '" & today & "', " & _
                "Mkdn_Pct = @mpct, Last_User = '" & thisUser & "' " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                "AND Year_Id = " & thisYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Plan_id = '" & thisPlan & "' "
            cmd = New SqlCommand(sql, con2)
            cmd.ExecuteNonQuery()
            con2.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR SAVING WEEK " & sql)
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Save_This_Period(ByVal prd)
        Try
            con.Open()
            sql = "DECLARE @pct decimal(18,4), @sales decimal(18,4), @allsales decimal(18,4) " & _
                "DECLARE @mpct decimal (18,4), @msales decimal(18,4), @mallsales decimal (18,4) " & _
                "SELECT @sales = (SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Plan_Id = '" &
                        thisPlan & "' AND Prd_Id = " & prd & ") " & _
                "SELECT @allsales = (SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Plan_Id = '" & thisPlan & "' AND Prd_Id = " & prd & ") " & _
                "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
                "SELECT @msales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Item_Sales " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " &
                        lastYear & " AND Prd_Id = " & prd & ") " & _
                "SELECT @mallsales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Item_Sales " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Dept = '" & thisDept & "') " & _
                "SELECT @mpct = COALESCE(@msales / (NULLIF(@sales,0) + NULLIF(@msales,0)),0) WHERE @msales <> 0 " & _
                "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
                    "AND Year_Id = " & thisYear & " AND Prd_Id = " & prd & " AND Week_Id = 0 " & _
                    "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "') " & _
                "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Mkdn_Pct, Last_Update, Last_User, Status) " & _
                "SELECT '" & thisPlan & "'," & thisYear & ",'" & thisStore & "','" & thisDept & "'," & prd & ",0,@sales," & _
                    "@pct,@mpct,'" & today & "', '" & thisUser & "','Active' " & _
                "ELSE " & _
                "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Mkdn_Pct = @mpct, Last_Update = '" & today & "', " & _
                    "Last_User = '" & thisUser & "' " & _
                    "WHERE Year_Id = " & thisYear & " AND Prd_Id = " & prd & " AND Week_Id = 0 " & _
                "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "' "
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            '' MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "ERROR SAVING PERIOD")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Save_This_Department()
        Try
            con.Open()                                               ' Update dept as pct of store
            sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint, @amt bigint " & _
                "SELECT @sales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales WHERE Str_Id = '1' AND Dept = 'AC' " & _
                    "AND Year_Id = " & thisYear & ") " & _
                "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales WHERE Str_Id = '1' AND Year_Id = " & thisYear & ") " & _
                "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
                "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
                        "AND Year_Id = " & thisYear & " AND Prd_Id = 0 AND Week_Id = 0 " & _
                        "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "') " & _
                     "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User, Status) " & _
                        "SELECT '" & thisPlan & "'," & thisYear & ",'" & thisStore & "','" & thisDept & "',0,0,@sales,@pct,'" &
                            today & "', '" & thisUser & "','Active' " & _
                    "ELSE " & _
                     "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Last_Update = '" & today & "', Last_User = '" & thisUser & "' " & _
                        "WHERE Year_Id = " & thisYear & " AND Prd_Id = 0 AND Week_Id = 0 " & _
                        "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
            For Each prdrow In tblPeriod.Rows                              ' Save Period Records
                thisPeriod = prdrow("Period")
                Call Save_This_Period(thisPeriod)
            Next
        Catch ex As Exception
            '' MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "ERROR SAVING DEPARTMENT")
        End Try
    End Sub

    Public Function getNumeric(value As String) As String
        Dim output As StringBuilder = New StringBuilder
        For i = 0 To value.Length - 1
            If IsNumeric(value(i)) Then
                output.Append(value(i))
            End If
        Next
        Return output.ToString()
    End Function
End Class