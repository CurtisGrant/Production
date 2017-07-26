Imports System.Data.SqlClient
Imports System.Xml
Imports System.Text
Imports Microsoft.VisualBasic

Public Class PlanMaintenance
    Public Shared conString, server, database, client, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared tbl, wkTbl, dayTbl, draftTbl As DataTable
    Public Shared oTest As Object
    Public Shared oTest2 As Object
    Public Shared today As Date = Date.Now
    Public Shared beginDate, endDate As Date
    Public Shared selectedYear, thisStore, selectedDept, thisPlan, thisBuyer, thisClass, planType, lastUpdate As String
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, lastYear, thisWeek As Int16
    Public Shared rowIndex As Int16
    Public Shared rnd As Double
    Public Shared theValue, adjustmentType As String
    Public Shared thisUser = Environment.UserName
    Public Shared planBy As String
    Public Shared thisPeriod, thisYear, periodSelected As Integer
    Public Shared percentagesBy As String
    Public Shared periodDates(50) As String
    Public Shared todaysPeriod As Integer
    Public Shared statusColumn As Integer
    Public Shared newDraftAmt As Integer
    Public Shared newDraftPlan As Decimal
    Public Shared theValueB4Change As String
    Public Shared thisPlanIsActive As Boolean
    Public Shared planStatus As String
    Public Shared onFormLoad As Boolean = True
    Public Shared tblRow, periodTotalRow As DataRow
    Public Shared formLoaded As Boolean = False
    Public Shared storeIndex, deptIndex, rndIndex As Integer
    Public Shared newDept, somethingChanged As Boolean

    Private Sub PlanMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.Location = New Point(0, 0)
            server = MainMenu.Server
            database = MainMenu.dBase
            client = MainMenu.Client_Id
            serverLabel.Text = MainMenu.serverLabel.Text
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con4 = New SqlConnection(conString)
            con5 = New SqlConnection(conString)

            Dim theAnswer As String = ""
            sql = "IF (EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Sales_Plan')) " & _
                "BEGIN " & _
                "SELECT 'It Exists'  AS ANS " & _
                "END"
            con.Open()
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                theAnswer = rdr("ANS")
            End While
            con.Close()

            If theAnswer = "" Then
                MsgBox("Run the set up for Sales Planning in RCSetup then try again")
                Exit Sub
            End If

            con.Open()
            sql = "SELECT DISTINCT Prd_Id FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Prd_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysPeriod = rdr(0)
            End While
            con.Close()

            con.Open()
            sql = "SELECT ID AS Str_Id FROM Stores WHERE Status = 'Active' ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr(0)
                cboStore.Items.Add(rdr(0))
            End While
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Dept FROM Sales_Summary Order By Dept"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            cboDept.Items.Add("ALL")
            While rdr.Read
                oTest = rdr(0)
                cboDept.Items.Add(rdr(0))
            End While
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Plan_Id FROM Sales_Plan ORDER BY Plan_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboPlan.Items.Add(rdr(0))
            End While
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Prd_Id FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Prd_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysPeriod = rdr(0)
            End While
            con.Close()

            thisYear = DatePart(DateInterval.Year, Date.Today)
            selectedYear = thisYear

            cboRnd.Items.Add(1)
            cboRnd.Items.Add(10)
            cboRnd.Items.Add(50)
            cboRnd.Items.Add(100)
            cboRnd.Items.Add(500)
            cboRnd.Items.Add(1000)
            cboRnd.SelectedIndex = 3
            rndIndex = 3
            rnd = CDbl(cboRnd.SelectedItem)
            cboStore.SelectedIndex = 0
            cboDept.SelectedIndex = 0
            storeIndex = 0
            deptIndex = 0
            selectedDept = cboDept.SelectedItem
            thisStore = cboStore.SelectedItem
            rdoPct.Checked = True
            adjustmentType = "Pct"

            con.Open()
            oTest = Nothing
            sql = "SELECT DISTINCT Plan_Id FROM Sales_Plan WHERE Year_Id = " & thisYear & " AND Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Plan_id")
            End While
            con.Close()

            con.Open()
            sql = "SELECT sDate, eDate FROM Calendar WHERE CONVERT(Date,GETDATE()) BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                beginDate = rdr("sDate")
                endDate = rdr("eDate")
            End While
            con.Close()

            If oTest = Nothing Then
                Dim ans As Integer = MessageBox.Show("OK to create new plan for " & thisYear & "?", "PLAN NOT FOUND", MessageBoxButtons.YesNo)
                If DialogResult = Windows.Forms.DialogResult.Yes Then
                    CreatePlan.Show()
                    Exit Sub
                End If
            End If

            If Not IsNothing(oTest) Then
                cboPlan.SelectedIndex = cboPlan.FindString(oTest)
                Call Create_OR_Load_Plan()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Create_OR_Load_Plan()
        'Try

        planBy = "Period"
        con.Open()
        sql = "SELECT DISTINCT Status, Year_Id FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        txtType.Text = "Draft"
        txtType.ForeColor = Color.Green
        thisPlanIsActive = False
        While rdr.Read
            oTest = rdr(0)
            If Not IsDBNull(oTest) Then
                planStatus = oTest
                If oTest = "Active" Then
                    txtType.Text = "Active"
                    txtType.ForeColor = Color.Red
                    thisPlanIsActive = True
                End If
                If oTest = "Inactive" Then
                    txtType.Text = "Inactive"
                    txtType.ForeColor = Color.Red
                End If
            End If
        End While
        con.Close()

        If thisPlanIsActive Then Call Check_For_Completed_Period()

        Call Load_The_Data()

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        '    If con.State = ConnectionState.Open Then con.Close()
        'End Try
    End Sub

    Private Sub Load_The_Data()
        somethingChanged = False
        Dim str As String
        Dim row As DataRow
        Dim i, clms As Int32
        Dim ly As Integer = 0
        Dim l2y As Integer = 0
        Dim l3y As Integer = 0
        Dim cnt As Integer = 0
        Dim prd As Integer
        Dim actSales, actVariance As Decimal
        priorYrs = 0
        cboDept.Enabled = True
        cboStore.Enabled = True
        '' ''thisStore = cboStore.SelectedItem
        '' ''selectedDept = cboDept.SelectedItem
        '' ''thisPlan = cboPlan.SelectedItem
        If Not IsNothing(thisplan) Then
            con.Open()
            sql = "SELECT DISTINCT Year_Id, Last_Update FROM Sales_Plan WHERE Plan_Id = '" & thisplan & "' " & _
                "AND Str_Id = '" & thisStore & "'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            Dim foundIt As Boolean = False
            While rdr.Read
                oTest = rdr("Year_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    selectedYear = rdr(0)
                    lastYear = selectedYear - 1
                    oTest = rdr("Last_Update")
                    If Not IsDBNull(oTest) Then txtLastUpdate.Text = oTest
                    foundIt = True
                End If
            End While
            If foundIt = False Then                          '       Clear the Grid and the two text boxes and bail out
                txtLastUpdate.Text = ""
                txtType.Text = ""
                '' tbl.Clear()
                con.Close()
                Exit Sub
            End If
            con.Close()
        End If
        con2 = New SqlConnection(conString)
        con3 = New SqlConnection(conString)
        con4 = New SqlConnection(conString)
        con5 = New SqlConnection(conString)

        tbl = New DataTable
        tbl.Reset()
        tbl.Columns.Clear()
        Me.Refresh()
        Dim column = New DataColumn()
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "Period"
        tbl.Columns.Add(column)
        Dim PrimaryKey(1) As DataColumn
        PrimaryKey(0) = tbl.Columns("Period")
        tbl.PrimaryKey = PrimaryKey
        tbl.Columns.Add("3 Year Average")
        tbl.Columns.Add("2 Year Average")
        tbl.Columns.Add(selectedYear & " Plan")
        tbl.Columns.Add("Plan Variance")
        tbl.Columns.Add("Draft Plan")
        tbl.Columns.Add("Draft Variance")
        tbl.Columns.Add(selectedYear & " Actual")
        str = selectedYear & " Actual"
        tbl.Columns.Add(selectedYear & " Variance")
        Dim x As Integer = tbl.Columns.Count
        con2.Open()
        sql = "SELECT DISTINCT Year_Id FROM Sales_Summary WHERE Act_Sales > 0 AND Str_Id = '" & thisStore & "' " & _
                "AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " ORDER BY Year_Id DESC"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr(0)
            If oTest < selectedYear Then
                If cnt = 0 Then
                    cnt += 1
                    ly = oTest                ' set ly to last year
                ElseIf cnt = 1 Then
                    cnt += 1
                    l2y = oTest               ' set l2y to year before last
                ElseIf cnt = 2 Then
                    cnt += 1
                    l3y = oTest               ' set l3y to three years ago
                End If
                priorYrs += 1
                str = rdr(0) & " Actual"
                tbl.Columns.Add(str)
                str = rdr(0) & " Variance"
                If cnt < 3 Then tbl.Columns.Add(str)
            End If
        End While
        tbl.Columns.Add("Status")
        Dim totalColumns = tbl.Columns.Count
        statusColumn = totalColumns
        'tbl.Columns.Add("Last Change")
        tbl.Columns.Add("Change Flag")
        con2.Close()

        con2.Open()
        sql = "SELECT Prd_Id, sDate, eDate FROM Calendar WHERE Year_ID = " & selectedYear & " AND Prd_Id > 0 AND Week_Id = 0 ORDER BY Prd_ID"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            periodDates(i) = rdr("sDate") & " - " & rdr("eDate")
            i += 1
            row = tbl.NewRow
            row(0) = i
            tbl.Rows.Add(row)
        End While
        Array.Resize(periodDates, i)
        con2.Close()

        newDept = True
        con2.Open()
        If selectedDept = "ALL" Then
            sql = "SELECT Prd_Id, Year_Id, ISNULL(SUM(Act_Sales),0) AS Actual FROM Sales_Summary " & _
            "WHERE Str_Id = '" & thisStore & "' AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " " & _
            "Group BY Prd_Id, Year_Id"
        Else
            sql = "SELECT Prd_Id, Year_Id, ISNULL(SUM(Act_Sales),0) AS Actual FROM Sales_Summary " & _
            "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' " & _
            "AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " " & _
            "Group BY Prd_Id, Year_Id"
        End If
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Prd_Id")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then prd = oTest
            row = tbl.Rows.Find(oTest)
            actSales = rdr("Actual")
            con3.Open()
            If selectedDept = "ALL" Then
                '   sql = "SELECT ISNULL(SUM(Amt),0) AS Amt, isnull(sum(case when draft_amt is not NULL then draft_Amt ELSE 0 END),0) AS Draft_Amt, " & _
                '       "'' AS Status, " & _
                '"(SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " " & _
                '"AND Week_Id > 0 AND Str_Id = '" & thisStore & "' AND Status = 'Active') As Plan_Amt " & _
                '"FROM Sales_Plan WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " " & _
                '"AND Week_Id > 0 AND Str_Id = '" & thisStore & "' AND Plan_Id = '" & thisPlan & "' group by status"
                sql = "SELECT ISNULL(SUM(Amt),0) AS Amt, CASE WHEN (SELECT ISNULL(SUM(Draft_Amt),0) FROM Sales_Plan " & _
                    "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' AND Year_Id = " & selectedYear & " " & _
                    "AND Prd_Id = " & oTest & " AND Week_Id > 0) > 0 " & _
                    "THEN ISNULL(SUM(CASE WHEN Draft_Amt IS NULL THEN Amt ELSE Draft_Amt END),0) END AS Draft_Amt, Status, " & _
                    "(SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " " & _
                    "AND Week_Id > 0 AND Str_Id = '1' AND Status = 'Active' ) As Plan_Amt FROM Sales_Plan " & _
                    "WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " AND Week_Id > 0 AND Str_Id = '" & thisStore & "' " & _
                    "AND Plan_Id = '" & thisPlan & "' " & _
                    "GROUP BY STATUS"
            Else
                '   sql = "SELECT ISNULL(SUM(Amt),0) AS Amt, (SELECT ISNULL(SUM(CASE WHEN Draft_Amt IS NULL THEN Amt ELSE Draft_Amt END),0) " & _
                '       "FROM Sales_Plan " & _
                '       "WHERE Plan_Id = '" & thisPlan & "' AND Year_Id = " & selectedYear & " AND Str_Id = '" & thisStore & "' " & _
                '       "AND Prd_Id = " & oTest & " AND Week_Id > 0 AND Dept = '" & selectedDept & "') AS Draft_Amt, " & _
                '       "Status, (SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " " & _
                '   "AND Week_Id > 0 AND Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' AND Status = 'Active' " & _
                '   ") As Plan_Amt " & _
                '"FROM Sales_Plan WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " " & _
                '"AND Week_Id > 0 AND Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' AND Plan_Id = '" & thisPlan & "' " & _
                '"group by status"
                sql = "SELECT ISNULL(SUM(Amt),0) AS Amt, CASE WHEN (SELECT ISNULL(SUM(Draft_Amt),0) FROM Sales_Plan " & _
                    "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' AND Year_Id = " & selectedYear & " " & _
                    "AND Prd_Id = " & oTest & " AND Week_Id > 0) > 0 " & _
                    "THEN ISNULL(SUM(CASE WHEN Draft_Amt IS NULL THEN Amt ELSE Draft_Amt END),0) END AS Draft_Amt, Status, " & _
                    "(SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " " & _
                    "AND Week_Id > 0 AND Str_Id = '1' AND Dept = '" & selectedDept & "' AND Status = 'Active' ) As Plan_Amt FROM Sales_Plan " & _
                    "WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & oTest & " AND Week_Id > 0 AND Str_Id = '" & thisStore & "' " & _
                    "AND Dept = '" & selectedDept & "' AND Plan_Id = '" & thisPlan & "' " & _
                    "GROUP BY STATUS"
            End If
            cmd = New SqlCommand(sql, con3)
            rdr2 = cmd.ExecuteReader
            Dim amt As Int32 = 0
            Dim plan_amt As Int32 = 0
            Dim draft_amt As Int32 = 0
            Dim status As String
            While rdr2.Read
                amt = rdr2("Amt")
                oTest = rdr2("Plan_Amt")
                If Not IsDBNull(oTest) Then plan_amt = oTest Else plan_amt = 0
                oTest = rdr2("Draft_Amt")
                If Not IsDBNull(oTest) Then draft_amt = CInt(oTest) Else draft_amt = 0
                row.Item("Draft Plan") = Format(draft_amt, "###,###,###")
                oTest = rdr2("Status")
                If Not IsDBNull(oTest) Then status = oTest Else status = ""
                row.Item(selectedYear & " Plan") = Format(plan_amt, "###,###,###")
                row.Item("Status") = status
            End While
            con3.Close()

            For Each column In tbl.Columns
                oTest = rdr("Year_Id")
                If rdr("Year_Id") <= selectedYear Then
                    str = rdr("Year_Id") & " Actual"                        ' this is where it blows up when there are fewer than 3 history
                    If rdr("Year_ID") = selectedYear And selectedYear = thisYear Then
                        If prd < todaysPeriod Then
                            row.Item(selectedYear & " Actual") = Format(rdr("Actual"), "###,###,###")
                        End If
                    Else : row.Item(str) = Format(rdr("Actual"), "###,###,###")
                    End If
                End If

            Next
        End While
        con2.Close()
        clms = tbl.Columns.Count - 2
        Dim Avg3YrTotal As Decimal = 0
        Dim Avg2YrTotal As Decimal = 0
        Dim tyPlanTotal As Decimal = 0
        Dim newPlanTotal As Decimal = 0
        Dim tyActualTotal As Decimal = 0
        Dim lyActualTotal As Decimal = 0
        Dim l2yActualTotal As Decimal = 0
        Dim l3yActualTotal As Decimal = 0
        Dim lyPeriodToDateTotal As Decimal = 0
        Dim draftPlanTotal As Decimal
        Dim tyPlan, draftPlan, tyActual, lyActual, l2yActual, l3yActual As Int32
        For Each row In tbl.Rows
            prd = row("Period")
            If priorYrs >= 1 Then
                oTest = row(ly & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "" Then
                        lyActual = oTest
                        lyActualTotal += CInt(oTest)
                    Else : lyActual = 0
                    End If
                End If
                If priorYrs >= 2 Then
                    oTest = row(l2y & " Actual")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            l2yActual = oTest
                            l2yActualTotal += CInt(oTest)
                        Else : l2yActual = 0
                        End If
                    End If
                End If
                If priorYrs = 3 Then
                    oTest = row(l3y & " Actual")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            l3yActual = oTest
                            l3yActualTotal += CInt(oTest)
                        Else : l3yActual = 0
                        End If
                    End If
                End If
                oTest = row(selectedYear & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "" Then
                        tyActual = oTest
                        tyActualTotal += CInt(oTest)
                        If prd < todaysPeriod Then lyPeriodToDateTotal += lyActual
                    Else : tyActual = 0
                    End If
                End If
            Else : lyActual = 0
            End If
            oTest = row(selectedYear & " Plan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then tyPlan = oTest Else tyPlan = 0
            End If
            If lyActual > 0 And tyPlan > 0 Then
                Dim p As Decimal = (tyPlan / lyActual) - 1
                row("Plan Variance") = Format(p, "###.0%")
            End If

            oTest = row("Draft Plan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                oTest = Replace(oTest, ",", "")
                If IsNumeric(oTest) Then
                    draftPlan = CInt(oTest)
                    draftPlanTotal += draftPlan
                    If lyActual > 0 And draftPlan > 0 Then
                        Dim p As Decimal = (draftPlan / lyActual) - 1
                        newPlanTotal += p
                        row("Draft Variance") = Format(p, "###.0%")
                    End If
                Else : draftPlanTotal += tyPlan
                End If
            End If

            If l2y > 0 Then
                oTest = row(lastYear & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    newDept = False                                                       ' turns off newDept flag when we have sales last year
                    oTest = row(selectedYear & " Plan")
                End If
            End If
            If tyActual <> 0 And lyActual <> 0 Then
                actVariance = (tyActual / lyActual) - 1
            Else : actVariance = 0
            End If
            ''If prd < todaysPeriod Then row(selectedYear & " Variance") = Format(actVariance, "###.0%")
            row(selectedYear & " Variance") = Format(actVariance, "###.0%")
            If selectedYear = thisYear And prd >= todaysPeriod Then row(selectedYear & " Variance") = Nothing
            If lyActual <> 0 And l2yActual <> 0 Then
                actVariance = (lyActual / l2yActual) - 1
            Else : actVariance = 0
            End If
            If ly > 0 Then row(ly & " Variance") = Format(actVariance, "###.0%")

            If l2yActual <> 0 And l3yActual <> 0 Then
                actVariance = (l2yActual / l3yActual) - 1
            Else : actVariance = 0
            End If
            If l2y > 0 Then row(l2y & " Variance") = Format(actVariance, "###.0%")

            If priorYrs >= 2 Then
                oTest = row(l2y & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then l2yActual = oTest Else l2yActual = 0
                row("2 Year Average") = Format(Math.Round((lyActual + l2yActual) / 2, MidpointRounding.AwayFromZero), "###,###,###")
            Else : l2yActual = 0
            End If
            If priorYrs >= 3 Then
                oTest = row(l3y & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) And IsNumeric(oTest) Then l3yActual = oTest Else oTest = 0
                If lyActual + l2yActual + l3yActual <> 0 Then
                    row("3 Year Average") = Format(Math.Round((lyActual + l2yActual + l3yActual) / 3, MidpointRounding.AwayFromZero), "###,###,###")
                End If
            End If
            oTest = row("3 Year Average")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then Avg3YrTotal += oTest
            End If
            oTest = row("2 Year Average")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then Avg2YrTotal += oTest
            End If
            oTest = row(selectedYear & " Plan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then tyPlanTotal += CDec(row(3))
            End If
            oTest = row("Plan Variance")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                oTest = Replace(oTest, "%", "")
                If oTest <> "" Then newPlanTotal += oTest
            End If
        Next '                                                                      Total row starts here
        Dim pct As Decimal
        row = tbl.NewRow
        row("Period") = "Total"
        row("3 Year Average") = Format(Avg3YrTotal, "###,###,###")
        row("2 Year Average") = Format(Avg2YrTotal, "###,###,###")
        row(selectedYear & " Plan") = Format(tyPlanTotal, "###,###,###")
        If tyPlanTotal > 0 And lyActualTotal > 0 Then
            pct = tyPlanTotal / lyActualTotal - 1
            row("Plan Variance") = Format(pct, "###.0%")
        End If
        ''If selectedDept <> "ALL" Then row("Draft Plan") = Format(draftPlanTotal, "###,###,###")
        row("Draft Plan") = Format(draftPlanTotal, "###,###,###")

        If draftPlanTotal > 0 And lyActualTotal > 0 Then
            pct = draftPlanTotal / lyActualTotal - 1
            row("Draft Variance") = Format(pct, "###.0%")
        End If

        If tyActualTotal > 0 And lyPeriodToDateTotal > 0 Then
            row(ly & " Actual") = Format(lyActualTotal, "###,###,###")
            pct = (tyActualTotal / lyPeriodToDateTotal) - 1                                 ' was lyactualtotal
            row(selectedYear & " Variance") = Format(pct, "##.0%")
        End If
        row(selectedYear & " Actual") = Format(tyActualTotal, "###,###,###")

        If lyActualTotal > 0 And l2yActualTotal > 0 Then
            row(l2y & " Actual") = Format(l2yActualTotal, "###,###,###")
            pct = (lyActualTotal / l2yActualTotal) - 1
            row(ly & " Variance") = Format(pct, "###.0%")
        End If
        If l2yActualTotal > 0 And l3yActualTotal > 0 Then
            row(l3y & " Actual") = Format(l3yActualTotal, "###,###,###")
            pct = (l2yActualTotal / l3yActualTotal) - 1
            row(l2y & " Variance") = Format(pct, "###.0%")
        End If
        tbl.Rows.Add(row)

        dgv1.Columns.Clear()
        dgv1.DataSource = tbl.DefaultView
        cnt = dgv1.RowCount
        dgv1.Rows(cnt - 1).ReadOnly = True                                    ' disable last row in dgv1
        For i = 0 To dgv1.Columns.Count - 1
            If i <> 6 Then dgv1.Columns(i).ReadOnly = True
        Next
        For i = 0 To dgv1.Rows.Count - 2
            dgv1.Rows(i).Cells(0).Style.ForeColor = Color.Blue
            dgv1.Rows(i).Cells(0).Style.Font = New Font("Ariel", 8, FontStyle.Underline)
        Next
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Dim lastColumn As Integer = dgv1.Columns.Count - 1
        dgv1.Columns(lastColumn).Visible = False
        dgv1.Columns(statusColumn - 1).Visible = False
        For i = 0 To dgv1.Columns.Count - 1
            dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        For i = 1 To dgv1.ColumnCount - 3
            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'If i = 6 And selectedDept <> "ALL" And planStatus = "Draft" Then
            '    dgv1.Columns(6).DefaultCellStyle.BackColor = Color.Cornsilk
            'End If
            'If selectedDept = "ALL" Then dgv1.Columns(i).ReadOnly = True
            dgv1.Columns(i).ReadOnly = True
            If dgv1.Columns(i).Name = "Draft Variance" Then
                If selectedDept <> "ALL" And planStatus = "Draft" Then
                    dgv1.Columns(6).DefaultCellStyle.BackColor = Color.Cornsilk
                    dgv1.Columns(6).ReadOnly = False
                End If
            End If
        Next
        dgv1.Columns(statusColumn).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv1.Columns(statusColumn).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        If txtType.Text = "Draft" And selectedDept <> "ALL" Then
            dgv1.Columns(6).ReadOnly = False
        End If
        dgv1.Rows(cnt - 1).Cells(6).Style.BackColor = Color.White

        If thisPlanIsActive Then
            dgv1.Columns(5).Visible = False
            dgv1.Columns(6).Visible = False
        End If
        formLoaded = True
    End Sub

    Private Sub btnCreatePlan_Click(sender As Object, e As EventArgs) Handles btnCreatePlan.Click
        Me.Close()

        CreatePlan.Show()                                                    ' Open the CreatePlan form and create the new plan there
        'If Not IsNothing(CreatePlan.planName) Then
        '    thisPlan = CreatePlan.planName
        '    cboPlan.Items.Add(thisPlan)
        '    cboPlan.SelectedIndex = cboPlan.FindString(thisPlan)
        '    Call Load_The_Data()
        '    Exit Sub
        'Else : Exit Sub
        'End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        lblProcessing.Visible = True
        Me.Refresh()
        Call Save_Period_Plan()
        ' MessageBox.Show(thisPlan & " saved", "Save Plan")
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub

    Private Sub Add_Sales_Summary_Records()
        Try
            Dim Year, prd, week, yrprd, prdwk, yeartoadd As Integer
            Dim sdate, edate As Date
            Dim store, dept, buyer, clss As String
            yeartoadd = selectedYear
            con2.Open()
            sql = "SELECT Year_Id, Prd_Id, Week_Id, sDate, eDate, YrPrd, PrdWk FROM Calendar " & _
                                "WHERE Week_Id > 0 and Year_Id = " & yeartoadd & " "
            cmd = New SqlCommand(sql, con2)
            rdr2 = cmd.ExecuteReader
            While rdr2.Read
                Year = rdr2("Year_Id")
                prd = rdr2("Prd_Id")
                week = rdr2("Week_Id")
                sdate = rdr2("sDate")
                edate = rdr2("eDate")
                prdwk = rdr2("PrdWk")
                yrprd = rdr2("YrPrd")
                con3.Open()
                sql = "SELECT ID AS Store FROM Stores WHERE Status = 'Active'"
                cmd = New SqlCommand(sql, con3)
                rdr3 = cmd.ExecuteReader
                While rdr3.Read
                    store = rdr3("Store")
                    con4.Open()
                    sql = "SELECT DISTINCT Dept, Buyer, Class FROM Sales_Summary w " & _
                        "JOIN Buyers b ON Buyer = ID WHERE Year_Id = " & yeartoadd - 1 & " AND Status = 'Active'"
                    cmd = New SqlCommand(sql, con4)
                    rdr4 = cmd.ExecuteReader
                    While rdr4.Read
                        dept = rdr4("Dept")
                        buyer = rdr4("Buyer")
                        clss = rdr4("Class")
                        con5.Open()
                        '' If dept = "AC" And prd = 2 And week = 2 Then MsgBox(dept & " " & prd & " " & week & " " & buyer)
                        sql = "IF NOT EXISTS (SELECT * FROM Sales_Summary WHERE Str_Id = '" & store & "' " & _
                            "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' " & _
                            "AND sDate = '" & sdate & "') " & _
                            "INSERT INTO Sales_Summary (Str_Id, Year_Id, Prd_Id, Week_Id, Dept, Buyer, Class, " & _
                                "sDate, eDate, YrPrd, Week_Num) " & _
                            "SELECT '" & store & "', '" & Year & "', '" & prd & "', '" & prdwk & "', '" &
                                dept & "', '" & buyer & "', '" & clss & "', '" & sdate & "', '" & edate & "', " &
                                yrprd & ", " & week & " "
                        cmd = New SqlCommand(sql, con5)
                        cmd.ExecuteNonQuery()
                        con5.Close()
                    End While
                    con4.Close()
                End While
                con3.Close()
            End While
            con2.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            If con3.State = ConnectionState.Open Then con.Close()
            If con4.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub Save_Period_Plan()
        Try
            Dim initialPlan As String = "_Initial"
            Dim pos As Integer = thisPlan.IndexOf(initialPlan)
            If pos > 0 Then Exit Sub '                                             Don't allow any changes to _Initial Plans

            lblSaving.Visible = True
            pb1.Visible = True
            pb1.Minimum = 1
            pb1.Step = 1
            pb1.Value = 1
            Dim stat As String = txtType.Text
            Dim numRows As Int16 = 0
            For Each row In dgv1.Rows
                oTest = row.cells("Draft Plan").value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then numRows += 1
            Next
            pb1.Maximum = numRows
            lastYear = selectedYear - 1
            Dim thisAmt As Int32

            con.Open()                                               ' Update dept as pct of store
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
            For Each row In tbl.Rows                         ' update Draft Plan for each period
                oTest = row("Change Flag")
                If IsDBNull(oTest) Or IsNothing(oTest) Then GoTo 100
                If oTest <> "Y" Then GoTo 100
                oTest = row(0)
                If oTest = "Total" Then GoTo 100
                oTest = row("Draft Plan")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "" Then
                        thisPeriod = row(0)
                        thisAmt = oTest
                        con.Open()
                        thisAmt = row("Draft Plan")
                        sql = "DECLARE @pct decimal(18,4), @sales decimal(18,4), @allsales decimal(18,4) " & _
                               "SELECT @sales = (SELECT ISNULL(SUM(Act_sales),0) FROM Sales_Summary " & _
                                   "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' AND Year_Id = " &
                                       lastYear & " AND Prd_Id = " & thisPeriod & ") " & _
                               "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary " & _
                                   "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Prd_Id = " & thisPeriod & ") " & _
                             "SELECT @pct = @sales / @allsales WHERE @sales > 0 " & _
                            "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
                               "AND Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id = 0 " & _
                               "AND Dept = '" & selectedDept & "' AND Str_Id = '" & thisStore & "') " & _
                            "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Draft_Amt, Pct, Last_Update, Last_User) " & _
                               "SELECT '" & thisPlan & "'," & selectedYear & ",'" & thisStore & "','" & selectedDept & "'," & thisPeriod & ",0," & _
                               thisAmt & ",@pct,'" & today & "', '" & thisUser & "' " & _
                           "ELSE " & _
                            "UPDATE Sales_Plan Set Draft_Amt = " & thisAmt & ", Last_Update = '" & today & "', Last_User = '" & thisUser & "' " & _
                               "WHERE Year_Id = " & selectedYear & " AND Prd_iD = " & thisPeriod & " AND Week_Id = 0 " & _
                               "AND Dept = '" & selectedDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "'"

                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        con.Close()

                        con.Open()
                        ''                                        round Draft Plan and update the week records
                        sql = "DECLARE @rnd AS Decimal " & _
                            "SET @rnd = (SELECT " & rnd & ") " & _
                            "UPDATE p1 SET Draft_Amt = ROUND((p2.Draft_Amt * p1.Pct) / @rnd,0) * @rnd FROM Sales_Plan p1 " & _
                                    "JOIN Sales_Plan p2 ON p2.Plan_Id = p1.Plan_Id AND p2.Str_Id = p1.Str_Id AND p2.Dept = p1.Dept " & _
                                    "AND p2.Year_Id = p1.Year_Id AND p2.Prd_Id = p1.Prd_Id " & _
                                    "WHERE p1.Plan_Id = '" & thisPlan & "' AND p1.Str_Id = '" & thisStore & "' AND p1.Dept = '" & selectedDept & "' " & _
                                    "AND p1.Year_Id = " & selectedYear & " AND p1.Prd_Id = " & thisPeriod & " AND p1.Week_Id > 0"

                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        ''                                      update Day_Sales_Plan
                        cmd = New SqlCommand("sp_RCAdmin_ReCalc_Day_Sales_Plan", con)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add("@planId", SqlDbType.VarChar).Value = thisPlan
                        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
                        cmd.Parameters.Add("@dept", SqlDbType.VarChar).Value = selectedDept
                        cmd.Parameters.Add("@year", SqlDbType.Int).Value = selectedYear
                        cmd.Parameters.Add("@period", SqlDbType.Int).Value = thisPeriod
                        cmd.Parameters.Add("@week", SqlDbType.Int).Value = 0
                        cmd.Parameters.Add("@amt", SqlDbType.Int).Value = thisAmt
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                End If
                pb1.PerformStep()
100:        Next

            Call Load_The_Data()
            lblSaving.Visible = False
            pb1.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub cboPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPlan.SelectedIndexChanged
        lblProcessing.Visible = True
        Me.Refresh()
        thisPlan = cboPlan.SelectedItem
        Call Create_OR_Load_Plan()
        lblProcessing.Visible = False

        Me.Refresh()
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Try
            If planStatus <> "Draft" Then
                MessageBox.Show("Only Draft Plans may be changed.", "ERROR!")
            End If
            If e.RowIndex > -1 Then rowIndex = e.RowIndex
            If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
            Dim foundRow As DataRow
            foundRow = tbl.Rows.Find(rowIndex + 1)
            thisPeriod = foundRow("Period")
            If newDept = False Then
                If thisPlanIsActive = True Then
                    MsgBox("Active Plans cannot be changed!")
                    foundRow("Draft Variance") = Nothing
                    Exit Sub
                End If

                If rowIndex + 1 < todaysPeriod And selectedYear <= thisYear Then
                    MsgBox("Sorry, you can't change a period plan once the period is complete.")
                    foundRow("Draft Variance") = Nothing
                    Exit Sub
                End If
                If rowIndex + 1 = todaysPeriod And selectedYear = thisYear Then                ' don't allow changes to current period
                    MsgBox("Sorry, you can't change the current period.")
                    foundRow("Draft Variance") = Nothing
                    Me.Refresh()
                    Exit Sub
                End If
            End If
            con.Open()
            sql = "SELECT Modified FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Dept = '" & selectedDept & "' " & _
                "AND Prd_Id = " & thisPeriod & " AND Str_Id = '" & thisStore & "' AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Modified")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    MessageBox.Show("Weekly variances are present for this period. To continue, click OK to close this message " &
                                    "and make changes at the week level.", "Change Not Allowed")
                    con.Close()
                    Exit Sub
                End If
            End While
            con.Close()

            Dim ly As Integer = selectedYear - 1
            Dim lyActualHeading As String = ly & " Actual"
            Dim lyactual, lyActualTotal, variance As Int32
            newDraftAmt = 0
            Dim formattedValue As String

            oTest = dgv1.Rows(rowIndex).Cells("Draft Variance").Value
            If Not IsNothing(oTest) And IsNumeric(oTest) Then variance = CInt(oTest)

            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                oTest = dgv1.Rows(rowIndex).Cells(lyActualHeading).Value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lyactual = CInt(oTest)

            End If

            If adjustmentType = "Pct" Then
                theValue = Math.Round(CDec(lyactual) + (CDec(lyactual) * CDec(variance) * 0.01), MidpointRounding.AwayFromZero)
                formattedValue = Format(variance * 0.01, "##.0%")
            Else
                theValue = Math.Round(CDec(lyactual) + CDec(variance), MidpointRounding.AwayFromZero)
                formattedValue = Format(variance, "$###,###,###")
            End If
            foundRow = tbl.Rows.Find(rowIndex + 1)
            foundRow("Change Flag") = "Y"
            somethingChanged = True
            theValue = Format(Math.Round(theValue / rnd, MidpointRounding.AwayFromZero) * rnd, "###,###,###")
            foundRow("Draft Plan") = theValue
            foundRow("Draft Variance") = formattedValue

            For Each rw In tbl.Rows
                oTest = rw(0)
                If oTest <> "Total" Then
                    oTest = rw("Draft Plan")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then
                            newDraftAmt += Replace(oTest, ",", "")
                        Else
                            oTest = rw(selectedYear & " Plan")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                If IsNumeric(oTest) Then
                                    newDraftAmt += rw(selectedYear & " Plan")
                                End If
                            End If
                        End If
                    End If
                    oTest = rw(lastYear & " Actual")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(Replace(oTest, ",", "")) Then lyActualTotal += CInt(Replace(oTest, ",", ""))
                    End If
                End If
            Next
            foundRow = tbl.Rows.Find("Total")
            Dim pct As Decimal = 0
            If Not IsNothing(foundRow) Then
                foundRow("Draft Plan") = Format(newDraftAmt, "###,###,###")
                oTest = foundRow(lastYear & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    pct = newDraftAmt / Replace(lyActualTotal, ",", "") - 1
                End If
                foundRow("Draft Variance") = Format(pct, "##.0%")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try

    End Sub

    Public Sub dgv1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseDown
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If columnIdx > -1 And rowIdx > -1 Then
                    oTest = dgv1.Rows(rowIdx).Cells(columnIdx).Value
                    If columnIdx = 0 And IsNumeric(oTest) Then
                        If somethingChanged Then
                            Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                                Case DialogResult.Yes
                                    Dim savePeriod As Integer = oTest
                                    Call Save_Period_Plan()
                                    oTest = savePeriod
                                Case Windows.Forms.DialogResult.Cancel
                                    Exit Sub
                            End Select
                        End If
                        periodSelected = oTest
                        planBy = "Week"
                        tblRow = tbl.Rows.Find(periodSelected)
                        ''oTest = tblRow(0)
                        periodTotalRow = tbl.Rows.Find("Total")
                        lastUpdate = Me.txtLastUpdate.Text
                        planType = Me.txtType.Text
                        Weeks.Show()
                    End If
                End If
            Case Windows.Forms.MouseButtons.Right

        End Select

    End Sub

    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim str As String
        Dim x As Integer = dgv1.Rows.Count - 1
        If columnIdx = 0 Then
            If rowIdx > -1 And rowIdx < x Then
                str = periodDates(rowIdx)
                With dgv1.Rows(rowIdx).Cells(0)
                    .ToolTipText = str
                End With
            End If
        End If
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

    Private Sub Save_This_Department()
        con.Open()                                               ' Update dept as pct of store
        sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint, @amt bigint " & _
            "SELECT @sales = (SELECT SUM(Act_Sales) FROM Sales_Summary WHERE Str_Id = '" & thisStore & "' AND Dept = 'AC' " & _
                "AND Year_Id = " & selectedYear & ") " & _
            "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & selectedYear & ") " & _
            "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales > 0 " & _
            "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
                    "AND Year_Id = " & selectedYear & " AND Prd_Id = 0 AND Week_Id = 0 " & _
                    "AND Dept = '" & selectedDept & "' AND Str_Id = '" & thisStore & "') " & _
                 "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User) " & _
                    "SELECT '" & thisPlan & "'," & selectedYear & ",'" & thisStore & "','" & selectedDept & "',0,0,@sales,@pct,'" &
                        today & "', '" & thisUser & "' " & _
                "ELSE " & _
                 "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Last_Update = '" & today & "', Last_User = '" & thisUser & "' " & _
                    "WHERE Year_Id = " & selectedYear & " AND Prd_Id = 0 AND Week_Id = 0 " & _
                    "AND Dept = '" & selectedDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "'"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Save_This_Period()
        'If thisPeriod < todaysPeriod Then
        con.Open()
        sql = "DECLARE @pct decimal(18,4), @sales decimal(18,4), @allsales decimal(18,4) " & _
            "SELECT @sales = (SELECT ISNULL(SUM(Act_sales),0) FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' AND Year_Id = " &
                    lastYear & " AND Prd_Id = " & thisPeriod & ") " & _
            "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Prd_Id = " & thisPeriod & ") " & _
          "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales > 0 " & _
         "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
            "AND Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id = 0 " & _
            "AND Dept = '" & selectedDept & "' AND Str_Id = '" & thisStore & "') " & _
         "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User) " & _
            "SELECT '" & thisPlan & "'," & selectedYear & ",'" & thisStore & "','" & selectedDept & "'," & thisPeriod & ",0,@sales," & _
                "@pct,'" & today & "', '" & thisUser & "' " & _
        "ELSE " & _
         "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Last_Update = '" & today & "', Last_User = '" & thisUser & "' " & _
            "WHERE Year_Id = " & selectedYear & " AND Prd_iD = " & thisPeriod & " AND Week_Id = 0 " & _
            "AND Dept = '" & selectedDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "'"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()
        'End If
    End Sub

    Private Sub Save_This_Week()
        Try

            con2.Open()
            '  SELECT COALESCE(dividend / NULLIF(divisor,0), 0) FROM sometable
            sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint,@amt bigint " & _
            "SELECT @sales = (SELECT SUM(Act_Sales) FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' AND Year_Id = " & lastYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & ") " & _
            "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' AND Year_Id = " & lastYear & " " & _
                "AND Prd_Id = " & thisPeriod & ") " & _
            "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales > 0 " & _
            "SELECT @amt = CONVERT(int,(@pct * @allsales)) " & _
            "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
                "AND Dept = '" & selectedDept & "' AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & ") AND @sales > 0 " & _
            "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User) " & _
            "SELECT '" & thisPlan & "','" & thisStore & "','" & selectedDept & "'," & selectedYear & "," & thisPeriod & "," & _
                thisWeek & ",@amt,@pct,'" & today & "','" & thisUser & "' " & _
            "ELSE " & _
            "UPDATE Sales_Plan SET Amt = @amt, Pct = @pct, Last_Update = '" & today & "', Last_User = '" & thisUser & "' " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & selectedDept & "' " & _
                "AND Year_Id = " & selectedYear & " " & _
                    "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Plan_id = '" & thisPlan & "'"
            cmd = New SqlCommand(sql, con2)
            cmd.ExecuteNonQuery()

            sql = "UPDATE Sales_Plan SET AMT = (SELECT SUM(AMT) FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & selectedDept & "' AND Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id > 0) " & _
                            "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & selectedDept & "' AND Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id = 0 "
            cmd = New SqlCommand(sql, con2)
            cmd.ExecuteNonQuery()
            con2.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()

        End Try
    End Sub

    Private Sub cboRnd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboRnd.SelectedIndexChanged
        rnd = cboRnd.SelectedItem
        rndIndex = cboRnd.SelectedIndex
    End Sub

    Private Sub Create_Other_Tables()
        Try
            lastYear = selectedYear - 1
            con.Open()
            sql = "DELETE FROM Buyer_Pct WHERE Plan_Year = " & selectedYear & ""
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()
            sql = "CREATE TABLE #t1a (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), Sales decimal(18,4)) " & _
                "INSERT INTO #t1a (Str_Id, Prd_Id, Dept, Buyer, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, CASE WHEN b.ID IS NULL THEN 'NA' ELSE w.Buyer END As Buyer, " & _
                    "ISNULL(SUM(Act_Sales),0) AS Sales FROM Sales_Summary w " & _
                    "LEFT JOIN Buyers b ON b.ID = w.Buyer " & _
                    "WHERE Year_Id = " & lastYear & " AND Act_Sales > 0 AND b.Status = 'Active' " & _
                    "GROUP BY Str_Id, Prd_Id, Dept, Buyer, ID"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE #t1 (Str_Id varchar(10) NOT NULL, Prd_id integer, Dept varchar(20), Buyer varchar(20), Sales decimal(18,4)) " & _
                "INSERT INTO #t1 (Str_Id, Prd_Id, Dept, Buyer, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, Buyer, SUM(Sales) As Sales " & _
                    "FROM #t1a GROUP BY Str_Id, Prd_Id, Dept, Buyer"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE #t2 (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Sales decimal(18,4)) " & _
                "INSERT INTO #t2 (Str_Id, Prd_Id, Dept, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, ISNULL(SUM(Act_Sales),0) AS Sales FROM Sales_Summary " & _
                "WHERE Year_Id = " & lastYear & " AND Act_Sales > 0 GROUP BY Str_Id, Prd_Id, Dept"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "INSERT INTO Buyer_Pct (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Pct) " & _
                "SELECT '" & selectedYear & "', #t1.Str_Id, #t1.Prd_Id, #t1.Dept, #t1.Buyer, #t1.Sales / " & _
                "(SELECT #t2.Sales FROM #t2 WHERE #t2.Str_Id = #t1.Str_Id AND #t2.Prd_Id = #t1.Prd_Id AND #t2.Dept = #t1.Dept) " & _
                "FROM #t1"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()
            sql = "DELETE FROM Class_PCT WHERE Plan_Year = " & selectedYear & ""
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()
            sql = "CREATE TABLE #t1a (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), " & _
                "Class varchar20), Sales decimal(18,4) INSERT INTO #t1a (Str_Id, Prd_Id, Dept, Buyer, Class, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, CASE WHEN Buyer IS NULL OR Buyer = '' THEN 'NA' ELSE Buyer END AS Buyer, " & _
                    "Class, ISNULL(SUM(Act_Sales),0) AS Sales FROM Sales_Summary " & _
                "WHERE Year_Id = " & lastYear & " AND Act_Sales > 0 GROUP BY Str_Id, Prd_Id, Dept, Buyer, Class"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE (Str_Id varchar(10), Prd_Id integer, Dept varchar(20), Buyer varchar(20), " & _
                "Class varchar(20) Sales INSERT INTO #t1a (Str_Id, Prd_Id, Dept, Buyer, Class, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, Buyer, Class, SUM(Sales) As Sales " & _
                    "FROM #t1a GROUP BY Str_Id, Prd_Id, Dept, Buyer, Class"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "INSERT INTO Class_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Class, Pct) " & _
                    "SELECT '" & selectedYear & "', Str_Id, Prd_Id, Dept, Buyer, Class, Sales / " & _
                    "(SELECT SUM(t1.Sales) FROM #t1 t1 WHERE t1.Str_Id = t2.Str_Id AND t1.Prd_Id = t2.Prd_Id AND t1.Dept = t2.Dept " & _
                        "AND t1.Buyer = t2.Buyer) FROM #t1 t2"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()
            sql = "CREATE TABLE #t1 (Prd_Id integer NOT NULL, Day integer, Sales decimal(18,4)) " & _
                "INSERT INTO #t1 (Prd_Id, Day, Sales) " & _
                "SELECT Prd_Id, DATEPART(dw,TRANS_DATE) AS Day, ISNULL(SUM(QTY * RETAIL),0) AS Sales " & _
                    "FROM Daily_Transaction_Log l " & _
                    "JOIN Calendar c ON TRANS_DATE BETWEEN sDate AND eDate AND PrdWk > 0 " & _
                    "WHERE QTY <> 0 AND RETAIL > 0 AND LOCATION = '1' " & _
                    "AND TYPE = 'Sold' AND DATEPART(dw,TRANS_DATE) > 1 " & _
                    "AND TRANS_DATE BETWEEN DATEADD(YEAR,-1,GETDATE()) AND GETDATE() " & _
                    "GROUP BY Prd_Id, DATEPART(dw,TRANS_DATE)"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()

            sql = "INSERT INTO DayOfWeekPct (Prd_Id, Day, Pct) " & _
                    "SELECT t1.Prd_Id, t1.Day, t1.Sales / (SELECT SUM(Sales) FROM #t1 t2 WHERE t2.Prd_Id = t1.Prd_Id) AS Pct " & _
                     "FROM #t1 t1 ORDER BY Prd_Id, Day"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub btnActivate_Click(sender As Object, e As EventArgs) Handles btnActivate.Click
        Try
            If txtType.Text <> "Draft" Then
                MsgBox("Only Draft Plans can be activated.")
                Exit Sub
            End If
            lblProcessing.Visible = True
            Me.Refresh()

            Call Save_Period_Plan()

            con.Open()                                                ' Get the Plan_Id for the Active plan for the year selected
            Dim ActivePlanId As String = ""
            sql = "SELECT Plan_Id FROM Sales_Plan WHERE Status = 'Active' AND Year_Id = " & selectedYear & " "
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                ActivePlanId = rdr("Plan_Id")
            End While
            con.Close()

            con.Open()                                                ' Deactivate the active plan
            sql = "UPDATE Sales_Plan SET Status = 'Inactive', Draft_Amt = Amt WHERE Status = 'Active' AND Year_Id = " & selectedYear & " " & _
                "AND Str_Id = '" & thisStore & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()
            sql = "UPDATE Sales_Plan SET Amt = Draft_Amt WHERE Plan_Id = '" & thisPlan & "' " & _
                "AND Str_Id = '" & thisStore & "' AND ISNULL(Draft_Amt,0) <> 0"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            Dim numbr As Integer = 0
            Dim newPlan As String
            con.Open()
            sql = "SELECT REPLACE(MAX(CONVERT(Integer,dbo.JustTheNumbers(Plan_Id)))," & selectedYear & ",'') FROM Sales_Plan " & _
                "WHERE Plan_Id LIKE '" & selectedYear & "_Plan%'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                If Not IsDBNull(rdr(0)) And Not IsNothing(rdr(0)) Then
                    If IsNumeric(rdr(0)) Then numbr = CInt(rdr(0))
                End If
            End While
            con.Close()

            numbr += 1
            newPlan = selectedYear.ToString & "_Plan" & numbr.ToString
            con.Open()
            sql = "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Last_Update, " & _
                "Last_User, Status, Origin, Draft_Amt) " & _
                "SELECT '" & newPlan & "', Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Last_Update, " & _
                "Last_User, 'Active', '" & thisPlan & "', NULL " & _
                "FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            con.Open()                                                              ' Set Mkdn_Pct to selected year's Active plan Mkdn_Pct
            sql = "UPDATE p1 SET p1.Mkdn_Pct = p2.Mkdn_Pct FROM Sales_Plan p1 " & _
                "JOIN Sales_Plan p2 ON p2.Str_Id = p1.Str_Id AND p2.Dept = p1.Dept " & _
                "AND p2.Year_Id = p1.Year_Id AND p2.Prd_Id = p1.Prd_Id AND p2.Week_Id = p1.Week_Id " & _
                "WHERE p2.Plan_Id = '" & ActivePlanId & "' AND p1.Plan_Id = '" & newPlan & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            txtType.Text = "Active"
            txtType.ForeColor = Color.Red
            Me.Refresh()
            cboPlan.Items.Add(newPlan)
            cboPlan.SelectedIndex = cboPlan.FindString(newPlan)

            '                                                      Update Sales_Summary
            Dim store As String = ""
            Dim dept As String = ""
            Dim buyer As String = ""
            Dim clss As String = ""
            Dim year, period, week As Integer
            Dim amt, mkdn As Decimal
            con.Open()
            sql = "SELECT s.Str_Id, s.Year_Id, c.Prd_Id, s.Week_Id, c.Dept, c.Buyer, c.Class, " & _
                "ISNULL(((s.Amt * c.Pct) * b.Pct),0) AS Amt, ISNULL(s.AMT * c.Pct * b.Pct,0) * ISNULL(s.Mkdn_Pct,0) AS PlanMkdn FROM Sales_Plan s " & _
                "JOIN Class_PCT c ON c.Str_Id = s.Str_Id AND c.Plan_Year = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.Dept = s.Dept " & _
                "JOIN Buyer_PCT b ON b.Plan_Year = s.Year_Id AND b.Str_Id = s.Str_Id AND b.Prd_Id = s.Prd_Id AND b.Dept = s.Dept AND b.Buyer = c.Buyer " & _
                "WHERE Plan_Id = '" & newPlan & "' AND s.Str_Id = c.Str_Id AND s.Week_Id > 0 ORDER BY s.Str_Id, s.Dept, s.Prd_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Str_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then store = oTest
                oTest = rdr("Year_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then year = oTest
                oTest = rdr("Prd_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then period = oTest
                oTest = rdr("Week_Id")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then week = oTest
                oTest = rdr("Dept")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest
                oTest = rdr("Buyer")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then buyer = oTest
                oTest = rdr("Class")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then clss = oTest
                oTest = rdr("amt")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then amt = oTest Else amt = 0
                End If
                oTest = rdr("PlanMkdn")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then mkdn = oTest Else mkdn = 0
                End If
                con2.Open()
                sql = "IF NOT EXISTS (SELECT Str_Id FROM Sales_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' " & _
                    "AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND " & _
                    "Year_Id = " & year & " AND Prd_Id = " & period & " AND Week_Id = " & week & ") " & _
                    "INSERT INTO Sales_Summary (Str_Id, Dept, Buyer, Class, Year_Id, Prd_Id, Week_Id, sDate, eDate, YrPrd, Week_Num, " & _
                    "Plan_Sales, Plan_Mkdn) " & _
                    "SELECT '" & store & "','" & dept & "','" & buyer & "','" & clss & "'," & year & "," & period & ",prdwk, " & _
                    "sDate,eDate,YrPrd,Week_Id," & amt & "," & mkdn & " FROM Calendar " & _
                    "WHERE Year_Id = " & year & " AND Prd_Id = " & period & " AND PrdWk = " & week & " " & _
                    "ELSE " & _
                    "UPDATE Sales_Summary SET Plan_Sales = " & CInt(amt) & ", Plan_Mkdn = " & CInt(mkdn) & " " & _
                    "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND Class = '" & clss & "' AND " & _
                    "Year_Id = " & year & " AND Prd_Id = " & period & " AND Week_Id = " & week & " "
                cmd = New SqlCommand(sql, con2)
                cmd.ExecuteNonQuery()
                con2.Close()
            End While

            con.Close()
            '' Call Load_The_Data()
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()

        End Try
    End Sub

    Private Sub btnSaveAs_Click(sender As Object, e As EventArgs) Handles btnSaveAs.Click
        Try
            lblProcessing.Visible = True
            Me.Refresh()
            Dim start As Integer
            Dim numbr As Integer = 0
            Dim test, newDraft As String
            con.Open()
            sql = "SELECT MAX(Plan_Id) FROM Sales_Plan WHERE Plan_Id LIKE '" & selectedYear & "_Draft%'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                If Not IsDBNull(rdr(0)) And Not IsNothing(rdr(0)) Then
                    test = rdr(0)
                    start = InStr(test, "Draft")
                    Dim final As String = Mid(test, start)
                    oTest = getNumeric(final)
                    If IsNumeric(oTest) Then numbr = CInt(oTest)
                End If
            End While
            con.Close()

            numbr += 1
            newDraft = selectedYear.ToString & "_Draft" & numbr.ToString
            con.Open()
            sql = "CREATE TABLE #t1 (Plan_Id varchar(30) NOT NULL, Str_Id varchar(10), Dept varchar(20), Year_Id integer, " & _
                "Prd_Id integer, Week_Id integer, Amt decimal(18,4), Pct decimal(18,4), Draft_Amt decimal(18,4), " & _
                "Mkdn_Pct decimal(18,4), Status varchar(10), Origin varchar(30)) " & _
                "INSERT INTO #t1 (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Draft_Amt, Mkdn_Pct, Status, Origin) " & _
                "SELECT '" & newDraft & "' AS Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, " & _
                "(SELECT Amt FROM Sales_Plan WHERE Status ='Active' AND Str_Id = s.Str_Id AND Dept = s.Dept " & _
                    "AND Year_Id = s.Year_Id AND Prd_Id =s.Prd_Id AND Week_Id = s.Week_Id) AS Amt, " & _
                "Pct, Draft_Amt, Mkdn_Pct, Status, Origin FROM Sales_Plan s WHERE Plan_Id = '" & thisPlan & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            sql = "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Draft_Amt, " & _
                    "Last_Update, Last_User, Mkdn_Pct, Status, Origin) " & _
                "SELECT Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Draft_Amt, '" & Date.Now & "', '" &
                    thisUser & "', Mkdn_Pct, 'Draft', '" & thisPlan & "' " & _
                "FROM #t1"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            ''sql = "UPDATE Sales_Plan SET Status = 'Inactive' WHERE Plan_Id = '" & thisPlan & "'"
            ''cmd = New SqlCommand(sql, con)
            ''cmd.ExecuteNonQuery()
            con.Close()

            cboPlan.Items.Add(newDraft)
            thisPlan = newDraft
            MessageBox.Show(thisPlan & " saved", "Save Plan")
            cboPlan.SelectedIndex = cboPlan.FindString(newDraft)

            'Call Create_OR_Load_Plan()
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()

        End Try
    End Sub

    Private Sub rdoAmt_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAmt.CheckedChanged
        adjustmentType = "Amt"
    End Sub

    Private Sub rdoPct_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPct.CheckedChanged
        adjustmentType = "Pct"
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

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem
        storeIndex = cboStore.SelectedIndex
        If formLoaded Then
            tbl.Clear()
            Me.Refresh()
            Call Load_The_Data()
        End If

    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        If somethingChanged Then
            Select Case MessageBox.Show("Do you wish to save changes before moving to next Department?", "CHANGE(S) DETECTED!",
                                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Call Save_Period_Plan()
                Case DialogResult.No
                Case Windows.Forms.DialogResult.Cancel
                    Exit Sub
            End Select
        End If
        selectedDept = cboDept.SelectedItem
        deptIndex = cboDept.SelectedIndex
        If formLoaded Then
            tbl.Clear()
            Call Load_The_Data()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            oTest = thisPlan
            Dim x As Integer = thisPlan.IndexOf("Plan")
            If thisPlan.IndexOf("Plan") > -1 Then
                MessageBox.Show("CANNOT DELETE " & thisPlan & "! ONLY Draft PLANS CAN BE DELETED.", "DELETE PLAN")
                Exit Sub
            End If
            Dim ans As DialogResult = MessageBox.Show("OK TO DELETE " & thisPlan & "?", "DELETE SALES PLAN", MessageBoxButtons.YesNo)
            If ans <> Windows.Forms.DialogResult.Yes Then Exit Sub
            con.Open()
            sql = "DELETE FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
            Dim idx As Integer = cboPlan.SelectedIndex
            cboPlan.Items.RemoveAt(idx)
            MessageBox.Show(thisPlan & " DELETED")
            Call Load_The_Data()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Check_For_Completed_Period()
        Dim prd As Integer = 0
        Dim lastPlanPrd, maxPrd As Integer
        con2.Open()
        sql = "SELECT MAX(Prd_Id) AS Prd_Id FROM Buyer_PCT WHERE Year_id = " & thisYear & " AND Str_Id = '" & thisStore & "' " & _
            "AND ISNULL(Projected_WksOH,0) > 0"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Prd_Id")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then lastPlanPrd = CInt(oTest)
            End If
        End While
        con2.Close()

        con2.Open()
        sql = "SELECT MAX(Prd_Id) AS Prd_Id FROM Sales_Plan WHERE Year_id = " & thisYear & " AND Str_Id = '" & thisStore & "'"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Prd_Id")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then maxPrd = CInt(oTest)
            End If
        End While
        ''If lastPlanPrd > 0 And maxPrd > 0 Then
        ''    For period As Integer = lastPlanPrd + 1 To maxPrd
        ''        Update_Closed_Period(period)
        ''    Next
        ''End If


    End Sub

    Private Sub Update_Closed_Period(ByVal period As Integer)
        Dim ans As DialogResult = MessageBox.Show("Period " & period & " just closed. OK to update plan now?", "SALES PLAN NEEDS TO BE UPDATED.", MessageBoxButtons.YesNo)
        If ans = Windows.Forms.DialogResult.No Then Exit Sub
        Dim tblDept As DataTable = New DataTable
        tblDept.Columns.Add("Dept")
        Dim row As DataRow
        con3.Open()
        sql = "SELECT ID FROM Departments WHERE Status = 'Active'"
        cmd = New SqlCommand(sql, con3)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tblDept.NewRow
            row(0) = rdr(0)
            tblDept.Rows.Add(row)
        End While
        con3.Close()

        Dim thisDept As String
        Dim thisEdate As Date
        Dim onHand As Decimal = 0
        Dim planSales As Int32 = 0
        Dim wks As Integer = 0
        Dim dteTbl As New DataTable
        dteTbl.Columns.Add("eDate")
        dteTbl.Columns.Add("PrdWk")
        Dim dRow As DataRow
        con3.Open()
        sql = "SELECT eDate, PrdWK FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " "
        cmd = New SqlCommand(sql, con3)
        rdr = cmd.ExecuteReader
        While rdr.Read
            dRow = dteTbl.NewRow
            dRow("eDate") = rdr("eDate")
            dRow("prdWk") = rdr("PrdWk")
            dteTbl.Rows.Add(dRow)
        End While
        con3.Close()
        For Each dRow In tblDept.Rows
            thisDept = dRow(0)
            con3.Open()                                     ' Buyer_Pct records
            sql = "SELECT s.Buyer, ISNULL(SUM(Act_Sales),0) AS Sales, ISNULL(SUM(Act_Inv_Retail),0) AS inv FROM Sales_Summary s " & _
                "JOIN Inv_Summary i ON i.Str_Id = s.Str_Id AND i.Dept = s.Dept AND i.Buyer = s.Buyer AND i.Class = s.Class " & _
                "AND i.eDate = s.eDate " & _
                "WHERE s.Str_Id = '" & thisStore & "' AND s.Dept = '" & thisDept & "' AND s.eDate = '" & dRow("eDate") & "' " & _
                "GROUP BY Buyer"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisBuyer = rdr("Buyer")
                onHand = rdr("inv")
                con4.Open()
                sql = "SELECT eDate, ISNULL(SUM(Plan_Sales),0) AS sales FROM Sales_Summary WHERE Str_Id = '" & thisStore & "' " & _
                    "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' AND eDate > '" & dRow("eDate") & "' " & _
                    "GROUP BY eDate"
                cmd = New SqlCommand(sql, con)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    planSales += rdr2("sales")
                    wks += 1
                End While
                If planSales >= onHand Then
                    con5.Open()
                    sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Year_Id = " & thisYear & " AND Str_Id = '" & thisStore & "' " & _
                        "AND Prd_Id = " & period & " AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "') " & _
                        "INSERT INTO Buyer_PCT (Year_Id, Str_Id, Prd_Id, Dept, Buyer, Projected_WksOH) " & _
                        "SELECT " & thisYear & ", '" & thisStore & "', " & period & ", '" & thisDept & "', '" & thisBuyer & "', " & wks & " " & _
                        "ELSE " & _
                        "UPDATE Buyer_PCT SET Projected_WksOH = " & wks & " WHERE Plan_Year = " & thisYear & " " & _
                        "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & " " & _
                        "AND Prd_Id = " & period & " "
                    cmd = New SqlCommand(sql, con5)
                    cmd.ExecuteNonQuery()
                    con5.Close()
                    Exit While
                End If
            End While
            con3.Close()

            '
            '
            '   change wksoh calc to same client_Item_Sales
            '   already know invetory for end of last period.  just select sum(planned salesand markdown) by week
            '   and loop record by record until sum(planned sales & markdown) >= inventory
            '
            ''con3.Open()
            ''sql = "CREATE TABLE #t1 (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), " & _
            ''    "Sales decimal(18,4), WksOH integer) " & _
            ''    "INSERT INTO #t1 (Str_Id, Prd_Id, Dept, Buyer, Sales, WksOH) " & _
            ''    "SELECT Str_Id, Prd_Id, Dept, CASE WHEN b.ID IS NULL THEN 'NA' ELSE w.Buyer END As Buyer, " & _
            ''        "ISNULL(SUM(Act_Sales),0) AS Sales, CONVERT(Integer,0) AS wksOH FROM Item_Sales w " & _
            ''        "LEFT JOIN Buyers b ON b.ID = w.Buyer " & _
            ''        "WHERE Str_Id = '" & thisStore & "' AND Prd_Id = " & period & " AND Act_Sales <> 0 " & _
            ''        "AND b.Status = 'Active' and dept = '" & thisDept & "' " & _
            ''        "GROUP BY Str_Id, Prd_Id, Dept, w.Buyer, ID"
            ''cmd = New SqlCommand(sql, con3)
            ''cmd.ExecuteNonQuery()

            ''sql = "CREATE TABLE #t2 (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), sales decimal(20)) " & _
            ''    "INSERT INTO #t2 (Str_Id, Prd_Id, Dept, sales) " & _
            ''    "SELECT Str_Id, Prd_Id, Dept, ISNULL(SUM(Sales),0) AS Sales FROM #t1 " & _
            ''    "GROUP BY Str_Id, Prd_Id, Dept"
            ''cmd = New SqlCommand(sql, con3)
            ''cmd.ExecuteNonQuery()
            ''sql = "INSERT INTO Buyer_Pct (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Pct, Projected_WksOH) " & _
            ''    "SELECT '" & thisYear & "', #t1.Str_Id, #t1.Prd_Id, #t1.Dept, #t1.Buyer, #t1.Sales / " & _
            ''    "(SELECT #t2.Sales FROM #t2 WHERE #t2.Str_Id = #t1.Str_Id AND #t2.Prd_Id = #t1.Prd_Id AND #t2.Dept = #t1.Dept), WksOH " & _
            ''    "FROM #t1"
            ''cmd = New SqlCommand(sql, con3)
            ''cmd.ExecuteNonQuery()

            sql = "DELETE FROM Class_Pct WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " "
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE #t1a (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), " & _
                "Class varchar(20), Sales decimal(18,4)) INSERT INTO #t1a (Str_Id, Prd_Id, Dept, Buyer, Class, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, CASE WHEN Buyer IS NULL OR Buyer = '' THEN 'NA' ELSE Buyer END AS Buyer, " & _
                    "Class, ISNULL(SUM(Act_Sales),0) AS Sales FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Prd_Id = " & period & " AND Act_Sales <> 0 " & _
                "GROUP BY Str_Id, Prd_Id, Dept, Buyer, Class"
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE #t6 (Str_Id varchar(10) NOT NULL, Prd_Id integer, Dept varchar(20), Buyer varchar(20), Class varchar(20), Sales decimal(18,4)) " & _
                "INSERT INTO #t6 (Str_Id, Prd_Id, Dept, Buyer, Class, Sales) " & _
                "SELECT Str_Id, Prd_Id, Dept, Buyer, Class, SUM(Sales) As Sales " & _
                    "FROM #t1a WHERE Sales <> 0 GROUP BY Str_Id, Prd_Id, Dept, Buyer, Class"
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            sql = "INSERT INTO Class_PCT (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Class, Pct) " & _
                    "SELECT '" & thisYear & "', Str_Id, Prd_Id, Dept, Buyer, Class, Sales / " & _
                    "(SELECT SUM(t6.Sales) FROM #t6 t6 WHERE t6.Str_Id = t2.Str_Id AND t6.Prd_Id = t2.Prd_Id AND t6.Dept = t2.Dept " & _
                        "AND t6.Buyer = t2.Buyer) FROM #t6 t2"
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()

            sql = "DELETE FROM DayOfWeekPct WHERE Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " "
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            sql = "CREATE TABLE #t1 (Prd_Id integer NOT NULL, Day integer, Sales decimal(18,4)) " & _
                "INSERT INTO #t1 (Prd_Id, Day, Sales) " & _
                "SELECT Prd_Id, DATEPART(dw,TRANS_DATE) AS Day, ISNULL(SUM(QTY * RETAIL),0) AS Sales " & _
                    "FROM Daily_Transaction_Log l " & _
                    "JOIN Calendar c ON TRANS_DATE BETWEEN sDate AND eDate AND Week_Id = 0 " & _
                    "WHERE QTY <> 0 AND RETAIL > 0 AND LOCATION = '1' " & _
                    "AND TYPE = 'Sold' c.Year_id = " & thisYear & " AND c.Prd_Id = " & thisPeriod & " " & _
                    "GROUP BY Prd_Id, DATEPART(dw,TRANS_DATE)"
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()

            sql = "INSERT INTO DayOfWeekPct (Prd_Id, Day, Pct) " & _
                    "SELECT t1.Prd_Id, t1.Day, t1.Sales / (SELECT SUM(Sales) FROM #t1 t2 WHERE t2.Prd_Id = t1.Prd_Id) AS Pct " & _
                     "FROM #t1 t1 ORDER BY Prd_Id, Day " & _
                     "WHERE "
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            con.Close()
            con3.Close()

            For Each dteRow In dteTbl.Rows
                thisEdate = dteRow("eDate")                          ' Week records
                thisWeek = dteRow("PrdWk")
                lastYear = thisYear - 1
                con4.Open()
                sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint,@amt decimal " & _
                    "DECLARE @rnd AS integer, @mpct decimal (18,4), @msales decimal(18,4), @mallsales decimal (18,4), @mamt decimal " & _
                    "SELECT @rnd = " & rnd & " " & _
                "SET @sales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                    "AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & ") " & _
                "SET @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                    "AND Prd_Id = " & period & ") " & _
                "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
                "SET @amt = CONVERT(int,(@pct * @allsales)) " & _
                "SET @msales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                    "AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & ") " & _
                "SET @mallsales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & lastYear & " " & _
                    "AND Prd_Id = " & period & ") " & _
                "SELECT @mpct = COALESCE(@msales / (NULLIF(@sales,0) + NULLIF(@msales,0)),0) WHERE @msales <> 0 " & _
                "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
                "AND Dept = '" & thisDept & "' AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & ") " & _
                "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Amt, Pct, Last_Update, Last_User, Status, " & _
                "Mkdn_Pct) " & _
                "SELECT '" & thisPlan & "','" & thisStore & "','" & thisDept & "'," & thisYear & "," & period & "," & _
                    thisWeek & ",@amt,@pct,'" & today & "','" & thisUser & "','Active',@mpct " & _
                "ELSE " & _
                "UPDATE Sales_Plan SET Amt = ROUND(@amt / @rnd,0) * @rnd, Pct = @pct, Last_Update = '" & today & "', " & _
                    "Mkdn_Pct = @mpct, Last_User = '" & thisUser & "' " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                    "AND Year_Id = " & thisYear & " " & _
                    "AND Prd_Id = " & period & " AND Week_Id = " & thisWeek & " AND Plan_id = '" & thisPlan & "'"
                cmd = New SqlCommand(sql, con4)
                cmd.ExecuteNonQuery()
                con4.Close()
            Next

            con3.Open()                                            ' Dept record
            sql = "DECLARE @pct decimal (18,4), @sales decimal(18,4), @allsales decimal(18,4), @plan bigint, @amt bigint " & _
               "SELECT @sales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary WHERE Str_Id = '1' AND Dept = 'AC' " & _
                   "AND Year_Id = " & thisYear & ") " & _
               "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary WHERE Str_Id = '1' AND Year_Id = " & thisYear & ") " & _
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
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            con3.Close()

            con3.Open()                                         ' Period record
            sql = "DECLARE @pct decimal(18,4), @sales decimal(18,4), @allsales decimal(18,4) " & _
                "DECLARE @mpct decimal (18,4), @msales decimal(18,4), @mallsales decimal (18,4) " & _
                "SELECT @sales = (SELECT ISNULL(SUM(Act_sales),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " &
                        lastYear & " AND Prd_Id = " & period & ") " & _
                "SELECT @allsales = (SELECT ISNULL(SUM(Act_Sales),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Prd_Id = " & period & ") " & _
                "SELECT @pct = COALESCE(@sales / NULLIF(@allsales,0),0) WHERE @sales <> 0 " & _
                "SELECT @msales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " &
                        lastYear & " AND Prd_Id = " & period & ") " & _
                "SELECT @mallsales = (SELECT ISNULL(SUM(Act_Mkdn),0) FROM Sales_Summary " & _
                    "WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & lastYear & " AND Dept = '" & thisDept & "') " & _
                "SELECT @mpct = COALESCE(@msales / (NULLIF(@sales,0) + NULLIF(@msales,0)),0) WHERE @msales <> 0 " & _
                "IF NOT EXISTS (SELECT * FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
                    "AND Year_Id = " & thisYear & " AND Prd_Id = " & period & " AND Week_Id = 0 " & _
                    "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "') " & _
                "INSERT INTO Sales_Plan (Plan_Id, Year_Id, Str_Id, Dept, Prd_Id, Week_Id, Amt, Pct, Mkdn_Pct, Last_Update, Last_User, Status) " & _
                "SELECT '" & thisPlan & "'," & thisYear & ",'" & thisStore & "','" & thisDept & "'," & period & ",0,@sales," & _
                    "@pct,@mpct,'" & today & "', '" & thisUser & "','Active' " & _
                "ELSE " & _
                "UPDATE Sales_Plan Set Amt = @sales, Pct = @pct, Mkdn_Pct = @mpct, Last_Update = '" & today & "', " & _
                    "Last_User = '" & thisUser & "' " & _
                    "WHERE Year_Id = " & thisYear & " AND Prd_Id = " & period & " AND Week_Id = 0 " & _
                "AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "' AND Plan_id = '" & thisPlan & "' "
            cmd = New SqlCommand(sql, con3)
            cmd.ExecuteNonQuery()
            con3.Close()
        Next
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                e.Cancel = False
            Case CloseReason.UserClosing
                thisPlan = Nothing
                If somethingChanged Then
                    Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        Case DialogResult.Yes
                            Call Save_Period_Plan()
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