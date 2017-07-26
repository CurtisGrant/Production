Imports System.Data.SqlClient
Imports System.Xml
Public Class Weeks
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared tbl, wkTbl, dayTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = Date.Now
    Public Shared thisStore, thisDept, thisPlan, thisBuyer, thisClass, todaysYrWk As String
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, selectedYear, lastYear, thisWeek, todaysYrPrd, thisYrPrd As Int16
    Public Shared rowIndex As Int16
    Public Shared rnd, draftPlanTotal, actualdraftplantotal, newTotal As Double
    Public Shared theValue As String
    Public Shared periodDraftAmt, newWeeksTotal, periodDraftTotal, beginAmt As Int32
    Public Shared thisUser = Environment.UserName
    Public Shared adjIn As String
    Public Shared thisPeriod, periodSelected, todaysPeriod, weekSelected, todaysWeekId As Integer
    Public Shared thisSdate, thisEdate As Date
    Public Shared weekDates(50) As String
    Public Shared statusColumn As Integer
    Public Shared clmCount, ly, l2y, l3y As Integer
    Public Shared periodRow, periodTotalRow, weekRow As DataRow
    Public Shared changesMade, thisPlanIsActive As Boolean
    Public Shared planStatus As String
    Public Shared formLoaded As Boolean = False
    Public Shared storeIndex, deptIndex, rndIndex As Integer

    Private Sub Weeks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            Me.Location = New Point(100, 250)
            conString = PlanMaintenance.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            tbl = PlanMaintenance.tbl

            selectedYear = PlanMaintenance.selectedYear
            thisPeriod = PlanMaintenance.periodSelected
            thisYrPrd = Microsoft.VisualBasic.Right(selectedYear, 2) * 100 + thisPeriod

            todaysPeriod = PlanMaintenance.todaysPeriod
            thisStore = PlanMaintenance.thisStore
            thisPlan = PlanMaintenance.thisPlan
            Me.txtPlan.Text = thisPlan
            thisDept = PlanMaintenance.selectedDept
            priorYrs = PlanMaintenance.priorYrs
            thisUser = PlanMaintenance.thisUser
            periodDraftAmt = PlanMaintenance.newDraftAmt
            storeIndex = PlanMaintenance.storeIndex
            deptIndex = PlanMaintenance.deptIndex
            rndIndex = PlanMaintenance.rndIndex
            periodRow = PlanMaintenance.tblRow
            planStatus = PlanMaintenance.planStatus
            oTest = periodRow(0)
            periodTotalRow = PlanMaintenance.periodTotalRow
            ' Dim plantype As String = PlanMaintenance.planType
            Me.txtType.Text = PlanMaintenance.planType
            'Dim lastupdate As String = PlanMaintenance.lastUpdate
            Me.txtLastUpdate.Text = PlanMaintenance.lastUpdate
            oTest = periodRow("Draft Plan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" And IsNumeric(oTest) Then
                    beginAmt = CInt(oTest)
                Else
                    oTest = periodRow(selectedYear & " Plan")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then beginAmt = CInt(oTest)
                    End If
                End If
            End If
            oTest = periodTotalRow("Draft Plan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                periodDraftTotal = CInt(Replace(oTest, ",", ""))
            End If
            cboRnd.Items.Add(1)
            cboRnd.Items.Add(10)
            cboRnd.Items.Add(50)
            cboRnd.Items.Add(100)
            cboRnd.Items.Add(500)
            cboRnd.Items.Add(1000)
            cboRnd.SelectedIndex = rndIndex
            rnd = cboRnd.SelectedItem


            If Me.txtType.Text = "Active" Then thisPlanIsActive = True Else thisPlanIsActive = False

            ''con.Open()
            ''sql = "SELECT sDate, eDate FROM Calendar WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & ""
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''While rdr.Read
            ''    thisSdate = rdr("sDate")
            ''    thisEdate = rdr("eDate")
            ''End While
            ''con.Close()
            lastYear = selectedYear - 1

            ''   ToolTip1.SetToolTip(txtPeriod, "Date Range" & ControlChars.NewLine & thisSdate & " - " & thisEdate)
            ''   txtPeriod.Text = thisPeriod & "  " & thisSdate & " - " & thisEdate

            con.Open()
            sql = "SELECT sDate, Week_Id, YrWk, YrPrd FROM Calendar WHERE CONVERT(Date,GETDATE()) BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisSdate = rdr("sDate")
                todaysWeekId = rdr("Week_Id")
                todaysYrWk = rdr("YrWk")
                todaysYrPrd = rdr("YrPrd")
            End While
            con.Close()

            txtStore.Text = PlanMaintenance.thisStore
            txtDept.Text = PlanMaintenance.selectedDept
            thisStore = PlanMaintenance.thisStore
            thisDept = PlanMaintenance.selectedDept
            rdoPct.Checked = True
            Call Show_Weeks_Plan()
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()

        End Try
    End Sub

    Private Sub Show_Weeks_Plan()
        Try
            '                                                         Create the week table columns
            changesMade = False
            Dim str As String
            Dim actual As Decimal
            Dim clmName(10) As String
            Dim cnt As Integer = 0
            Dim pct As Double
            Dim row As DataRow
            wkTbl = New DataTable
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Week"
            wkTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = wkTbl.Columns("Week")
            wkTbl.PrimaryKey = PrimaryKey
            wkTbl.Columns.Add("3 Year Average")
            wkTbl.Columns.Add("2 Year Average")
            wkTbl.Columns.Add(selectedYear & " Plan")
            wkTbl.Columns.Add("Plan Variance")
            wkTbl.Columns.Add("Draft Plan")
            wkTbl.Columns.Add("Draft Variance")
            wkTbl.Columns.Add(selectedYear & " Actual")
            wkTbl.Columns.Add(selectedYear & " Variance")
            con2.Open()
            sql = "SELECT DISTINCT Year_Id FROM Item_Sales WHERE Act_Sales <> 0 " & _
                "AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " ORDER BY Year_Id DESC"
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr(0)
                If oTest < selectedYear Then
                    If cnt = 0 Then
                        cnt += 1
                        ly = oTest
                    ElseIf cnt = 1 Then
                        cnt += 1
                        l2y = oTest
                    ElseIf cnt = 2 Then
                        cnt += 1
                        l3y = oTest
                    End If
                    str = rdr("Year_Id") & " Actual"
                    wkTbl.Columns.Add(str)
                    str = rdr("Year_Id") & " Variance"
                    If cnt < 3 Then wkTbl.Columns.Add(str)
                End If
            End While
            con2.Close()
            Array.Resize(clmName, clmCount)
            Dim d As Integer = clmName.Length
            wkTbl.Columns.Add("Status")
            statusColumn = wkTbl.Columns.Count - 1
            'wkTbl.Columns.Add("Last Change")
            wkTbl.Columns.Add("Change Flag")
            '                                                   Create a new row for each week of the selected period
            con2.Open()
            cnt = 0
            sql = "SELECT DISTINCT Week_Id, sDate, eDate FROM Calendar WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " " & _
                "AND Week_Id > 0"
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = wkTbl.NewRow
                row(0) = rdr("Week_Id")
                wkTbl.Rows.Add(row)
                weekDates(cnt) = rdr("sDate") & " - " & rdr("eDate")
                cnt += 1
            End While
            ' Array.Resize(weekDates, cnt)
            con2.Close()


            '                                                   Get actuals from Item_Sales
            con2.Open()
            If thisDept = "ALL" Then
                sql = "SELECT c.Week_Num, c.Year_Id, ISNULL(SUM(Act_Sales),0) As Actual " & _
                      "FROM Calendar c LEFT JOIN Item_Sales w ON w.eDate = c.eDate AND c.Week_Id > 0 " & _
                        "WHERE Str_Id = '" & thisStore & "' AND w.sDate < '" & thisSdate & "' AND w.Prd_Id = " & thisPeriod & " " & _
                        "AND w.Year_Id BETWEEN '" & selectedYear - 3 & "' AND '" & selectedYear & "' AND c.Week_Id > 0 " & _
                        "GROUP by c.Week_Num, c.Year_Id"
            Else
                sql = "SELECT c.Week_Num, c.Year_Id, ISNULL(SUM(Act_Sales),0) As Actual " & _
                      "FROM Calendar c LEFT JOIN Item_Sales w ON w.eDate = c.eDate AND c.Week_Id > 0 " & _
                        "WHERE w.Str_Id = '" & thisStore & "' AND w.Dept = '" & thisDept & "' " & _
                        "AND w.Prd_Id = " & thisPeriod & " AND w.Year_Id BETWEEN '" & selectedYear - 3 & "' AND '" & selectedYear & "' " & _
                        "AND c.Week_Id > 0 GROUP by c.Week_Num, c.Year_Id"
            End If
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            Dim yrID As Integer
            While rdr.Read
                oTest = rdr("Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then actual = oTest
                oTest = rdr("Week_Num")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then thisWeek = oTest
                row = wkTbl.Rows.Find(thisWeek)
                If Not IsNothing(row) Then
                    yrID = rdr("Year_Id")
                    str = yrID & " Actual"
                    oTest = actual
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> 0 Then row(str) = Format(oTest, "###,###,###")
                    End If
                End If
            End While
            con2.Close()

            Dim newTotal As Integer = 0
            For Each row In wkTbl.Rows
                con3.Open()
                thisWeek = row("Week")
                If thisDept = "ALL" Then
                    sql = "SELECT Plan_Id, ISNULL(SUM(Amt),0) AS Amt, 0 AS Pct, " & _
                        "ISNULL(SUM(CASE WHEN Draft_Amt IS NULL THEN Amt ELSE Draft_Amt END),0) AS Draft_Amt, " & _
                        "MIN(Status) AS Status, MIN(Last_Update) AS Last_Update, " & _
                        "(SELECT ISNULL(SUM(Amt),0) FROM Sales_Plan s " & _
                            "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.PrdWk = s.Week_Id " & _
                            "WHERE s.Year_Id = " & selectedYear & " AND s.Prd_Id = " & thisPeriod & " AND c.Week_id = " & thisWeek & "  " & _
                                "AND Str_Id = '" & thisStore & "' AND Status = 'Active') AS Plan_Amt " & _
                        "FROM Sales_Plan s " & _
                            "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.PrdWk = s.Week_Id " & _
                            "WHERE s.Year_Id = " & selectedYear & " AND s.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & " " & _
                                "AND Str_Id = '" & thisStore & "' AND Plan_Id = '" & thisPlan & "' " & _
                                "GROUP BY Plan_Id"
                Else
                    sql = "SELECT Plan_Id, ISNULL(Amt,0) AS Amt, ISNULL(Pct,0), ISNULL(Draft_Amt,0) AS Draft_Amt, Status, Last_Update, " & _
                    "(SELECT ISNULL(Amt,0) FROM Sales_Plan s " & _
                        "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.PrdWk = s.Week_Id " & _
                        "WHERE s.Year_Id = " & selectedYear & " AND s.Prd_Id = " & thisPeriod & " AND c.Week_id = " & thisWeek & "  " & _
                            "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Status = 'Active') AS Plan_Amt " & _
                    "FROM Sales_Plan s " & _
                        "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.PrdWk = s.Week_Id " & _
                        "WHERE s.Year_Id = " & selectedYear & " AND s.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & " " & _
                            "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Plan_Id = '" & thisPlan & "'"
                End If

                cmd = New SqlCommand(sql, con3)
                rdr2 = cmd.ExecuteReader
                Dim plan_amt As Int32 = 0
                Dim draft_amt As Int32 = 0
                Dim periodAmt As Int32 = 0
                Dim status As String = ""
                oTest = periodRow("Draft Plan")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then periodAmt = CInt(Replace(periodRow("Draft Plan"), ",", ""))
                While rdr2.Read
                    oTest = rdr2("Plan_Id")
                    oTest = rdr2("Plan_Amt")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then plan_amt = oTest Else plan_amt = 0
                    oTest = rdr2(2)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then pct = oTest Else pct = 0
                    End If
                    oTest = rdr2("Draft_Amt")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) And CInt(oTest) > 0 Then
                            draft_amt = oTest
                        End If
                    End If
                    If draft_amt <> 0 Then
                        draft_amt = Math.Round(draft_amt / rnd, MidpointRounding.AwayFromZero) * rnd
                        newTotal += draft_amt
                        row.Item("Draft Plan") = Format(draft_amt, "###,###,###")
                    Else : newTotal += plan_amt
                    End If
                    oTest = rdr2("Status")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then status = oTest Else status = ""
                    row.Item(selectedYear & " Plan") = Format(plan_amt, "###,###,###")
                    row.Item("Status") = status
                End While
                con3.Close()
            Next



            '                                                   Get selected year's plan amount
            Dim clms As Integer = wkTbl.Columns.Count - 2
            Dim Avg3YrTotal, Avg2YrTotal, tyPlanTotal, lyActualTotal, l2yActualTotal, l3yActualTotal As Decimal
            Dim tyPlan, draftPlan, tyActual, tyActualTotal, lyActual, l2yActual, l3yActual As Int32
            Dim tyActualWeekToDateTotal, lyWeekToDateTotal As Int32
            draftPlanTotal = 0
            actualdraftplantotal = 0
            For Each row In wkTbl.Rows
                Try
                    If priorYrs >= 1 Then
                        oTest = row(ly & " Actual")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then
                                lyActual = CInt(oTest)
                                lyActualTotal += CDec(oTest)
                            Else : lyActual = 0
                            End If
                        Else : lyActual = 0
                        End If

                        If priorYrs >= 2 Then
                            oTest = row(l2y & " Actual")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                If IsNumeric(oTest) Then
                                    l2yActual = CInt(oTest)
                                Else : l2yActual = 0
                                End If
                            Else : l2yActual = 0
                            End If
                        End If
                        If priorYrs >= 3 Then
                            oTest = row(l3y & " Actual")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                If IsNumeric(oTest) Then
                                    l3yActual = CInt(oTest)
                                Else : l3yActual = 0
                                End If
                            Else : l3yActual = 0
                            End If
                        End If

                    Else : lyActual = 0
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "If priorYrs >= 1 Then")
                End Try

                Try
                    oTest = row(selectedYear & " Actual")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then
                            tyActual = oTest
                            tyActualTotal += CInt(oTest)
                            If CInt(row("Week")) < todaysWeekId Then lyWeekToDateTotal += lyActual
                            If CInt(row("Week")) < todaysWeekId Then tyActualWeekToDateTotal += tyActual
                        Else : tyActual = 0
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "oTest = row(selectedYear &  Actual)")
                End Try

                Try
                    oTest = row(selectedYear & " Plan")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then tyPlan = CInt(oTest) Else tyPlan = 0
                    End If
                    If lyActual > 0 And tyPlan > 0 Then
                        Dim p As Decimal = (tyPlan / lyActual) - 1
                        row("Plan Variance") = Format(p, "###.0%")
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "oTest = row(selectedYear &  Plan)")
                End Try

                Try
                    oTest = row("Draft Plan")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then
                            draftPlan = CInt(oTest)
                            draftPlanTotal += draftPlan
                            actualdraftplantotal += draftPlan
                            If lyActual > 0 And draftPlan > 0 Then
                                Dim p As Decimal = (draftPlan / lyActual) - 1
                                row("Draft Variance") = Format(p, "###.0%")
                            End If
                        End If
                    Else : draftPlanTotal += tyPlan

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "oTest = row(Draft Plan)")
                End Try

                Try
                    Dim actVariance As Decimal = 0
                    If tyActual > 0 Then
                        actVariance = (tyActual / lyActual) - 1
                    Else : actVariance = 0
                    End If
                    If CInt(row("Week")) < todaysWeekId Then
                        row(selectedYear & " Variance") = Format(actVariance, "###.0%")
                    End If

                    If lyActual <> 0 And l2yActual <> 0 Then
                        actVariance = (lyActual / l2yActual) - 1
                    Else : actVariance = 0
                    End If
                    row(ly & " Variance") = Format(actVariance, "###.0%")

                    If l2yActual <> 0 And l3yActual <> 0 Then
                        actVariance = (l2yActual / l3yActual) - 1
                    Else : actVariance = 0
                    End If
                    row(l2y & " Variance") = Format(actVariance, "###.0%")
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Dim actVariance As Decimal")
                End Try

                Try
                    If priorYrs >= 2 Then
                        oTest = row(l2y & " Actual")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then
                                l2yActualTotal += CInt(oTest)
                                l2yActual = CInt(oTest)
                            Else : l2yActual = 0
                            End If
                            row("2 Year Average") = Format(Math.Round((lyActual + l2yActual) / 2, MidpointRounding.AwayFromZero), "###,###,###")
                        End If
                    Else : l2yActual = 0

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "If prioryrs >=2")
                End Try

                Try
                    If priorYrs >= 3 Then
                        oTest = row(l3y & " Actual")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If IsNumeric(oTest) Then l3yActualTotal += CInt(oTest)
                        End If
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) And IsNumeric(oTest) Then l3yActual = CInt(oTest) Else l3yActual = 0
                        If lyActual <> 0 And l2yActual <> 0 And l3yActual <> 0 Then
                            row("3 Year Average") = Format(Math.Round((lyActual + l2yActual + l3yActual) / 3, MidpointRounding.AwayFromZero), "###,###,###")
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "If prioryrs >=3")
                End Try


                oTest = row("3 Year Average")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then Avg3YrTotal += CInt(oTest)
                End If
                oTest = row("2 Year Average")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then Avg2YrTotal += CInt(oTest)
                End If
                oTest = row(selectedYear & " Plan")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then tyPlanTotal += CInt(oTest)
                End If
                oTest = row("Plan Variance")
            Next

            '                                                                Total row starts here
            row = wkTbl.NewRow
            row("Week") = "Total"
            row("3 Year Average") = Format(Avg3YrTotal, "###,###,###")
            row("2 Year Average") = Format(Avg2YrTotal, "###,###,###")
            row(selectedYear & " Plan") = Format(tyPlanTotal, "###,###,###")
            If tyPlanTotal > 0 And lyActualTotal > 0 Then
                pct = tyPlanTotal / lyActualTotal - 1
                row("Plan Variance") = Format(pct, "###.0%")
            End If
            row("Draft Plan") = Format(newTotal, "###,###,###")   ' was actualdraftplantotal
            If draftPlanTotal > 0 And lyActualTotal > 0 Then
                pct = (draftPlanTotal / lyActualTotal) - 1
                row("Draft Variance") = Format(pct, "###.0%")
            End If

            If tyActualTotal > 0 And lyActualTotal > 0 Then
                ''pct = (tyActualTotal / lyWeekToDateTotal) - 1
                pct = (tyActualWeekToDateTotal / lyWeekToDateTotal) - 1
                row(selectedYear & " Variance") = Format(pct, "##.0%")
            End If

            row(selectedYear & " Actual") = Format(tyActualTotal, "###,###,###")
            row(ly & " Actual") = Format(lyActualTotal, "###,###,###")
            If priorYrs >= 2 Then row(l2y & " Actual") = Format(l2yActualTotal, "###,###,###")
            If priorYrs >= 3 Then row(l3y & " Actual") = Format(l3yActualTotal, "###,###,###")

            If lyActualTotal > 0 And l2yActualTotal > 0 Then
                pct = (lyActualTotal / l2yActualTotal) - 1
                row(ly & " Variance") = Format(pct, "###.0%")
            End If

            If l2yActualTotal > 0 And l3yActualTotal > 0 Then
                pct = (l2yActualTotal / l3yActualTotal) - 1
                row(l2y & " Variance") = Format(pct, "###.0%")
            End If
            wkTbl.Rows.Add(row)

            Dim bindingSource As New BindingSource
            bindingSource.DataSource = wkTbl
            dgv1.DataSource = bindingSource
            cnt = dgv1.RowCount
            dgv1.Rows(cnt - 1).ReadOnly = True
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv1.Columns(dgv1.Columns.Count - 1).Visible = False
            For i = 0 To dgv1.ColumnCount - 1
                dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).ReadOnly = True
                If dgv1.Columns(i).Name = "Draft Variance" Then
                    If thisDept <> "ALL" And planStatus = "Draft" Then
                        dgv1.Columns(6).DefaultCellStyle.BackColor = Color.Cornsilk
                        dgv1.Columns(6).ReadOnly = False
                    End If
                End If
            Next
            If thisPlanIsActive Then
                dgv1.Columns("Draft Plan").Visible = False
                dgv1.Columns("Draft Variance").Visible = False
            Else
                dgv1.Columns("Draft Plan").Visible = True
                dgv1.Columns("Draft Variance").Visible = True
            End If
            dgv1.Columns(statusColumn).Visible = False
            formLoaded = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Load_Data")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            If con3.State = ConnectionState.Open Then con3.Close()
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
            If planStatus <> "Draft" Then
                MessageBox.Show("Only Draft Plans may be cjhanged.", "ERROR!")
                ''foundRow("Draft Variance") = Nothing
                Exit Sub
            End If
            If e.RowIndex > -1 Then rowIndex = e.RowIndex
            If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
            Dim foundRow As DataRow
            oTest = dgv1.Rows(rowIndex).Cells(0).Value
            foundRow = wkTbl.Rows.Find(oTest)
            If oTest = "Total" Then
                foundRow("Draft Variance") = Nothing
                Exit Sub
            End If



            If thisPlanIsActive = True Then
                MessageBox.Show("ACTIVE PLANS CANNOT BE CHANGED!", "CHANGE WEEK PLAN")
                foundRow("Draft Variance") = Nothing
                Exit Sub
            End If

            If thisYrPrd < todaysYrPrd Then
                MessageBox.Show("PRIOR PERIODS CANNOT BE CHANGED!", "CHANGE WEEK PLAN")
                foundRow("Draft Variance") = Nothing
                Exit Sub
            End If

            If thisYrPrd = todaysYrPrd And oTest <= todaysWeekId Then
                MessageBox.Show("CURRENT OR PRIOR WEEKS CANNOT BE CHANGED!", "CHANGE WEEK PLAN")
                foundRow("Draft Variance") = Nothing
                Exit Sub
            End If
            Dim lyActual, lyActualTotal, variance As Int32
            Dim thisWk As Integer
            Dim rnd As Double = CDbl(cboRnd.SelectedItem)
            newTotal = 0

            Dim formattedValue As String = ""
            oTest = dgv1.Rows(rowIndex).Cells("Draft Variance").Value
            If Not IsNothing(oTest) And IsNumeric(oTest) Then variance = CInt(oTest)
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                changesMade = True
                thisWk = dgv1.Rows(rowIndex).Cells(0).Value
                oTest = dgv1.Rows(rowIndex).Cells(ly & " Actual").Value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then lyActual = CInt(oTest)
                If rdoPct.Checked Then
                    theValue = Math.Round(CDec(lyActual) + (CDec(lyActual) * CDec(variance) * 0.01), MidpointRounding.AwayFromZero)
                    formattedValue = Format(variance * 0.01, "##.0%")
                Else
                    theValue = Math.Round(CDec(lyActual) + CDec(variance), MidpointRounding.AwayFromZero)
                    formattedValue = Format(variance, "$###,###,###")
                End If
                foundRow = wkTbl.Rows.Find(thisWk)

                ''  foundRow("Change Flag") = "Y"

                theValue = Format(Math.Round(theValue / rnd, MidpointRounding.AwayFromZero) * rnd, "###,###,###")
                foundRow("Draft Plan") = theValue
                foundRow("Draft Variance") = formattedValue
            End If

            ''dgv1.Rows(rowIndex).Cells("Draft Plan").Value = formattedValue             ' was theValue
            For Each rw In wkTbl.Rows
                oTest = rw(0)
                If oTest <> "Total" Then
                    oTest = rw("Draft Plan")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then
                            newTotal += CInt(Replace(oTest, ",", ""))
                        End If
                    Else
                        If IsNumeric(rw(selectedYear & " Plan")) Then
                            newTotal += rw(selectedYear & " Plan")
                        End If
                    End If
                    oTest = rw(lastYear & " Actual")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        oTest = Replace(oTest, ",", "")
                        If IsNumeric(oTest) Then lyActualTotal += CInt(Replace(oTest, ",", ""))
                    End If
                End If
            Next
            foundRow = wkTbl.Rows.Find("Total")                          ' total line code here
            If Not IsNothing(foundRow) Then
                foundRow("Draft Plan") = Format(newTotal, "###,###,###")
                oTest = Replace(lyActualTotal, ",", "")
                Dim pct As Decimal = 0
                If IsNumeric(oTest) And oTest <> 0 Then
                    pct = newTotal / Replace(lyActualTotal, ",", "") - 1
                End If
                foundRow("Draft Variance") = Format(pct, "##.0%")

                periodDraftAmt = newTotal
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim x As Integer = dgv1.Rows.Count - 1
        Dim str As String
        If columnIdx = 0 Then
            If rowIdx > -1 And rowIdx < x Then
                str = weekDates(rowIdx)
                With dgv1.Rows(rowIdx).Cells(0)
                    .ToolTipText = str
                End With
            End If
        End If
    End Sub

    Private Sub dgv1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellClick
        Try
            If e.RowIndex > -1 And e.ColumnIndex > -1 Then
                rowIndex = e.RowIndex
                columnIndex = e.ColumnIndex
                oTest = dgv1.Rows(rowIndex).Cells(0).Value
                If oTest = "Total" Then Exit Sub
                thisWeek = CInt(oTest)
                Dim promoTbl As DataTable = New DataTable
                Dim promoRow As DataRow
                promoTbl.Columns.Add("Year")
                promoTbl.Columns.Add("Comment")
                For i As Integer = 0 To priorYrs
                    con.Open()
                    sql = "SELECT Comment FROM Promotions p JOIN Calendar c ON p.YrWks = c.YrWks " & _
                        "WHERE c.Year_Id = " & selectedYear - i & " And c.Week_id = " & thisWeek & " AND Str_Id = '" & thisStore & "'"
                    cmd = New SqlCommand(sql, con)
                    rdr = cmd.ExecuteReader
                    While rdr.Read
                        If Not IsNothing(rdr(0)) Then
                            promoRow = promoTbl.NewRow
                            promoRow(0) = selectedYear - i
                            promoRow(1) = rdr(0)
                            promoTbl.Rows.Add(promoRow)
                        End If
                    End While
                    con.Close()
                Next
                lblPromotions.Visible = True
                dgv2.Visible = True
                dgv2.DataSource = promoTbl.DefaultView
                '' dgv2.AutoResizeColumns()
                Dim column As DataGridViewColumn = dgv2.Columns(0)
                column.Width = 60
                column = dgv2.Columns(1)
                column.Width = 1100
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Save_Week_Plan()
        Try
            If changesMade = False Then Exit Sub
            Dim thisWeek, thisAmt As Int32
            For Each row In wkTbl.Rows
                oTest = row(0)
                If oTest = "Total" Then GoTo 100
                oTest = row("Draft Plan")

                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    Dim x As String = Replace(oTest.ToString, ",", "")
                    thisAmt = Convert.ToInt32(x)
                    thisWeek = row(0)
                    con.Open()
                    sql = "IF NOT EXISTS (SELECT Plan_Id FROM Sales_Plan p JOIN Calendar c ON c.Year_Id = p.Year_Id AND c.Prd_Id = p.Prd_Id AND c.PrdWk = p.Week_Id " & _
                        "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' " & _
                        "AND p.Year_Id = " & selectedYear & " AND p.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & " " & _
                        "AND Dept = '" & thisDept & "') " & _
                        "INSERT INTO Sales_Plan (Plan_Id, Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Draft_Amt, Last_Update, Last_User, Modified) " & _
                        "SELECT '" & thisPlan & "','" & thisStore & "','" & thisDept & "'," & selectedYear & "," & thisPeriod & ", PrdWk, " & _
                        thisAmt & ",'" & today & "','" & thisUser & "', 1 FROM Calendar WHERE Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " " & _
                            "AND Week_Id = " & thisWeek & " " & _
                        "ELSE " & _
                        "UPDATE p SET Draft_Amt = " & thisAmt & ", Last_Update = '" & today & "', Last_User = '" & thisUser & "', " & _
                        "Modified = 1 FROM Sales_Plan p " & _
                        "JOIN Calendar c ON c.Year_Id = p.Year_Id AND c.Prd_Id = p.Prd_Id AND c.PrdWk = p.Week_id " & _
                        "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND p.Year_Id = " & selectedYear & " AND p.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & ""
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                    cmd = New SqlCommand("sp_RCAdmin_ReCalc_Day_Sales_Plan", con)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("@planId", SqlDbType.VarChar).Value = thisPlan
                    cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
                    cmd.Parameters.Add("@dept", SqlDbType.VarChar).Value = thisDept
                    cmd.Parameters.Add("@year", SqlDbType.Int).Value = selectedYear
                    cmd.Parameters.Add("@period", SqlDbType.Int).Value = thisPeriod
                    cmd.Parameters.Add("@week", SqlDbType.Int).Value = thisWeek
                    cmd.Parameters.Add("@amt", SqlDbType.Int).Value = thisAmt
                    cmd.ExecuteNonQuery()
                    con.Close()
                End If
100:        Next
            ''                                                    Roll week plan into period plan

            Dim temp As DataTable = New DataTable
            Dim trow As DataRow
            temp.Columns.Add("Plan_Id")
            temp.Columns.Add("Str_Id")
            temp.Columns.Add("Dept")
            temp.Columns.Add("Year_Id")
            temp.Columns.Add("Prd_Id")
            temp.Columns.Add("Week_Id")
            temp.Columns.Add("Amt")
            temp.Columns.Add("Draft_Amt")
            Dim Total As Int32 = 0
            Dim pct As Decimal
            con.Open()
            sql = "SELECT Plan_Id, Str_Id, Year_Id, Prd_Id, Week_Id, ISNULL(Amt,0) AS Amt, ISNULL(Pct,0) AS Pct, ISNULL(Draft_Amt,0) AS Draft_Amt " & _
                "FROM Sales_Plan WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & selectedYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' AND Plan_Id = '" & thisPlan & "' ORDER BY Week_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                If rdr("Week_Id") = 0 Then
                    Total = rdr("Amt")
                Else
                    oTest = rdr("Draft_Amt")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If rdr("Draft_Amt") > 0 Then
                            Total = Total - rdr("Amt") + rdr("Draft_Amt")  ' here
                        End If
                        trow = temp.NewRow
                        trow("Plan_Id") = rdr("Plan_Id")
                        trow("Str_Id") = rdr("Str_Id")
                        trow("Year_Id") = rdr("Year_Id")
                        trow("Prd_Id") = rdr("Prd_Id")
                        trow("Week_Id") = rdr("Week_Id")
                        trow("Amt") = rdr("Amt")
                        trow("Draft_Amt") = rdr("Draft_Amt")
                        temp.Rows.Add(trow)
                    End If
                End If
            End While
            con.Close()

            con.Open()
            sql = "UPDATE Sales_Plan SET Draft_Amt = " & Total & " WHERE Str_Id = '" & thisStore & "' AND Year_Id = " & selectedYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' AND Plan_Id = '" & thisPlan & "' AND Week_Id = 0"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()

            Dim store, plan As String
            Dim year, period, week
            For Each row In temp.Rows
                store = row("Str_Id")
                plan = row("Plan_Id")
                year = row("Year_Id")
                period = row("Prd_Id")
                week = row("Week_Id")
                If row("Draft_Amt") > 0 Then
                    pct = row("Draft_Amt") / Total
                Else
                    pct = row("Amt") / Total
                End If
                con.Open()
                sql = "UPDATE Sales_Plan SET Pct = " & pct & " WHERE Plan_Id = '" & plan & "' AND Str_Id = '" & store & "' " & _
                    "AND Year_Id = " & year & " AND Prd_Id = " & period & " AND Week_Id = " & week & " "
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            Dim foundRow As DataRow
            foundRow = wkTbl.Rows.Find("Total")
            periodRow("Draft Plan") = Format(newTotal, "###,###,###")                  ' this line populates the Draft Plan field in the PlanMaintenance form
            periodDraftTotal -= beginAmt
            periodDraftTotal += newTotal
            periodTotalRow("Draft Plan") = Format(periodDraftTotal, "###,###,###")
            oTest = foundRow("Draft Variance")
            periodRow("Draft Variance") = foundRow("Draft Variance")
            changesMade = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SAVE CHANGES")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Call Save_Week_Plan()
            MessageBox.Show(thisPlan & " changes saved.", "Save Changes")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub rdoAmt_CheckedChanged(sender As Object, e As EventArgs) Handles rdoAmt.CheckedChanged
        adjIn = "Amt"
    End Sub

    Private Sub rdoPct_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPct.CheckedChanged
        adjIn = "Pct"
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
                            Call Save_Week_Plan()
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

    Private Sub cboRnd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboRnd.SelectedIndexChanged
        rnd = cboRnd.SelectedItem
        Call Show_Weeks_Plan()
    End Sub
End Class