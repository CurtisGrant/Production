Imports System
Imports System.Text
Imports System.Data.SqlClient
Public Class Days
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared tbl, wkTbl, dayTbl, dayTbl2, promoTbl, workTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = Date.Now
    Public Shared selectedYear, thisStore, thisDept, thisPlan, thisBuyer, thisClass As String
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, lastYear, thisWeek As Int16
    Public Shared rowIndex As Int16
    Public Shared rnd As Double
    Public Shared theValue As String
    Public Shared newDraftAmt, newWeeksTotal As Int32
    Public Shared thisUser = Environment.UserName
    Public Shared adjIn As String
    Public Shared thisPeriod, periodSelected, todaysPeriod, todaysWeekId As Integer
    Public Shared thisSdate, thisEdate As Date
    Public Shared periods(50) As String
    Public Shared statusColumn As Integer
    Public Shared clmCount, ly, l2y, l3y As Integer
    Public Shared tblRow, tblRow2 As DataRow
    Public Shared changesMade As Boolean = False

    Private Sub Days_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            Me.Location = New Point(150, 300)
            con = PlanMaintenance.con
            con2 = PlanMaintenance.con2
            con3 = PlanMaintenance.con3
            selectedYear = PlanMaintenance.selectedYear
            thisPeriod = PlanMaintenance.periodSelected
            todaysPeriod = PlanMaintenance.todaysPeriod
            thisStore = PlanMaintenance.thisStore
            thisPlan = PlanMaintenance.thisPlan
            thisDept = PlanMaintenance.selectedDept
            priorYrs = PlanMaintenance.priorYrs
            thisUser = PlanMaintenance.thisUser
            tblRow = Weeks.weekRow
            thisWeek = tblRow("Week")
            cboPlan.Items.Add(thisPlan)
            cboPlan.SelectedIndex = 0
            cboStore.Items.Add(thisStore)
            cboStore.SelectedIndex = 0
            txtType.Text = PlanMaintenance.txtType.Text
            txtLastUpdate.Text = PlanMaintenance.txtLastUpdate.Text
            DateTimePicker1.Format = DateTimePickerFormat.Custom
            DateTimePicker1.CustomFormat = "MM/dd/yyyy"
            Call Load_Data()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Load_Data()
        Try
            changesMade = False
            con.Open()
            Dim amt, week, total As Integer
            Dim dte, dt As Date
            Dim day, month As Integer
            Dim dow As String
            Dim dept As String = ""
            Dim pct As Double
            Dim row, row2 As DataRow
            Dim array(8) As Integer
            Dim array2(8) As Integer
            Dim foundRow As DataRow
            oTest = tblRow("Draft Plan")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                amt = CInt(Replace(oTest, ",", ""))
            Else
                oTest = tblRow(selectedYear & " Plan")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    amt = CInt(Replace(oTest, ",", ""))
                End If
            End If
            week = tblRow("Week")
            sql = "SELECT sDate, eDate, Prd_Id, Week_Id FROM Calendar WHERE Year_Id = " & selectedYear & " AND Week_Id = " & week & " "
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                dte = rdr("sDate")
                thisSdate = rdr("sDate")
                thisEdate = rdr("eDate")
                lblPrdWk = "Period " & rdr("Prd_Id") & " Week " & rdr("Week_Id")
            End While
            con.Close()
            DateTimePicker1.Value = thisSdate
            dayTbl = New DataTable
            dayTbl.Columns.Clear()
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Dept"
            dayTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = dayTbl.Columns("Dept")
            dayTbl.PrimaryKey = PrimaryKey

            dayTbl2 = New DataTable
            dayTbl2.Columns.Clear()
            Dim column2 = New DataColumn()
            column2.DataType = System.Type.GetType("System.String")
            column2.ColumnName = "Dept"
            dayTbl2.Columns.Add(column2)
            Dim PrimaryKey2(1) As DataColumn
            PrimaryKey2(0) = dayTbl2.Columns("Dept")
            dayTbl2.PrimaryKey = PrimaryKey2

            For i As Integer = 0 To 6
                ''dayTbl.Columns.Add(DateAdd(DateInterval.Day, i, dte))
                dt = DateAdd(DateInterval.Day, i, dte)
                day = DatePart(DateInterval.Day, dt)
                month = DatePart(DateInterval.Month, dt)
                dow = dt.ToString("ddd")
                ''dayTbl.Columns.Add(dow & " " & month & "/" & day)
                dayTbl.Columns.Add(dow)
                dayTbl2.Columns.Add(dow)
            Next
            dayTbl.Columns.Add("Total")
            dayTbl2.Columns.Add("Total")
            row = dayTbl.NewRow
            row(0) = "% WK"
            row2 = dayTbl2.NewRow
            row2(0) = " "
            ''con.Open()
            ''sql = "SELECT Pct FROM DayOfWeekPct WHERE Prd_Id = " & thisPeriod & " ORDER BY Day"
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''Dim cnt As Integer = 1
            ''While rdr.Read
            ''    row(cnt) = Format(rdr("Pct"), "###.0%")
            ''    cnt += 1
            ''End While
            ''con.Close()
            dayTbl.Rows.Add(row)
            dayTbl2.Rows.Add(row2)
            con.Open()
            sql = "SELECT ID AS Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = dayTbl.NewRow
                row2 = dayTbl2.NewRow
                oTest = rdr("Dept")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then thisDept = oTest
                total = 0
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest
                row("Dept") = dept
                row2("Dept") = dept
                For i As Integer = 0 To 6
                    con2.Open()
                    sql = "DECLARE @dayAmt int, @planAmt int " & _
                        "SET @dayAmt = 0 " & _
                        "SET @dayAmt = (SELECT ISNULL(Amt,0) FROM Day_Sales_Plan WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & i + 1 & ") " & _
                        "SET @planAmt = (SELECT s.Amt * d.Pct AS Amt FROM Sales_Plan s " & _
                        "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.PrdWk = s.Week_Id " & _
                        "JOIN DayOfWeekPct d ON d.Year_Id = s.Year_Id AND d.Prd_Id = s.Prd_Id AND Day = " & i + 1 & " " & _
                            "AND s.Str_Id = d.Str_Id " & _
                        "WHERE Plan_Id = '" & thisPlan & "' AND s.Str_Id = '" & thisStore & "' AND Dept = '" & dept & "' " & _
                        "AND s.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & ") " & _
                        "SELECT CASE WHEN @dayAmt <> 0 THEN @dayAmt ELSE @planAmt END AS Amt"
                    cmd = New SqlCommand(sql, con2)
                    rdr2 = cmd.ExecuteReader
                    While rdr2.Read
                        oTest = rdr2("Amt")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = CInt(oTest) Else amt = 0
                        ''oTest = rdr2("Pct")
                        ''If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = CDec(oTest) Else pct = 0
                        total += amt
                        row(i + 1) = Format(amt, "###,##0")
                        array(i) += amt
                        array(7) += amt
                    End While
                    con2.Close()
                Next
                row("Total") = Format(total, "###,###")
                dayTbl.Rows.Add(row)
                dayTbl2.Rows.Add(row2)
            End While
            con.Close()
            row = dayTbl.NewRow
            row2 = dayTbl2.NewRow
            row(0) = "Total"
            foundRow = dayTbl.Rows.Find("% WK")
            For i As Integer = 0 To 7
                row(i + 1) = Format(array(i), "###,###,###")
                pct = array(i) / array(7)
                If i < 7 Then foundRow(i + 1) = Format(pct, "##.0%")
            Next
            dayTbl.Rows.Add(row)

            workTbl = dayTbl
            dgv1.DataSource = dayTbl.DefaultView
            dgv1.AutoResizeColumns()
            dgv1.Rows(0).DefaultCellStyle.BackColor = Color.LightGray
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).ReadOnly = True
            Next
            dgv2.DataSource = dayTbl2.DefaultView
            dgv2.AutoResizeColumns()
            dgv2.Rows(0).ReadOnly = True
            For i = 1 To dgv2.ColumnCount - 1
                dgv2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv2.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv2.Columns(i).DefaultCellStyle.BackColor = Color.Cornsilk
            Next
            dgv2.Columns(0).ReadOnly = True
            dgv2.Columns(8).ReadOnly = True
            promoTbl = New DataTable
            promoTbl.Columns.Add("Year")
            promoTbl.Columns.Add("Comment")
            promoTbl.Columns.Add("eDate")
            con.Open()
            sql = "SELECT Year_Id, eDate, Comment FROM Promotions WHERE Week_Id = " & thisWeek & " " & _
                "AND Year_Id >= " & selectedYear - 2 & " AND Str_Id = '" & thisStore & "' ORDER BY Year_Id DESC"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = promoTbl.NewRow
                row(0) = rdr("Year_Id")
                row(1) = rdr("Comment")
                row(2) = rdr("eDate")
                promoTbl.Rows.Add(row)
            End While
            con.Close()
            dgv3.DataSource = promoTbl.DefaultView
            Dim clm As DataGridViewColumn = dgv3.Columns(0)
            clm.Width = 40
            clm = dgv3.Columns(1)
            clm.Width = 1000
            dgv3.Columns(2).Visible = False
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()

        End Try
    End Sub

    ''Public Sub dgv2_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv2.CellValueChanged
    Private Sub dgv2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv2.CellValueChanged
        Try
            Dim amt, planDeptTotal, dayTotal, deptTotal As Integer
            Dim totalRow As DataRow
            If e.RowIndex > 0 Then rowIndex = e.RowIndex
            If e.ColumnIndex > 0 Then columnIndex = e.ColumnIndex
            If rowIndex < 1 Or columnIndex < 1 Then Exit Sub
            oTest = dgv2.Rows(rowIndex).Cells(columnIndex).Value
            If Not IsNothing(oTest) And IsNumeric(oTest) Then
                changesMade = True
                tblRow2 = dayTbl2.Rows(rowIndex)
                totalRow = dayTbl2.Rows.Find("Total")
                oTest = tblRow2(columnIndex)
                amt = CInt(tblRow2(columnIndex))
                For Each row In dayTbl2.Rows
                    If row(0) <> "Total" Then
                        oTest = row(columnIndex)
                        If IsNumeric(oTest) Then
                            amt = CInt(Replace(oTest, ",", ""))
                            dayTotal += amt
                        End If
                    End If
                Next
                planDeptTotal = 0
                deptTotal = 0
                For i As Integer = 1 To 7
                    If IsNumeric(tblRow2(i)) Then
                        amt = CInt(Replace(tblRow2(i), ",", ""))
                        deptTotal += amt
                    End If
                Next
                tblRow2(8) = Format(deptTotal, "###,###,###")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        Try
            If changesMade = False Then Exit Sub
            Dim total As Integer = 0
            Dim planTotalRow As DataRow
            For Each row In dayTbl2.Rows
                oTest = row(8)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(oTest) Then
                        total += CInt(oTest)
                    End If
                End If
            Next
            If total > 0 Then
                MsgBox("Total must be zero before changes can be saved!")
                Exit Sub
            End If
            Dim amt As Integer
            Dim cnt As Integer = -1
            For Each row In dayTbl2.Rows
                cnt += 1
                tblRow = workTbl.Rows.Find(row("Dept"))
                planTotalRow = workTbl.Rows.Find("Total")
                For i As Integer = 1 To 7
                    oTest = row(i)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) And oTest <> 0 Then
                            amt = CInt(oTest)
                            planTotalRow(i) = Format(CInt(Replace(planTotalRow(i), ",", "")) + amt, "###,###,###")
                            thisDept = row("Dept")
                            con.Open()
                            ''sql = "Select Amt * d.Pct AS Amt FROM Sales_Plan s " & _
                            ''    "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.Prd_Id = s.Prd_Id AND c.PrdWk = s.Week_Id " & _
                            ''    "JOIN DayOfWeekPct d ON d.Prd_Id = s.Prd_Id AND Day = " & i & " " & _
                            ''    "WHERE Plan_Id = '" & thisPlan & "' AND s.Dept = '" & thisDept & "' AND c.Week_Id = " & thisWeek & " "
                            sql = "DECLARE @dayAmt int, @planAmt int " & _
                                   "SET @dayAmt = 0 " & _
                                   "SET @dayAmt = (SELECT ISNULL(Amt,0) FROM Day_Sales_Plan WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                                   "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & i & ") " & _
                                   "SET @planAmt = (SELECT s.Amt * d.Pct AS Amt FROM Sales_Plan s " & _
                                   "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.PrdWk = s.Week_Id " & _
                                   "JOIN DayOfWeekPct d ON d.Prd_Id = s.Prd_Id AND Day = " & i & " " & _
                                   "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                                   "AND s.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & ") " & _
                                   "SELECT CASE WHEN @dayAmt <> 0 THEN @dayAmt ELSE @planAmt END AS Amt"
                            cmd = New SqlCommand(sql, con)
                            rdr = cmd.ExecuteReader
                            While rdr.Read
                                oTest = rdr(0)
                                If IsNumeric(oTest) Then amt += rdr(0)
                                dgv1.Item(i, cnt).Style.BackColor = Color.Yellow
                                dgv1.Item(i, cnt).Style.Font = New Font("Sans Serif", 8.25, FontStyle.Bold)
                                tblRow(i) = Format(amt, "###,###,###")
                            End While
                            con.Close()
                            con.Open()
                            sql = "IF NOT EXISTS (SELECT * FROM Day_Sales_Plan WHERE Str_id = '" & thisStore & "' " & _
                                "AND Dept = '" & thisDept & "' AND Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " " & _
                                "AND Week_Id = " & thisWeek & " AND Day = " & i & ") " & _
                                "INSERT INTO Day_Sales_Plan (Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Day, Amt) " & _
                                "SELECT '" & thisStore & "', '" & thisDept & "', " & selectedYear & ", " & thisPeriod & ", " &
                                thisWeek & ", " & i & ", " & amt & " " & _
                                "ELSE " & _
                                "Update Day_Sales_Plan Set Amt = " & amt & " WHERE Str_Id = '" & thisStore & "' " & _
                                "AND Dept = '" & thisDept & "' AND Year_Id = " & selectedYear & " AND Prd_Id = " & thisPeriod & " " & _
                                "AND Week_Id = " & thisWeek & " AND Day = " & i & " "
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                            con.Close()
                        End If
                    End If
                Next
            Next
            For Each row In promoTbl.Rows
                Dim comment As String = row(1).ToString
                con.Open()
                sql = "IF NOT EXISTS (SELECT Comment FROM Promotions WHERE Str_Id = '" & thisStore & "' " & _
                    "AND eDate = '" & CDate(row(2)) & "') " & _
                    "INSERT INTO Promotions (Str_Id, eDate, YrWks, Comment) " & _
                    "SELECT '" & thisStore & "', '" & CDate(row(2)) & "', YrWks, '" & comment & "' " & _
                    "FROM Calendar WHERE eDate = '" & CDate(row(2)) & "' " & _
                    "ELSE " & _
                    "UPDATE Promotions SET Comment = '" & comment & "' WHERE Str_Id = '" & thisStore & "' AND eDate = '" & CDate(row(2)) & "' " & _
                    "AND comment <> '" & comment & "' AND eDate >= '" & CDate(row(2)) & "'"
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            changesMade = False
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cboPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPlan.SelectedIndexChanged
        Try
            thisPlan = cboPlan.SelectedItem
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        Try
            thisStore = cboStore.SelectedItem
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Try
            Dim dte As Date = CDate(DateTimePicker1.Value)
            con.Open()
            sql = "SELECT sDate, eDate FROM Calendar WHERE '" & dte & "' BETWEEN sDate AND eDate"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisSdate = rdr("sDate")
                thisEdate = rdr("eDate")
            End While
            con.Close()
        Catch ex As Exception

        End Try
    End Sub

    Public Function getValue(value As Object) As String
        Dim output As Integer
        Select Case value
            Case IsDBNull(value)
                output = 0
            Case IsNothing(value)
                output = 0
            Case value = ""
                output = 0
            Case IsNumeric(value)
                output = CInt(value)
            Case Len(value) > 0 And Not IsNumeric(value)
                output = CStr(value)
        End Select
        Return output
    End Function

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