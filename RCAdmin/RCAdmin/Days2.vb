Imports System
Imports System.Xml
Imports System.Text
Imports System.Data.SqlClient
Public Class Days2
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared tbl, wkTbl, dayTbl, dayTbl2, promoTbl, workTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = Date.Now
    Public Shared selectedYear, thisDept, thisBuyer, thisClass As String
    Public Shared thisPlan As String = ""
    Public Shared thisStore As String = ""
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, lastYear, thisWeek, thisWeekId As Int16
    Public Shared rowIndex As Int16
    Public Shared weekDates(50) As String
    Public Shared rnd As Double
    Public Shared theValue As String
    Public Shared newDraftAmt, newWeeksTotal, deptTotal As Int32
    Public Shared thisUser = Environment.UserName
    Public Shared adjIn As String
    Public Shared thisPeriod, periodSelected, todaysYear, todaysPeriod, todaysWeekId As Integer
    Public Shared thisSdate, thisEdate, todaysSdate As Date
    Public Shared periods(50) As String
    Public Shared statusColumn As Integer
    Public Shared clmCount, ly, l2y, l3y As Integer
    Public Shared tblRow, tblRow2 As DataRow
    Public Shared changesMade As Boolean = False
    Public Shared formLoaded As Boolean
    Public Shared thisYear As Integer
    Public Shared canBeReset As Boolean = False
    Private Shared canModifyStore As Boolean
    Public Shared foundRow As DataRow

    Private Sub Days2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            formLoaded = False
            Me.Location = New Point(150, 100)
            lblPrdWk.Text = Nothing
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Sales_Plan ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYear.Items.Add(rdr("Year_Id"))
            End While
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Str_Id FROM DayOfWeekPct ORDER BY Str_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStore.Items.Add(rdr("Str_Id"))
            End While
            con.Close()
            cboStore.SelectedIndex = 0
            thisStore = cboStore.SelectedItem

            con.Open()
            sql = "SELECT Year_Id, Prd_Id, Week_Id, sDate FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysYear = rdr("Year_Id")
                todaysPeriod = rdr("Prd_Id")
                todaysWeekId = rdr("Week_Id")
                todaysSdate = rdr("sDate")
            End While
            con.Close()
            Dim x As Integer = cboYear.FindString(todaysYear)
            cboYear.SelectedIndex = x
            x = cboDate.FindString(todaysSdate)
            cboDate.SelectedIndex = x
            formLoaded = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR LOADING FORM")
        End Try
    End Sub
    Private Sub Load_Data()
        Try
            If thisStore = "" And Len(thisPlan) > 0 Then
                MessageBox.Show("SELECT A STORE TO CONTINUE", "LOAD DATA")
                Exit Sub
            End If
            If Len(thisStore) > 0 And thisPlan = "" Then
                MessageBox.Show("SELECT A SALESPLAN TO CONTINUE", "LOAD DATA")
                Exit Sub
            End If

            canModifyStore = True
            changesMade = False
            Dim amt, total As Integer
            Dim totlPct As Decimal = 0
            Dim dept As String = ""
            Dim pct As Double
            Dim row As DataRow
            Dim array(8) As Integer
            Dim array2(8) As Integer
            Dim dte, dt As Date
            Dim day, month As Integer
            Dim dow As String
            Dim firstDay As Integer = Nothing
            dayTbl = New DataTable
            dayTbl.Columns.Clear()
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Dept"
            dayTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = dayTbl.Columns("Dept")
            dayTbl.PrimaryKey = PrimaryKey
            dayTbl.Columns.Add("Description")
            dayTbl2 = New DataTable
            dayTbl2.Columns.Clear()
            Dim column2 = New DataColumn()
            column2.DataType = System.Type.GetType("System.String")
            column2.ColumnName = "Dept"
            dayTbl2.Columns.Add(column2)
            Dim PrimaryKey2(1) As DataColumn
            PrimaryKey2(0) = dayTbl2.Columns("Dept")
            dayTbl2.PrimaryKey = PrimaryKey2
            dte = cboDate.SelectedItem
            For i As Integer = 0 To 6
                dt = DateAdd(DateInterval.Day, i, dte)
                day = DatePart(DateInterval.Day, dt)
                month = DatePart(DateInterval.Month, dt)
                dow = dt.ToString("ddd")
                dayTbl.Columns.Add(dow)
                dayTbl2.Columns.Add(dow)
            Next
            dayTbl.Columns.Add("Total")
            dayTbl2.Columns.Add("Total")
            row = dayTbl.NewRow
            row(0) = "% WK"
            dayTbl.Rows.Add(row)

            ''con.Open()
            ''sql = "SELECT Day, Pct FROM DayOfWeekPct WHERE Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " "
            ''cmd = New SqlCommand(sql, con)
            ''rdr = cmd.ExecuteReader
            ''row = dayTbl.Rows(0)
            ''While rdr.Read
            ''    day = rdr("Day")
            ''    pct = rdr("Pct")
            ''    If IsNumeric(day) Then
            ''        totlPct += pct
            ''        row(day + 1) = Format(pct, "###.0%")
            ''    End If
            ''End While
            ''row("Total") = Format(totlPct, "###.0%")
            ''con.Close()

            con.Open()
            sql = "SELECT ID AS Dept, Description FROM Departments WHERE Status = 'Active' ORDER BY Dept"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = dayTbl.NewRow
                oTest = rdr("Dept")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then thisDept = oTest
                total = 0
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest
                row("Dept") = dept
                oTest = rdr("Description")
                row("Description") = rdr("Description")
                For i As Integer = 0 To 6
                    con2.Open()
                    sql = "DECLARE @dayAmt int, @planAmt int, @flag varchar(1) " & _
                        "SET @dayAmt = (SELECT ISNULL(Amt,0) FROM Day_Sales_Plan WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & i + 1 & ") " & _
                        "SET @flag = (SELECT CASE WHEN @dayAmt IS NULL THEN 'N' ELSE 'Y' END) " & _
                        "SET @planAmt = (SELECT DISTINCT s.Amt * d.Pct AS Amt FROM Sales_Plan s " & _
                        "JOIN Calendar c ON c.Year_Id = s.Year_Id AND c.PrdWk = s.Week_Id AND c.Week_Id > 0 " & _
                        "JOIN DayOfWeekPct d ON d.Year_Id = s.Year_Id AND d.Prd_Id = s.Prd_Id AND Day = " & i + 1 & " " & _
                            "AND s.Str_Id = d.Str_Id " & _
                        "WHERE Plan_Id = '" & thisPlan & "' AND s.Str_Id = '" & thisStore & "' AND Dept = '" & dept & "' " & _
                        "AND s.Prd_Id = " & thisPeriod & " AND s.Week_Id = " & thisWeek & ") " & _
                        "SELECT CASE WHEN @dayAmt IS NULL THEN @planAmt ELSE @dayAmt END AS Amt, @flag AS Flag"
                    cmd = New SqlCommand(sql, con2)
                    rdr2 = cmd.ExecuteReader
                    While rdr2.Read
                        oTest = rdr2("Amt")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then amt = CInt(oTest) Else amt = 0
                        oTest = rdr2("Flag")
                        If oTest = "Y" Then
                            canBeReset = True
                        End If
                        total += amt
                        row(i + 2) = Format(amt, "###,##0")
                        array(i) += amt
                        array(7) += amt
                    End While
                    con2.Close()
                Next
                row("Total") = Format(total, "###,###")
                dayTbl.Rows.Add(row)
            End While
            con.Close()
            Dim pctRow As DataRow = dayTbl.Rows(0)
            row = dayTbl.NewRow
            row(0) = "Total"
            For i As Integer = 0 To 7
                row(i + 2) = Format(array(i), "###,###0")
                pctRow(i + 2) = Format(array(i) / array(7), "###.0%")
            Next
            dayTbl.Rows.Add(row)
            row = dayTbl.Rows(0)
            '
            '     Determine if any changes have been made at the department level and set firstI accordingly
            '
            Dim firstI As Integer = 0
            con.Open()
            sql = "SELECT DISTINCT Day, Pct INTO #t1 FROM Day_Sales_Plan WHERE Year_Id = " & selectedYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " " & _
                "SELECT Day, COUNT(*) AS Count INTO #t2 FROM #t1 GROUP BY Day " & _
                "SELECT MAX(Count) AS Count FROM #t2"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Count")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    firstI = 1
                    canModifyStore = False
                End If
            End While
            con.Close()

            workTbl = dayTbl
            dgv1.DataSource = dayTbl.DefaultView
            For i As Integer = firstI To dgv1.Rows.Count - 2
                dgv1.Rows(i).Cells(0).Style.ForeColor = Color.Blue
                dgv1.Rows(i).Cells(0).Style.Font = New Font("Ariel", 8, FontStyle.Underline)
            Next
            dgv1.AutoResizeColumns()
            dgv1.Rows(0).DefaultCellStyle.BackColor = Color.LightGray

            For i = 1 To dgv1.ColumnCount - 1
                If i > 1 Then
                    dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
                dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            dgv1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgv1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv1.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv1.Columns(0).ReadOnly = False
            promoTbl = New DataTable
            promoTbl.Columns.Add("Year")
            promoTbl.Columns.Add("Comment")
            promoTbl.Columns.Add("eDate")
            con.Open()
            sql = "SELECT Year_Id, eDate, Comment FROM Promotions WHERE Week_Id = " & thisWeekId & " " & _
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
            clm.Width = 60
            clm = dgv3.Columns(1)
            clm.Width = 1000
            dgv3.Columns(2).Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD DATA")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()

        End Try
    End Sub

    Private Sub tb_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) And Not Char.IsDigit(e.KeyChar) And e.KeyChar <> "." And e.KeyChar <> "-" Then
            e.Handled = True
        End If
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
                    If columnIdx = 0 Then
                        If rowIdx = 0 Then
                            If Not canModifyStore Then
                                MessageBox.Show("Store level modifications not permitted once department percentages have been modified.",
                                                "DEPARTMENT LEVEL MODIFICATIONS DETECTED!")
                                Exit Sub
                            End If
                            lblProcessing.Visible = True
                            Me.Refresh()
                            foundRow = dayTbl(0)
                            If canModifyStore Then Days4.Show()
                            lblProcessing.Visible = False
                        Else
                            lblProcessing.Visible = True
                            Me.Refresh()
                            thisDept = oTest
                            foundRow = dayTbl.Rows.Find(thisDept)
                            oTest = foundRow(9)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                If IsNumeric(oTest) Then
                                    deptTotal = CInt(oTest)
                                End If
                            End If
                            Days3.Show()
                            lblProcessing.Visible = False
                        End If
                    End If
                End If

            Case Windows.Forms.MouseButtons.Right

        End Select

    End Sub
    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Try
            Dim dayTotal As Integer
            Dim x As String
            Dim pct As Decimal
            Dim mltplr As Decimal = 0
            Dim wkpct As Decimal = 0
            If e.RowIndex > 0 Then rowIndex = e.RowIndex
            If e.ColumnIndex > 0 Then columnIndex = e.ColumnIndex
            oTest = dgv1.Rows(rowIndex).Cells(columnIndex).Value
            If IsNumeric(oTest) Then mltplr = CDec(oTest * 0.01)
            If rowIndex = 0 And columnIndex > 1 And columnIndex < 9 Then
                Dim lr As Integer = dayTbl.Rows.Count - 1
                If IsNumeric(oTest) Then
                    changesMade = True
                    For r As Integer = 1 To lr - 1
                        Dim thiscell As Integer = dayTbl.Rows(r).Item(columnIndex)
                        dayTotal += CInt(mltplr * dayTbl.Rows(r).Item(9))
                        dayTbl.Rows(r).Item(columnIndex) = Format(CInt(mltplr * dayTbl.Rows(r).Item(9)), "###,###,###")
                        dayTbl.Rows(0).Item(columnIndex) = Format(mltplr, "###.0%")
                    Next
                    For c As Integer = 2 To 8
                        x = Replace(dayTbl.Rows(0).Item(c), "%", "").ToString
                        pct = Convert.ToDecimal(x)
                        wkpct += (pct * 0.01)
                    Next
                    dayTbl.Rows(lr).Item(columnIndex) = Format(dayTotal, "###,###,###")
                    dayTbl.Rows(0).Item(9) = Format(wkpct, "###.0%")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR ENTERING DATA IN DATAGRIDVIEW", "dgv2_CellEnter")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub


    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            thisStore = cboStore.SelectedItem
            If formLoaded Then Call Load_Data()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SELECT STORE")
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

    Private Sub cboDate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDate.SelectedIndexChanged
        Try
            thisSdate = cboDate.SelectedItem
            con.Open()
            sql = "SELECT Year_Id, Prd_Id, Week_Id, PrdWk FROM Calendar WHERE '" & thisSdate & "' BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                selectedYear = rdr("Year_Id")
                thisPeriod = rdr("Prd_Id")
                thisWeek = rdr("PrdWk")
                thisWeekId = rdr("Week_Id")
            End While
            con.Close()
            lblPrdWk.Text = "[Period " & thisPeriod & " Week " & thisWeekId & "]"
            Me.Refresh()
            Call Load_Data()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYear.SelectedIndexChanged
        thisYear = cboYear.SelectedItem
        con.Open()
        sql = "SELECT DISTINCT Plan_Id FROM Sales_Plan WHERE Status = 'Active' and Year_Id= " & thisYear & " "
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            thisPlan = rdr("Plan_Id")
        End While
        con.Close()
        cboDate.Items.Clear()
        con.Open()
        sql = "SELECT sDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Week_Id > 0 ORDER BY sDate DESC"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboDate.Items.Add(rdr("sDate"))
        End While
        con.Close()
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs)
        Try
            Dim rows As Integer = dgv1.Rows.Count
            Dim colms As Integer = dgv1.Columns.Count
            For r As Integer = 0 To rows - 1
                For colm As Integer = 1 To 7
                    oTest = dgv1.Rows(r).Cells(colm).Style.ForeColor.Name
                    If oTest = "Red" Then
                        con.Open()
                        sql = "DELETE FROM Day_Sales_Plan WHERE Str_Id ='" & thisStore & "' AND Year_Id = " & thisYear & " " & _
                            "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & colm & " "
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                Next
            Next
            Call Load_Data()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "RESET TO ACTIVE PLAN VALUES")
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