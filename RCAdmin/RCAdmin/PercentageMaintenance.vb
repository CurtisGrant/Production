Imports System.Data.SqlClient
Imports System.Xml
Public Class PercentageMaintenance
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared wkTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = Date.Now
    Public Shared lastYear As Integer = DatePart(DateInterval.Year, Today) - 1
    Public Shared thisYear, thisStore, thisPeriod, thisPlan, thisDept, thisBuyer, thisClass, thisTable, thisField As String
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, rowIndex, tRows As Int16
    Public Shared tNew As Decimal
    Public Shared theValue As String
    Public Shared thisUser = Environment.UserName
    Public Shared planBy As String
    Public Shared periodSelected, numberPeriods As Integer
    Public Shared buyers(50) As String
    Public Shared changesMade As Boolean
    Private Sub PercentageMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            Me.Location = New Point(300, 200)
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            'periodSelected = PlanMaintenance.periodSelected
            con = New SqlConnection(conString)
            Dim cnt As Integer = 0

            con.Open()
            sql = "SELECT DISTINCT Plan_Year FROM Buyer_PCT ORDER BY Plan_Year"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYear.Items.Add(rdr(0))
            End While
            con.Close()
            cboYear.SelectedIndex = 0
            thisPeriod = 1
            con.Open()
            sql = "SELECT ID FROM Stores ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStore.Items.Add(rdr(0))
            End While
            con.Close()
            cboStore.SelectedIndex = 0
            thisStore = cboStore.SelectedItem
            con.Open()
            sql = "SELECT ID FROM Departments ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr(0)
                cboDept.Items.Add(rdr(0))
            End While
            con.Close()
            cboDept.SelectedIndex = 0
            thisDept = cboDept.SelectedItem
            cnt = 0
            con.Open()
            sql = "SELECT ID FROM Buyers ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                buyers(cnt) = rdr(0)
                cnt += 1
            End While
            con.Close()
            Array.Resize(buyers, cnt)
            ''thisPlan = cboYr.SelectedItem
            thisYear = cboYear.SelectedItem

            If rdoBuyer.Checked = True Then
                thisTable = "Buyer_PCT"
                thisField = "Buyer"
            Else
                thisTable = "DayOfWeekPCT"
                thisField = "Day"
            End If
            numberPeriods = 0
            con.Open()
            sql = "SELECT DISTINCT Prd_Id FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id > 0 ORDER BY Prd_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                ' cboPeriod.Items.Add(rdr(0))
                numberPeriods += 1
            End While
            con.Close()
            'cboPeriod.SelectedIndex = 0
            'thisPeriod = cboPeriod.SelectedItem
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''Public Sub dgv1_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        '' MsgBox("you're in cell enter row " & rowIndex & " column " & columnIndex)
        Dim change As Decimal
        If rdoBuyer.Checked = True And columnIndex = 7 And e.RowIndex < tRows Then
            oTest = dgv1.Rows(rowIndex).Cells(7).Value
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                change = oTest * 0.01
                tNew = 0
                dgv1.Rows(rowIndex).Cells(6).Value = Format(change, "###.0%")
                For i As Integer = 0 To tRows - 1
                    oTest = dgv1.Rows(i).Cells(6).Value
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        changesMade = True
                        oTest = Replace(oTest, "%", "")
                        If oTest <> "" And IsNumeric(oTest) Then
                            oTest = CDec(oTest)
                            tNew += oTest
                        End If
                    End If
                Next
            End If
        End If
        If rdoDay.Checked = True Then
            If e.RowIndex < tRows Then
                oTest = dgv1.Rows(rowIndex).Cells(columnIndex).Value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    change = CDec(Replace(oTest, "%", "")) * 0.01
                    tNew = 0
                    Dim findThis As String = "%"
                    Dim foundIt As Integer = oTest.indexof(findThis)
                    If foundIt = 0 Then
                        dgv1.Rows(rowIndex).Cells(columnIndex).Value = Format(change, "###.0%")
                    End If
                    For i As Integer = 1 To 7
                        oTest = dgv1.Rows(rowIndex).Cells(i).Value
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            changesMade = True
                            tNew += CDec(Replace(oTest, "%", ""))
                        End If
                    Next
                    dgv1.Rows(rowIndex).Cells(8).Value = Format(tNew * 0.01, "###.0%")
                End If
            End If
        End If
        If rdoBuyer.Checked = True Then
            dgv1.Rows(tRows).Cells(6).Value = Format(tNew * 0.01, "###.0%")
            dgv1.Rows(tRows + 1).Cells(6).Value = Format(1 - tNew * 0.01, "###.0%")
        Else
            Dim dgvrow As DataGridViewRow = Me.dgv1.Rows(rowIndex)
            If tNew < 99.6 Or tNew > 100.4 Then
                dgvrow.DefaultCellStyle.BackColor = Color.Salmon
            Else
                dgvrow.DefaultCellStyle.BackColor = Color.White
            End If
        End If
        Me.Refresh()
    End Sub

    Public Sub dgv1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim ht As DataGridView.HitTestInfo
            ht = Me.dgv1.HitTest(e.X, e.Y)
            Dim rowIdx As Int16 = ht.RowIndex
            Dim columnIdx As Int16 = ht.ColumnIndex
            '' MsgBox("you're in mousedown row " & rowIdx & " column " & columnIdx)
            If columnIdx = 0 And rdoBuyer.Checked = True Then
                thisBuyer = dgv1.Rows(rowIdx).Cells(0).Value
                Classes.Show()
            End If
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        If rdoBuyer.Checked = True Then
            Dim message, title As String
            ' Dim r As Integer = numberPeriods - cboPeriod.SelectedIndex








            Dim r As Integer = 1















            Dim rx As Integer = r
            message = "Percentages can be saved to the next " & r & " periods. Click OK button to save screen to next " & r & " periods "
            message &= "or enter a number up to " & r & " to save to instead."
            title = "Save Percentages"
            oTest = InputBox(message, title)
            If oTest <> "" Then
                rx = CInt(oTest)
                If rx > r Then
                    MsgBox("The number can't exceed " & r & "!")
                    Exit Sub
                End If
            Else : Exit Sub
            End If
            For i As Integer = thisPeriod To rx
                For Each row In wkTbl.Rows
                    oTest = row(6)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If Not IsNothing(oTest) And Not IsDBNull(oTest) And row(0) <> "Difference" And row(0) <> "Total" Then
                            thisField = row(0).ToString
                            theValue = CDec(Replace(oTest, "%", "") * 0.01)
                            con.Open()
                            sql = "IF NOT EXISTS (SELECT * FROM Buyer_PCT WHERE Plan_Year = '" & thisYear & "' " & _
                                    "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                    "AND Buyer = '" & thisField & "') " & _
                                "INSERT INTO Buyer_Pct (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, PCT) " & _
                                    "SELECT '" & thisYear & "', '" & thisStore & "', " & i & ", '" & thisDept & "', '" & _
                                    thisField & "', " & theValue & " " & _
                                    "ELSE " & _
                                "UPDATE Buyer_Pct SET PCT = " & theValue & " WHERE Plan_Year = '" & thisYear & "' " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                "AND Buyer = '" & thisField & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                            con.Close()
                        End If
                    End If
                Next
            Next
        Else
            Dim total, pct As Decimal
            For Each row In wkTbl.Rows
                total = 0
                thisField = row(0)           ' Period
                For Day As Integer = 1 To 7
                    oTest = row(Day)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        total += CDec(Replace(oTest, "%", ""))
                    End If
                Next
                If total < 99.6 Or total > 100.4 Then
                    MsgBox(total & " - Percentage distribution across days must equal 100%. Period " & row(0) & " was not saved.")
                Else
                    For Day As Integer = 1 To 7
                        oTest = row(Day)
                        pct = Nothing
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then pct = CDec(Replace(oTest, "%", "")) * 0.01
                        con.Open()
                        sql = "IF NOT EXISTS (SELECT * FROM DayOfWeekPct WHERE Prd_Id = " & row(0) & " AND Day = " & Day & ") " & _
                            "INSERT INTO DayOfWeekPct (Prd_Id, Day, Pct) " & _
                            "SELECT " & row(0) & ", " & Day & ", " & pct & " " & _
                            "ELSE " & _
                            "UPDATE DayOfWeekPct SET Pct = " & pct & " " & _
                            "WHERE Prd_Id = " & row(0) & " AND Day = " & Day & ""
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        con.Close()
                        'End If
                    Next
                End If
            Next
        End If
    End Sub

    Private Sub rdoBuyer_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBuyer.CheckedChanged
        If rdoBuyer.Checked = True Then
            planBy = "Buyer"
            rdoDay.Checked = False
        End If
    End Sub

    Private Sub rdoDay_CheckedChanged(sender As Object, e As EventArgs) Handles rdoDay.CheckedChanged
        If rdoDay.Checked = True Then
            planBy = "Day"
            rdoBuyer.Checked = False
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
        MainMenu.Show()

    End Sub

    Private Sub btnEdit_1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs)
        thisDept = cboDept.SelectedItem
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs)
        thisStore = cboStore.SelectedItem
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            changesMade = False
            Dim tAvg52 As Decimal = 0
            Dim tAvg26 As Decimal = 0
            Dim tAvg12 As Decimal = 0
            Dim tAvg4 As Decimal = 0
            Dim tCurrent As Decimal = 0
            Dim tInv As Decimal = 0
            Dim firstRecord As Boolean = True

            Dim heading, tblName, thisField As String
            If rdoBuyer.Checked Then
                heading = "Buyer"
                tblName = "Buyer_PCT"
                Dim row As DataRow
                wkTbl = New DataTable
                wkTbl.Clear()
                Dim column = New DataColumn()
                column.DataType = System.Type.GetType("System.String")
                column.ColumnName = heading
                wkTbl.Columns.Add(column)
                Dim PrimaryKey(1) As DataColumn
                PrimaryKey(0) = wkTbl.Columns(heading)
                wkTbl.PrimaryKey = PrimaryKey
                wkTbl.Columns.Add("52 Week Average")
                wkTbl.Columns.Add("26 Week Average")
                wkTbl.Columns.Add("12 Week Average")
                wkTbl.Columns.Add("4 Week Average")
                wkTbl.Columns.Add("Current Plan%")
                wkTbl.Columns.Add("New Plan%")
                wkTbl.Columns.Add("Enter New %")
                con.Open()
                sql = "SELECT ID AS Buyer FROM Buyers WHERE Status = 'Active' ORDER BY ID"
                wkTbl.Columns.Add("Current OnHand %")
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    thisField = rdr(heading)
                    con2.Open()
                    'If heading = "Buyer" Then
                    sql = "DECLARE @wk52 date, @wk26 date, @wk12 date, @wk4 date " & _
                    "SELECT @wk52 = DATEADD(week,-52,GETDATE()) " & _
                    "SELECT @wk26 = DATEADD(week,-26,GETDATE()) " & _
                    "SELECT @wk12 = DATEADD(week,-12,GETDATE()) " & _
                    "SELECT @wk4 = DATEADD(week,-4,GETDATE()) " & _
                    "CREATE TABLE #t1 (" & heading & " varchar(20) NOT NULL, Avg52 decimal(18,4)) INSERT INTO #t1 (Buyer, Avg52) " & _
                    "SELECT " & heading & ", SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                        "WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg52 " & _
                        "FROM Item_Sales WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & thisDept & "' " & _
                            "AND " & heading & " = '" & thisField & "' GROUP BY " & heading & " " & _
                    "CREATE TABLE #t2 (" & heading & " varchar(20) NOT NULL, Avg26 decimal(18,4)) INSERT INTO #t2 (Buyer, Avg26) " & _
                    "SELECT " & heading & ", SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                        "WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg26 " & _
                        "FROM Item_Sales WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' " & _
                            " AND Dept = '" & thisDept & "' " & _
                            "AND " & heading & " = '" & thisField & "' GROUP BY " & heading & " " & _
                    "CREATE TABLE #t3 (" & heading & " varchar(20) NOT NULL, Avg12 decimal(18,4)) INSERT INTO #t3 (Buyer, Avg12) " & _
                    "SELECT " & heading & ", SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                        "WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg12 " & _
                        "FROM Item_Sales WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & thisDept & "' " & _
                            "AND " & heading & " = '" & thisField & "' GROUP BY " & heading & " " & _
                    "CREATE TABLE #t4 (" & heading & " varchar(20) NOT NULL, Avg4 decimal(18,4)) INSERT INTO #t4 (Buyer, Avg4) " & _
                    "SELECT " & heading & ", SUM(Act_Sales) / (SELECT ISNULL(SUM(Act_Sales),0) FROM Item_Sales " & _
                        "WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "') AS Avg4 " & _
                        "FROM Item_Sales WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & thisDept & "' " & _
                            "AND " & heading & " = '" & thisField & "' GROUP BY " & heading & " " & _
                    "CREATE TABLE #t5 (" & heading & " varchar(20) NOT NULL, pctInv decimal(18,4)) INSERT INTO #t5 (Buyer, pctInv) " & _
                    "SELECT Buyer, SUM(Act_Inv_Retail) / (SELECT SUM(Act_Inv_Retail) FROM Item_Sales " & _
                        "WHERE CONVERT(VARCHAR,GETDATE(),101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "') AS pctInv " & _
                        "FROM Item_Sales w INNER JOIN Buyers b ON b.ID = w.Buyer " & _
                            "WHERE CONVERT(VARCHAR,GETDATE(),101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "' AND Status = 'Active' " & _
                            "GROUP BY Buyer " & _
                    "SELECT #t1." & heading & ", #t1.Avg52, #t2.Avg26, #t3.Avg12, #t4.Avg4, " & _
                        "(SELECT PCT FROM " & heading & "_PCT p WHERE Plan_Year = '" & thisPlan & "' " & _
                            "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " " & _
                            "AND p." & heading & " = #t1." & heading & " " & _
                        "AND Dept = '" & thisDept & "') AS Prd1, #t5.pctInv " & _
                        "FROM #t1 JOIN #t2 ON #t2." & heading & " = #t1." & heading & " " & _
                            "JOIN #t3 ON #t3." & heading & " = #t1." & heading & " " & _
                            "JOIN #t4 ON #t4." & heading & " = #t1." & heading & " " & _
                            "JOIN #t5 ON #t5.Buyer = #t1." & heading & " ORDER BY " & heading & ""

                    cmd = New SqlCommand(sql, con2)
                    rdr2 = cmd.ExecuteReader

                    While rdr2.Read
                        oTest = rdr2(0)
                        'If heading = "Day" And firstRecord Then
                        '    If CInt(oTest) > 1 Then
                        '        For i As Integer = 1 To CInt(oTest) - 1
                        '            row = wkTbl.NewRow
                        '            row(0) = i
                        '            wkTbl.Rows.Add(row)
                        '            firstRecord = False
                        '        Next
                        '    End If
                        'End If
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            row = wkTbl.NewRow
                            oTest = rdr2(0)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row(0) = rdr2(0)
                            oTest = rdr2("Avg52")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                row(1) = Format(rdr2(1), "###.0%")
                                tAvg52 += CDec(oTest)
                            End If
                            oTest = rdr2("Avg26")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                row(2) = Format(rdr2(2), "###.0%")
                                tAvg26 += CDec(oTest)
                            End If
                            oTest = rdr2("Avg12")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                row(3) = Format(rdr2(3), "###.0%")
                                tAvg12 += CDec(oTest)
                            End If
                            oTest = rdr2("Avg4")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                row(4) = Format(rdr2(4), "###.0%")
                                tAvg4 += rdr2(4)
                            End If

                            oTest = rdr2("Prd1")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                row(5) = Format(rdr2(5), "###.0%")
                                tCurrent += rdr2(5)
                            End If
                            oTest = rdr2("pctInv")
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                row(8) = Format(oTest, "###.0%")
                                tInv += rdr2(6)
                            End If
                            wkTbl.Rows.Add(row)
                        End If
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
                row(8) = Format(Math.Round(tInv, 1, MidpointRounding.AwayFromZero), "###.#%")
                wkTbl.Rows.Add(row)
                tRows = wkTbl.Rows.Count - 1
                row = wkTbl.NewRow
                row(0) = "Difference"
                wkTbl.Rows.Add(row)
            Else                                                                       ' Day of Week is checked
                firstRecord = False
                Dim prevPreiod As Integer = 0
                Dim dow As Integer
                Dim row As DataRow
                wkTbl = New DataTable
                wkTbl.Clear()
                Dim column As New DataColumn
                column.DataType = System.Type.GetType("System.String")
                column.ColumnName = "Period"
                wkTbl.Columns.Add(column)
                Dim PrimaryKey(1) As DataColumn
                PrimaryKey(0) = wkTbl.Columns("Period")
                wkTbl.PrimaryKey = PrimaryKey
                wkTbl.Columns.Add("SUN")
                wkTbl.Columns.Add("MON")
                wkTbl.Columns.Add("TUE")
                wkTbl.Columns.Add("WED")
                wkTbl.Columns.Add("THU")
                wkTbl.Columns.Add("FRI")
                wkTbl.Columns.Add("SAT")
                wkTbl.Columns.Add("Total")
                con.Open()
                sql = "SELECT DISTINCT Prd_Id FROM DayOFWeekPct"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    row = wkTbl.NewRow
                    row(0) = rdr("Prd_Id")
                    wkTbl.Rows.Add(row)
                    tRows = wkTbl.Rows.Count - 1
                End While
                con.Close()

                con.Open()
                sql = "SELECT Prd_Id, Day, Pct FROM DayOfWeekPct ORDER BY Prd_ID, Day"
                cmd = New SqlCommand(sql, con)
                rdr = cmd.ExecuteReader
                While rdr.Read
                    oTest = rdr("Prd_Id")
                    row = wkTbl.Rows.Find(oTest)
                    dow = rdr("Day")
                    row(dow) = Format(rdr("Pct"), "###.0%")
                End While
                con.Close()
            End If
            dgv1.DataSource = wkTbl.DefaultView
            dgv1.AutoResizeColumns()
            dgv1.Columns(0).ReadOnly = True
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Format = "p2"
                If i = 7 And rdoBuyer.Checked = True Then dgv1.Columns(7).DefaultCellStyle.BackColor = Color.Cornsilk
                If rdoBuyer.Checked = True Then dgv1.Columns(i).ReadOnly = True
            Next
            tNew = 0
            dgv1.Columns(7).ReadOnly = False
            dgv1.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
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