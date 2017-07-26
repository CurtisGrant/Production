Imports System.Data.SqlClient
Public Class DayPct2
    Public Shared con As SqlConnection
    Public Shared sql As String
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared tbl As DataTable
    Public Shared thisDept As String
    Public Shared thisYear As Integer
    Public Shared thisPeriod As Integer
    Public Shared thisYrPrd As Integer
    Public Shared todaysPeriod As Integer
    Public Shared todaysYrPrd As Integer
    Public Shared thisWeek As Integer
    Public Shared thisSdate, todaysSdate, maxEdate As Date
    Public Shared thisStore As String
    Public Shared foundRow As DataRow
    Public Shared oTest As Object
    Public Shared changesMade As Boolean
    Private Sub DayPct2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblProcessing.Visible = True

        Call Load_Data()
        lblProcessing.Visible = False
    End Sub

    Private Sub Load_data()
        con = DayPct.con
        thisYear = DayPct.thisYear
        thisPeriod = DayPct.thisPeriod
        thisYrPrd = CInt((thisYear - 2000) & thisPeriod)
        todaysYrPrd = DayPct.todaysYrPrd
        todaysPeriod = DayPct.todaysPeriod
        thisStore = DayPct.thisStore
        foundRow = DayPct.foundRow
        changesMade = False
        lblServer.Text = MainMenu.serverLabel.Text

        con.Open()
        sql = "SELECT MAX(eDate) AS eDate FROM Calendar WHERE YrPrd < " & todaysYrPrd & " AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            maxEdate = rdr("eDate")
        End While
        con.Close()

        tbl = New DataTable
        tbl.Columns.Add("Period " & thisPeriod)
        tbl.Columns.Add("Sun")
        tbl.Columns.Add("Mon")
        tbl.Columns.Add("Tue")
        tbl.Columns.Add("Wed")
        tbl.Columns.Add("Thu")
        tbl.Columns.Add("Fri")
        tbl.Columns.Add("Sat")
        tbl.Columns.Add("Total")
        Dim row As DataRow
        row = tbl.NewRow
        row(0) = "Values"
        Dim totl As Integer = 0
        For i = 1 To 8
            row(i) = foundRow(i)
        Next
        tbl.Rows.Add(row)
        row = tbl.NewRow
        tbl.Rows.Add("New Percentage")
        tbl.Rows.Add("")
        con.Open()
        sql = "SELECT Year_Id, Day AS DOW, Pct FROM DayOfWeekPct WHERE Str_Id = '" & thisStore & "' AND Year_Id < " & thisYear & " " & _
            "AND Prd_Id = " & thisPeriod & " ORDER BY Year_Id DESC, DOW"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        Dim yr, dw, sold, yrs As Integer
        Dim pct, totalPct, avg, sat As Decimal
        Dim prevYear As Integer = 0
        totl = 0
        totalPct = 0
        Dim arry(3, 8) As Decimal
        row = tbl.NewRow
        While rdr.Read
            yr = rdr("Year_Id")
            If prevYear = 0 Then prevYear = yr
            If yr <> prevYear Then
                sat = 1 - totalPct
                row(7) = Format(sat, "##0.0%")
                row(8) = Format(totalPct + sat, "##0.0%")
                tbl.Rows.Add(row)
                row = tbl.NewRow
                prevYear = yr
                totl = 0
                totalPct = 0
                yrs += 1
            End If
            row(0) = prevYear
            dw = CInt(rdr("DOW"))
            ''sold = CInt(rdr("Sold"))
            pct = rdr("Pct")
            totl += sold
            If dw < 7 Then totalPct += pct
            arry(yrs, 7) += pct
            arry(yrs, dw - 1) += pct
            If pct > 0 Then row(dw) = Format((pct), "##0.0%")
        End While
        con.Close()
        row(8) = Format(Math.Ceiling(totalPct), "##0.0%")
        yrs += 1
        If yrs < 4 Then tbl.Rows.Add(row)
        If yrs > 3 Then yrs = 3

        row = tbl.NewRow
        row(0) = "Average"
        Dim tot As Decimal
        For i = 1 To 8
            Dim test1 As Decimal = arry(0, i - 1)
            Dim test2 As Decimal = arry(1, i - 1)
            Dim test3 As Decimal = arry(2, i - 1)
            tot = (arry(0, i - 1) + arry(1, i - 1) + arry(2, i - 1))
            avg = tot / yrs
            If avg > 0 Then row(i) = Format(avg, "##0.0%")
        Next
        tbl.Rows.Add(row)
        dgv1.DataSource = tbl.DefaultView
        dgv1.AutoResizeColumns()
        For i = 1 To 8
            If i < 8 Then dgv1.Rows(1).Cells(i).Style.BackColor = Color.Cornsilk
            dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next i
        dgv1.Rows(0).ReadOnly = True
        For i = 2 To dgv1.Rows.Count - 1
            dgv1.Rows(i).ReadOnly = True
        Next
        dgv1.Rows(1).Cells(0).ReadOnly = True
        dgv1.Rows(1).Cells(8).ReadOnly = True
        dgv1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        'Try
        Dim pct, adjTotal, planPct, planTotal As Decimal
        Dim adjRow, planRow As DataRow
        If e.RowIndex < 1 Or e.ColumnIndex < 1 Then Exit Sub
        oTest = dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        If Not IsNothing(oTest) And IsNumeric(oTest) Then
            adjRow = tbl.Rows(e.RowIndex)
            planRow = tbl.Rows(0)
            planTotal = Replace(planRow(8), "%", "")
            If thisYrPrd <= todaysYrPrd Then
                MessageBox.Show("CANNOT CHANGE PLAN FOR CURRENT OR PRIOR PERIOD!", "ENTER ADJUSTMENT PERCENTAGES")
                adjRow(e.ColumnIndex) = Nothing
                Exit Sub
            End If
            changesMade = True
            pct = CInt(oTest) * 0.01
            adjRow(e.ColumnIndex) = Format(pct, "##0.0%")
            oTest = planRow(e.ColumnIndex)
            If IsDBNull(oTest) Then oTest = 0
            planPct = CDec(Replace(oTest, "%", "") * 0.01)
            planPct += pct
            ''planRow(e.ColumnIndex) = Format(pct, "##0%")
            ''planRow(8) = Format(planPct, "##0%")
            adjTotal = 0
            For i As Integer = 1 To 7
                oTest = adjRow(i)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    oTest = Replace(oTest, "%", "")
                    If IsNumeric(oTest) Then
                        pct = CDec(oTest) * 0.01
                        adjTotal += pct
                    End If
                End If
            Next
            ''planRow(8) = Format(planPct, "##0%")

            adjRow(8) = Format(adjTotal, "##0.0%")
            dgv1.Refresh()
        End If
        'Catch ex As Exception
        '    MessageBox.Show("ERROR ENTERING DATA IN DATAGRIDVIEW", "dgv1_CellEnter")
        '    If con.State = ConnectionState.Open Then con.Close()
        'End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        lblProcessing.Visible = True
        Me.Refresh()

        Dim pct As Decimal
        Dim adjRow As DataRow = tbl.Rows(1)
        oTest = adjRow(8)
        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
            oTest = Replace(oTest, "%", "")
            If oTest <> 100 Then
                MessageBox.Show("Total Adjustments must equal 100%", "SAVE CHANGES")
                Exit Sub
            End If
        End If
        For i = 1 To 7
            foundRow(i) = adjRow(i)
            oTest = adjRow(i)

            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                oTest = Replace(oTest, "%", "")
                pct = CDec(oTest) * 0.01
                con.Open()
                sql = "IF NOT EXISTS (SELECT * FROM DayOfWeekPct WHERE Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod &
                    " AND Str_Id = '" & thisStore & "' AND Day = " & i & ") " & _
                    "INSERT INTO DayOfWeekPct (Year_Id, Str_Id, Prd_Id, Day, Pct) " & _
                    "SELECT " & thisYear & ", '" & thisStore & "', " & thisPeriod & ", " & i & ", " & pct & " " & _
                    "ELSE " & _
                    "UPDATE DayOfWeekPct SET Pct = " & pct & " WHERE Str_Id = '" & thisStore & "' AND Year_Id = " &
                        thisYear & " AND Prd_Id = " & thisPeriod & " AND Day = " & i & " "
                cmd = New SqlCommand(sql, con)
                cmd.ExecuteNonQuery()
            End If
            con.Close()
        Next

        changesMade = False
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub
End Class