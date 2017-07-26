Imports System.Data.SqlClient
Public Class Days3
    Public Shared con As SqlConnection = Days2.con
    Public Shared sql As String
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared tbl As DataTable
    Public Shared thisDept As String
    Public Shared thisYear As Integer
    Public Shared thisPeriod As Integer
    Public Shared thisWeek As Integer
    Public Shared thisSdate As Date
    Public Shared todaysSdate As Date
    Public Shared thisStore As String
    Public Shared foundRow As DataRow
    Public Shared oTest As Object
    Public Shared changesMade As Boolean
    Public Shared deptTotal As Int32

    Private Sub Days3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            thisDept = Days2.thisDept
            thisYear = Days2.thisYear
            thisPeriod = Days2.thisPeriod
            thisWeek = Days2.thisWeek
            thisSdate = Days2.thisSdate
            todaysSdate = Days2.todaysSdate
            thisStore = Days2.thisStore
            foundRow = Days2.foundRow
            deptTotal = Days2.deptTotal
            changesMade = False
            ServerLabel.Text = MainMenu.serverLabel.Text
            tbl = New DataTable
            tbl.Columns.Add(foundRow(1))
            Dim dte As Date = Days2.thisSdate
            Dim i As Integer
            For i = 0 To 6
                Dim dt As Date = DateAdd(DateInterval.Day, i, dte)
                Dim Day As Integer = DatePart(DateInterval.Day, dt)
                Dim Month As Integer = DatePart(DateInterval.Month, dt)
                Dim dow As String = dt.ToString("ddd")
                tbl.Columns.Add(dow)
            Next
            tbl.Columns.Add("Total")
            Dim row As DataRow
            row = tbl.NewRow
            row(0) = "Values"
            Dim totl As Integer = 0
            For i = 2 To 9
                row(i - 1) = foundRow(i)
            Next
            tbl.Rows.Add(row)
            row = tbl.NewRow
            tbl.Rows.Add("Adjustments")
            tbl.Rows.Add("")

            con.Open()
            sql = "SELECT DISTINCT Year_Id, DATEPART(dw,Trans_Date) AS DOW, ISNULL(SUM(Qty * Retail),0) AS Sold FROM Daily_Transaction_Log l " & _
                "JOIN Calendar c ON CONVERT(Date,Trans_Date) BETWEEN sDate AND eDate AND Week_Id > 0 " & _
                "WHERE Type = 'Sold' AND Year_Id < " & thisYear & " AND Prd_Id = " & thisPeriod & " AND PrdWk = " & thisWeek & " " & _
                "AND Location = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                "GROUP BY Year_Id, DATEPART(dw,Trans_Date) ORDER BY Year_Id DESC, DATEPART(dw,Trans_Date)"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            Dim yr, dw, sold, yrs, avg As Integer
            Dim prevYear As Integer = 0
            totl = 0
            Dim arry(8) As Integer
            row = tbl.NewRow
            While rdr.Read
                yr = rdr("Year_Id")
                If prevYear = 0 Then prevYear = yr
                If yr <> prevYear Then
                    row(8) = Format(totl, "###,###,###")
                    tbl.Rows.Add(row)
                    row = tbl.NewRow
                    prevYear = yr
                    totl = 0
                    yrs += 1
                End If
                row(0) = prevYear
                dw = CInt(rdr("DOW"))
                sold = CInt(rdr("Sold"))
                totl += sold
                arry(7) += sold
                arry(dw - 1) += sold
                row(dw) = Format(sold, "###,###,###")
            End While
            con.Close()

            row(8) = Format(totl, "###,###,###")
            tbl.Rows.Add(row)
            yrs += 1
            row = tbl.NewRow
            row(0) = "Average"
            For i = 1 To 8
                avg = arry(i - 1) / yrs
                row(i) = Format(avg, "###,###,###")
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
        Catch ex As Exception

        End Try
    End Sub

    'Restrict Draft Variance to numbers only
    Private Sub dgv1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then
            Dim tb As TextBox = TryCast(e.Control, TextBox)
            AddHandler tb.KeyPress, AddressOf tb_KeyPress
        End If
    End Sub

    Private Sub tb_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) And Not Char.IsDigit(e.KeyChar) And e.KeyChar <> "." And e.KeyChar <> "-" Then
            e.Handled = True
        End If
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Try
            Dim amt, adjTotal As Integer
            Dim adjRow, planRow As DataRow
            If e.RowIndex < 1 Or e.ColumnIndex < 1 Then Exit Sub
            oTest = dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            If Not IsNothing(oTest) And IsNumeric(oTest) Then
                adjRow = tbl.Rows(e.RowIndex)
                planRow = tbl.Rows(0)

                'If thisSdate <= todaysSdate Then
                '    MessageBox.Show("CANNOT CHANGE PLAN FOR CURRENT OR PRIOR PERIOD!", "ENTER ADJUSTMENT AMOUNTS")
                '    adjRow(e.ColumnIndex) = Nothing
                '    Exit Sub
                'End If

                changesMade = True
                amt = CInt(Replace(oTest, ",", ""))
                adjRow(e.ColumnIndex) = Format(amt, "###,###,##0")
                adjTotal = 0
                For i As Integer = 1 To 7
                    oTest = adjRow(i)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(oTest) Then
                            amt = CInt(Replace(oTest, ",", ""))
                            adjTotal += amt
                        End If
                    End If
                Next
                adjRow(8) = Format(adjTotal, "###,###,##0")
                dgv1.Refresh()
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR ENTERING DATA IN DATAGRIDVIEW", "dgv1_CellValueChanged")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        lblProcessing.Visible = True
        Me.Refresh()
        Dim amt, adj As Integer
        Dim pct As Decimal
        Dim row As DataRow = tbl.Rows(0)
        Dim adjRow As DataRow = tbl.Rows(1)
        oTest = adjRow(8)
        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
            If oTest <> 0 Then
                MessageBox.Show("Total Adjustments must equal 0", "Save Changes")
                lblProcessing.Visible = False
                Exit Sub
            End If
        End If

        con.Open()
        For i = 1 To 7
            oTest = adjRow(i)
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                adj = CInt(oTest)
            Else : adj = 0
            End If
            oTest = row(i)
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                amt = row(i) + adj
                pct = (row(i) + adj) / row(8)
            End If
            foundRow(i + 1) = Format(amt, "###,###,###")
            sql = "IF NOT EXISTS (SELECT * FROM Day_Sales_Plan WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                "AND Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & i & ") " & _
                "INSERT INTO Day_Sales_Plan (Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Day, Amt, Pct) " & _
                "SELECT '" & thisStore & "','" & thisDept & "'," & thisYear & "," & thisPeriod & "," & thisWeek & "," & i &
                "," & amt & "," & pct & " " & _
                "ELSE " & _
                "UPDATE Day_Sales_Plan SET Amt = " & amt & ", Pct = " & pct & " " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & thisYear & " " & _
                "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & i & " "
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            For Each rw As DataRow In Days2.dayTbl.Rows

            Next
        Next
        con.Close()
        changesMade = False
        lblProcessing.Visible = False
        Me.Refresh()
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