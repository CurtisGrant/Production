Imports System.Data.SqlClient
Public Class Days4
    Public Shared con As SqlConnection = Days2.con
    Private Shared sql As String
    Private Shared cmd As SqlCommand
    Private Shared rdr As SqlDataReader
    Public Shared tbl, amtTbl, pctTbl, dayAmtTbl, dayPctTbl As DataTable
    Public Shared headRow, foundRow, totalRow As DataRow
    Public Shared thisDept, thisDeptDescr As String
    Public Shared thisYear As Integer
    Public Shared thisPeriod As Integer
    Public Shared thisYrPrd As Integer
    Public Shared thisYrWk As Integer
    Public Shared thisWeek As Integer
    Public Shared thisSdate As Date
    Public Shared todaysSdate As Date
    Public Shared thisStore As String
    Private Shared oTest As Object
    Private Shared changesMade As Boolean
    Private Shared deptTotal As Int32
    Private Sub Days4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Days2.lblProcessing.Visible = False
        'lblProcessing.Visible = True
        'Me.Refresh()
        thisDept = Days2.thisDept
        thisYear = Days2.thisYear
        thisPeriod = Days2.thisPeriod
        thisWeek = Days2.thisWeek
        thisSdate = Days2.thisSdate
        todaysSdate = Days2.todaysSdate
        thisStore = Days2.thisStore
        deptTotal = Days2.deptTotal
        tbl = Days2.dayTbl
        dayAmtTbl = Days2.dayTbl
        ''foundRow = tbl.Rows.Find("% WK")
        foundRow = Days2.foundRow
        con.Open()
        sql = "SELECT YrPrd, YrWk FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            thisYrPrd = rdr("YrPrd")
            thisYrWk = rdr("YrWk")
        End While
        con.Close()
        Call Load_Data("")

    End Sub

    Private Sub Load_Data(ByVal thisDept As String)
        Try


            Dim i, j As Integer
            Dim deptName As String = ""
            Dim totlPct As Decimal = 0
            changesMade = False
            lblServer.Text = MainMenu.serverLabel.Text
            amtTbl = New DataTable
            pctTbl = New DataTable
            amtTbl.Columns.Add(deptName)
            pctTbl.Columns.Add(" ")
            For i = 2 To 9
                amtTbl.Columns.Add(tbl.Columns(i).ColumnName)
                pctTbl.Columns.Add(tbl.Columns(i).ColumnName)
            Next

            Dim row As DataRow
            row = pctTbl.NewRow
            row(0) = "Values"
            Dim totl As Integer = 0
            For i = 2 To 8
                oTest = foundRow(i)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    oTest = Replace(oTest, "%", "")
                    If IsNumeric(oTest) Then
                        totlPct += CDec(oTest)
                        row(i - 1) = foundRow(i)
                    End If
                End If
            Next
            row(8) = Format(totlPct * 0.01, "###%")
            pctTbl.Rows.Add(row)
            pctTbl.Rows.Add("New Values")
            con.Open()
            sql = "SELECT DISTINCT Year_Id, DATEPART(dw,Trans_Date) AS DOW, ISNULL(SUM(Qty * Retail),0) AS Sold FROM Daily_Transaction_Log l " & _
                "JOIN Calendar c ON CONVERT(Date,Trans_Date) BETWEEN sDate AND eDate AND Week_Id > 0 " & _
                "WHERE Type = 'Sold' AND Year_Id < " & thisYear & " AND Prd_Id = " & thisPeriod & " AND PrdWk = " & thisWeek & " " & _
                "AND Location = '" & thisStore & "' " & _
                "GROUP BY Year_Id, DATEPART(dw,Trans_Date) ORDER BY Year_Id DESC, DATEPART(dw,Trans_Date)"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            Dim yr, dw, sold, yrs, avg As Integer
            Dim prevYear As Integer = 0
            totl = 0
            Dim arry(8) As Integer
            row = amtTbl.NewRow
            While rdr.Read
                yr = rdr("Year_Id")
                If prevYear = 0 Then prevYear = yr
                If yr <> prevYear Then
                    row(8) = totl
                    amtTbl.Rows.Add(row)
                    row = amtTbl.NewRow
                    prevYear = yr
                    totl = 0
                End If
                row(0) = prevYear
                dw = CInt(rdr("DOW"))
                sold = CInt(rdr("Sold"))
                totl += sold
                arry(7) += sold
                arry(dw - 1) += sold
                row(dw) = sold
            End While
            con.Close()

            row(8) = totl
            amtTbl.Rows.Add(row)
            row = amtTbl.NewRow
            row(0) = "Average"
            For i = 1 To 8
                yrs = 0
                For j = 0 To amtTbl.Rows.Count - 1
                    oTest = amtTbl.Rows(j).Item(i)
                    If IsNumeric(oTest) Then
                        If oTest > 0 Then yrs += 1
                    End If
                Next
                If yrs > 0 Then
                    avg = arry(i - 1) / yrs
                    row(i) = avg
                End If
            Next
            amtTbl.Rows.Add(row)

            Dim pct As Decimal
            For Each row In amtTbl.Rows
                Dim newRow = pctTbl.NewRow
                newRow(0) = row(0)
                For i = 1 To 7
                    If Not IsDBNull(row(8)) And Not IsDBNull(row(i)) Then
                        pct = row(i) / row(8)
                        newRow(i) = Format(pct, "###.0%")
                    End If
                Next
                newRow(8) = "100%"
                pctTbl.Rows.Add(newRow)
                If row(0) = "Values" Then
                    pctTbl.Rows.Add("New Values")
                    ''pctTbl.Rows.Add("")
                End If
            Next
            dgv1.DataSource = pctTbl.DefaultView
            dgv1.AutoResizeColumns()
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                If i = 8 Then dgv1.Columns(8).ReadOnly = True
                If i < 8 Then
                    dgv1.Rows(1).Cells(i).Style.BackColor = Color.Cornsilk
                End If
                dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            For i = 0 To dgv1.RowCount - 1
                If i <> 1 Then dgv1.Rows(i).ReadOnly = True
            Next
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "BAD ERROR")
        End Try
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Dim pctTotal As Decimal
        If e.RowIndex <> 1 Then Exit Sub
        If e.ColumnIndex < 1 Or e.ColumnIndex > 7 Then Exit Sub
        oTest = dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        If Not IsNothing(oTest) And IsNumeric(oTest) Then
            'If thisSdate <= todaysSdate Then
            '    MessageBox.Show("CANNOT CHANGE PLAN FOR CURRENT OR PRIOR WEEKS!", "ERROR")
            '    dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Nothing
            '    Exit Sub
            'End If
            dgv1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Format(CDec(oTest * 0.01), "###.0%")
            pctTotal = 0
            For i As Integer = 1 To 7
                oTest = dgv1.Rows(e.RowIndex).Cells(i).Value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(Replace(oTest, "%", "")) Then
                        pctTotal += (CDec(Replace(oTest, "%", "")) * 0.01)
                        changesMade = True
                    End If
                End If
            Next
            dgv1.Rows(e.RowIndex).Cells(8).Value = Format(pctTotal, "###.0%")
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()

        Dim pct As Decimal
        Dim val, amt As Integer
        Dim dept As String
        oTest = dgv1.Rows(1).Cells(8).Value
        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
            If Replace(oTest, "%", "") <> 100 Then
                MessageBox.Show("New Values Total does not equal 100%!", "ERROR")
                Exit Sub
            End If
        End If
        con.Open()
        Dim lastRow As Integer = dayAmtTbl.Rows.Count - 1
        For i As Integer = 1 To 7
            Dim totl As Integer = 0
            oTest = pctTbl.Rows(1).Item(i)
            dayAmtTbl.Rows(0).Item(i + 1) = oTest
            For j As Integer = 1 To lastRow - 1
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(Replace(oTest, "%", "")) Then
                        dept = dayAmtTbl.Rows(j).Item(0)
                        pct = Replace(oTest, "%", "") * 0.01
                        amt = dayAmtTbl.Rows(j).Item(9)
                        val = amt * pct
                        totl += val
                        '' dayAmtTbl.Rows(1).Item(i) = Format(val, "###,###,###")
                        sql = "IF NOT EXISTS (SELECT * FROM Day_Sales_Plan WHERE Str_Id = '" & thisStore & "' " & _
                            "AND Dept = '" & dept & "' AND Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " " & _
                            "AND Week_Id = " & thisWeek & " AND Day = " & i & ") " & _
                            "INSERT INTO Day_Sales_Plan (Str_Id, Dept, Year_Id, Prd_Id, Week_Id, Day, Amt, Pct) " & _
                            "SELECT '" & thisStore & "','" & dept & "'," & thisYear & "," & thisPeriod & "," & _
                            thisWeek & "," & i & "," & val & "," & pct & " " & _
                            "ELSE " & _
                            "UPDATE Day_Sales_Plan SET Amt = " & val & ", Pct = " & pct & " " & _
                            "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & dept & "' AND Year_Id = " & thisYear & " " & _
                            "AND Prd_Id = " & thisPeriod & " AND Week_Id = " & thisWeek & " AND Day = " & i & " "
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                
                    End If
                    changesMade = False
                End If
            Next
            dayAmtTbl.Rows(lastRow).Item(i + 1) = Format(totl, "###,###,###")
        Next
        For r As Integer = 1 To Days2.dayTbl.Rows.Count - 1
            oTest = Days2.dayTbl.Rows(r).Item(0)
            For c As Integer = 2 To 8
                oTest = pctTbl.Rows(1).Item(c)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    pct = Replace(oTest, "%", "") * 0.01
                Else : pct = 0
                End If
                oTest = Days2.dayTbl.Rows(r).Item(9)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    amt = Replace(oTest, ",", "")
                Else : amt = 0
                End If
                Days2.dayTbl.Rows(r).Item(c + 1) = Format(amt * pct, "###,###")
            Next
        Next
        con.Close()
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