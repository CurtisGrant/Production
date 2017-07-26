Imports System.Data.SqlClient
Public Class WeeksOnHand
    Private con, con2, con3 As SqlConnection
    Private cmd As SqlCommand
    Private rdr, rdr2 As SqlDataReader
    Private sql As String
    Private selectedYear, thisYear As Integer
    Private thisStore As String = Nothing
    Private thisDept As String = Nothing
    Private thisBuyer As String = Nothing
    Private tbl As DataTable
    Private oTest As Object
    Private statusColumn As Integer
    Private periodDates(53) As String
    Private todaysPeriod As Integer
    Private newDept, changesMade As Boolean

    Private Sub WeeksOnHand_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        selectedYear = DatePart(DateInterval.Year, Date.Today)
        thisDept = Nothing
        thisBuyer = Nothing
        thisStore = Nothing
        con = MainMenu.con
        con2 = New SqlConnection(MainMenu.conString)

        con.Open()
        sql = "SELECT DISTINCT Year_Id FROM Buyer_PCT ORDER BY Year_Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboYr.Items.Add(rdr("Year_Id"))
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT ID FROM Buyers WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboBuyer.Items.Add(rdr("ID"))
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT ID FROM Stores WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboStr.Items.Add(rdr("ID"))
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT ID FROM Departments WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboDept.Items.Add(rdr("ID"))
        End While
        con.Close()

        con.Open()
        sql = "SELECT Prd_Id FROM Calendar WHERE CONVERT(Date,GETDATE()) BETWEEN sDate AND eDate AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            todaysPeriod = rdr("Prd_Id")
        End While
        con.Close()

        thisYear = DatePart(DateInterval.Year, Date.Today)
        cboYr.SelectedIndex = cboYr.FindString(thisYear)
        lblServer.Text = MainMenu.serverLabel.Text
        cboBuyer.SelectedIndex = 0
        cboStr.SelectedIndex = 0
        cboDept.SelectedIndex = 0
    End Sub

    Private Sub Load_Data()
        changesMade = False
        Dim cnt As Integer = 0
        Dim ly As Integer = 0
        Dim l2y As Integer = 0
        Dim l3y As Integer = 0
        Dim plan, priorYrs, wks, prd, yr As Integer
        Dim str As String
        Dim i As Integer = 0
        Dim lastYear As Integer = selectedYear - 1
        tbl = New DataTable
        Dim cc As Integer = tbl.Rows.Count
        Dim row As DataRow
        Dim column As New DataColumn()
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
        tbl.Columns.Add("New Plan WksOH")
        tbl.Columns.Add("New Plan Variance")
        tbl.Columns.Add(selectedYear & " WksOH")
        str = selectedYear & " WksOH"
        tbl.Columns.Add(selectedYear & " Variance")
        Dim x As Integer = tbl.Columns.Count
        If con.State = ConnectionState.Closed Then con.Open()
        sql = "SELECT DISTINCT Year_Id FROM Buyer_Pct WHERE Str_Id = '" & thisStore & "' " & _
            "AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " ORDER BY Year_Id DESC"
        cmd = New SqlCommand(sql, con)
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
                str = rdr(0) & " WksOH"
                tbl.Columns.Add(str)
                str = rdr(0) & " Variance"
                If cnt < 3 Then tbl.Columns.Add(str)
            End If
        End While
        con.Close()

        con.Open()
        sql = "SELECT Prd_Id, sDate, eDate FROM Calendar WHERE Year_ID = " & selectedYear & " AND Prd_Id > 0 AND Week_Id = 0 ORDER BY Prd_ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            periodDates(i) = rdr("sDate") & " - " & rdr("eDate")
            i += 1
            row = tbl.NewRow
            row(0) = i
            tbl.Rows.Add(row)
        End While
        con.Close()

        Array.Resize(periodDates, i)
        newDept = True
        '
        '                                            Change table from Item_Sales to Buyer_PCT below
        '
        con.Open()
        'sql = "IF OBJECT_ID('tempdb.dbo.#t1','U') IS NOT NULL DROP TABLE #t1 " & _
        '    "SELECT Prd_Id, Year_Id, eDate, CASE WHEN Act_WksOH IS NULL THEN Projected_WksOH " & _
        '    "ELSE Act_WksOH END AS wks INTO #t1 FROM Item_Sales WHERE Str_Id = '" & thisStore & "' " & _
        '    "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
        '    "AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " AND eDate < GETDATE() " & _
        '    "SELECT Year_Id, Prd_Id , AVG(wks) AS wks FROM #t1 WHERE ISNULL(wks,0) > 0 GROUP BY Year_Id, Prd_Id"
        sql = "SELECT Prd_Id, Year_Id, CASE WHEN Act_WksOH IS NULL THEN ISNULL(Projected_WksOH,0) " & _
          "ELSE Act_WksOH END AS wks FROM Buyer_Pct WHERE Str_Id = '" & thisStore & "' " & _
          "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
          "AND Year_Id BETWEEN " & selectedYear - 3 & " AND " & selectedYear & " ORDER BY Year_Id, Prd_Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr("Prd_Id")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then prd = oTest
            row = tbl.Rows.Find(oTest)
            wks = rdr("wks")
            yr = rdr("Year_Id")
            ''con2.Open()
            ''sql = "SELECT Plan_WksOH FROM Buyer_PCT WHERE Str_Id = '" & thisStore & "' AND Prd_id = " & prd & " " & _
            ''    "AND Year_Id = " & selectedYear & " AND Buyer = '" & thisBuyer & "' AND Dept = '" & thisDept & "'"
            ''cmd = New SqlCommand(sql, con2)
            ''rdr2 = cmd.ExecuteReader
            ''While rdr2.Read
            ''    oTest = rdr2("Plan_WksOH")
            ''    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then row.Item(selectedYear & " Plan") = oTest

            ''End While
            ''con2.Close()
            ''For Each column In tbl.Columns
            ''    oTest = rdr("Year_Id")
            ''    If rdr("Year_Id") <= selectedYear Then
            ''        str = rdr("Year_Id") & " WksOH"
            ''        If rdr("Year_ID") = selectedYear Then
            ''            If prd < todaysPeriod Then row.Item(selectedYear & " WksOH") = wks
            ''        Else : row.Item(str) = wks
            ''        End If
            ''    End If
            ''Next
            oTest = rdr("Year_Id")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                str = oTest & " WksOH"
                row(str) = wks
            End If
            Dim clms As Integer = tbl.Columns.Count
            Dim Avg3YrTotal As Decimal = 0
            Dim Avg2YrTotal As Decimal = 0
            Dim tyPlanTotal As Decimal = 0
            Dim newPlanTotal As Decimal = 0
            Dim tyActualTotal As Decimal = 0
            Dim lyActualTotal As Decimal = 0
            Dim l2yActualTotal As Decimal = 0
            Dim l3yActualTotal As Decimal = 0
            Dim lyPeriodToDateTotal As Decimal = 0
            Dim draftPlanTotal, actVariance As Decimal
            Dim tyPlan, draftPlan, tyActual, lyActual, l2yActual, l3yActual As Int32
            For Each row In tbl.Rows
                prd = row("Period")
                If priorYrs >= 1 Then
                    oTest = row(ly & " WksOH")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            lyActual = oTest
                            lyActualTotal += CInt(oTest)
                        Else : lyActual = 0
                        End If
                    End If
                    If priorYrs >= 2 Then
                        oTest = row(l2y & " WksOH")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If oTest <> "" Then
                                l2yActual = oTest
                                l2yActualTotal += CInt(oTest)
                            Else : l2yActual = 0
                            End If
                        End If
                    End If
                    If priorYrs = 3 Then
                        oTest = row(l3y & " WksOH")
                        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                            If oTest <> "" Then
                                l3yActual = oTest
                                l3yActualTotal += CInt(oTest)
                            Else : l3yActual = 0
                            End If
                        End If
                    End If
                    oTest = row(selectedYear & " WksOH")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            tyActual = oTest
                            tyActualTotal += CInt(oTest)
                            If prd < todaysPeriod Then lyPeriodToDateTotal += lyActual
                            If prd >= todaysPeriod Then row(selectedYear & " WksOH") = Nothing ' added this 9/6
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

                oTest = row("New Plan WksOH")
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

                oTest = row(lastYear & " WksOH")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    newDept = False                                                       ' turns off newDept flag when we have sales last year
                    oTest = row(selectedYear & " Plan")
                End If

                If tyActual <> 0 And lyActual <> 0 Then
                    actVariance = (tyActual / lyActual) - 1
                Else : actVariance = 0
                End If
                If prd < todaysPeriod Then row(selectedYear & " Variance") = Format(actVariance, "###.0%")

                If lyActual <> 0 And l2yActual <> 0 Then
                    actVariance = (lyActual / l2yActual) - 1
                Else : actVariance = 0
                End If
                row(ly & " Variance") = Format(actVariance, "###.0%")

                If l2yActual <> 0 And l3yActual <> 0 Then
                    actVariance = (l2yActual / l3yActual) - 1
                Else : actVariance = 0
                End If
                If l2y > 0 Then
                    row(l2y & " Variance") = Format(actVariance, "###.0%")
                End If

                If priorYrs >= 2 Then
                    oTest = row(l2y & " WksOH")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then l2yActual = oTest Else l2yActual = 0
                    row("2 Year Average") = Format(Math.Round((lyActual + l2yActual) / 2, MidpointRounding.AwayFromZero), "###,###,###")
                Else : l2yActual = 0
                End If
                If priorYrs >= 3 Then
                    oTest = row(l3y & " WksOH")
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
            Next '                                                       total row starts here             
        End While
        con.Close()

        Dim hdr As String
        For Each row In tbl.Rows
            str = selectedYear & " WksOH"
            oTest = row(str)
            If Not IsDBNull(oTest) Then
                If row(str) = 0 Then row(str) = Nothing
            End If
        Next
        dgv1.DataSource = tbl.DefaultView
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv1.Columns(0).ReadOnly = True
        For i = 1 To dgv1.ColumnCount - 1
            dgv1.Columns(i).ReadOnly = True
            dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            hdr = dgv1.Columns(i).HeaderText
            If IsNumeric(Microsoft.VisualBasic.Left(hdr, 4)) Then
                If InStr(hdr, "Variance") Then dgv1.Columns(i).Visible = False
            End If
        Next
        dgv1.Columns("New Plan WksOH").DefaultCellStyle.BackColor = Color.Cornsilk
        dgv1.Columns("New Plan WksOH").ReadOnly = False

    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Dim foundRow As DataRow
        Dim draft, plan As Integer
        Dim variance As Decimal
        foundRow = tbl.Rows.Find(e.RowIndex + 1)
        Dim thisPeriod As Integer = foundRow(0)
        If e.RowIndex + 1 < todaysPeriod And selectedYear <= thisYear Then
            MessageBox.Show("Period Plan cannot be changed once the period is complete.", "ERROR!")
            foundRow("New Plan WksOH") = Nothing
            Exit Sub
        End If
        If e.RowIndex + 1 = todaysPeriod And selectedYear = thisYear Then
            MessageBox.Show("The current period cannot be changed.", "ERROR!")
            foundRow("New Plan WksOH") = Nothing
            Exit Sub
            Me.Refresh()
        End If
        draft = foundRow("New Plan WksOH")
        oTest = foundRow(selectedYear & " Plan")
        If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
            variance = (draft / plan) - 1
            foundRow("New Plan Variance") = Format(variance, "###.0%")
        End If
        changesMade = True
    End Sub
    Private Sub cboYr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYr.SelectedIndexChanged
        selectedYear = cboYr.SelectedItem
        If IsNothing(thisBuyer) Or IsNothing(thisDept) Or IsNothing(thisStore) Then Exit Sub
        If changesMade Then
            Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Call Save_Changes()
            End Select
        End If
        tbl.Reset()
        Call Load_Data()
    End Sub

    Private Sub cboStr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStr.SelectedIndexChanged
        thisStore = cboStr.SelectedItem
        If IsNothing(thisBuyer) Or IsNothing(thisYear) Or IsNothing(thisDept) Then Exit Sub
        If changesMade Then
            Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Call Save_Changes()
            End Select
        End If
        Call Load_Data()
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
        If IsNothing(thisYear) Or IsNothing(thisBuyer) Or IsNothing(thisStore) Then Exit Sub
        If changesMade Then
            Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Call Save_Changes()
            End Select
        End If
        Call Load_Data()
    End Sub

    Private Sub cboBuyer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboBuyer.SelectedIndexChanged
        thisBuyer = cboBuyer.SelectedItem
        If IsNothing(thisDept) Or IsNothing(thisStore) Or IsNothing(thisYear) Then Exit Sub
        If changesMade Then
            Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Call Save_Changes()
            End Select
        End If
        Call Load_Data()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        Dim prd As String
        con.Open()
        For Each row In tbl.Rows
            prd = row(0)
            oTest = row("New Plan WksOH")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) And oTest > 0 Then
                    sql = "UPDATE Buyer_PCT SET Plan_WksOH = " & CInt(oTest) & " WHERE Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' AND Year_Id = " & selectedYear & " " & _
                        "AND Prd_Id = " & CInt(prd) & " "
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                End If
            End If
        Next
        con.Close()
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