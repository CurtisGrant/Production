Imports System.Data.SqlClient

Public Class MarkdownWeeks
    Public Shared thisYear As Integer
    Public Shared thisPlan As String
    Public Shared thisPeriod As Integer
    Public Shared todaysPeriod As Integer
    Public Shared thisStore As String
    Public Shared thisDept As String
    Public Shared thisBuyer As String
    Public Shared thisClass As String
    Public Shared con As SqlConnection
    Public Shared con2 As SqlConnection
    Public Shared con3 As SqlConnection
    Public Shared thisPlanIsActive As Boolean
    Private Shared todaysWeekId As Integer
    Private Shared oTest As Object
    Private Shared sql As String
    Private Shared weekDates(5) As String
    Private Shared rdr, rdr2, rdr3 As SqlDataReader
    Private Shared wkTbl As DataTable
    Private Shared columnIndex, rowIndex As Integer
    Private Shared cmd As SqlCommand
    Private Shared theValueB4Change As String
    Private Shared changesMade As Boolean
    Public Shared conString As String

    Private Sub MkdnWeeks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        thisPlan = MarkdownPlan.thisPlan
        thisYear = MarkdownPlan.thisYear
        thisPeriod = MarkdownPlan.periodSelected
        todaysPeriod = MarkdownPlan.todaysPeriod
        thisStore = MarkdownPlan.thisStore
        thisDept = MarkdownPlan.thisDept
        thisBuyer = MarkdownPlan.thisBuyer
        thisClass = MarkdownPlan.thisClass
        conString = MarkdownPlan.constring
        ''con = MarkdownPlan.con
        ''con2 = MarkdownPlan.con2
        ''con3 = MarkdownPlan.con3
        con = New SqlConnection(conString)
        con2 = New SqlConnection(conString)
        con3 = New SqlConnection(conString)
        serverLabel.Text = MainMenu.serverLabel.Text
        thisPlanIsActive = MarkdownPlan.thisPlanIsActive
        cboYear.Items.Add(MarkdownPlan.cboYear.SelectedItem)
        cboStore.Items.Add(MarkdownPlan.cboStore.SelectedItem)
        cboYear.SelectedIndex = 0
        cboStore.SelectedIndex = 0
        con.Open()
        sql = "SELECT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboDept.Items.Add(rdr("Dept"))
        End While
        con.Close()
        thisDept = MarkdownPlan.thisDept
        Dim idx As Integer = cboDept.FindString(thisDept)
        cboDept.SelectedIndex = idx
        con.Open()
        sql = "SELECT PrdWk FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Week_Id > 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            todaysWeekId = rdr("PrdWk")
        End While
        con.Close()

        ''Call Show_Weeks_Plan()
    End Sub

    Private Sub Show_Weeks_Plan()
        'Try
        changesMade = False
        Dim str As String
        Dim row As DataRow
        wkTbl = New DataTable
        Dim column = New DataColumn()
        Dim priorYrs As Integer = 0
        Dim thisYrPrd As Integer = MarkdownPlan.thisYrPrd
        Dim firstYrpPrd As Integer = thisYrPrd - 300
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "Week"
        wkTbl.Columns.Add(column)
        Dim PrimaryKey(1) As DataColumn
        PrimaryKey(0) = wkTbl.Columns("Week")
        wkTbl.PrimaryKey = PrimaryKey
        wkTbl.Columns.Add("3 Year Average")
        wkTbl.Columns.Add("2 Year Average")
        wkTbl.Columns.Add(thisYear & " Plan")
        wkTbl.Columns.Add("New Plan")
        wkTbl.Columns.Add("Actual Variance")
        wkTbl.Columns.Add(thisYear & " Actual")
        con2.Open()
        sql = "SELECT DISTINCT Year_Id FROM Calendar ORDER BY Year_Id DESC"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr(0)
            If oTest < thisYear Then
                priorYrs += 1
                If priorYrs < 4 Then
                    str = rdr(0) & " Actual"
                    wkTbl.Columns.Add(str)
                End If
            End If
        End While
        con2.Close()
        wkTbl.Columns.Add("Status")
        '                                                   Create a new row for each week of the selected period
        con2.Open()
        Dim cnt As Integer = 0
        sql = "SELECT Week_Id, sDate, eDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id = " & thisPeriod & " " & _
            "AND PrdWk > 0"
        cmd = New SqlCommand(sql, con2)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = wkTbl.NewRow
            row(0) = rdr(0)
            wkTbl.Rows.Add(row)
            weekDates(cnt) = rdr("sDate") & " - " & rdr("eDate")
            cnt += 1
        End While
        Array.Resize(weekDates, cnt)
        con2.Close()
        '                                                   Get actuals from Item_Sales
        con2.Open()
        sql = "SELECT Week_Num, Year_Id, ISNULL(SUM(Act_Sales),0) As Sales, ISNULL(SUM(Act_Mkdn),0) AS Mkdn, " & _
                "CASE WHEN ISNULL(SUM(Act_Mkdn),0) + ISNULL(SUM(Act_Sales),0) = 0 THEN 0 " & _
                "ELSE ISNULL(SUM(Act_Mkdn) / (SUM(Act_Sales) + SUM(Act_Mkdn)),0) END AS MkdnPct  " & _
                "FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Prd_Id = " & thisPeriod & " " & _
                "AND YrPrd BETWEEN " & firstYrpPrd & " AND " & thisYrPrd & " " & _
               "GROUP by Week_Num, Year_Id ORDER BY Year_Id"
        cmd = New SqlCommand(sql, con2)
        rdr2 = cmd.ExecuteReader
        Dim year, week As Int32
        Dim sales, mkdn As Int32
        Dim pct As Decimal
        While rdr2.Read
            oTest = rdr2("Week_Num")
            row = wkTbl.Rows.Find(oTest)
            year = rdr2("Year_Id")
            week = rdr2("Week_Num")
            sales = rdr2("Sales")
            mkdn = rdr2("Mkdn")
            pct = rdr2("MkdnPct")
            For Each column In wkTbl.Columns
                oTest = rdr2("Year_Id")
                str = rdr2("Year_Id") & " Actual"
                row.Item(str) = Format(pct, "##.0%")
            Next
            con3.Open()
            sql = "SELECT Mkdn_Pct FROM Calendar c " & _
                "JOIN Sales_Plan p ON c.Year_Id = p.Year_Id AND c.Prd_Id = p.Prd_Id AND c.PrdWk = p.Week_Id AND c.Week_Id > 0 " & _
                "WHERE Plan_Id = '" & thisPlan & "' AND p.Year_Id = " & year & " AND p.Prd_Id = " & thisPeriod & " " & _
                "AND c.Week_Id =" & week & " AND Str_id = '" & thisStore & "' " & _
                "AND Dept = '" & thisDept & "'"
            cmd = New SqlCommand(sql, con3)
            rdr3 = cmd.ExecuteReader
            While rdr3.Read
                str = year & " Plan"
                row(str) = Format(rdr3(0), "##.0%")
            End While
            con3.Close()
        End While
        con2.Close()

        Dim clms As Int16 = wkTbl.Columns.Count - 1
        Dim yr3, yr2, yr1, avg3Yr, avg2Yr, lyActual, tyActual, actVariance As Decimal
        Dim lastYear As Integer = thisYear - 1
        For Each rw In wkTbl.Rows
            If clms >= 8 Then
                oTest = rw(8)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    str = Replace(oTest, "%", "")
                    yr3 = CDec(str * 0.01)
                End If
            End If
            If clms > 7 Then
                oTest = rw(7)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    str = Replace(oTest, "%", "")
                    yr2 = CDec(str * 0.01)
                End If
            End If
            If clms > 6 Then
                oTest = rw(6)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    str = Replace(oTest, "%", "")
                    yr1 = CDec(str * 0.01)
                End If
            End If
            If clms >= 8 Then
                avg3Yr = (yr3 + yr2 + yr1) / 3
                rw(1) = Format(avg3Yr, "##.0%")
            End If
            If clms > 7 Then
                avg2Yr = (yr2 + yr1) / 2
                rw(2) = Format(avg2Yr, "##.0%")
            End If
            If clms >= 6 Then
                oTest = rw(lastYear & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "" Then lyActual = CDec(Replace(oTest, "%", "")) * 0.01 Else lyActual = 0
                End If
                oTest = rw(thisYear & " Actual")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "" Then
                        tyActual = CDec(Replace(oTest, "%", "")) * 0.01
                    Else : tyActual = 0
                    End If
                End If
            Else : lyActual = 0
            End If
            actVariance = tyActual - lyActual
            rw("Actual Variance") = Format(actVariance, "###.0%")
        Next
        Dim bindingSource As New BindingSource
        bindingSource.DataSource = wkTbl
        dgv1.DataSource = bindingSource
        dgv1.AutoResizeColumns()
        dgv1.Columns(dgv1.Columns.Count - 1).Visible = False                   ' hide status column
        For i = 1 To dgv1.ColumnCount - 1
            dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            If i = 4 Then dgv1.Columns(4).DefaultCellStyle.BackColor = Color.Cornsilk
            dgv1.Columns(i).ReadOnly = True
        Next
        dgv1.Columns(4).ReadOnly = False

        'Catch ex As Exception
        '    If con.State = ConnectionState.Open Then con.Close()
        '    If con2.State = ConnectionState.Open Then con2.Close()
        '    MessageBox.Show(ex.Message, "LOAD DATE")
        'End Try
    End Sub

    'Restrict Draft Variance to numbers only
    Private Sub dgv1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then
            Dim tb As TextBox = TryCast(e.Control, TextBox)
            If Me.dgv1.CurrentCell.ColumnIndex = 4 Then
                AddHandler tb.KeyPress, AddressOf tb_KeyPress
            End If
        End If
    End Sub

    Private Sub tb_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) And Not Char.IsDigit(e.KeyChar) And e.KeyChar <> "." And e.KeyChar <> "-" Then
            e.Handled = True
        End If
    End Sub

    ''Public Sub dgv1_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        oTest = dgv1.Rows(rowIndex).Cells(0).Value
        If thisPeriod < todaysPeriod Then
            MessageBox.Show("PLAN CONNOT BE CHANGED FOR COMPLETED PERIODS!", "CHANGE MARKDOWN PLAN")
            Dim row As DataRow = wkTbl.Rows.Find(oTest)
            row(4) = Nothing
            Exit Sub
        End If
        If thisPeriod = todaysPeriod Then
            If oTest <= todaysWeekId Then
                MessageBox.Show("PLAN CANNOT BE CHANGED FOR CURRENT OR PRIOR WEEKS!", "CHANGE MARKDOWN PLAN")
                Dim row As DataRow = wkTbl.Rows.Find(oTest)
                row(4) = Nothing
                Exit Sub
            End If
        End If
        If thisPlanIsActive = True Then
            MsgBox("Sorry, active Plans can not be changed!")
            Call Show_Weeks_Plan()
            Exit Sub
        End If
        oTest = dgv1.Rows(rowIndex).Cells(4).Value
        If columnIndex = 4 Then
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                changesMade = True
                oTest = Format(CDec(oTest) * 0.01, "##.0%")
                dgv1.Rows(rowIndex).Cells(3).Value = oTest
            End If
        End If
    End Sub

    Public Sub dgv1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim ht As DataGridView.HitTestInfo
            ht = Me.dgv1.HitTest(e.X, e.Y)
            Dim rowIdx As Int16 = ht.RowIndex
            Dim columnIdx As Int16 = ht.ColumnIndex
            If rowIdx > -1 And columnIdx = 4 Then
                oTest = dgv1.Rows(rowIdx).Cells(columnIdx).Value
            End If
        End If
    End Sub


    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim x As Integer = dgv1.Rows.Count
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

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        Try
            Dim thisWeek As Int16
            Dim lastYear As Int16 = thisYear - 1
            thisBuyer = "*"
            thisClass = "*"
            Dim wks As Integer = 0
            Dim totl As Decimal = 0
            For Each row In dgv1.Rows
                wks += 1
                oTest = row.cells(0).value
                If oTest = "Total" Then GoTo 100
                oTest = row.cells(4).value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    Dim thisPct As Decimal = Replace(oTest, "%", "") * 0.01
                    totl += thisPct
                    thisWeek = row.cells(0).value
                    con.Open()
                    sql = "UPDATE p Set Mkdn_Pct = " & thisPct & " FROM Sales_Plan p " & _
                        "JOIN Calendar c ON c.Year_Id = p.Year_Id AND c.Prd_Id = p.Prd_Id AND c.PrdWk = p.Week_Id AND c.Week_Id > 0 " & _
                        "WHERE Plan_Id = '" & thisPlan & "' AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND p.Year_Id = " & thisYear & " AND p.Prd_Id = " & thisPeriod & " AND c.Week_Id = " & thisWeek & ""
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                    con.Close()
                End If
100:        Next
            con.Close()
            '                                      Roll average week plan into period plan
            con.Open()
            sql = "UPDATE Sales_Plan SET Mkdn_Pct = " & totl / wks & " WHERE Plan_Id = '" & thisPlan & "' AND Year_Id = " & thisYear & " " & _
                "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Prd_Id = " & thisPeriod & " AND Week_Id = 0"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
            changesMade = False
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            MessageBox.Show(ex.Message, "SAVE CHANGES")
        End Try
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
        Call Show_Weeks_Plan()
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