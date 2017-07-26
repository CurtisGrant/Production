Imports System.Data.SqlClient
Imports System.Xml
Public Class MarkdownPlan
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared thisStore, thisDept, thisPlan, thisBuyer, thisClass As String
    Public Shared thisYear, firstYear As Integer
    Public Shared thisPeriod, periodSelected, thisYrPrd As Integer
    Public Shared todaysPeriod As Integer
    Public Shared thisPlanIsActive As Boolean
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Private Shared cmd As SqlCommand
    Private Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Private Shared tbl, wkTbl, dayTbl As DataTable
    Private Shared oTest As Object
    Private Shared today As Date = Date.Now
    Private Shared shortToday = Date.Today
    Private Shared priorYrs As Int16 = 0
    Private Shared columnIndex, lastYear, thisWeek As Int16
    Private Shared rowIndex As Int16
    Private Shared rnd As Double
    Private Shared perodDates(50) As String
    Private Shared theValue As String
    Private Shared thisUser = Environment.UserName
    Private Shared planBy As String
    Private Shared percentagesBy As String
    Private Shared periodDates(50) As String
    Private Shared thisPrd As Integer
    Private Shared formLoaded As Boolean
    Private Shared changesMade As Boolean

    Private Sub MarkdownPlan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            formLoaded = False
            serverLabel.Text = MainMenu.serverLabel.Text

            Me.Location = New Point(100, 150)
            lblPlan.Text = PlanMaintenance.thisPlan
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)

            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Sales_Plan WHERE Status = 'Active' ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYear.Items.Add(rdr("Year_Id"))
            End While
            oTest = DatePart(DateInterval.Year, Date.Today).ToString
            cboYear.SelectedIndex = cboYear.FindString(oTest)
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Str_Id FROM Stores WHERE Status = 'Active' ORDER BY Str_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Str_Id")
                cboStore.Items.Add(rdr(0))
            End While
            cboStore.SelectedIndex = 0
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Dept FROM Departments WHERE Status = 'Active' ORDER BY Dept"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr(0)
                cboDept.Items.Add(rdr(0))
            End While
            cboDept.SelectedIndex = 0
            con.Close()

            con.Open()
            sql = "SELECT DISTINCT Plan_Id FROM Sales_Plan WHERE Year_Id =" & thisYear & " AND Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisPlan = rdr("Plan_Id")
            End While
            con.Close()

            con.Open()
            sql = "SELECT Prd_Id FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                todaysPeriod = rdr("Prd_Id")
            End While
            con.Close()
           
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD FORM")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub Load_The_Data()
        Try
            lblProcessing.Visible = True
            Me.Refresh()
            changesMade = False
            Dim str As String
            Dim row As DataRow
            Dim cnt As Integer = 0
            lastYear = thisYear - 1
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con4 = New SqlConnection(conString)
            con5 = New SqlConnection(conString)
            tbl = New DataTable
            tbl.Columns.Clear()
            dgv1.DataSource = tbl.DefaultView
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Period"
            tbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = tbl.Columns("Period")
            tbl.PrimaryKey = PrimaryKey
            tbl.Columns.Add("3 Year Average")
            tbl.Columns.Add("2 Year Average")
            tbl.Columns.Add(thisYear & " Plan")
            tbl.Columns.Add("New Plan")
            tbl.Columns.Add(thisYear & " Actual")
            con2.Open()
            sql = "SELECT DISTINCT Year_Id FROM Calendar ORDER BY Year_Id DESC"
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Year_Id")
                If oTest < thisYear And cnt < 3 Then
                    priorYrs += 1
                    str = rdr("Year_Id") & " Actual"
                    tbl.Columns.Add(str)
                    cnt += 1
                End If
            End While
            tbl.Columns.Add("Status")
            Dim totalColumns = tbl.Columns.Count
            con2.Close()

            con2.Open()
            sql = "SELECT YrPrd FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisYrPrd = rdr("YrPrd")
            End While
            con2.Close()

            con2.Open()
            sql = "SELECT Prd_Id, sDate, eDate FROM Calendar WHERE Year_ID = " & thisYear & " AND Prd_Id > 0 AND Week_Id = 0 ORDER BY Prd_ID"
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            Dim i As Int32 = 0
            While rdr.Read
                periodDates(i) = rdr("sDate") & " - " & rdr("eDate")
                i += 1
                row = tbl.NewRow
                row(0) = rdr(0)
                tbl.Rows.Add(row)
            End While
            con2.Close()

            con2.Open()
            sql = "SELECT Year_Id, Prd_Id, SUM(ISNULL(Act_Sales,0)) AS Sales, SUM(ISNULL(Act_Mkdn,0)) AS Mkdn INTO #t1 FROM Sales_Summary " & _
                "WHERE Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND YrPrd < " & thisYrPrd & " " & _
                "AND Year_Id >= " & firstYear & " AND Year_Id <= " & thisYear & " " & _
                "GROUP BY Year_Id, Prd_Id " & _
                "SELECT Year_Id, Prd_Id, Sales, Mkdn, CASE WHEN Sales + Mkdn = 0 THEN 0 " & _
                "ELSE Mkdn / (CONVERT(Decimal(18,4),Sales) + CONVERT(Decimal(18,4),Mkdn)) END AS MkdnPct FROM #t1"

            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            Dim sales, mkdn, year, period, clms As Int32
            Dim pct, planPct As Decimal
            While rdr.Read
                oTest = rdr("Prd_Id")
                row = tbl.Rows.Find(oTest)
                year = rdr("Year_Id")
                period = rdr("Prd_Id")
                sales = rdr("Sales")
                mkdn = rdr("Mkdn")
                pct = rdr("MkdnPct")
                For Each column In tbl.Columns
                    oTest = rdr("Year_Id")
                    str = rdr("Year_Id") & " Actual"
                    row.Item(str) = Format(pct, "##.0%")
                Next
                con3.Open()
                sql = "SELECT Mkdn_Pct As Plan_Pct " & _
                    "FROM Sales_Plan WHERE Plan_Id = '" & thisPlan & "' AND Year_Id = " & thisYear & " " & _
                    "AND Prd_Id = " & period & " AND Week_Id = 0 AND Dept = '" & thisDept & "' AND Str_Id = '" & thisStore & "'"
                cmd = New SqlCommand(sql, con3)
                rdr3 = cmd.ExecuteReader
                While rdr3.Read
                    planPct = 0
                    str = thisYear & " Plan"
                    oTest = rdr3("Plan_Pct")
                    If Not IsDBNull(oTest) Then planPct = oTest
                    row(str) = Format(planPct, "###.0%")
                End While
                con3.Close()
            End While
            con2.Close()

            clms = tbl.Columns.Count - 2
            Dim yr3, yr2, yr1, avg3Yr, avg2Yr As Decimal
            Dim tyPlan, tyActual, lyActual As Decimal
            For Each rw In tbl.Rows
                If clms >= 8 Then
                    oTest = rw(8)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            str = Replace(oTest, "%", "")
                            yr3 = CDec(str * 0.01)
                        End If
                    End If
                End If
                If clms > 7 Then
                    oTest = rw(7)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            str = Replace(oTest, "%", "")
                            yr2 = CDec(str * 0.01)
                        End If
                    End If
                End If
                If clms > 6 Then
                    oTest = rw(6)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If oTest <> "" Then
                            str = Replace(oTest, "%", "")
                            yr1 = CDec(str * 0.01)
                        End If
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
                oTest = rw(thisYear & " Plan")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If oTest <> "" Then tyPlan = CDec(Replace(oTest, "%", "")) * 0.01 Else tyPlan = 0
                End If
            Next
            dgv1.DataSource = tbl.DefaultView
            dgv1.AutoResizeColumns()
            For i = 0 To dgv1.Rows.Count - 1
                dgv1.Rows(i).Cells(0).Style.ForeColor = Color.Blue
                dgv1.Rows(i).Cells(0).Style.Font = New Font("Ariel", 8, FontStyle.Underline)
            Next
            dgv1.Columns(totalColumns - 1).Visible = False                          ' hide status column
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv1.Columns(i).ReadOnly = True
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                If i = 4 Then
                    dgv1.Columns(4).ReadOnly = False
                    dgv1.Columns(4).DefaultCellStyle.BackColor = Color.Cornsilk
                End If
            Next
            formLoaded = True
            lblProcessing.Visible = False
            Me.Refresh()
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
            MessageBox.Show(ex.Message, "Load Form")
        End Try
    End Sub

    Private Sub Save_This_Department()
        con2.Open()                                               ' Update dept as pct of store
        sql = "UPDATE Markdown_Plan SET Amt = (SELECT SUM(Amt) FROM MarkDown_Plan WHERE Plan_Id = '" & thisPlan & "' " & _
            "AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' AND Year_Id = " & thisYear & " " & _
            "AND Prd_Id = 0 AND Week_Id = 0)"
        cmd = New SqlCommand(sql, con2)
        cmd.ExecuteNonQuery()
        con2.Close()
    End Sub

    Private Sub Save_This_Period()
        Try
            Dim period As Integer
            Dim pct As Decimal
            con2.Open()
            For Each row In tbl.Rows
                period = CInt(row(0))
                oTest = row("New Plan")
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If IsNumeric(Replace(oTest, "%", "")) Then
                        pct = CDec(Replace(oTest, "%", "")) * 0.01
                        sql = "UPDATE Sales_Plan SET Mkdn_Pct = " & pct & " WHERE Plan_Id = '" & thisPlan & "' " & _
                            "AND Prd_Id = " & period & " "
                        cmd = New SqlCommand(sql, con2)
                        cmd.ExecuteNonQuery()
                    End If
                End If
            Next
            con2.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SAVE CHANGES")
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        lblProcessing.Visible = True
        Me.Refresh()
        Call Save_Changes()
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub

    Public Sub Save_Changes()
        Try
            Call Save_This_Period()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "SAVE CHANGES")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
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

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        oTest = dgv1.Rows(rowIndex).Cells(0).Value
        If rowIndex + 1 < todaysPeriod Then
            MessageBox.Show("PLAN CANNOT BE CHANGED FOR COMPLETED PERIODS!", "CHANGE MARKDOWN PLAN")
            Dim row As DataRow = tbl.Rows.Find(oTest)
            row(4) = Nothing
            Exit Sub
        End If
        Dim pct As Decimal
        If columnIndex = 4 Then
            Dim row As DataRow = tbl.Rows(rowIndex)
            oTest = dgv1.Rows(rowIndex).Cells(4).Value
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                changesMade = True
                pct = CDec(oTest) * 0.01
                oTest = Format(pct, "##.0%")
                row(4) = oTest

            End If
        End If
    End Sub

    Public Sub dgv1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim ht As DataGridView.HitTestInfo
            ht = Me.dgv1.HitTest(e.X, e.Y)
            Dim rowIdx As Int16 = ht.RowIndex
            Dim columnIdx As Int16 = ht.ColumnIndex
            If columnIdx = 0 And rowIdx > -1 Then
                periodSelected = dgv1.Rows(rowIdx).Cells(0).Value
                planBy = "Week"
                MarkdownWeeks.Show()
            End If
        End If
    End Sub


    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
        If formLoaded = False Then Exit Sub
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        Dim rowIdx As Int16 = ht.RowIndex
        Dim columnIdx As Int16 = ht.ColumnIndex
        Dim x As Integer = dgv1.Rows.Count
        Dim str As String
        If columnIdx = 0 Then
            If rowIdx > -1 And rowIdx < x Then
                str = periodDates(rowIdx)
                With dgv1.Rows(rowIdx).Cells(0)
                    .ToolTipText = str
                End With
            End If
        End If
    End Sub

    Private Sub cboYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYear.SelectedIndexChanged
        Try
            thisYear = cboYear.SelectedItem
            firstYear = thisYear - 3
            con2.Open()
            sql = "SELECT DISTINCT Plan_Id FROM Sales_Plan WHERE Status = 'Active' AND Year_Id = " & thisYear & " "
            cmd = New SqlCommand(sql, con2)
            rdr = cmd.ExecuteReader
            While rdr.Read
                thisPlan = rdr("Plan_Id")
            End While
            con2.Close()
            If formLoaded Then Call Load_The_Data()
        Catch ex As Exception
            If con2.State = ConnectionState.Open Then con2.Close()
            MessageBox.Show(ex.Message, "SELECT YEAR")
            If con2.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem
        If formLoaded Then Call Load_The_Data()
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
        Call Load_The_Data()
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