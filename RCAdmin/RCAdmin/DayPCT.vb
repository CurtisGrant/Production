Imports System.Xml
Imports System.Data.SqlClient
Public Class DayPct
    Public Shared con, con2, con3 As SqlConnection
    Public Shared rdr, rdr2, rdr3 As SqlDataReader
    Public Shared conString, sql As String
    Public Shared cmd As SqlCommand
    Public Shared thisYear, thisYrPrd, thisPeriod, numberPeriods As Integer
    Public Shared thisStore, thisDept, thisTable, thisBuyer As String
    Public Shared oTest As Object
    Public Shared periodDates(50) As String
    Public Shared formLoaded As Boolean
    Public Shared dayTbl, dayTbl2 As DataTable
    Public Shared row, row2, tblRow, tblRow2, foundRow As DataRow
    Public Shared rowIndex, columnIndex, todaysPeriod, todaysYrPrd As Integer
    Public Shared formLoad, changesMade As Boolean
    Private Sub DayPct_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            serverLabel.Text = MainMenu.serverLabel.Text
            formLoaded = False
            Me.Location = New Point(300, 200)
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con3 = New SqlConnection(conString)
            con = New SqlConnection(conString)
            Dim cnt As Integer = 0

            con.Open()
            sql = "SELECT DISTINCT Year_Id FROM Sales_Plan ORDER BY Year_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboYear.Items.Add(rdr(0))
            End While
            con.Close()

            con.Open()
            sql = "SELECT Prd_Id, YrPrd FROM Calendar WHERE GETDATE() BETWEEN sDate AND eDate AND Week_Id > 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read()
                todaysPeriod = rdr("Prd_Id")
                todaysYrPrd = rdr("YrPrd")
            End While
            con.Close()

            thisYear = DatePart(DateInterval.Year, Date.Today())
            cboYear.SelectedIndex = cboYear.FindString(thisYear)

            con.Open()
            sql = "SELECT ID FROM Stores WHERE Status = 'Active' ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                cboStore.Items.Add(rdr(0))
            End While
            con.Close()
            cboStore.SelectedIndex = 0
            thisStore = cboStore.SelectedItem

            Call Load_Data()
            formLoaded = True
        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            MessageBox.Show(ex.Message, "Load Form")
        End Try
    End Sub

    Private Sub Load_Data()
        Try
            Dim dow As Integer
            Dim pct, total As Decimal
            Dim totalArray(6) As Decimal
            dayTbl = New DataTable
            dayTbl.Clear()
            Dim column As New DataColumn
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Period"
            dayTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = dayTbl.Columns("Period")
            dayTbl.PrimaryKey = PrimaryKey
            dayTbl2 = New DataTable
            Dim column2 As New DataColumn
            column2.DataType = System.Type.GetType("System.String")
            column2.ColumnName = "Period"
            dayTbl2.Columns.Add(column2)
            Dim PrimaryKey2(1) As DataColumn
            PrimaryKey2(0) = dayTbl2.Columns("Period")
            dayTbl2.PrimaryKey = PrimaryKey2
            dayTbl.Columns.Add("Sun")
            dayTbl.Columns.Add("Mon")
            dayTbl.Columns.Add("Tue")
            dayTbl.Columns.Add("Wed")
            dayTbl.Columns.Add("Thu")
            dayTbl.Columns.Add("Fri")
            dayTbl.Columns.Add("Sat")
            dayTbl.Columns.Add("Total")
            dayTbl2.Columns.Add("Sun")
            dayTbl2.Columns.Add("Mon")
            dayTbl2.Columns.Add("Tue")
            dayTbl2.Columns.Add("Wed")
            dayTbl2.Columns.Add("Thu")
            dayTbl2.Columns.Add("Fri")
            dayTbl2.Columns.Add("Sat")
            dayTbl2.Columns.Add("Total")

            numberPeriods = 0
            con.Open()
            sql = "SELECT Prd_Id, sDate, eDate FROM Calendar WHERE Year_Id = " & thisYear & " AND Prd_Id > 0 AND Week_Id = 0"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                periodDates(numberPeriods) = rdr("sDate") & " - " & rdr("eDate")
                row = dayTbl.NewRow
                row2 = dayTbl2.NewRow
                row(0) = rdr("Prd_Id")
                row2(0) = rdr("Prd_Id")
                dayTbl.Rows.Add(row)
                dayTbl2.Rows.Add(row2)
                numberPeriods += 1
            End While
            con.Close()

            con.Open()
            sql = "SELECT Prd_Id, Day, Pct FROM DayOfWeekPct WHERE Year_Id = " & thisYear & " " & _
                "AND Str_Id = '" & thisStore & "' ORDER BY Prd_ID, Day"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("Prd_Id")
                row = dayTbl.Rows.Find(oTest)
                dow = rdr("Day")
                oTest = rdr("Pct")
                If IsNumeric(oTest) Then
                    row(dow) = Format(Decimal.Round(oTest, 3, MidpointRounding.AwayFromZero), "##0.0%")
                    total = CDec(oTest)
                    totalArray(dow - 1) += CDec(Math.Round(oTest, 2))
                Else : row(dow) = Nothing
                End If
            End While
            con.Close()

            For Each row As DataRow In dayTbl.Rows
                total = 0
                For i As Integer = 1 To 7
                    oTest = row(i)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        If IsNumeric(Replace(oTest, "%", "")) Then
                            total += CDec(Replace(oTest, "%", ""))
                        End If
                    End If
                Next
                row(8) = Format(total * 0.01, "###%")
            Next

            Dim tavg As Decimal = 0
            row = dayTbl.NewRow
            row(0) = "Avg"
            For i As Integer = 0 To 6
                If totalArray(i) > 0 Then
                    row(i + 1) = Format(totalArray(i) / CDec(numberPeriods), "###%")
                    tavg += totalArray(i)
                End If
            Next
            row(8) = Format(tavg / CDec(numberPeriods), "###%")
            dayTbl.Rows.Add(row)
            dgv1.DataSource = dayTbl.DefaultView
            dgv1.AutoResizeColumns()
            Dim clm As DataGridViewColumn = dgv1.Columns(0)
            clm.Width = 50
            For i As Integer = 1 To dgv1.Rows.Count - 2
                dgv1.Rows(i).Cells(0).Style.ForeColor = Color.Blue
                dgv1.Rows(i).Cells(0).Style.Font = New Font("Ariel", 8, FontStyle.Underline)
            Next

            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ''dgv1.Columns(i).DefaultCellStyle.Format = "p2"
                dgv1.Columns(i).ReadOnly = True
            Next
            ''dgv1.Columns(7).ReadOnly = False
            ''dgv2.DataSource = dayTbl2.DefaultView
            ''dgv2.AutoResizeColumns()
            ''clm = dgv2.Columns(0)
            clm.Width = 50
            ''For i As Integer = 1 To dgv2.ColumnCount - 1
            ''    dgv2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''    dgv2.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''    dgv2.Columns(i).DefaultCellStyle.Format = "p2"
            ''    If i < 8 Then
            ''        dgv2.Columns(i).DefaultCellStyle.BackColor = Color.Cornsilk
            ''    End If
            ''Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD DATA")
            If con.State = ConnectionState.Open Then con.Close()
        End Try
    End Sub

    'Restrict Draft Variance to numbers only
    Private Sub dgv2_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs)
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

    ''Private Sub dgv2_CellTabkey(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    ''    Try
    ''        oTest = dgv2.Rows(e.RowIndex).Cells(e.ColumnIndex)

    ''    Catch ex As Exception

    ''    End Try
    ''End Sub

    ''Public Sub dgv2_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv2.CellValueChanged
    ''Private Sub dgv2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv2.CellValueChanged
    ''    Try
    ''        Dim chng, planTotal, changeTotal, valu As Decimal
    ''        If e.RowIndex > 0 Then rowIndex = e.RowIndex
    ''        If e.ColumnIndex > 0 Then columnIndex = e.ColumnIndex
    ''        If rowIndex < 0 Or columnIndex < 0 Then Exit Sub
    ''        tblRow2 = dayTbl2.Rows(rowIndex)
    ''        If tblRow2(0) <= todaysPeriod Then
    ''            MessageBox.Show("CURRENT OR PREVIOUS PERIODS CANNOT BE CHANGED!", "CHANGE PERCENT")
    ''            tblRow2(columnIndex) = Nothing
    ''            Exit Sub
    ''        End If
    ''        oTest = dgv2.Rows(rowIndex).Cells(columnIndex).Value
    ''        If Not IsNothing(oTest) And IsNumeric(oTest) Then
    ''            changesMade = True

    ''            oTest = tblRow2(columnIndex)
    ''            chng = CDec(tblRow2(columnIndex) * 0.01)
    ''            tblRow2(columnIndex) = Format(chng, "##0%")
    ''            For i As Integer = 1 To 7
    ''                oTest = tblRow2(i)
    ''                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
    ''                    oTest = Replace(oTest, "%", "")
    ''                    chng = CDec(oTest * 0.01)
    ''                    changeTotal += chng
    ''                End If
    ''            Next
    ''            tblRow2(8) = Format(changeTotal, "##0%")
    ''            dgv2.Refresh()
    ''        End If
    ''    Catch ex As Exception
    ''        MessageBox.Show(ex.Message, "Cell Enter")
    ''        If con.State = ConnectionState.Open Then con.Close()
    ''    End Try
    ''End Sub

    Public Sub dvg1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseMove
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
                        lblProcessing.Visible = True
                        Me.Refresh()
                        thisPeriod = oTest
                        foundRow = dayTbl.Rows.Find(thisPeriod)
                        DayPct2.Show()
                        lblProcessing.Visible = False
                        Me.Refresh()
                    End If
                End If
            Case Windows.Forms.MouseButtons.Right

        End Select

    End Sub
    Private Sub cboYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYear.SelectedIndexChanged
        thisYear = cboYear.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem
        If formLoaded Then Call Load_Data()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        Try
            If changesMade = False Then Exit Sub
            Dim total As Decimal = 0
            For Each rw In dayTbl2.Rows
                oTest = rw(8)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    oTest = Replace(oTest, "%", "")
                    If IsNumeric(oTest) Then
                        total += CDec(oTest) * 0.01
                    End If
                End If
            Next
            If total <> 1 Then
                MessageBox.Show("PERIOD TOTAL MUST EQUAL 100% BEFORE CHANGES CAN BE SAVED!", "SAVE CHANGES")
                Exit Sub
            End If

            Dim pct As Decimal
            Dim rowCounter As Integer = 0
            For Each rw In dayTbl2.Rows
                If IsNumeric(rw(0)) Then
                    thisPeriod = rw(0)
                    oTest = rw(8)
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        Dim foundRow As DataRow = dayTbl.Rows.Find(thisPeriod)
                        For i As Integer = 1 To 7
                            pct = Nothing
                            oTest = rw(i)
                            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                                pct = CDec(Replace(oTest, "%", "") * 0.01)
                                foundRow(i) = rw(i)                                               ' set Planned percent to Revised Percent
                                dgv1.Item(i, rowCounter).Style.BackColor = Color.Yellow
                                dgv1.Item(i, rowCounter).Style.Font = New Font("Sans Serif", 8.25, FontStyle.Bold)
                            End If
                            con.Open()
                            sql = "UPDATE DayOfWeekPct SET Pct = " & pct & " WHERE Year_Id >= " & thisYear & " " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " AND Day = " & i & " "
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                            con.Close()
                            rw(i) = Nothing

                        Next
                        rw(8) = Nothing
                    End If
                End If
                rowCounter += 1
            Next
            changesMade = False

        Catch ex As Exception
            If con.State = ConnectionState.Open Then con.Close()
            MessageBox.Show(ex.Message, "SAVE CHANGES")
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