Imports System.Data.SqlClient
Imports System.Xml
Public Class Classes
    Public Shared conString, server, database, sql, sql2, sql3 As String
    Public Shared con, con2, con3, con4, con5 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr, rdr2, rdr3, rdr4, rdr5 As SqlDataReader
    Public Shared tbl, wkTbl, dayTbl As DataTable
    Public Shared oTest As Object
    Public Shared today As Date = Date.Now
    Public Shared thisYear, thisStore, thisDept, thisPlan, thisBuyer, thisClass As String
    Public Shared priorYrs As Int16 = 0
    Public Shared columnIndex, lastYear, thisWeek, tRows As Int16
    Public Shared rowIndex As Int16
    Public Shared rnd As Double
    Public Shared tNew As Decimal
    Public Shared theValue As String
    Public Shared thisUser = Environment.UserName
    Public Shared planBy As String
    Public Shared thisPeriod, periodSelected, numberPeriods As Integer
    Public Shared percentagesBy As String
    Public Shared periods(50) As String
    Public Shared todaysPeriod As Integer
    Public Shared somethingChanged As Boolean

    Private Sub Classes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        serverLabel.Text = MainMenu.serverLabel.Text
        Me.Location = New Point(100, 150)
        con = PercentageMaintenance.con
        con2 = PercentageMaintenance.con2
        con3 = PercentageMaintenance.con3
        thisYear = PercentageMaintenance.thisYear
        thisPeriod = PercentageMaintenance.thisPeriod
        thisStore = PercentageMaintenance.thisStore
        thisPlan = PercentageMaintenance.thisPlan
        thisDept = PercentageMaintenance.thisDept
        thisBuyer = PercentageMaintenance.thisBuyer
        thisUser = PercentageMaintenance.thisUser
        numberPeriods = PercentageMaintenance.numberPeriods
        Dim i As Integer
        For i = 0 To PercentageMaintenance.cboStore.Items.Count - 1
            Me.cboStr.Items.Add(PercentageMaintenance.cboStore.Items(i))
        Next
        Me.cboStr.SelectedIndex = PercentageMaintenance.cboStore.SelectedIndex
        Me.cboStr.SelectedItem = PercentageMaintenance.cboStore.SelectedItem
        For i = 0 To PercentageMaintenance.cboDept.Items.Count - 1
            Me.cboDept.Items.Add(PercentageMaintenance.cboDept.Items(i))
        Next
        Me.cboDept.SelectedIndex = PercentageMaintenance.cboDept.SelectedIndex
        Me.cboDept.SelectedItem = PercentageMaintenance.cboDept.SelectedItem
        con.Open()
        sql = "SELECT ID FROM Buyers WHERE Status = 'Active' ORDER BY Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            oTest = rdr(0)
            cboBuyer.Items.Add(rdr(0))
            If thisBuyer = rdr(0) Then cboBuyer.SelectedItem = thisBuyer
        End While
        con.Close()

        Call Show_Class_Plan()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Call Show_Class_Plan()
    End Sub

    Private Sub Show_Class_Plan()
        '                                                         Create the class table columns
        Try
            somethingChanged = False
            Dim clss As String
            Dim tAvg52 As Decimal = 0
            Dim tAvg26 As Decimal = 0
            Dim tAvg12 As Decimal = 0
            Dim tAvg4 As Decimal = 0
            Dim tCurrent As Decimal = 0
            Dim tInv As Decimal = 0
            Dim firstRecord As Boolean = True
            Dim row As DataRow
            wkTbl = New DataTable
            wkTbl.Clear()
            Dim column = New DataColumn()
            column.DataType = System.Type.GetType("System.String")
            column.ColumnName = "Class"
            wkTbl.Columns.Add(column)
            Dim PrimaryKey(1) As DataColumn
            PrimaryKey(0) = wkTbl.Columns("Class")
            wkTbl.PrimaryKey = PrimaryKey
            wkTbl.Columns.Add("52 Week Average")
            wkTbl.Columns.Add("26 Week Average")
            wkTbl.Columns.Add("12 Week Average")
            wkTbl.Columns.Add("4 Week Average")
            wkTbl.Columns.Add("Current Plan%")
            wkTbl.Columns.Add("New Plan%")
            wkTbl.Columns.Add("Enter New %")
            wkTbl.Columns.Add("Current On Hand %")
            con.Open()
            '' ''sql = "SELECT DISTINCT Class FROM Class_PCT WHERE Plan_Year = '" & thisPlan & "' " & _
            '' ''        "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " AND Dept = '" & thisDept & "' " & _
            '' ''        "AND Buyer = '" & thisBuyer & "' ORDER BY Class"
            sql = "SELECT ID AS Class FROM Classes WHERE Dept = '" & thisDept & "' AND Status = 'Active'"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                ''pb1.PerformStep()
                clss = rdr("Class")
                row = wkTbl.NewRow
                row(0) = clss
                wkTbl.Rows.Add(row)
                con2.Open()

                sql = "DECLARE @wk52 date, @wk26 date, @wk12 date, @wk4 date " & _
                "SELECT @wk52 = DATEADD(week,-52,GETDATE()) " & _
                "SELECT @wk26 = DATEADD(week,-26,GETDATE()) " & _
                "SELECT @wk12 = DATEADD(week,-12,GETDATE()) " & _
                "SELECT @wk4 = DATEADD(week,-4,GETDATE()) " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Item_Sales " & _
                    "WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0)  AS Avg52 " & _
                    "INTO #t1 FROM Item_Sales WHERE eDate >= @wk52 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Item_Sales " & _
                    "WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0) AS Avg26 " & _
                    "INTO #t2 FROM Item_Sales WHERE eDate >= @wk26 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Item_Sales " & _
                    "WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0) AS Avg12 " & _
                    "INTO #t3 FROM Item_Sales WHERE eDate >= @wk12 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                "SELECT Class, COALESCE(SUM(Act_Sales) / (SELECT NULLIF(SUM(Act_Sales),0) FROM Item_Sales " & _
                    "WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' AND Dept = '" & thisDept & "' " & _
                        "AND Buyer = '" & thisBuyer & "'),0) AS Avg4 " & _
                    "INTO #t4 FROM Item_Sales WHERE eDate >= @wk4 AND Str_Id = '" & thisStore & "' " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND Class = '" & clss & "' GROUP BY Class " & _
                 "SELECT Class, SUM(Act_Inv_Retail) / (SELECT SUM(Act_Inv_Retail) FROM Item_Sales " & _
                        "WHERE CONVERT(VARCHAR,GETDATE(),101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "') AS pctInv " & _
                        "INTO #t5 FROM Item_Sales w INNER JOIN Buyers b ON b.ID = w.Buyer " & _
                            "WHERE CONVERT(VARCHAR,GETDATE(),101) BETWEEN sDate AND eDate AND Dept = '" & thisDept & "' AND Status = 'Active' " & _
                            "GROUP BY Class " & _
                "SELECT #t1.Class, #t1.Avg52, #t2.Avg26, #t3.Avg12, #t4.Avg4, " & _
                    "(SELECT PCT FROM Class_PCT p WHERE Plan_Year = '" & thisPlan & "' " & _
                        "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & thisPeriod & " " & _
                        "AND Dept = '" & thisDept & "' AND Buyer = '" & thisBuyer & "' " & _
                        "AND p.Class = #t1.Class) AS Prd1, #t5.pctInv " & _
                    "FROM #t1 JOIN #t2 ON #t2.Class = #t1.Class " & _
                        "JOIN #t3 ON #t3.Class = #t1.Class " & _
                        "JOIN #t4 ON #t4.Class = #t1.Class " & _
                        "JOIN #t5 ON #t5.Class = #t1.Class ORDER BY Class"


                'pb1.Visible = False
                cmd = New SqlCommand(sql, con2)
                rdr2 = cmd.ExecuteReader
                While rdr2.Read
                    oTest = rdr2(0)
                 
                    oTest = rdr2("Avg52")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(1) = Format(rdr2(1), "###.0%")
                        tAvg52 += oTest
                    End If
                    oTest = rdr2("Avg26")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(2) = Format(rdr2(2), "###.0%")
                        tAvg26 += oTest
                    End If
                    oTest = rdr2("Avg12")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(3) = Format(rdr2(3), "###.0%")
                        tAvg12 += oTest
                    End If
                    oTest = rdr2("Avg4")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(4) = Format(rdr2(4), "###.0%")
                        tAvg4 += oTest
                    End If
                    oTest = rdr2("Prd1")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(5) = Format(rdr2(5), "###.0%")
                        tCurrent += oTest
                    End If
                    oTest = rdr2("pctInv")
                    If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                        row(8) = Format(oTest, "###.0%")
                        tInv += rdr2(6)
                    End If
                    '' ''wkTbl.Rows.Add(row)
                    '' ''End If
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
            dgv1.DataSource = wkTbl.DefaultView
            dgv1.AutoResizeColumns()
            For i = 1 To dgv1.ColumnCount - 1
                dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgv1.Columns(i).DefaultCellStyle.Format = "p2"
                If i = 5 Then dgv1.Columns(7).DefaultCellStyle.BackColor = Color.Cornsilk
                dgv1.Columns(i).ReadOnly = True
            Next
            dgv1.Columns(7).ReadOnly = False
            dgv1.Refresh()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "LOAD FORM")
            If con.State = ConnectionState.Open Then con.Close()
            If con2.State = ConnectionState.Open Then con2.Close()
        End Try
    End Sub
    ''Public Sub dgv1_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        Dim change As Decimal
        If columnIndex = 7 Then
            oTest = dgv1.Rows(rowIndex).Cells(7).Value
            change = oTest * 0.01
            dgv1.Rows(rowIndex).Cells(6).Value = Format(change, "###.#%")
            tNew = 0
            For i As Integer = 0 To dgv1.Rows.Count - 3
                oTest = dgv1.Rows(i).Cells(6).Value
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    somethingChanged = True
                    oTest = Replace(oTest, "%", "")
                    If oTest <> "" And IsNumeric(oTest) Then
                        oTest = CDec(oTest)
                        tNew += oTest
                    End If
                End If
            Next
        End If
        dgv1.Rows(tRows).Cells(6).Value = Format(tNew * 0.01, "###.#%")
        dgv1.Rows(tRows + 1).Cells(6).Value = Format(1 - tNew * 0.01, "###.0%")
    End Sub

    Private Sub Save_Class_Plan()
        Dim message, title As String
        Dim r As Integer = numberPeriods - cboPrd.SelectedIndex + 1
        If r > 6 Then r = 6
        Dim rx As Integer = r
        message = "Percentages can be saved to the next " & r & " periods. Enter a number up to 6 to copy "
        message &= "instead or click 'OK' to save to the next " & r & " periods."
        title = "Save Percentages"
        oTest = InputBox(message, title)
        If oTest <> "" Then
            rx = CInt(oTest)
            If rx > r Then
                ''MsgBox("The number can't exceed " & r & "!")
                MessageBox.Show("THE NUMBER CAN'T EXCEED " & r & "!", "SAVE CHANGES")
                Exit Sub
            End If
        End If
        Dim thisClass As String
        Dim lastYear As Int16 = thisYear - 1
        For i As Integer = thisPeriod To rx
            For Each row In wkTbl.Rows
                oTest = row(6)
                If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                    If Not IsNothing(oTest) And Not IsDBNull(oTest) And row(0) <> "Total" Then
                        thisClass = row(0).ToString
                        theValue = CDec(Replace(oTest, "%", "") * 0.01)
                        con.Open()
                        sql = "IF NOT EXISTS (SELECT * FROM Class_PCT WHERE Plan_Year = '" & thisPlan & "' " & _
                                "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                                "AND BUYER = '" & thisBuyer & "' AND Class = '" & thisClass & "') " & _
                            "INSERT INTO Class_Pct (Plan_Year, Str_Id, Prd_Id, Dept, Buyer, Class, PCT) " & _
                                "SELECT '" & thisPlan & "', '" & thisStore & "', " & i & ", '" & thisDept & "', '" & _
                                thisBuyer & "', '" & thisClass & "', " & theValue & " " & _
                                "ELSE " & _
                            "UPDATE Class_PCT SET PCT = " & theValue & " WHERE Plan_Year = '" & thisPlan & "' " & _
                            "AND Str_Id = '" & thisStore & "' AND Prd_Id = " & i & " AND Dept = '" & thisDept & "' " & _
                            "AND Buyer = '" & thisBuyer & "' AND Class = '" & thisClass & "'"
                        cmd = New SqlCommand(sql, con)
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                End If
            Next
        Next
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Class_Plan()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        If somethingChanged Then
            Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Call Save_Class_Plan()
                    Me.Close()
                Case DialogResult.No
                    Me.Close()
            End Select
        End If
    End Sub


    Private Sub cboStr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStr.SelectedIndexChanged
        thisStore = cboStr.SelectedItem
    End Sub

    Private Sub cboYr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboYr.SelectedIndexChanged
        thisYear = cboYr.SelectedItem
    End Sub

    Private Sub cboPrd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPrd.SelectedIndexChanged
        thisPeriod = cboPrd.SelectedItem
    End Sub

    Private Sub cboDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDept.SelectedIndexChanged
        thisDept = cboDept.SelectedItem
    End Sub

    Private Sub cboBuyer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboBuyer.SelectedIndexChanged
        thisBuyer = cboBuyer.SelectedItem
    End Sub


    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                e.Cancel = False
            Case CloseReason.UserClosing
                If somethingChanged Then
                    Select Case MessageBox.Show("Do you wish to save changes before exiting?", "CHANGE(S) DETECTED!",
                                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        Case DialogResult.Yes
                            Call Save_Class_Plan()
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