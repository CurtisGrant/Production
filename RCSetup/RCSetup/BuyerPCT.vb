Imports System.Data.SqlClient
Public Class BuyerPCT
    Private con, con2, con3, con4, con5 As SqlConnection
    Private cmd As SqlCommand
    Private sql, thisStore As String
    Private rdr As SqlDataReader
    Private deptTbl, buyerTbl, dateTbl As DataTable
    Private somethingChanged As Boolean
    Private oTest As Object
    Public conString As String

    Private Sub BuyerPCT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        somethingChanged = False
        conString = DBAdmin.conString
        con = New SqlConnection(constring)
        con2 = New SqlConnection(conString)
        con3 = New SqlConnection(conString)
        con4 = New SqlConnection(conString)
        con5 = New SqlConnection(conString)
        deptTbl = New DataTable
        deptTbl.Columns.Add("Dept")
        deptTbl.Columns.Add("WksOH")
        deptTbl.Columns.Add("Turns")
        buyerTbl = New DataTable
        buyerTbl.Columns.Add("Buyer")
        dateTbl = New DataTable
        dateTbl.Columns.Add("Year")
        dateTbl.Columns.Add("Period")
        Dim row As DataRow
        con.Open()
        sql = "SELECT ID FROM Departments WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = deptTbl.NewRow
            row("Dept") = rdr("ID")
            deptTbl.Rows.Add(row)
        End While
        con.Close()

        con.Open()
        sql = "SELECT ID FROM Buyers WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = buyerTbl.NewRow
            row("Buyer") = rdr("ID")
            buyerTbl.Rows.Add(row)
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT Year_Id, Prd_Id FROM Calendar WHERE Prd_Id > 0 AND Year_Id <= Datepart(Year,GetDate()) ORDER BY Year_Id, Prd_Id"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = dateTbl.NewRow
            row("Year") = rdr("Year_Id")
            row("Period") = rdr("Prd_Id")
            dateTbl.Rows.Add(row)
        End While
        con.Close()

        con.Open()
        sql = "SELECT ID FROM Stores WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            cboStore.Items.Add(rdr("ID"))
        End While
        con.Close()

        cboStore.SelectedIndex = 0
        dgv1.DataSource = deptTbl.DefaultView
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv1.Columns(0).ReadOnly = True
        dgv1.Columns("Turns").ReadOnly = True
        dgv1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns("WksOH").DefaultCellStyle.BackColor = Color.Cornsilk
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        MessageBox.Show("Copy all of the sp's in SQL/SalesPlan folder to client database then 'Execute' to build them.",
                        "PAUSE TO SETUP DATABASE")
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        If IsNothing(thisStore) Or thisStore = "" Then
            MessageBox.Show("Select a Store to Continue.", "ERROR")
            Exit Sub
        End If

        lblProcessing.Visible = True
        Call Create_Tables()
        Call Create_Plan()

        con.Open()
        Dim thisYear As Integer = DatePart(DateInterval.Year, Date.Today)
        Dim dept As String = ""
        Dim wks As Integer
        lblHeading.Text = "Updating Plan Sales and Markdowns, Plan and Projected Inventory and Projected Week On Hand"
        Me.Refresh()
        cmd = New SqlCommand("sp_OptionalModule_ProjectedWksOH", con)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 480
        cmd.ExecuteNonQuery()
        For Each row In deptTbl.Rows
            oTest = row("Dept")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then dept = oTest
            oTest = row("WksOH")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then wks = CInt(oTest)
            lblHeading.Text = "Updating Weekly Summary for Dept " & dept
            Me.Refresh()
            cmd = New SqlCommand("sp_RCSetup_UpdateWeeklySummary", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
            cmd.Parameters.Add("@dept", SqlDbType.VarChar).Value = dept
            cmd.Parameters.Add("@wks", SqlDbType.Int).Value = wks
            cmd.CommandTimeout = 480
            cmd.ExecuteNonQuery()
            lblHeading.Text = "Updating Buyer_PCT Weeks on Hand for " & dept
            Me.Refresh()
            cmd = New SqlCommand("sp_RCSetup_UpdateBuyerWksOH", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
            cmd.Parameters.Add("@dept", SqlDbType.VarChar).Value = dept
            cmd.ExecuteNonQuery()
        Next

        con.Close()

        lblProcessing.Visible = False
        somethingChanged = False
        Me.Close()
    End Sub

    Private Sub Create_Tables()
        lblHeading.Text = "Initializing Tables. This step will take a few minutes."
        Me.Refresh()
        Dim theDate As Date
        Dim row, sRow As DataRow
        Dim store As String
        Dim cmd As SqlCommand
        Dim sql As String
        Dim rdr As SqlDataReader
        Dim rdr2 As SqlDataReader
        Dim rdr3 As SqlDataReader
        Dim rdr4 As SqlDataReader
        Dim rdr5 As SqlDataReader
        Dim clss As String
        Dim cnt As Integer = 0

        con.Open()
        cmd = New SqlCommand("sp_RCSetup_CreateSalesPlanTables", con)                    ' Create SQL Sales Plan Tables
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
        cmd.Parameters.Add("@year", SqlDbType.Int).Value = DatePart(DateInterval.Year, Date.Today)
        cmd.ExecuteNonQuery()
        con.Close()

        lblHeading.Text = "Updating Buyer_PCT Table"
        Me.Refresh()
        con.Open()
        cmd = New SqlCommand("sp_RCSetup_UpdateBuyerPCT", con)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
        cmd.CommandTimeout = 480
        cmd.ExecuteNonQuery()
        con.Close()
        lblHeading.Text = "Updating Class_PCT Table"
        Me.Refresh()
        con.Open()
        cmd = New SqlCommand("sp_RCSetup_UpdateClassPCT", con)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
        cmd.CommandTimeout = 480
        cmd.ExecuteNonQuery()
        con.Close()
        lblHeading.Text = "Creating DayOfWeekPCT Table"
        Me.Refresh()
        con.Open()
        cmd = New SqlCommand("sp_RCSetup_UpdateDayOfWeekPCT", con)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
        cmd.CommandTimeout = 480
        cmd.ExecuteNonQuery()
        con.Close()


        ''lblHeading.Text = "Updating Actual Week On Hand"
        ''lblProcessing.Visible = True
        ''Me.Refresh()
        '                         Code moved to sp_RCSetup_UpdateWeeklySummary
        ' ''
        ' ''                       Update Act_Inv_Cost and Act_Inv_Retail in Weekly_Summary
        ' ''
        ''con.Open()
        ''sql = "SELECT eDate, Str_Id, Dept, Buyer, Class, ISNULL(SUM(End_OH * Retail),0) AS OnHand INTO #t1 FROM Item_Inv i " & _
        ''        "JOIN Item_Master m ON m.Item_No = i.Item_No " & _
        ''        "GROUP BY eDate, Str_Id, Dept, Buyer, Class " & _
        ''        "UPDATE w SET w.Act_Inv_Retail = t.OnHand FROM Weekly_Summary w " & _
        ''        "JOIN #t1 t ON t.eDate = w.edate AND t.Str_Id = w.Str_Id AND t.Dept = w.Dept " & _
        ''        "AND t.Buyer = w.Buyer AND t.Class = w.Class"
        ''cmd = New SqlCommand(sql, con)
        ''cmd.ExecuteNonQuery()
        ''sql = "SELECT eDate, Str_Id, Dept, Buyer, Class, ISNULL(SUM(End_OH * Cost),0) AS OnHand INTO #t2 FROM Item_Inv i " & _
        ''        "JOIN Item_Master m ON m.Item_No = i.Item_No " & _
        ''        "GROUP BY eDate, Str_Id, Dept, Buyer, Class " & _
        ''        "UPDATE w SET w.Act_Inv_Cost = t.OnHand FROM Weekly_Summary w " & _
        ''        "JOIN #t2 t ON t.eDate = w.edate AND t.Str_Id = w.Str_Id AND t.Dept = w.Dept " & _
        ''        "AND t.Buyer = w.Buyer AND t.Class = w.Class"
        ''cmd = New SqlCommand(sql, con)
        ''cmd.ExecuteNonQuery()
        ''sql = "SELECT DISTINCT Str_Id FROM Weekly_Summary ORDER BY Str_Id"
        ''cmd = New SqlCommand(sql, con)
        ''rdr = cmd.ExecuteReader
        ''While rdr.Read
        ''    store = rdr("Str_Id")
        ''    con2.Open()
        ''    sql = "SELECT DISTINCT Dept FROM Weekly_Summary WHERE Str_Id = '" & store & "' ORDER By Dept"
        ''    cmd = New SqlCommand(sql, con2)
        ''    rdr2 = cmd.ExecuteReader
        ''    While rdr2.Read
        ''        Dim dept = rdr2("Dept")
        ''        lblHeading.Text = "Updating Department " & dept
        ''        Me.Refresh()
        ''        con3.Open()
        ''        sql = "SELECT DISTINCT Buyer FROM Weekly_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' Order BY Buyer"
        ''        cmd = New SqlCommand(sql, con3)
        ''        rdr3 = cmd.ExecuteReader
        ''        While rdr3.Read
        ''            Dim buyer As String = rdr3("Buyer")
        ''            sql = "SELECT DISTINCT Class FROM Weekly_Summary WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' " & _
        ''                "AND Buyer = '" & buyer & "' ORDER BY Class"
        ''            con4.Open()
        ''            cmd = New SqlCommand(sql, con4)
        ''            rdr4 = cmd.ExecuteReader
        ''            While rdr4.Read
        ''                clss = rdr4("Class")
        ''                Dim invTable As New DataTable
        ''                Dim inv, totl, val As Decimal
        ''                Dim indx2 As Integer
        ''                invTable.Columns.Add("eDate", Type.GetType("System.DateTime"))
        ''                invTable.Columns.Add("Inv", Type.GetType("System.Decimal"))

        ''                con5.Open()
        ''                sql = "SELECT eDate, ISNULL(Act_Inv_Retail,0) AS Inv FROM Weekly_Summary " & _
        ''                    "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' " & _
        ''                    "AND eDate BETWEEN (SELECT MIN(eDate) FROM Weekly_Summary WHERE Str_Id = '" & store & "' " & _
        ''                    "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND [Class] = '" & clss & "') " & _
        ''                    "AND (SELECT MAX(eDate) FROM Weekly_Summary WHERE Str_Id = '" & store & "' " & _
        ''                    "AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' AND [Class] = '" & clss & "') " & _
        ''                    "AND Class = '" & clss & "' AND ISNULL(Act_Inv_Retail,0) > 0 " & _
        ''                    "AND ISNULL(Act_Sales,0) > 0 ORDER BY eDate"
        ''                cmd = New SqlCommand(sql, con5)
        ''                rdr5 = cmd.ExecuteReader
        ''                While rdr5.Read
        ''                    row = invTable.NewRow
        ''                    row("eDate") = rdr5("eDate")
        ''                    inv = rdr5("Inv")
        ''                    row("Inv") = rdr5("Inv")
        ''                    invTable.Rows.Add(row)
        ''                End While
        ''                con5.Close()

        ''                Dim salesTable As New DataTable
        ''                Dim Column As DataColumn = New DataColumn
        ''                Column.DataType = System.Type.GetType("System.DateTime")
        ''                Column.ColumnName = "eDate"
        ''                salesTable.Columns.Add(Column)
        ''                Dim PrimaryKey(1) As DataColumn
        ''                PrimaryKey(0) = salesTable.Columns("eDate")
        ''                salesTable.PrimaryKey = PrimaryKey
        ''                salesTable.Columns.Add("Sales", Type.GetType("System.Decimal"))

        ''                con5.Open()
        ''                sql = "SELECT eDate, ISNULL(Act_Sales,0) Sales FROM Weekly_Summary " & _
        ''                     "WHERE Str_Id = '" & store & "' AND Dept = '" & dept & "' AND Buyer = '" & buyer & "' " & _
        ''                     "AND Class = '" & clss & "' ORDER BY eDate"
        ''                cmd = New SqlCommand(sql, con5)
        ''                rdr5 = cmd.ExecuteReader
        ''                While rdr5.Read
        ''                    sRow = salesTable.NewRow
        ''                    sRow("eDate") = rdr5("eDate")
        ''                    sRow("Sales") = rdr5("Sales")
        ''                    oTest = rdr5("eDate")
        ''                    val = rdr5("Sales")
        ''                    salesTable.Rows.Add(sRow)
        ''                End While
        ''                con5.Close()
        ''                For Each iRow As DataRow In invTable.Rows
        ''                    theDate = iRow("eDate")
        ''                    inv = iRow("Inv")
        ''                    totl = 0
        ''                    cnt = 0
        ''                    indx2 = FindRow(salesTable, theDate)
        ''                    If indx2 > -1 Then
        ''                        Do While totl < inv And indx2 < salesTable.Rows.Count
        ''                            sRow = salesTable(indx2)
        ''                            oTest = sRow("eDate")
        ''                            val = sRow("Sales")
        ''                            totl += val
        ''                            indx2 += 1
        ''                            cnt += 1
        ''                        Loop
        ''                        If totl >= inv Then
        ''                            con5.Open()
        ''                            sql = "UPDATE Weekly_Summary SET Act_WksOH = " & cnt & " WHERE Str_Id = '" & store & "' " & _
        ''                                "AND Dept = '" & dept & "' AND Class = '" & clss & "' AND Buyer = '" & buyer & "' " & _
        ''                                "AND eDate = '" & theDate & "'"
        ''                            cmd = New SqlCommand(sql, con5)
        ''                            cmd.ExecuteNonQuery()
        ''                            con5.Close()
        ''                        End If
        ''                    End If
        ''                Next
        ''                ' MsgBox("Ended idx2=" & indx2 & " rows=" & salesTable.Rows.Count & " " & store & " " & dept & " " & buyer & " " & clss)
        ''            End While
        ''            con4.Close()
        ''            ' MsgBox("Endedxx " & store & " " & dept & " " & buyer & " " & clss)
        ''        End While
        ''        con3.Close()
        ''    End While
        ''    con2.Close()
        ''End While
        ''con.Close()
        lblProcessing.Visible = False
    End Sub

    Private Sub Create_Plan()
        lblHeading.Text = "Creating Sales Plan"
        Me.Refresh()
        con.Open()
        cmd = New SqlCommand("sp_RCSetup_CreateFirstSalesPlan", con)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
        cmd.Parameters.Add("@thisYear", SqlDbType.Int).Value = DatePart(DateInterval.Year, Date.Today)
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        somethingChanged = True
        If e.ColumnIndex = 1 Then
            Dim turns As Decimal
            oTest = deptTbl.Rows(e.RowIndex).Item("WksOH")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If IsNumeric(oTest) Then
                    turns = 52 / CInt(oTest)
                    deptTbl.Rows(e.RowIndex).Item("Turns") = Format(turns, "###.00")
                End If
            End If
        End If
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

    Function FindRow(ByVal dt As DataTable, ByVal eDate As String) As Integer
        For i As Integer = 0 To dt.Rows.Count - 1
            If dt.Rows(i)("eDate") = eDate Then Return i
        Next
        Return -1
    End Function

End Class