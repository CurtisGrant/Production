Imports System.Data.SqlClient
Imports System.Xml
Imports System.String
Imports Microsoft.VisualBasic
Public Class TableMaintenance
    Public Shared tableName As String
    Public Shared server, database, conString, sql As String
    Private con, con2 As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared rowIndex, columnIndex, NewStatusColumn As Integer
    Public Shared otest As Object
    Public Shared dt As DataTable
    Public Shared somethingChanged, hoursChanged As Boolean

    Private Sub TableMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim open, close As DateTime
            serverLabel.Text = MainMenu.serverLabel.Text
            tableName = MainMenu.tableName
            Me.Text = tableName
            Me.Location = New Point(100, 50)
            somethingChanged = False
            hoursChanged = False
            conString = MainMenu.conString
            con = New SqlConnection(conString)
            con2 = New SqlConnection(conString)
            con.Open()
            dt = New DataTable
            Dim row As DataRow
            dt.Columns.Add("ID")
            dt.Columns.Add("Description")
            If tableName = "Stores" Then
                dt.Columns.Add("Open", GetType(System.String))
                dt.Columns.Add("Close", GetType(System.String))
                dt.Columns.Add("Inv_Loc", GetType(System.String))
            End If
            dt.Columns.Add("Status")
            dt.Columns.Add("Last Change Date", GetType(System.DateTime))
            dt.Columns.Add("Last Change User")
            dt.Columns.Add("Orig Date", GetType(System.DateTime))
            dt.Columns.Add("Changed")
            sql = "SELECT * FROM " & tableName & " ORDER BY ID"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                row = dt.NewRow
                row("ID") = rdr("ID")
                otest = rdr("Description")
                If Not IsDBNull(otest) Then
                    row("Description") = Clean(otest)
                End If
                row("Status") = rdr("Status")
                otest = rdr("Last_Change_Date")
                If Not IsDBNull(otest) And Not IsNothing(otest) Then row("Last Change Date") = CDate(otest)
                row("Last Change User") = rdr("Last_Change_User")
                otest = rdr("Orig_Date")
                If Not IsDBNull(otest) And Not IsNothing(otest) Then row("Orig Date") = CDate(otest)
                If tableName = "Stores" Then
                    otest = rdr("Open")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then
                        row("Open") = String.Format("{0:T}", otest)
                    End If
                    otest = rdr("Close")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then
                        row("Close") = String.Format("{0:T}", otest)
                    End If
                    otest = rdr("Inv_Loc")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then
                        row("Inv_Loc") = otest
                    End If
                End If
                dt.Rows.Add(row)
            End While
            con.Close()

            dgv.DataSource = dt.DefaultView
            dgv.AutoResizeColumns()
            NewStatusColumn = 5
            For i = 0 To 5
                If i <> 2 Then dgv.Columns(i).ReadOnly = True
                dgv.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            dgv.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgv.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'dgv.Columns(6).DefaultCellStyle.Format = "hhhh:mm tt"
            'dgv.Columns(7).DefaultCellStyle.Format = "hhhh:mm tt"
            dgv.Columns("Status").ReadOnly = False
            dgv.Columns("Open").Visible = False
            dgv.Columns("Close").Visible = False
            dgv.Columns("Changed").Visible = False
            ''Dim cbo As New DataGridViewComboBoxColumn
            ''cbo.HeaderText = "New Status"
            ''cbo.Name = "cbo"
            ''cbo.MaxDropDownItems = 2
            ''cbo.Items.Add("Active")
            ''cbo.Items.Add("Inactive")
            ''dgv.Columns.Add(cbo)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub dgv_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim ht As DataGridView.HitTestInfo
            ht = Me.dgv.HitTest(e.X, e.Y)
            Dim rowIdx As Int16 = ht.RowIndex
            Dim columnIdx As Int16 = ht.ColumnIndex
        End If
    End Sub

    Private Sub dgv_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        If columnIndex < 9 Then
            Dim val As String = dt.Rows(rowIndex).Item(columnIndex).ToString
            Dim columnName As String = dgv.Columns(columnIndex).Name
            Dim xDate As DateTime
            Select Case columnName
                Case "Status"
                    If val <> "Active" And val <> "Inactive" Then
                        MessageBox.Show("Status must be either 'Active' or 'Inactive'", "ERROR!")
                        dt.Rows(rowIndex).Item(columnIndex) = Nothing
                        dgv.Refresh()
                        Exit Sub
                    End If
                Case "Open"
                    otest = dt.Rows(rowIndex).Item(columnIndex)
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then
                        If IsNumeric(otest) Then
                            If otest < 12 Then
                                xDate = Convert.ToDateTime(otest & " AM")
                            Else
                                xDate = Convert.ToDateTime(otest & " PM")
                            End If
                        Else
                            xDate = Convert.ToDateTime(otest)
                        End If
                    End If
                    otest = Format(xDate, "hhhh:mm tt")
                    Dim idx As Integer = otest.indexof(":")
                    If idx > 0 Then
                        Dim minut As String = Mid(otest.trim(" "), idx + 2, 2)
                        If minut <> "00" And minut <> "15" And minut <> "30" And minut <> "45" Then
                            MsgBox("Entry not allowed. 00, 15, 30, 45 ONLY!")
                            dt.Rows(rowIndex).Item(columnIndex) = Nothing
                            xDate = Nothing
                        End If
                    End If
                    dt.Rows(rowIndex).Item(columnIndex) = Format(xDate, "hhhh:mm tt")
                    Me.Refresh()
                    ''Exit Sub

                Case "Close"
                    otest = dt.Rows(rowIndex).Item(columnIndex)
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then
                        If IsNumeric(otest) Then
                            If otest < 12 Then
                                xDate = Convert.ToDateTime(otest & " AM")
                            Else
                                xDate = Convert.ToDateTime(otest & " PM")
                            End If
                        Else
                            xDate = Convert.ToDateTime(otest)
                        End If
                    End If
                    otest = Format(xDate, "hhhh:mm tt")
                    Dim idx As Integer = otest.indexof(":")
                    If idx > 0 Then
                        Dim minut As String = Mid(otest.trim(" "), idx + 2, 2)
                        If minut <> "00" And minut <> "15" And minut <> "30" And minut <> "45" Then
                            MsgBox("Entry not allowed. 00, 15, 30, 45 ONLY!")
                            dt.Rows(rowIndex).Item(columnIndex) = Nothing
                            xDate = Nothing
                        End If
                    End If
                    dt.Rows(rowIndex).Item(columnIndex) = Format(xDate, "hhhh:mm tt")
                    Me.Refresh()
                    ''Exit Sub
            End Select
          
            somethingChanged = True
            If columnName = "Open" Or columnName = "Close" Then hoursChanged = True
            dt.Rows(rowIndex).Item("Changed") = "Y"
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
        MainMenu.Show()
    End Sub

    ''Private Sub dgv_EditingControlShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgv.EditingControlShowing
    ''    If dgv.CurrentCell.ColumnIndex = NewStatusColumn Then
    ''        Dim combo As ComboBox = CType(e.Control, ComboBox)
    ''        rowIndex = dgv.CurrentRow.Index
    ''        If (combo IsNot Nothing) Then
    ''            RemoveHandler combo.SelectionChangeCommitted, New EventHandler(AddressOf ComboBox_SelectionChangeCommitted)
    ''            AddHandler combo.SelectionChangeCommitted, New EventHandler(AddressOf ComboBox_SelectionChangeCommitted)
    ''        End If
    ''    End If
    ''End Sub

    ''Private Sub ComboBox_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ''    Dim combo As ComboBox = CType(sender, ComboBox)
    ''    dgv.Rows(rowIndex).Cells(2).Value = combo.SelectedItem
    ''    somethingChanged = True
    ''    dt.Rows(rowIndex).Item("Changed") = "Y"
    ''    otest = combo.SelectedItem
    ''    dt.Rows(rowIndex).Item("Status") = combo.SelectedItem
    ''    dgv.Refresh()
    ''End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        lblProcessing.Visible = True
        Me.Refresh()
        Dim id As String = ""
        Dim descr As String = ""
        Dim status As String = ""
        Dim val As String = ""
        Dim invLoc As String = ""
        Dim now As Date = Date.Now
        Dim open, close As DateTime
        Dim user As String = Environment.UserName
        For Each row In dt.Rows
            otest = row("Changed")
            If Not IsDBNull(otest) Then
                If row("Changed") = "Y" Then
                    otest = row("ID")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then id = otest
                    otest = row("DESCRIPTION")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then descr = otest Else descr = ""
                    otest = row("STATUS")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then status = otest Else status = ""
                    If tableName = "Stores" Then
                        otest = row("Inv_Loc")
                        If Not IsDBNull(otest) And Not IsNothing(otest) Then
                            invLoc = otest
                            sql = "UPDATE " & tableName & " SET Description = '" & descr & "', Status = '" & status & "', " & _
                        "Inv_Loc = '" & invLoc & "', " & _
                        "Last_Change_Date = '" & now & "', Last_Change_User = '" & user & "' WHERE ID = '" & id & "'"
                        End If
                    Else
                        sql = "UPDATE " & tableName & " SET Description = '" & descr & "', Status = '" & status & "', " & _
                        "Last_Change_Date = '" & now & "', Last_Change_User = '" & user & "' WHERE ID = '" & id & "'"
                    End If
                    con.Open()
                    
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                    con.Close()

                    If tableName = "Stores" Then
                        otest = row("Open")
                        If Not IsDBNull(otest) And Not IsNothing(otest) Then
                            ''If IsNumeric(otest) Then
                            ''    val = Convert.ToDateTime(otest & " AM")
                            ''Else
                            ''    val = Convert.ToDateTime(otest)
                            ''End If
                            open = DateTime.Parse(otest)
                            hoursChanged = True
                        End If
                        otest = row("Close")
                        If Not IsDBNull(otest) And Not IsNothing(otest) Then
                            ''If IsNumeric(otest) Then
                            ''    val = Convert.ToDateTime(otest & "PM")
                            ''Else
                            ''    val = Convert.ToDateTime(otest)
                            ''End If
                            close = DateTime.Parse(otest)
                            If close < open Then close = DateAdd(DateInterval.Hour, 24, close)
                            hoursChanged = True
                        End If
                        If hoursChanged Then
                            con.Open()
                            sql = "UPDATE " & tableName & " SET Last_Change_Date = '" & now & "', Last_Change_User = '" & user & "', " & _
                        "[Open] = '" & open & "', [Close] = '" & close & "' " & _
                        "WHERE ID = '" & id & "'"
                            cmd = New SqlCommand(sql, con)
                            cmd.ExecuteNonQuery()
                            con.Close()
                        End If
                        If hoursChanged Then Call Compute_Hour_Pct(id, open, close)
                    End If
                End If
            End If
        Next
        somethingChanged = False
        con.Close()

        con.Open()
        Dim badTimes As Integer = 0
        sql = "SELECT COUNT(*) AS cnt FROM Hour_Pct WHERE Pct IS NULL"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            badTimes = rdr("cnt")
        End While
        con.Close()

        con.Open()
        Dim goodDate As String = ""
        If badTimes > 0 Then
            sql = "SELECT Str_Id, CONVERT(Varchar,MIN(TimeOfDay)) AS time FROM Hour_Pct WHERE Pct IS NOT NULL " & _
                "GROUP BY Str_Id"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                goodDate = rdr("time")
            End While
            MessageBox.Show("No sales were found prior to " & goodDate & ". Change Open time and try again.", "ERROR! BAD OPEN TIME!")
        End If
        con.Close()
        badTimes = 0

        con.Open()
        sql = "SELECT COUNT(*) AS cnt FROM Hour_Pct WHERE Pct IS NULL"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            badTimes = rdr("cnt")
        End While
        con.Close()

        con.Open()
        If badTimes > 0 Then
            sql = "SELECT CONVERT(Varchar,MAX(TimeOfDay)) AS time FROM Hour_Pct WHERE Pct IS NOT NULL"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader()
            While rdr.Read
                goodDate = rdr("time")
            End While
            MessageBox.Show("No sales were found after " & goodDate & ". Change Open time and try again.", "ERROR! BAD OPEN TIME!")
        End If
        con.Close()
        lblProcessing.Visible = False
        Me.Refresh()
    End Sub

    Private Sub Compute_Hour_Pct(ByVal id As String, open As DateTime, close As DateTime)
        con2.Open()
        Dim ox As DateTime = open.ToLongTimeString
        Dim cx As DateTime = close.ToLongTimeString
        cmd = New SqlCommand("dbo.sp_Update_Hour_Pct", con2)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = id
        cmd.Parameters.Add("@open", SqlDbType.VarChar).Value = open
        cmd.Parameters.Add("@close", SqlDbType.VarChar).Value = close
        cmd.ExecuteNonQuery()
        con2.Close()
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
    Private Function Clean(ByVal oTest)
        ''oTest = Replace(oTest, "'", "")
        Return oTest
    End Function

    Private Sub btnHours_Click(sender As Object, e As EventArgs) Handles btnHours.Click
        StoreHours.Show()
    End Sub
End Class