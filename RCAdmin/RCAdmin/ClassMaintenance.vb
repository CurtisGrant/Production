Imports System.Data.SqlClient
Imports System.Xml
Public Class ClassMaintenance
    Public Shared tableName As String
    Public Shared server, database, conString, sql As String
    Public con As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared rowIndex, columnIndex, NewStatusColumn As Integer
    Public Shared otest As Object
    Public Shared dt As DataTable
    Public Shared NewStatusColumnAdded As Boolean = False
    Public Shared somethingChanged As Boolean

    Private Sub ClassMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        somethingChanged = False
        serverLabel.Text = MainMenu.serverLabel.Text
        tableName = MainMenu.tableName
        Me.Location = New Point(100, 50)
        conString = MainMenu.conString
        con = New SqlConnection(conString)
        con.Open()
        sql = "SELECT ID FROM Departments WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            cbodept.Items.Add(rdr(0))
        End While
        con.Close()
    End Sub

    Private Sub cbodept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbodept.SelectedIndexChanged
        con.Open()
        dt = New DataTable
        dt.Clear()
        Dim row As DataRow
        dt.Columns.Add("ID")
        dt.Columns.Add("Description")
        dt.Columns.Add("Status")
        dt.Columns.Add("Last Change Date", GetType(System.DateTime))
        dt.Columns.Add("Last Change User")
        dt.Columns.Add("Orig Date", GetType(System.DateTime))
        dt.Columns.Add("Changed")
        sql = "SELECT * FROM Classes WHERE Dept = '" & cbodept.SelectedItem & "' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        ''Dim adapter As SqlDataAdapter = New SqlDataAdapter(sql, con)
        ''adapter.Fill(dt)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = dt.NewRow
            row("ID") = rdr("ID")
            row("Description") = rdr("Description")
            row("Status") = rdr("Status")
            otest = rdr("Last_Change_Date")
            If Not IsDBNull(otest) And Not IsNothing(otest) Then row("Last Change Date") = CDate(otest)
            row("Last Change User") = rdr("Last_Change_User")
            otest = rdr("Orig_Date")
            If Not IsDBNull(otest) And Not IsNothing(otest) Then row("Orig Date") = CDate(otest)
            dt.Rows.Add(row)
        End While
        con.Close()
        dgv.Columns.Clear()
        dgv.DataSource = dt.DefaultView
        dgv.AutoResizeColumns()
        For i = 0 To dgv.Columns.Count - 1
            If i <> 2 Then dgv.Columns(i).ReadOnly = True
        Next
        dgv.Columns("Changed").Visible = False
        ''Dim cbo As New DataGridViewComboBoxColumn
        ''cbo.HeaderText = "New Status"
        ''cbo.Name = "cbo"
        ''cbo.MaxDropDownItems = 3
        ''cbo.Items.Add("Active")
        ''cbo.Items.Add("Inactive")
        ''dgv.Columns.Add(cbo)
        ''NewStatusColumn = dgv.Columns.Count - 1
        ''Dim rx As Integer = dgv.Rows.Count - 1
        ''For r As Integer = 0 To dgv.Rows.Count - 1
        ''    otest = dgv.Rows(r).Cells(2).Value
        ''    If IsDBNull(otest) Then
        ''        dgv.Rows(r).Cells(7).Value = "Active"
        ''    Else
        ''        dgv.Rows(r).Cells(7).Value = otest
        ''    End If
        ''Next
        ''NewStatusColumnAdded = True
    End Sub


    Public Sub dgv_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim ht As DataGridView.HitTestInfo
            ht = Me.dgv.HitTest(e.X, e.Y)
            Dim rowIdx As Int16 = ht.RowIndex
            Dim columnIdx As Int16 = ht.ColumnIndex
        End If
    End Sub

    ''Public Sub dgv_CellEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellValueChanged
    Private Sub dgv_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv.CellValueChanged
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        Dim columnName As String = dgv.Columns(columnIndex).Name
        Dim val As String = dt.Rows(rowIndex).Item(columnIndex).ToString
        Select Case columnName
            Case "Status"
                If val <> "Active" And val <> "Inactive" Then
                    MessageBox.Show("Status must be either 'Active' or 'Inactive'", "ERROR!")
                    dt.Rows(rowIndex).Item(columnIndex) = Nothing
                    dgv.Refresh()
                    Exit Sub
                End If
        End Select
        dt.Rows(rowIndex).Item("Changed") = "Y"
        dgv.Refresh()
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
    ''    otest = combo.SelectedItem
    ''    dgv.Rows(rowIndex).Cells(2).Value = combo.SelectedItem
    ''    somethingChanged = True
    ''    dt.Rows(rowIndex).Item("Changed") = "Y"
    ''    dgv.Refresh()
    ''End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Call Save_Changes()
    End Sub

    Private Sub Save_Changes()
        con.Open()
        Dim id As String = ""
        Dim descr As String = ""
        Dim status As String = ""
        Dim now As Date = Date.Now
        Dim user As String = Environment.UserName
        For Each row In dt.Rows
            otest = row("Changed")
            If Not IsDBNull(otest) Then
                If otest = "Y" Then
                    otest = row("ID")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then id = otest
                    otest = row("DESCRIPTION")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then descr = otest Else descr = ""
                    otest = row("STATUS")
                    If Not IsDBNull(otest) And Not IsNothing(otest) Then status = otest Else status = ""
                    sql = "UPDATE Classes SET Status = '" & status & "', " & _
                        "Last_Change_Date = '" & now & "', Last_Change_User = '" & user & "' WHERE ID = '" & id & "'"
                    cmd = New SqlCommand(sql, con)
                    cmd.ExecuteNonQuery()
                End If
            End If
        Next
        con.Close()
        somethingChanged = False
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
End Class