Imports System.Data.SqlClient
Public Class ControlsAdmin
    Public Shared con As SqlConnection = MainMenu.con
    Private Shared cmd As SqlCommand
    Private Shared rdr As SqlDataReader
    Private Shared sql As String
    Private Shared tbl As DataTable
    Private Shared oTest As Object
    Private Shared user As String = Environment.UserName
    Private Shared selectedID, selectedParam As String
    Private Shared rowIdx, columnIdx As Integer
    Private Shared prevRow As Integer = -1
    Private Sub ControlsAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblServer.Text = "Server: " & MainMenu.Server & "  Database: " & MainMenu.dBase
        Call Load_Data()
    End Sub
    Private Sub Load_Data()
        tbl = New DataTable
        tbl.Columns.Add("ID", GetType(System.String))
        tbl.Columns.Add("Parameter", GetType(System.String))
        tbl.Columns.Add("Value", GetType(System.String))
        tbl.Columns.Add("Value2", GetType(System.String))
        tbl.Columns.Add("Value3", GetType(System.String))
        tbl.Columns.Add("Value4", GetType(System.String))
        tbl.Columns.Add("Value5", GetType(System.String))
        tbl.Columns.Add("Last_Change_Date", GetType(System.DateTime))
        tbl.Columns.Add("Last_Change_User", GetType(System.String))
        tbl.Columns.Add("Orig_Date", GetType(System.String))
        Dim row As DataRow
        con.Open()
        sql = "SELECT * FROM Controls ORDER BY ID, Parameter"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            row = tbl.NewRow
            row("ID") = rdr("ID")
            row("Parameter") = rdr("Parameter")
            oTest = rdr("Value")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Value") = oTest
            oTest = rdr("Value2")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Value2") = oTest
            oTest = rdr("Value3")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Value3") = oTest
            oTest = rdr("Value4")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Value4") = oTest
            oTest = rdr("Value5")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Value5") = oTest
            oTest = rdr("Last_Change_Date")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Last_Change_Date") = CDate(oTest)
            oTest = rdr("Last_Change_User")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then row("Last_Change_User") = oTest
            oTest = rdr("Orig_Date")
            If Not IsDBNull(oTest) AndAlso Not IsNothing(oTest) Then
                oTest = FormatDateTime(oTest, vbShortDate) : row("Orig_Date") = oTest
            End If
            tbl.Rows.Add(row)
        End While
        con.Close()

        dgv1.DataSource = tbl.DefaultView
        dgv1.AutoResizeColumns()


    End Sub

    Public Sub dgv1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv1.MouseDown
        Dim ht As DataGridView.HitTestInfo
        ht = Me.dgv1.HitTest(e.X, e.Y)
        rowIdx = ht.RowIndex
        columnIdx = ht.ColumnIndex
        Dim id As String = dgv1.Rows(rowIdx).Cells(0).Value
        Dim param As String = dgv1.Rows(rowIdx).Cells(1).Value
        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                If columnIdx = 0 And rowIdx > -1 Then
                    ''  dgv1.Rows(rowIdx).DefaultCellStyle.BackColor = Color.Yellow
                    Call Get_Changes(id, param)
                    selectedID = dgv1.Rows(rowIdx).Cells(0).Value
                    selectedParam = dgv1.Rows(rowIdx).Cells(1).Value
                End If

        End Select
    End Sub

    Private Sub Get_Changes(ByVal id As String, param As String)
        txtID.Text = Nothing : txtParam.Text = Nothing : txtVal.Text = Nothing : txtVal2.Text = Nothing
        txtVal3.Text = Nothing : txtVal4.Text = Nothing : txtVal5.Text = Nothing
        con.Open()
        sql = "SELECT * FROM Controls WHERE ID = '" & id & "' AND Parameter = '" & param & "'"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            txtID.Text = rdr("ID") : txtParam.Text = rdr("Parameter") : txtVal.Text = rdr("Value")
            oTest = rdr("Value2")
            If Not IsDBNull(oTest) Then txtVal2.Text = CStr(oTest)
            oTest = rdr("Value3")
            If Not IsDBNull(oTest) Then txtVal3.Text = CStr(oTest)
            oTest = rdr("Value4")
            If Not IsDBNull(oTest) Then txtVal4.Text = CStr(oTest)
            oTest = rdr("Value5")
            If Not IsDBNull(oTest) Then txtVal5.Text = CStr(oTest)
        End While
        con.Close()
        Me.Refresh()
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim id, param As String
        Dim val, val2, val3, val4, val5 As Object
        Dim dte As Date = Date.Today
        con.Open()
        id = txtID.Text
        param = txtParam.Text
        val = txtVal.Text
        val2 = txtVal2.Text
        val3 = txtVal3.Text
        val4 = txtVal4.Text
        val5 = txtVal5.Text
        sql = "IF NOT EXISTS (SELECT ID FROM Controls WHERE ID = '" & id & "' AND PArameter = '" & param & "') " & _
            "INSERT INTO Controls (ID, Parameter, Value, Value2, Value3, Value4, Value5, Orig_Date) " & _
            "SELECT '" & id & "', '" & param & "', '" & val & "', '" & val2 & "', '" & val3 & "', '" &
                val4 & "', '" & val5 & "', CONVERT(Date,GETDATE()) " & _
            "ELSE " & _
            "UPDATE Controls SET Parameter = '" & param & "', Value = '" & val & "', " & _
                "Value2 = '" & val2 & "', Value3 = '" & val3 & "', Value4 = '" & val4 & "', Value5 = '" & val5 & "', " & _
                "Last_Change_Date = '" & dte & "', Last_Change_User = '" & user & "' " & _
            "WHERE ID = '" & id & "' AND Parameter = '" & param & "'"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        con.Close()
        txtID.Text = Nothing : txtParam.Text = Nothing : txtVal.Text = Nothing : txtVal2.Text = Nothing
        txtVal3.Text = Nothing : txtVal4.Text = Nothing : txtVal5.Text = Nothing
        Call Load_Data()
        Me.Refresh()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim ans As Integer = MessageBox.Show("Are you sure?", "DELETE CONTROL", MessageBoxButtons.YesNo)
        If ans = DialogResult.Yes Then
            con.Open()
            sql = "DELETE FROM Controls WHERE ID = '" & selectedID & "' AND Parameter = '" & selectedParam & "'"
            cmd = New SqlCommand(sql, con)
            cmd.ExecuteNonQuery()
            con.Close()
            Dim row As DataGridViewRow = dgv1.Rows(rowIdx)
            dgv1.Rows.Remove(row)
            txtID.Text = Nothing : txtParam.Text = Nothing : txtVal.Text = Nothing : txtVal2.Text = Nothing
            txtVal3.Text = Nothing : txtVal4.Text = Nothing : txtVal5.Text = Nothing
        End If
    End Sub
End Class