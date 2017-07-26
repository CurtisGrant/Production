Imports System.Data.SqlClient
Public Class StoreHours
    Public Shared server, database, conString, sql As String
    Private Shared con, con2 As SqlConnection
    Private Shared cmd As SqlCommand
    Private Shared rdr, rdr2 As SqlDataReader
    Private Shared oTest As Object
    Private Shared thisStore As String
    Private Shared tbl As DataTable
    Private Shared row As DataRow
    Private Shared fromDate, thruDate As Date

    Private Sub StoreHours_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conString = MainMenu.conString
        con = New SqlConnection(conString)
        con2 = New SqlConnection(conString)
        con.Open()
        sql = "SELECT ID FROM Stores WHERE Status = 'Active' ORDER BY ID"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        cboStore.Items.Add("All")
        While rdr.Read
            cboStore.Items.Add(rdr("ID"))
        End While
        con.Close()

        Dim sDate As Date
        con.Open()
        sql = "SELECT DISTINCT sDate FROM Sales_Summary WHERE sDate IS NOT NULL Order BY sDate"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            sDate = rdr("sDate")
            cboFrom.Items.Add(sDate)
            cboThru.Items.Add(DateAdd(DateInterval.Day, 6, sDate))
        End While
        con.Close()

        con.Open()
        sql = "SELECT DISTINCT lastUpdate AS Date FROM Store_Hours"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            txtLastUpdate.Text = rdr("Date")
        End While
        con.Close()
    End Sub

    Private Sub Load_The_Hours(ByVal store As String)
        Dim dte As Date
        Dim xDate As DateTime
        Dim timeOnly As TimeSpan
        Dim inStoreHours As Boolean = False
        Dim dteOffset As DateTimeOffset
        tbl = New DataTable
        Dim column As New DataColumn()
        column.DataType = System.Type.GetType("System.String")
        column.ColumnName = "DOW"
        tbl.Columns.Add(column)
        Dim PrimaryKey(1) As DataColumn
        PrimaryKey(0) = tbl.Columns("DOW")
        tbl.PrimaryKey = PrimaryKey
        column = New DataColumn("Day", GetType(System.String))
        tbl.Columns.Add(column)
        column = New DataColumn("Open", GetType(System.String))
        column.AllowDBNull = True
        tbl.Columns.Add(column)
        column = New DataColumn("Close", GetType(System.String))
        column.AllowDBNull = True
        tbl.Columns.Add(column)

        For i As Integer = 0 To 6
            row = tbl.NewRow
            dte = Date.Parse(GetPreviouSunday(Date.Today, i))
            row("DOW") = i
            dteOffset = New DateTimeOffset(dte, TimeZoneInfo.Local.GetUtcOffset(dte))
            row("Day") = dte.ToString("ddd")
            con.Open()
            sql = "SELECT MIN(TimeOfDay) AS opn, MAX(TimeOfDay) AS clse FROM Store_Hours " & _
                "WHERE Str_Id = '" & thisStore & "' " & " AND Day = " & i + 1 & ""
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader
            While rdr.Read
                oTest = rdr("opn")
                If Not IsDBNull(oTest) Then
                    inStoreHours = True
                    row("Open") = oTest
                    oTest = rdr("clse")
                    If Not IsDBNull(oTest) Then row("Close") = oTest
                Else
                    con2.Open()
                    sql = "SELECT MIN(CONVERT(Time,TRANS_DATE)) AS MIN, MAX(CONVERT(Time,TRANS_DATE)) AS MAX FROM Daily_Transaction_Log " & _
                        "WHERE DATEPART(dw,TRANS_DATE) = " & i + 1 & " " & _
                        "AND CONVERT(Date,TRANS_DATE) BETWEEN '" & fromDate & "' AND '" & thruDate & "' " & _
                        "AND (STORE = '" & thisStore & "' OR '" & thisStore & "' = 'All')"
                    cmd = New SqlCommand(sql, con2)
                    rdr2 = cmd.ExecuteReader
                    While rdr2.Read
                        If Not IsDBNull(rdr2("MIN")) Then row("Open") = rdr2("MIN") Else row("Open") = Nothing
                        If Not IsDBNull(rdr2("MAX")) Then row("Close") = rdr2("MAX") Else row("Close") = Nothing
                    End While
                    con2.Close()
                End If

                
            End While
            con.Close()
            tbl.Rows.Add(row)
        Next

        dgv1.DataSource = tbl.DefaultView
        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv1.Columns(0).ReadOnly = True
        dgv1.Columns(1).ReadOnly = True
    End Sub

    Private Sub dgv1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellValueChanged
        Dim rowIndex, columnIndex As Integer
        If e.RowIndex > -1 Then rowIndex = e.RowIndex
        If e.ColumnIndex > -1 Then columnIndex = e.ColumnIndex
        If columnIndex < 9 Then
            Dim val As String = tbl.Rows(rowIndex).Item(columnIndex).ToString
            Dim columnName As String = dgv1.Columns(columnIndex).Name
            Dim xDate As DateTime
            Dim timeOnly As TimeSpan
            Dim inputOK As Boolean
            Select Case columnName
                Case "Open"
                    oTest = tbl.Rows(rowIndex).Item(columnIndex).ToString
                    Dim idx As Integer = oTest.indexof(":")
                    If idx > 0 Then
                        oTest = tbl.Rows(rowIndex).Item(columnIndex)
                        inputOK = (oTest Like "##:##")
                        If inputOK Then
                            xDate = Convert.ToDateTime(oTest)
                            timeOnly = xDate.TimeOfDay
                        End If
                    Else
                        MessageBox.Show("Enter time in 24 hour format ie. ""21:00"" for 9:00 PM.", "ERROR!")
                        Exit Sub
                    End If
                    tbl.Rows(rowIndex).Item(columnIndex) = Format(xDate, "hhhh:mm tt")
                    '' tbl.Rows(rowIndex).Item("openTime") = timeOnly
                    Me.Refresh()
                Case "Close"
                    oTest = tbl.Rows(rowIndex).Item(columnIndex).ToString
                    Dim idx As Integer = oTest.indexof(":")
                    If idx > 0 Then
                        oTest = tbl.Rows(rowIndex).Item(columnIndex)
                        inputOK = (oTest Like "##:##")
                        If inputOK Then
                            xDate = Convert.ToDateTime(oTest)
                            timeOnly = xDate.TimeOfDay
                        End If
                    Else
                        MessageBox.Show("Enter time in 24 hour format ie. ""21:00"" for 9:00 PM.", "ERROR!")
                        Exit Sub
                    End If
                    tbl.Rows(rowIndex).Item(columnIndex) = Format(xDate, "hhhh:mm tt")
                    Me.Refresh()
            End Select
        End If
    End Sub

    Private Sub cboStore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStore.SelectedIndexChanged
        thisStore = cboStore.SelectedItem
        Call Load_The_Hours(thisStore)
    End Sub

    Function GetPreviouSunday(fromDate As Date, var As Integer) As Date
        Return fromDate.AddDays(var - fromDate.DayOfWeek)
    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If thisStore = "All" Then
            For Each item In cboStore.Items
                thisStore = item
                If thisStore <> "All" Then Call Update_Store_Hours(thisStore)
            Next
        Else
            Call Update_Store_Hours(thisStore)
        End If
        txtLastUpdate.Text = DateAndTime.Now
        Me.Refresh()
    End Sub
    Private Sub Update_Store_Hours(ByVal thisStore As String)
        Dim dayofweek As Integer
        Dim open As Date = fromDate
        Dim close As Date = thruDate
        Dim opentime As String = ""
        Dim closetime As String = ""
        con.Open()
        For Each row As DataRow In tbl.Rows
            opentime = ""
            closetime = ""
            oTest = row("DOW")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then DayOfWeek = oTest
            End If
            oTest = row("Open")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then opentime = Convert.ToDateTime(oTest) Else opentime = ""
            End If
            oTest = row("Close")
            If Not IsDBNull(oTest) And Not IsNothing(oTest) Then
                If oTest <> "" Then closetime = Convert.ToDateTime(oTest) Else closetime = ""
            End If
            If opentime <> "" And closetime <> "" Then
                ''If close() < open Then close = DateAdd(DateInterval.Hour, 24, close)
                cmd = New SqlCommand("sp_Update_Store_Hours", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("@store", SqlDbType.VarChar).Value = thisStore
                cmd.Parameters.Add("@dayofweek", SqlDbType.Int).Value = dayofweek + 1
                cmd.Parameters.Add("@fromdate", SqlDbType.Date).Value = fromDate
                cmd.Parameters.Add("@thrudate", SqlDbType.Date).Value = thruDate
                cmd.Parameters.Add("@open", SqlDbType.VarChar).Value = opentime
                cmd.Parameters.Add("@close", SqlDbType.VarChar).Value = closetime
                cmd.ExecuteNonQuery()
            End If
        Next
        con.Close()
    End Sub

    Private Sub cboFrom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFrom.SelectedIndexChanged
        fromDate = cboFrom.SelectedItem
    End Sub

    Private Sub cboThru_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboThru.SelectedIndexChanged
        thruDate = cboThru.SelectedItem

    End Sub

    Private Sub btnCompute_Click(sender As Object, e As EventArgs)
        Dim dayofweek As Integer
        Dim open As Date = fromDate
        Dim close As Date = thruDate
        Dim opentime, closetime As String
        Dim xyz As String
        con.Open()
        For Each row As DataRow In tbl.Rows
            oTest = row("Open")
            If Not IsDBNull(oTest) Then
                xyz = open & " " & oTest
                ''openLong = Convert.ToDateTime(xyz)
            End If
            oTest = row("Close")
            If Not IsDBNull(oTest) Then
                xyz = close & " " & Convert.ToDateTime(oTest)
                ''closeLong = Convert.ToDateTime(xyz)
            End If
            ''If closeLong < openLong Then closeLong = DateAdd(DateInterval.Hour, 24, close)
            dayofweek = row("Day")
            cmd = New SqlCommand("sp_Update_Store_Hours", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("@store", thisStore)

        Next
    End Sub
End Class