Imports System.Xml
Imports System.Data.SqlClient

Public Class Calendar
    Public Shared conString, sql As String
    Public Shared con As SqlConnection
    Public Shared cmd As SqlCommand
    Public Shared rdr As SqlDataReader
    Public Shared oTest As Object
    Public Shared dt As DataTable
    Private Sub Calendar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fld As String = ""
        Dim val As String = ""
        Dim server As String = ""
        Dim database As String = ""
        Dim path As String = ""
        Dim xmlPath As String
        Dim xmlReader As XmlTextReader = New XmlTextReader("c:\RTRAC.xml")
        While xmlReader.Read
            Select Case xmlReader.NodeType
                Case XmlNodeType.Element
                    fld = xmlReader.Name
                Case XmlNodeType.Text
                    val = xmlReader.Value
                Case XmlNodeType.EndElement
                    If fld = "Server" Then server = val
                    If fld = "Database" Then database = val
                    If fld = "Path" Then path = val
            End Select
        End While
        xmlPath = path & "\Items.xml"
        conString = "Server=" & server & ";Initial Catalog=" & database & ";Integrated Security=True"
        con = New SqlConnection(conString)
        For i As Integer = 2000 To 2050
            cboYear.Items.Add(i)
        Next
        cboYear.SelectedIndex = 13
    End Sub

    Private Sub btnPeriod_Click(sender As Object, e As EventArgs) Handles btnPeriod.Click
        Dim sDate, eDate As Date
        Dim year As Integer
        '' ''If Date.TryParseExact(txtYrBegin.Text.ToString(), "yyyy-mm-dd", _
        '' ''                 System.Globalization.CultureInfo.CurrentCulture, _
        '' ''                 Globalization.DateTimeStyles.None, test) Then
        '' ''    sDate = CDate(txtYrBegin.Text.ToString)
        '' ''Else
        '' ''    MsgBox("Enter the date the business year begins as yyyy-mm-dd and try again!")
        '' ''    Exit Sub
        '' ''End If
        '' ''If Date.TryParseExact(txtYrEnd.Text.ToString(), "yyyy-mm-dd", _
        '' ''                 System.Globalization.CultureInfo.CurrentCulture, _
        '' ''                 Globalization.DateTimeStyles.None, test) Then
        '' ''    eDate = CDate(txtYrEnd.Text.ToString)
        '' ''Else
        '' ''    MsgBox("Enter the date the business year ends as yyyy-mm-dd and try again!")
        '' ''    Exit Sub
        '' ''End If
        ' oTest = txtYear
        If IsNumeric(oTest) Then year = CInt(oTest)
        Dim yrprd As Integer = Microsoft.VisualBasic.Right(year, 2) & "00"
        con.Open()
        sql = "IF NOT EXISTS (SELECT * FROM Calendar WHERE sDate = '" & sDate & "' AND eDate = '" & eDate & "') " & _
            "INSERT INTO Calendar Year_Id, Prd_Id, Week_Id, sDate, eDate, YrPrd, YrWks, PrdWk " & _
            "SELECT " & year & ",0,0,'" & sDate & "','" & eDate & "'," & yrprd & "," & yrprd & ",0"
        cmd = New SqlCommand(sql, con)
        cmd.ExecuteNonQuery()
        dt = New DataTable
        Dim row As DataRow
        Dim hasData As Boolean = False
        dt.Columns.Add("Prd_Id")
        dt.Columns.Add("sDate")
        dt.Columns.Add("eDate")
        con.Open()
        sql = "SELECT Prd_Id, sDate,eDate FROM Calendar WHERE Year_Id = " & year & " AND Week_Id = 0"
        cmd = New SqlCommand(sql, con)
        rdr = cmd.ExecuteReader
        While rdr.Read
            If Not IsDBNull(rdr(0)) Then
                hasData = True
                row = dt.NewRow
                row(0) = rdr("Prd_Id")
                row(1) = rdr("sDate")
                row(2) = rdr("eDate")
            End If
        End While
        con.Close()
        If Not hasData Then
            row = dt.NewRow
            row(0) = 1
            row(1) = sDate
        End If
        '   dgv.DataSource = dt.DefaultView
    End Sub
End Class